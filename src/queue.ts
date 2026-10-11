import { RenderError } from './errors.ts';
import { ServiceError } from './service-error.ts';
import { RenderService } from './render-service.ts';
import type { ServiceStore } from './store-interface.ts';

export class DurableQueue {
  private timer?: NodeJS.Timeout;
  private running = new Set<Promise<void>>();
  private stopping = false;
  private ticking = false;
  private nextCleanup = 0;
  readonly store: ServiceStore;
  readonly renderer: RenderService;
  readonly onError: (error: unknown) => void;
  constructor(store: ServiceStore, renderer: RenderService, onError: (error: unknown) => void) {
    this.store = store; this.renderer = renderer; this.onError = onError;
  }
  start() {
    this.timer = setInterval(() => { void this.tick(); }, 250);
    this.timer.unref();
    void this.tick();
  }
  async tick() {
    if (this.stopping || this.ticking) return;
    this.ticking = true;
    try {
      if (Date.now() >= this.nextCleanup) { await this.store.cleanup(); this.nextCleanup = Date.now() + 60_000; }
      while (!this.stopping && this.renderer.available) {
        // Reserve shared admission while an asynchronous store claims a job.
        this.renderer.active++;
        let job;
        try { job = await this.store.claim(); } finally { this.renderer.active--; }
        if (!job) break;
        const pending = (async () => {
          try {
            const result = await this.renderer.render(job.template, JSON.parse(job.payload));
            await this.store.succeed(job, result);
          } catch (error) {
            const code = error instanceof RenderError || error instanceof ServiceError ? error.code : 'RENDER_FAILED';
            await this.store.fail(job, code, ['CONVERSION_FAILED', 'RENDER_FAILED', 'OBJECT_UNAVAILABLE'].includes(code));
          }
        })().catch(this.onError).finally(() => this.running.delete(pending));
        this.running.add(pending);
      }
    } catch (error) { this.onError(error); } finally { this.ticking = false; }
  }
  async stop() {
    this.stopping = true;
    clearInterval(this.timer);
    // A claim already in flight must finish before closing its store.
    while(this.ticking) await new Promise(resolve=>setTimeout(resolve,10));
    await Promise.all(this.running);
  }
}
