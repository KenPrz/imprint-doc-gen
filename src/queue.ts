import { RenderError } from './errors.ts';
import { ServiceError } from './service-error.ts';
import { RenderService } from './render-service.ts';
import { StandaloneStore } from './store.ts';

export class DurableQueue {
  private timer?: NodeJS.Timeout;
  private running = new Set<Promise<void>>();
  private stopping = false;
  private nextCleanup = 0;
  readonly store: StandaloneStore;
  readonly renderer: RenderService;
  readonly onError: (error: unknown) => void;
  constructor(store: StandaloneStore, renderer: RenderService, onError: (error: unknown) => void) {
    this.store = store; this.renderer = renderer; this.onError = onError;
  }
  start() {
    this.timer = setInterval(() => this.tick(), 250);
    this.timer.unref();
    this.tick();
  }
  tick() {
    if (this.stopping) return;
    try {
      if (Date.now() >= this.nextCleanup) { this.store.cleanup(); this.nextCleanup = Date.now() + 60_000; }
      while (this.renderer.available) {
        const job = this.store.claim();
        if (!job) break;
        const pending = (async () => {
          try {
            const result = await this.renderer.render(job.template, JSON.parse(job.payload));
            this.store.succeed(job, result);
          } catch (error) {
            const code = error instanceof RenderError || error instanceof ServiceError ? error.code : 'RENDER_FAILED';
            this.store.fail(job, code, ['CONVERSION_FAILED', 'RENDER_FAILED'].includes(code));
          }
        })().catch(this.onError).finally(() => this.running.delete(pending));
        this.running.add(pending);
      }
    } catch (error) { this.onError(error); }
  }
  async stop() {
    this.stopping = true;
    clearInterval(this.timer);
    await Promise.all(this.running);
  }
}
