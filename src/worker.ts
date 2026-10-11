import { configuredStore } from './config-store.ts';
import { RenderService } from './render-service.ts';
import { DurableQueue } from './queue.ts';
import { WebhookDispatcher } from './webhooks.ts';
const store = await configuredStore();
if (!store) throw new Error('Independent workers require DATABASE_URL and S3_BUCKET');
const queue = new DurableQueue(
  store,
  new RenderService(Number(process.env.IMPRINT_CONCURRENCY ?? 2), process.env.GOTENBERG_URL),
  () => console.error('Queue operation failed; retrying'),
);
const webhooks = process.env.WEBHOOK_URL
  ? new WebhookDispatcher(store, {
      url: process.env.WEBHOOK_URL,
      secret: process.env.WEBHOOK_SECRET ?? '',
      allowHttp: process.env.WEBHOOK_ALLOW_HTTP === 'true',
    })
  : undefined;
queue.start();
webhooks?.start();
// Queue timers are unref'ed for embedded use; the independent worker must remain alive.
const keepAlive = setInterval(() => {}, 60000);
let closing = false;
for (const signal of ['SIGTERM', 'SIGINT'])
  process.once(signal, () => {
    if (closing) return;
    closing = true;
    void (async () => {
      await queue.stop();
      await webhooks?.stop();
      await store.close();
      clearInterval(keepAlive);
    })().catch(() => (process.exitCode = 1));
  });
console.log('Imprint worker ready');
