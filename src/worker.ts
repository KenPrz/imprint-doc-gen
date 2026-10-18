import { workerMetricsServer } from './worker-metrics.ts';
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
const metrics = process.env.WORKER_METRICS_PORT ? workerMetricsServer(queue.renderer,()=>store.queued(),process.env.WORKER_METRICS_TOKEN ?? '') : undefined;
if (metrics) {
  const port = Number(process.env.WORKER_METRICS_PORT);
  if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error('Invalid WORKER_METRICS_PORT');
  await new Promise<void>((resolve,reject)=>{metrics.once('error',reject);metrics.listen(port,process.env.WORKER_METRICS_HOST ?? '0.0.0.0',resolve);});
}
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
      if (metrics) await new Promise<void>((resolve,reject)=>metrics.close(error=>error?reject(error):resolve()));
      await store.close();
      clearInterval(keepAlive);
    })().catch(() => (process.exitCode = 1));
  });
console.log('Imprint worker ready');
