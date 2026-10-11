import { createApp } from './app.ts';
const app = createApp({ apiKey: process.env.IMPRINT_API_KEY ?? '', converterUrl: process.env.GOTENBERG_URL, logger: true });
await app.listen({ host: process.env.HOST ?? '127.0.0.1', port: Number(process.env.PORT ?? 3000) });
for (const signal of ['SIGTERM', 'SIGINT']) process.once(signal, () => { void app.close().then(() => process.exit(0)); });
