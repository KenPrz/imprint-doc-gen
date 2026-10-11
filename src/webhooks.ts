import { createHmac, randomUUID } from 'node:crypto';
import type { PostgresStore } from './postgres-store.ts';
export interface WebhookConfig {
  url: string;
  secret: string;
  allowHttp?: boolean;
}
export class WebhookDispatcher {
  private timer?: NodeJS.Timeout;
  private pending?: Promise<void>;
  private stopping = false;
  readonly store: PostgresStore;
  readonly config: WebhookConfig;
  constructor(store: PostgresStore, config: WebhookConfig) {
    const url = new URL(config.url);
    if (
      url.username ||
      url.password ||
      url.hash ||
      !['https:', ...(config.allowHttp ? ['http:'] : [])].includes(url.protocol) ||
      config.secret.length < 32
    )
      throw new Error('Webhook requires HTTPS and a secret of at least 32 characters');
    this.store = store;
    this.config = config;
  }
  start() {
    this.timer = setInterval(() => {
      if (!this.pending && !this.stopping)
        this.pending = this.tick()
          .catch(() => {})
          .finally(() => {
            this.pending = undefined;
          });
    }, 500);
    this.timer.unref();
  }
  async tick() {
    const event = await this.store.transaction(async (c) => {
      const now = this.store.now();
      const row = (
        await c.query(
          'SELECT * FROM webhooks WHERE delivered_at IS NULL AND attempts<8 AND available_at<=$1 AND (lease_until IS NULL OR lease_until<=$1) ORDER BY created_at,id FOR UPDATE SKIP LOCKED LIMIT 1',
          [now],
        )
      ).rows[0];
      if (!row) return;
      const lease = randomUUID();
      await c.query('UPDATE webhooks SET lease=$1,lease_until=$2,attempts=attempts+1 WHERE id=$3', [
        lease,
        now + 30000,
        row.id,
      ]);
      return { ...row, lease };
    });
    if (!event) return;
    const timestamp = String(Math.floor(Date.now() / 1000)),
      signature = createHmac('sha256', this.config.secret)
        .update(`${timestamp}.${event.payload}`)
        .digest('hex');
    let success = false;
    try {
      const response = await fetch(this.config.url, {
        method: 'POST',
        redirect: 'error',
        signal: AbortSignal.timeout(5000),
        headers: {
          'content-type': 'application/json',
          'x-imprint-event-id': event.id,
          'x-imprint-timestamp': timestamp,
          'x-imprint-signature': `sha256=${signature}`,
        },
        body: event.payload,
      });
      success = response.ok;
      await response.body?.cancel();
    } catch {
      /* Store only a stable error code, never response bodies or endpoint credentials. */
    }
    await this.store.pool.query(
      'UPDATE webhooks SET delivered_at=$1,last_error=$2,available_at=$3,lease=NULL,lease_until=NULL WHERE id=$4 AND lease=$5',
      [
        success ? this.store.now() : null,
        success ? null : 'DELIVERY_FAILED',
        this.store.now() + Math.min(3600000, 1000 * 2 ** event.attempts),
        event.id,
        event.lease,
      ],
    );
  }
  async stop() {
    this.stopping = true;
    clearInterval(this.timer);
    await this.pending;
  }
}
