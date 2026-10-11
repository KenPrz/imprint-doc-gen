// Server-side client: keep API keys in your PMS backend, never in browser code.
export interface Job {
  id: string; templateId: string; versionId: string;
  status: 'queued' | 'running' | 'succeeded' | 'failed';
  error: string | null; format: 'docx' | 'xlsx' | 'pdf' | null; downloadUrl: string | null;
}
export class ImprintClient {
  readonly endpoint: string;
  readonly apiKey: string;
  constructor(endpoint: string, apiKey: string) { this.endpoint = endpoint; this.apiKey = apiKey; }
  private async request(path: string, method = 'GET', body?: unknown, idempotencyKey?: string, timeout = 35_000) {
    const response = await fetch(new URL(path, this.endpoint), {
      method, signal: AbortSignal.timeout(timeout),
      headers: { authorization: `Bearer ${this.apiKey}`, ...(body === undefined ? {} : { 'content-type': 'application/json' }), ...(idempotencyKey ? { 'idempotency-key': idempotencyKey } : {}) },
      body: body === undefined ? undefined : JSON.stringify(body),
    });
    if (!response.ok) throw new Error(`Imprint HTTP ${response.status}: ${(await response.text()).slice(0, 1000)}`);
    return response;
  }
  async upload(name: string, template: Uint8Array): Promise<{ templateId: string; versionId: string }> {
    return (await this.request('/v1/templates', 'POST', { name, template: Buffer.from(template).toString('base64') })).json();
  }
  async publish(templateId: string, versionId: string): Promise<void> {
    await this.request(`/v1/templates/${encodeURIComponent(templateId)}/publish`, 'POST', { versionId });
  }
  async submit(input: { templateId: string; versionId?: string; data: Record<string, unknown>; attachments?: Record<string, Uint8Array>; output?: 'docx' | 'xlsx' | 'pdf' }, idempotencyKey: string): Promise<Job> {
    const attachments = Object.fromEntries(Object.entries(input.attachments ?? {}).map(([name, bytes]) => [name, Buffer.from(bytes).toString('base64')]));
    return (await this.request('/v1/jobs', 'POST', { ...input, attachments }, idempotencyKey)).json();
  }
  async status(jobId: string): Promise<Job> { return (await this.request(`/v1/jobs/${encodeURIComponent(jobId)}`)).json(); }
  async wait(jobId: string, timeoutMs = 120_000): Promise<Job> {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      const job = await this.status(jobId);
      if (job.status === 'succeeded') return job;
      if (job.status === 'failed') throw new Error(`Document job ${job.id} failed: ${job.error}`);
      await new Promise(resolve => setTimeout(resolve, 500));
    }
    throw new Error(`Still processing job ${jobId}; save its ID and poll again later`);
  }
  async download(jobId: string): Promise<Buffer> {
    return Buffer.from(await (await this.request(`/v1/jobs/${encodeURIComponent(jobId)}/result`)).arrayBuffer());
  }
}
