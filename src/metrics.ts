// Fixed label sets keep telemetry bounded and exclude tenant IDs, paths and data.
const buckets = [0.01, 0.05, 0.1, 0.25, 0.5, 1, 2.5, 5, 10, 30, 120];
export class Metrics {
  private requests = new Map<string, number>();
  private renders = { succeeded: 0, failed: 0 };
  private durations = buckets.map(() => 0);
  private sum = 0;
  private bytes = 0;
  request(status: number) {
    const label = `${Math.floor(status / 100)}xx`;
    this.requests.set(label, (this.requests.get(label) ?? 0) + 1);
  }
  render(ok: boolean, seconds: number, bytes = 0) {
    this.renders[ok ? 'succeeded' : 'failed']++;
    this.sum += seconds;
    this.bytes += bytes;
    buckets.forEach((limit, n) => {
      if (seconds <= limit) this.durations[n]!++;
    });
  }
  text(active: number, capacity: number, queued = 0): string {
    const count = this.renders.succeeded + this.renders.failed;
    return [
      '# HELP imprint_http_requests_total Completed HTTP responses by status class.',
      '# TYPE imprint_http_requests_total counter',
      ...[...this.requests].map(([status, n]) => `imprint_http_requests_total{status="${status}"} ${n}`),
      '# TYPE imprint_renders_total counter',
      ...Object.entries(this.renders).map(([result, n]) => `imprint_renders_total{result="${result}"} ${n}`),
      '# TYPE imprint_render_duration_seconds histogram',
      ...buckets.map(
        (limit, n) => `imprint_render_duration_seconds_bucket{le="${limit}"} ${this.durations[n]}`,
      ),
      `imprint_render_duration_seconds_bucket{le="+Inf"} ${count}`,
      `imprint_render_duration_seconds_sum ${this.sum}`,
      `imprint_render_duration_seconds_count ${count}`,
      '# TYPE imprint_output_bytes_total counter',
      `imprint_output_bytes_total ${this.bytes}`,
      '# TYPE imprint_active_renders gauge',
      `imprint_active_renders ${active}`,
      '# TYPE imprint_render_capacity gauge',
      `imprint_render_capacity ${capacity}`,
      '# TYPE imprint_queued_jobs gauge',
      `imprint_queued_jobs ${queued}`,
      '# TYPE imprint_process_resident_bytes gauge',
      `imprint_process_resident_bytes ${process.memoryUsage().rss}`,
      '# TYPE imprint_process_uptime_seconds gauge',
      `imprint_process_uptime_seconds ${process.uptime()}`,
      '',
    ].join('\n');
  }
}
