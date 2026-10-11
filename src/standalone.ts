import { spawn, type ChildProcess } from 'node:child_process';
import { setTimeout as delay } from 'node:timers/promises';
import { fileURLToPath } from 'node:url';

// tini is PID 1. This supervisor forwards shutdown in dependency order, drains jobs,
// and makes either child exiting fail the container so Docker can restart it.
if ((process.env.IMPRINT_API_KEY ?? '').length < 24) throw new Error('IMPRINT_API_KEY must contain at least 24 characters');
const converterUrl = 'http://127.0.0.1:3001';
let stopping = false;
let api: ChildProcess | undefined;
const children = new Set<ChildProcess>();
function launch(command: string, args: string[], env = process.env) {
  const child = spawn(command, args, { stdio: 'inherit', detached: true, env });
  children.add(child);
  child.once('error', error => { console.error('Child startup failed', error.message); void shutdown(1); });
  child.once('exit', () => { children.delete(child); if (!stopping) void shutdown(1); });
  return child;
}
function signal(child: ChildProcess, value: NodeJS.Signals) {
  if (child.pid) { try { process.kill(-child.pid, value); } catch { /* Already exited. */ } }
}
async function stop(child: ChildProcess | undefined, timeout: number) {
  if (!child || !children.has(child)) return;
  await new Promise<void>(resolve => {
    const timer = setTimeout(() => { signal(child, 'SIGKILL'); resolve(); }, timeout);
    child.once('exit', () => { clearTimeout(timer); resolve(); });
    signal(child, 'SIGTERM');
  });
}
async function shutdown(code: number) {
  if (stopping) return;
  stopping = true;
  await stop(api, 100_000);
  await stop(converter, 10_000);
  for (const child of children) signal(child, 'SIGKILL');
  process.exit(code);
}
const converter = launch('gotenberg', [
  '--api-bind-ip=127.0.0.1', '--api-port=3001', '--api-body-limit=20MB',
  `--api-timeout=${process.env.GOTENBERG_API_TIMEOUT ?? '60s'}`,
  '--api-disable-download-from=true', '--webhook-disable=true', '--chromium-disable-routes=true',
  '--libreoffice-max-queue-size=8', '--libreoffice-restart-after=1', '--libreoffice-deny-list=.*',
]);
for (const sig of ['SIGTERM', 'SIGINT'] as const) process.once(sig, () => { void shutdown(0); });
let ready = false;
for (let attempt = 0; attempt < 60 && !stopping; attempt++) {
  try {
    const response = await fetch(`${converterUrl}/health`, { signal: AbortSignal.timeout(1000) });
    ready = response.ok; await response.body?.cancel();
    if (ready) break;
  } catch { /* Startup is bounded below. */ }
  await delay(500);
}
if (!ready) { console.error('Converter did not become ready'); await shutdown(1); }
if (!stopping) api = launch(process.execPath, [fileURLToPath(new URL('./server.js', import.meta.url))], {
  ...process.env, GOTENBERG_URL: converterUrl, IMPRINT_DATA_DIR: process.env.IMPRINT_DATA_DIR ?? '/data',
});
