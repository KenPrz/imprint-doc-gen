import type { StandaloneStore } from './store.ts';
type Methods =
  | 'authenticate'
  | 'listTenants'
  | 'createTenant'
  | 'rotateKey'
  | 'template'
  | 'listTemplates'
  | 'versions'
  | 'upload'
  | 'publish'
  | 'deleteTemplate'
  | 'enqueue'
  | 'job'
  | 'result'
  | 'deleteJob'
  | 'claim'
  | 'succeed'
  | 'fail'
  | 'cleanup'
  | 'close';
export type ServiceStore = {
  [K in Methods]: (
    ...args: Parameters<StandaloneStore[K]>
  ) => ReturnType<StandaloneStore[K]> | Promise<ReturnType<StandaloneStore[K]>>;
} & { kind: 'sqlite' | 'postgres'; ready(): void | Promise<void>; queued(): number | Promise<number> };
