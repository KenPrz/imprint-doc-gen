import { RenderError } from './errors.ts';
import { forbidden } from './scope.ts';

export function validateData(data: unknown): asserts data is Record<string, unknown> {
  if (!data || typeof data !== 'object' || Array.isArray(data)) throw new RenderError('INVALID_DATA', 'Data must be a JSON object');
  let count = 0; const visiting = new WeakSet<object>();
  const visit = (value: unknown, depth: number) => {
    if (++count > 100_000 || depth > 64) throw new RenderError('LIMIT_EXCEEDED', 'Data exceeds its complexity budget');
    if (value === null || typeof value === 'string' || typeof value === 'boolean' || (typeof value === 'number' && Number.isFinite(value))) return;
    if (typeof value !== 'object' || visiting.has(value)) throw new RenderError('INVALID_DATA', 'Data must contain only acyclic JSON values');
    if (!Array.isArray(value) && ![Object.prototype, null].includes(Object.getPrototypeOf(value))) throw new RenderError('INVALID_DATA', 'Data must contain plain objects');
    visiting.add(value);
    for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) {
      if (Array.isArray(value) && key === 'length') continue;
      if (forbidden.has(key) || !Object.hasOwn(descriptor, 'value')) throw new RenderError('INVALID_DATA', 'Accessors and prototype keys are not supported');
      visit(descriptor.value, depth + 1);
    }
    visiting.delete(value);
  };
  visit(data, 0);
}
