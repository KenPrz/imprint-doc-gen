export function withArrayLength(value: unknown, count: number): unknown {
  if (Array.isArray(value)) return value.length ? Array.from({ length: count }, (_, index) => withArrayLength(value[index % value.length], count)) : [];
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, withArrayLength(item, count)]));
  return value;
}
