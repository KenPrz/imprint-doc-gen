import { createHash } from 'node:crypto';

export interface DocumentComponent {
  id: string;
  version: string;
  source: string;
}
export interface ComponentExpansion {
  source: string;
  lock: { id: string; version: string; sha256: string }[];
}
const identity = /^[A-Za-z0-9][A-Za-z0-9_.-]{0,127}$/;
const marker = /\{\{>\s*([A-Za-z0-9][A-Za-z0-9_.-]{0,127})@([A-Za-z0-9][A-Za-z0-9_.-]{0,127})\s*\}\}/g;
function references(source: string) {
  if (typeof source !== 'string' || source.length > 1_000_000)
    throw new Error('Component source exceeds budget or is not text');
  const refs = [...source.matchAll(marker)].map((m) => ({
    key: `${m[1]}@${m[2]}`,
    start: m.index,
    end: m.index + m[0].length,
  }));
  if (refs.length > 10000 || source.replace(marker, '').includes('{{>'))
    throw new Error('Invalid component reference or reference budget exceeded');
  return refs;
}
/** Immutable exact-version plain-text components. Office XML fragments must not be passed here. */
export class ComponentRegistry {
  private readonly entries = new Map<string, DocumentComponent>();
  constructor(components: readonly DocumentComponent[]) {
    if (!Array.isArray(components) || components.length > 10000)
      throw new Error('Component count exceeds budget');
    let total = 0;
    for (const component of components) {
      if (
        !component ||
        typeof component.id !== 'string' ||
        typeof component.version !== 'string' ||
        !identity.test(component.id) ||
        !identity.test(component.version)
      )
        throw new Error('Invalid component identity');
      const key = `${component.id}@${component.version}`;
      if (this.entries.has(key)) throw new Error(`Duplicate component: ${key}`);
      references(component.source);
      total += component.source.length;
      if (total > 10_000_000) throw new Error('Registry source budget exceeded');
      this.entries.set(key, { ...component });
    }
    // Validate every entry, including currently unreachable dependencies.
    const visited = new Set<string>();
    const visit = (key: string, path: Set<string>) => {
      if (path.has(key)) throw new Error(`Component cycle: ${[...path, key].join(' -> ')}`);
      if (path.size >= 64) throw new Error('Component depth exceeds 64');
      const item = this.entries.get(key);
      if (!item) throw new Error(`Missing component: ${key}`);
      if (visited.has(key)) return;
      for (const ref of references(item.source)) visit(ref.key, new Set([...path, key]));
      visited.add(key);
    };
    for (const key of this.entries.keys()) visit(key, new Set());
  }
  expand(source: string): ComponentExpansion {
    const used = new Set<string>();
    let operations = 0;
    const expand = (text: string, depth: number): string => {
      if (depth > 64) throw new Error('Component depth exceeds 64');
      let output = '',
        cursor = 0;
      for (const ref of references(text)) {
        if (++operations > 10000) throw new Error('Component expansion count exceeds 10000');
        const item = this.entries.get(ref.key);
        if (!item) throw new Error(`Missing component: ${ref.key}`);
        used.add(ref.key);
        output += text.slice(cursor, ref.start) + expand(item.source, depth + 1);
        if (output.length > 1_000_000) throw new Error('Expanded source exceeds budget');
        cursor = ref.end;
      }
      output += text.slice(cursor);
      if (output.length > 1_000_000) throw new Error('Expanded source exceeds budget');
      return output;
    };
    const expanded = expand(source, 0);
    return {
      source: expanded,
      lock: [...used].sort().map((key) => {
        const item = this.entries.get(key)!;
        return {
          id: item.id,
          version: item.version,
          sha256: createHash('sha256').update(item.source).digest('hex'),
        };
      }),
    };
  }
  /** Exact-version transitive impact; callers decide whether to adopt a new version. */
  impact(id: string, version: string, templates: Readonly<Record<string, string>>) {
    const key = `${id}@${version}`;
    if (!this.entries.has(key)) throw new Error(`Missing component: ${key}`);
    if (Object.keys(templates).length > 10000) throw new Error('Template count exceeds budget');
    return Object.keys(templates)
      .sort()
      .filter((name) => this.expand(templates[name]!).lock.some((c) => c.id === id && c.version === version));
  }
}
