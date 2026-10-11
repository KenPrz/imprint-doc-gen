import type { Format } from './common.ts';

export const bands = {
  docx: [
    ['Text fragmentation', 'Increasing splits within placeholders; escaping, Unicode and significant whitespace'],
    ['Picture attachments', 'Increasing picture placements, distinct attachments, shared placeholder media and varied extents'],
    ['Tables', 'Increasing rows and columns with cell-local placeholders'],
    ['Inline structure', 'Hyperlinks, bookmarks and content controls surrounding split placeholders'],
    ['Header and footer stories', 'Repeated header/footer bindings and images with part-local relationships'],
    ['Repeated table rows', 'Increasing item counts, per-item images and row-local loop scope'],
    ['Nested repetition', 'Increasing group counts with nested child items and parent scope values'],
    ['Page and section structure', 'Increasing sections, orientation changes and page breaks'],
    ['Package obfuscation', 'Namespace aliases, nonsequential relationship IDs, XML comments and ZIP entry order'],
    ['Combined bounded stress', 'All preceding features plus increasing long-content and repeated-row load'],
  ],
  xlsx: [
    ['Cells and types', 'Increasing text bindings with numbers, booleans, Unicode and formula-like literal strings'],
    ['Shared and rich strings', 'Increasing placeholder fragmentation in shared rich-text strings'],
    ['Picture attachments', 'Increasing drawing anchors, distinct images and shared placeholder media'],
    ['Worksheet layout', 'Increasing styled cells and merged ranges, frozen panes and explicit dimensions'],
    ['Formula preservation', 'Increasing formulas with relative, absolute, mixed and cross-sheet references'],
    ['Repeated rows', 'Increasing item counts, typed numeric output and rebased row formulas'],
    ['Multiple worksheets', 'Increasing sheets with shared strings and nontrivial names'],
    ['Dependent ranges', 'Increasing range consumers: tables, filters, names, validation and conditional formatting'],
    ['Package obfuscation', 'Namespace aliases, nonsequential relationship IDs, XML comments and ZIP entry order'],
    ['Combined bounded stress', 'All preceding features plus increasing repeated rows and populated cells'],
  ],
} as const;
export const amount = (level: number, band: number) => Math.max(0, Math.min(10, level - (band - 1) * 10));
export function recipe(format: Format, level: number) {
  if (!Number.isInteger(level) || level < 1 || level > 100) throw new Error('Level must be 1..100');
  const tier = Math.ceil(level / 10);
  return { id: `${format}-${String(level).padStart(3, '0')}`, format, level, tier, step: (level - 1) % 10 + 1,
    title: bands[format][tier - 1]![0], focus: bands[format][tier - 1]![1],
    dimensions: Array.from({ length: 10 }, (_, i) => amount(level, i + 1)),
    progression: 'Previous bands stay at severity 10; the current band increases from 1 to 10. Difficulty is a curriculum, not a universal mathematical ordering.',
  };
}
export function inputData(level: number) {
  const rows = 2 + amount(level, 6) + 15 * amount(level, 10);
  return {
    title: 'Imprint Doc Gen — Café & Tea <2026>', customer: 'Zoë 山田 & Sons',
    spaced: '  preserve these spaces  ', literal: '=1+1', enabled: false, count: 0,
    items: Array.from({ length: rows }, (_, i) => ({ name: `Item ${i + 1} — Ω`, quantity: i + 1, price: 12.5, image: i % 2 ? 'blue' : 'red' })),
    groups: Array.from({ length: 1 + amount(level, 7) }, (_, i) => ({ name: `Group ${i + 1}`, items: [{ name: `Child ${i + 1}.1`, image: 'red' }, { name: `Child ${i + 1}.2`, image: 'blue' }] })),
  };
}
export type Data = ReturnType<typeof inputData>;
