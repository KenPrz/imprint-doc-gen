import { readPackage } from '../../src/package.ts';
import { snapshot } from '../corpus/inspect.ts';

// Content-only oracle: Carbone may omit explicit cell coordinates, emit numeric
// data as text and leave formula caches stale. Formula/type correctness is tested
// separately against fixed corpus goldens and independent LibreOffice calculation.
export async function xlsxContent(bytes: Uint8Array): Promise<string[]> {
  const content = snapshot(await readPackage(bytes), 'xlsx')['xl/worksheets/sheet1.xml'] as { cells: { value: string; formula: string | null }[] };
  return content.cells.filter(cell => !cell.formula && cell.value !== '').map(cell => cell.value);
}
