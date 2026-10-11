import { readFile } from 'node:fs/promises';
import { inspectTemplate } from '../src/inspect.ts';
const path = process.argv[2];
if (!path) throw new Error('Usage: node scripts/inspect.ts <template>');
const report = await inspectTemplate(await readFile(path));
console.log(JSON.stringify(report, null, 2));
if (!report.valid) process.exitCode = 1;
