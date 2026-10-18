#!/usr/bin/env node
import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { executeAssurance } from '../../src/assurance-task.ts';
const [input, output, ...rest] = process.argv.slice(2);
if (!input || !output || rest.length) {
  console.error(
    'Usage: npm run assurance -- request.json evidence.json (same payload as POST /v1/assurance)',
  );
  process.exitCode = input === '--help' ? 0 : 2;
} else {
  try {
    if (resolve(input) === resolve(output)) throw new Error('Evidence must not overwrite the request');
    const bytes = await readFile(input);
    if (bytes.length > 12 * 1024 * 1024) throw new Error('Request exceeds 12 MiB');
    const result = await executeAssurance(JSON.parse(bytes.toString('utf8')));
    await writeFile(output, JSON.stringify(result, null, 2) + '\n', { flag: 'wx', mode: 0o600 });
    console.log(`Wrote ${output}`);
  } catch (error) {
    console.error(error instanceof Error ? error.message : 'Assurance operation failed');
    process.exitCode = 1;
  }
}
