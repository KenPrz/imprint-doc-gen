#!/usr/bin/env node
import { readFile, writeFile } from 'node:fs/promises';
import { resolve, dirname, extname } from 'node:path';
import { compile } from './index.ts';
import { RenderError } from './errors.ts';

const args = process.argv.slice(2);
if (args.includes('--help') || args.length < 3) {
  console.log('Usage: imprint-doc-gen template.<format> data.json output.<format> [attachments.json]\nAttachment manifest: {"logo":"./logo.png"}; paths are relative to the manifest.');
  process.exitCode = args.includes('--help') ? 0 : 2;
} else {
  try {
    if (args.length > 4) throw new RenderError('ARGUMENTS', 'Too many arguments');
    const [template, data, output, manifest] = args;
    if (resolve(template!) === resolve(output!)) throw new RenderError('ARGUMENTS', 'Output must not overwrite the input template');
    const attachments: Record<string, Uint8Array> = Object.create(null);
    if (manifest) {
      const entries = JSON.parse(await readFile(manifest, 'utf8')) as Record<string, string>;
      for (const [key, path] of Object.entries(entries)) attachments[key] = await readFile(resolve(dirname(manifest), path));
    }
    const compiled = await compile(await readFile(template!));
    if (extname(output!).toLowerCase() !== `.${compiled.format}`) throw new RenderError('ARGUMENTS', `Output extension must be .${compiled.format}`);
    const result = await compiled.render(JSON.parse(await readFile(data!, 'utf8')), { attachments });
    await writeFile(output!, result);
    console.log(`Rendered ${output} (${result.length} bytes)`);
  } catch (error) {
    console.error(JSON.stringify({ error: error instanceof RenderError ? error.code : 'RENDER_FAILED', message: (error as Error).message, part: error instanceof RenderError ? error.part : undefined }));
    process.exitCode = 1;
  }
}
