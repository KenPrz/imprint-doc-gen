import { readFile, writeFile, mkdir, copyFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
const pages = [
  ['start', 'Getting started'],
  ['syntax', 'Template syntax'],
  ['images', 'Images & barcodes'],
  ['formats', 'Format support'],
  ['api', 'API & SDK'],
  ['pdf', 'PDF operations'],
  ['deployment', 'Deployment'],
  ['operations', 'Operations'],
  ['roadmap', 'Roadmap & validation'],
] as const;
const source = 'docs/site',
  target = 'artifacts/docs';
await rm(target, { recursive: true, force: true });
await mkdir(target, { recursive: true });
const escape = (s: string) => s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('"', '&quot;');
const file = (slug: string) => (slug === 'start' ? 'index.html' : `${slug}.html`);
const index: { title: string; file: string; text: string }[] = [];
const bodies = new Map<string, string>();
for (const [slug, title] of pages) {
  const body = await readFile(join(source, 'pages', `${slug}.html`), 'utf8');
  bodies.set(file(slug), body);
  index.push({ title, file: file(slug), text: body.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ') });
}
for (const [slug, title] of pages) {
  const body = bodies.get(file(slug))!,
    headings = [...body.matchAll(/<h2 id="([^"]+)">([^<]+)<\/h2>/g)];
  const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="description" content="Imprint Doc Gen: ${escape(title)}. Self-hosted document generation with TypeScript and Node.js."><meta name="color-scheme" content="light"><title>${escape(title)} · Imprint Doc Gen</title><link rel="stylesheet" href="style.css"><script src="site.js" defer></script></head><body><a class="skip" href="#main">Skip to content</a><aside class="sidebar"><a class="brand" href="index.html">imprint<span aria-hidden="true">.</span><small>DOCUMENT GENERATION</small></a><button class="menu" aria-expanded="false" aria-controls="navigation">Menu</button><label for="search" class="search-label">Search documentation</label><input id="search" type="search" placeholder="Fields, images, deployment…" autocomplete="off"><div class="search-results" aria-live="polite"></div><nav id="navigation" aria-label="Documentation">${pages.map(([s, t]) => `<a href="${file(s)}"${s === slug ? ' aria-current="page"' : ''}>${escape(t)}</a>`).join('')}</nav><footer>Preview release · Node.js 24<br><a href="https://github.com/KenPrz/imprint-doc-gen">Source repository ↗</a></footer></aside><div class="layout"><main id="main">${body}<footer class="bottom">Imprint Doc Gen · Independently authored templates and examples. See format support for current limits.</footer></main><aside class="toc" aria-label="On this page"><strong>On this page</strong>${headings.map(([, id, text]) => `<a href="#${id}">${text}</a>`).join('')}</aside></div></body></html>`;
  for (const [, href] of body.matchAll(/href="([^"]+)"/g)) {
    if (/^https?:/.test(href!)) continue;
    const [path, fragment] = href!.split('#');
    const linked = bodies.get(path || file(slug));
    if (!linked || (fragment && !linked.includes(`id="${fragment}"`)))
      throw new Error(`Broken documentation link: ${slug} -> ${href}`);
  }
  await writeFile(join(target, file(slug)), html);
}
for (const asset of ['style.css', 'site.js']) await copyFile(join(source, asset), join(target, asset));
await writeFile(join(target, 'search.json'), JSON.stringify(index));
await writeFile(join(target, '.nojekyll'), '');
await copyFile(join(target, 'index.html'), join(target, '404.html'));
console.log(`Built ${pages.length} documentation pages; internal links validated`);
