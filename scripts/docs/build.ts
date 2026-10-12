import { readFile, writeFile, mkdir, copyFile, cp, rm } from 'node:fs/promises';
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
  const navigation = pages
    .map(
      ([s, t], n) =>
        `<a href="${file(s)}"${s === slug ? ' aria-current="page"' : ''}><span class="nav-index" aria-hidden="true">${String(n + 1).padStart(2, '0')}</span><span>${escape(t)}</span></a>`,
    )
    .join('');
  const moon =
    '<svg class="moon" viewBox="0 0 24 24" aria-hidden="true"><path d="M20.5 14A8.5 8.5 0 0 1 10 3.5 8.5 8.5 0 1 0 20.5 14Z"/></svg>';
  const sun =
    '<svg class="sun" viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2m0 16v2M2 12h2m16 0h2M5 5l1.5 1.5m11 11L19 19M5 19l1.5-1.5m11-11L19 5"/></svg>';
  const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <meta name="description" content="Imprint Doc Gen: ${escape(title)}. Self-hosted document generation with TypeScript and Node.js.">
  <meta name="color-scheme" content="light dark">
  <title>${escape(title)} · Imprint Doc Gen</title>
  <script src="theme-init.js"></script>
  <link rel="preload" href="fonts/inter-latin-wght-normal.woff2" as="font" type="font/woff2" crossorigin>
  <link rel="preload" href="fonts/fraunces-latin-wght-normal.woff2" as="font" type="font/woff2" crossorigin>
  <link rel="stylesheet" href="fonts.css">
  <link rel="stylesheet" href="style.css">
  <script src="site.js" defer></script>
</head>
<body>
  <a class="skip" href="#main">Skip to content</a>
  <header class="masthead">
    <div class="identity"><a class="brand" href="index.html">imprint<span aria-hidden="true">.</span></a><span class="brand-note">Document generation / Docs</span></div>
    <div class="header-actions">
      <a class="source-link" href="https://github.com/KenPrz/imprint-doc-gen">Source repository <span aria-hidden="true">↗</span></a>
      <button class="theme-toggle" type="button" aria-label="Use dark theme">${moon}${sun}</button>
      <button class="menu" type="button" aria-expanded="false" aria-controls="navigation">Menu</button>
    </div>
  </header>
  <aside class="sidebar">
    <label for="search" class="search-label">Search documentation</label>
    <input id="search" type="search" placeholder="Fields, images, deployment…" autocomplete="off" aria-controls="search-results">
    <div id="search-results" class="search-results" aria-live="polite"></div>
    <nav id="navigation" aria-label="Documentation"><p class="nav-label">Documentation</p>${navigation}</nav>
    <footer>Preview release · Node.js 24<br><a href="https://github.com/KenPrz/imprint-doc-gen">Source repository ↗</a></footer>
  </aside>
  <div class="layout">
    <main id="main">${body}<footer class="bottom">Imprint Doc Gen · Independently authored templates and examples. See format support for current limits.</footer></main>
    <aside class="toc" aria-label="On this page"><strong>On this page</strong>${headings.map(([, id, text]) => `<a href="#${id}">${text}</a>`).join('')}</aside>
  </div>
</body>
</html>`;
  for (const [, href] of body.matchAll(/href="([^"]+)"/g)) {
    if (/^https?:/.test(href!)) continue;
    const [path, fragment] = href!.split('#');
    const linked = bodies.get(path || file(slug));
    if (!linked || (fragment && !linked.includes(`id="${fragment}"`)))
      throw new Error(`Broken documentation link: ${slug} -> ${href}`);
  }
  await writeFile(join(target, file(slug)), html);
}
for (const asset of ['style.css', 'site.js', 'theme-init.js', 'fonts.css'])
  await copyFile(join(source, asset), join(target, asset));
await cp(join(source, 'fonts'), join(target, 'fonts'), { recursive: true });
await writeFile(join(target, 'search.json'), JSON.stringify(index));
await writeFile(join(target, '.nojekyll'), '');
await copyFile(join(target, 'index.html'), join(target, '404.html'));
console.log(`Built ${pages.length} documentation pages; internal links validated`);
