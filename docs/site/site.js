const toggle = document.querySelector('.menu');
toggle?.addEventListener('click', () => {
  const open = document.querySelector('.sidebar').classList.toggle('open');
  toggle.setAttribute('aria-expanded', String(open));
});
const search = document.querySelector('#search'),
  results = document.querySelector('.search-results');
let index;
search?.addEventListener('input', async () => {
  const query = search.value.trim().toLowerCase();
  results.replaceChildren();
  if (query.length < 2) return;
  try {
    index ??= await (await fetch('search.json')).json();
    if (query !== search.value.trim().toLowerCase()) return;
    for (const page of index
      .filter((page) => (page.title + ' ' + page.text).toLowerCase().includes(query))
      .slice(0, 8)) {
      const a = document.createElement('a');
      a.href = page.file;
      a.textContent = page.title;
      results.append(a);
    }
    if (!results.childElementCount) results.textContent = 'No matching pages.';
  } catch {
    results.textContent = 'Search unavailable. Use the navigation links.';
  }
});
