const toggle = document.querySelector('.menu');
const sidebar = document.querySelector('.sidebar');
function closeMenu() {
  sidebar.classList.remove('open');
  toggle.setAttribute('aria-expanded', 'false');
}
toggle.addEventListener('click', () => {
  const open = sidebar.classList.toggle('open');
  toggle.setAttribute('aria-expanded', String(open));
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && sidebar.classList.contains('open')) {
    closeMenu();
    toggle.focus();
  }
});
const themeButton = document.querySelector('.theme-toggle');
function labelTheme() {
  themeButton.setAttribute(
    'aria-label',
    document.documentElement.classList.contains('dark') ? 'Use light theme' : 'Use dark theme',
  );
}
labelTheme();
themeButton.addEventListener('click', () => {
  const root = document.documentElement;
  const next = root.classList.contains('dark') ? 'light' : 'dark';
  root.classList.toggle('dark', next === 'dark');
  try {
    localStorage.setItem('imprint-docs-theme', next);
  } catch {
    /* Storage is optional. */
  }
  labelTheme();
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
