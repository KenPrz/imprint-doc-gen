// Apply the saved theme before paint. Light is the portfolio's default.
document.documentElement.classList.add('js');
try {
  const saved = localStorage.getItem('imprint-docs-theme');
  const dark = saved === 'dark' || (saved === 'system' && matchMedia('(prefers-color-scheme: dark)').matches);
  document.documentElement.classList.toggle('dark', dark);
} catch {
  /* Documentation stays usable when browser storage is unavailable. */
}
