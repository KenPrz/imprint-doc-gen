# Bundled fonts implementation plan

**Goal:** Ship the redistributable Google Fonts collection in both PDF-capable images, support operator-supplied fonts, and report actual coverage of Carbone's supplied font list.

**Architecture:** A shared Python build utility installs a checksum-verified Google Fonts snapshot with its licenses. Both converter Dockerfiles use it in a disposable build stage. The final images include fonts, provenance, and a fontconfig inventory compared against a pinned Carbone reference list; no runtime downloads are needed.

**Tech stack:** Docker multi-stage builds, Python standard library, fontconfig, LibreOffice/Gotenberg, Node test tooling.

**Spec:** User request in this session to bundle fonts from https://carbone.io/img/doc/carbone-cloud-font-list.txt in the project's images. Commercial fonts without redistribution rights are supplied separately; a public font-name list does not provide font files or licenses.

## Global constraints

- Preserve the existing pinned Gotenberg base, pdfcpu patch, non-root execution, and read-only runtime support.
- No Carbone runtime dependency or copying of its commercial font binaries.
- Do not commit customer documents, font binaries, or independently licensed local font files.
- Pin the upstream archive to a commit and SHA-256; retain every installed family's license.
- Do not count fontconfig fallback matches as installed fonts or claim full Carbone parity.

## Review focus

- A missing requested face must appear as missing, even if fontconfig substitutes another family.
- A corrupt download or archive traversal must fail before installing fonts.
- Variable fonts and family/style aliases must appear accurately in the inventory; presence is not a universal rendering guarantee.
- Local font files must be copied into both final images and detected after cache regeneration.
- Conversion must work offline as the normal non-root user with a read-only container filesystem.

### Task 1: Install and audit a pinned font collection

**Files:** `docker/fonts/manage.py`, `docker/fonts/sources.lock.json`, `docker/fonts/carbone-cloud-font-list.txt`, `docker/fonts/custom/README.md`, `docker/fonts/custom/.gitignore`, `tests/fonts_test.py`.

**Interfaces:** `manage.py install --archive PATH --destination DIR`; archive optional (download pinned source when omitted). `manage.py inventory --output DIR --require FAMILY ...` writes installed-fonts.json and carbone-coverage.json using fc-list, never fallback-only fc-match. `--strict` fails on missing reference entries.

- [x] Test checksum rejection, safe archive selection, license retention, missing-font detection, normal/bold/italic and variable-instance aliases, and strict coverage failure.
- [x] Implement streamed archive download and extraction of only the ofl/apache/ufl font-family files; keep license and metadata files alongside fonts.
- [x] Include source metadata, installed file counts, explicit present/missing face lists, and supplied-font instructions.
- [x] Run `python3 -m unittest discover -s tests -p fonts_test.py` (7 tests passed).

### Task 2: Bundle fonts in both images and validate PDF conversion

**Files:** `Dockerfile.gotenberg`, `Dockerfile.standalone`, `src/standalone.ts`, `scripts/font-smoke.py`, `.github/workflows/corpus.yml`.

**Interfaces:** Both images contain `/usr/local/share/fonts/imprint/`, `/usr/local/share/imprint/fonts/`, and `/usr/local/lib/imprint-fonts/manage.py`. `REQUIRED_FONTS` build argument optionally enforces extra family requirements. Smoke script sends a synthetic DOCX containing regular/bold/italic text in representative newly installed fonts to the real converter and checks embedded PDF font names.

- [x] Run the smoke check against the old image and observe missing requested fonts (all four families substituted with Noto Serif).
- [x] Add identical font stages, custom font copies, cache creation, and inventory generation to both builds.
- [x] Build both images; verify exact installed fonts and successful offline conversion under read-only runtime constraints.
- [x] Add catalog unit checks and converter smoke tests to the existing CI jobs; allow sufficient time for the large font snapshot.
- [x] Preload LibreOffice before readiness and extend standalone startup polling. Cold font initialization exceeded the 5-second request budget; with preloading, both images pass immediately after readiness under that same deadline.

### Task 3: Document the shipped coverage and extension path

**Files:** `README.md`, `docs/site/pages/deployment.html`, `docs/standalone-guide.txt`, `docs/font-validation.json`.

- [x] Explain source pinning, expected image size increase, catalog paths, exact coverage limits, custom font build inputs, and rebuild/recreate requirements.
- [x] Record measured image/catalog coverage and PDF validation results without customer data in `docs/font-validation.json`.
- [x] Run documentation build, relevant tests, Docker checks, and inspect final diff. Do not publish or replace the running application.

## Outcome

Both images contain 4,232 font files including base-image fonts; 4,771 of the
6,361 reference face names match. The other 1,590 are reported explicitly.
Regular/bold/italic PDF checks, existing API smoke checks, and timeout recovery
passed. Source archive integrity and negative required-font/parity checks passed.
Individual operator-supplied font files still require their own visual validation.

Implementation proceeds in this session under the user's request to add image support. No separate execution-method approval is needed.
