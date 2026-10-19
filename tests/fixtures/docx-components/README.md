# Independent DOCX component fixtures

`host.docx` and `clause.docx` were authored by the committed `generate.py` using python-docx 1.2.0, not the assembly implementation. They contain synthetic text and a generated solid-color PNG; no third-party document content. The normal python-docx defaults, themes, style catalog, numbering definitions, and DrawingML are retained. Regeneration requires `python-docx==1.2.0`.

The host has two whole-paragraph `{{>clause@1}}` markers. The clause exercises a heading, direct bold/color formatting, two numbered paragraphs, a styled grid table, and an embedded PNG. Input component SHA-256 pins are calculated from the fixture bytes by the tests.

Run `node --test tests/docx-components.test.ts` for structural, reproducibility, and rejection tests. Run `node tests/fixtures/docx-components/verify-office.ts /absolute/output/directory` with LibreOffice and Poppler installed for independent conversion assertions and a page PNG for visual review. Snap LibreOffice requires the output directory to be under the user's home rather than `/tmp`.

The independent converter exposed a list-continuation defect despite distinct style IDs: imported style display names also had to be unique. The acceptance script requires numbering 1/2 in both inserted clauses, plus both tables and both images. It does not establish Microsoft Word desktop acceptance or arbitrary DOCX compatibility.
