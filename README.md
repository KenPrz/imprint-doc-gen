# imprint-doc-gen

[![Document corpus](https://github.com/KenPrz/imprint-doc-gen/actions/workflows/corpus.yml/badge.svg)](https://github.com/KenPrz/imprint-doc-gen/actions/workflows/corpus.yml)

Independent document generation in TypeScript and Node.js. Render a DOCX or XLSX template with JSON data and image attachments using the SDK, CLI or authenticated HTTP API. Convert the result to PDF through a private LibreOffice service.

The template compiler and renderer are implemented in this repository. Carbone and Docxtemplater are not runtime dependencies. This is an early service foundation, with **DOCX levels 1–100 and XLSX levels 1–50 passing**; it is not feature-complete or template-syntax compatible with either product.

## Current capabilities

- Scalar placeholders such as `{{customer.name}}`, including tags split across formatted text runs.
- Text replacement in document bodies, tables, hyperlinks, content controls, headers and footers.
- Replacement of existing pictures using alternative text such as `{{image:logo}}` and PNG/JPEG attachments.
- Independent image placements even when pictures share their original media; existing drawing dimensions are preserved.
- Repeated table rows and paragraph blocks, nested object scopes, parent lookups and per-item images.
- Reusable compiled templates, JSON-only data and explicit errors for missing or unsupported bindings.
- XLSX shared/rich strings, native numeric and boolean cells, and anchored picture replacement.
- Authenticated DOCX/XLSX/PDF API with bounded worker execution and request concurrency.

**Not implemented yet:** XLSX row loops and formula/range rewriting, DOCX multiline/tabbed replacement values, inline image insertion, SVG, strict OOXML, tenant accounts, durable jobs, template storage or billing. `null` becomes empty text in DOCX and a blank cell for whole-cell XLSX bindings; `false` and `0` remain meaningful values.

## Quick start: Docker development

Requires Docker Engine and Docker Compose v2.

```bash
git clone https://github.com/KenPrz/imprint-doc-gen.git
cd imprint-doc-gen
docker compose up --build -d
docker compose exec app node scripts/smoke.ts
```

The API listens at **http://127.0.0.1:8088**. Development uses the API key `imprint-local-development-key`, unless `IMPRINT_API_KEY` is set in your environment or `.env`.

```bash
curl http://127.0.0.1:8088/health
docker compose logs -f app
docker compose exec app npm run typecheck
docker compose exec app npm test
docker compose down
```

Compose automatically combines [`compose.yml`](compose.yml) with [`compose.override.yml`](compose.override.yml). Development runs TypeScript in Node watch mode, mounts source/tests/scripts read-only and makes the corpus writable. The converter has no published host port.

## Production configuration

```bash
cp .env.example .env
# Set IMPRINT_API_KEY in .env to a generated secret of at least 24 characters.
docker compose -f compose.yml up --build -d
```

Passing `-f compose.yml` excludes the development override. The production image contains compiled JavaScript and production dependencies, runs as a non-root user, and uses a read-only filesystem, bounded temporary storage and resource limits. Startup requires an API key. Put a TLS reverse proxy in front of the loopback-bound API when deploying.

Node and Gotenberg base images are pinned by digest. The converter image replaces bundled pdfcpu 0.15.0 with verified release 0.16.1. Linux amd64 and arm64 builds are defined; local validation covers amd64.

**Known conversion limitation:** LibreOffice still stalls on DOCX stress reference level 93. Compose now recycles LibreOffice after every conversion attempt, including timeouts, so the following document can render without a manual restart. A timed-out request returns JSON with HTTP `504`; the failing document is not retried automatically or reported as converted. Recycling trades some throughput for predictable recovery. See [recovery evidence](docs/converter-recovery.json).

## Template and image bindings

In Word, add `{{customer.name}}` as ordinary text. To replace an image, insert a placeholder picture and set its alternative-text description to `{{image:logo}}`. Pass the replacement bytes under the attachment key `logo`. This replaces the picture while preserving its placement and dimensions.

### DOCX loops

To repeat paragraph blocks, put the opening and closing markers in separate paragraphs with no other content:

```text
{{#each items}}
{{name}} — {{quantity}}
{{/each}}
```

The marker paragraphs are removed. To repeat a **single table row**, put `{{#each items}}` and `{{/each}}` inside that same row, usually in its first and last cells. Other bindings in the row use the current item. Nested blocks can appear inside repeated rows or other blocks.

```json
{"items": [{"name": "First item", "quantity": 2, "image": "red"}]}
```

A picture with alternative text `{{image:image}}` resolves the current item's `image` field to an attachment key. Inside a loop, image paths always use the current scope; use `{{image:../logo}}` to read a parent's `logo` field. Outside loops, direct attachment keys such as `{{image:logo}}` remain supported.

- `{{name}}` reads the current object; `{{../name}}` reads its parent. There is no implicit fallback to root data.
- Loops accept arrays of objects. Empty arrays remove the repeated region; missing/null arrays and non-object items produce errors.
- Text and image bindings are compiled once. Data containing `{{...}}` stays literal and is never interpreted as another template.
- Expansion is limited to 16 nested loops, 10,000 total iterations per document, 8 MiB per XML part and 32 MiB per expanded package.
- Multi-row loop regions, inline text loops, repeated bookmarks/comments/notes, tracked changes, vertical merges, bound content controls and VML/OLE content require later adapters and fail explicitly. Basic content controls and DrawingML pictures get fresh identities when repeated.

### XLSX cells and pictures

Use the same scalar tags in Excel cells and the same `{{image:logo}}` alternative text on existing pictures. The renderer follows workbook relationships across worksheets, including hidden sheets, and handles tags split across rich-text runs or stored in shared strings.

| Cell template | JSON value | Output |
| --- | --- | --- |
| `{{count}}` | `12.5` | Numeric cell |
| `{{enabled}}` | `false` | Boolean cell |
| `{{optional}}` | `null` | Blank cell, with its style retained |
| `{{label}}` | `"=1+1"` | Literal text, never a formula |
| `Total: {{count}}` | `12.5` | Text `Total: 12.5` |

Only a tag occupying the entire cell gets native typing. Other replacements remain text; rich-text replacements use the first placeholder character's formatting. XLSX strings support line breaks and tabs and are limited to 32,767 UTF-16 code units. Text is escaped for both XML and Excel's `_xHHHH_` representation.

Existing picture anchors, sizes, crop and rotation are preserved. Pictures can use one-cell, two-cell or absolute anchors. Replacing a bound picture does not change another picture sharing its original image.

Existing formulas and their cached results are preserved. Changed worksheets request automatic full recalculation when opened in Excel or converted by LibreOffice; the SDK does **not** calculate new formula results. Row loops, formula placeholders, date conversion and bound phonetic strings require later adapters. A workbook containing XLSX loop markers fails explicitly.

See the [corpus guide](docs/corpus-guide.txt) for implemented and planned semantics.

## SDK and CLI

For local development without the API container, use Node.js 24:

```bash
npm ci --ignore-scripts
npm run build
npm run render -- template.docx data.json output.docx attachments.json
npm run render -- template.xlsx data.json output.xlsx attachments.json
```

The optional `attachments.json` maps attachment keys to image paths relative to that file:

```json
{"logo": "./logo.png"}
```

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { compile } from './dist/index.js';

const template = await compile(await readFile('template.docx'));
const output = await template.render(
  { customer: { name: 'Example Company' } },
  { attachments: { logo: await readFile('logo.png') } },
);
await writeFile('output.docx', output);
```

Compiled templates can be reused for concurrent renders. Compilation copies input buffers; the template language does not execute JavaScript. The SDK and CLI detect DOCX/XLSX from the package and preserve that format. `compiled.format` reports the result type; the CLI requires a matching output extension. The HTTP endpoint also supports PDF conversion.

## HTTP API

`GET /health` requires no authentication. Submit rendering requests to `POST /v1/render` with these headers:

```http
Authorization: Bearer <IMPRINT_API_KEY>
Content-Type: application/json
```

```json
{
  "template": "<base64 DOCX or XLSX bytes>",
  "data": {"customer": {"name": "Example Company"}},
  "attachments": {"logo": "<base64 PNG or JPEG bytes>"},
  "output": "docx"
}
```

`output` can be `docx`, `xlsx` or `pdf`. Omit it to keep the detected template format. DOCX↔XLSX conversion is not supported; a mismatched format returns `422`. Successful responses contain the binary file with its matching content type and filename. The API admits two concurrent jobs and returns `503` with `Retry-After` when full.

| HTTP status | Meaning |
| --- | --- |
| 400 | Request schema validation failed |
| 401 | Missing or incorrect API key |
| 413 | Request or document exceeds a resource limit |
| 422 | Invalid data, template or unsupported feature |
| 502 | PDF converter unavailable or conversion failed |
| 503 | Render capacity is full |
| 504 | Render or conversion deadline exceeded |

## Test corpus and CI

Browse [`corpus/index.html`](corpus/index.html) locally for **100 DOCX and 100 XLSX levels**. Each level contains a template, JSON input, attachment mapping and reference document: 400 document files in total.

```bash
npm test
npm run corpus:generate
npm run corpus:validate
npm run corpus:render
npm run corpus:score
docker compose exec app npm run corpus:office
```

The renderer never reads reference documents. The separate scorer reports passes, failures, unsupported cases and unrun cases; exit code `2` means incomplete coverage, not a successful full corpus run. Add `-- --all` to `corpus:office` to attempt all 400 files; stress cases can still reach their conversion deadline. For the recovery integration check, start a separate Compose project with `GOTENBERG_API_TIMEOUT=5s` and run `scripts/converter-recovery.ts` inside its network, as shown in the CI workflow. The normal converter deadline is 60 seconds; the client bounds the complete PDF response at 65 seconds and 32 MiB.

The [first GitHub CI run](https://github.com/KenPrz/imprint-doc-gen/actions/runs/38117554855) passed type checking, build, tests, deterministic regeneration, fixture validation, renderer scoring and Compose configuration validation. That run did not exercise LibreOffice conversion. A separate Docker conversion CI job now builds the production services, checks DOCX/PDF responses and verifies timeout recovery; its result is reported by the workflow badge above.

Measured evidence (fixture package and full Office conversion counts are the milestone-1 baseline):

| Check | Result |
| --- | --- |
| Automated tests | 82 passed |
| Fixture validation | 200 cases / 400 document packages passed |
| DOCX renderer | All 100 levels passed |
| XLSX renderer | Levels 1–50 passed; 51–100 unsupported |
| LibreOffice fixture conversion | 385 converted, 3 failed, 12 not run |
| Converter recovery | Repeated timeout → successful conversion, including the production API |
| Selected XLSX PDF comparison | Levels 1, 20, 30, 40 and 50: all 10 pages matched at 96 DPI |
| Selected DOCX PDF comparison | Levels 1, 20, 50, 60, 70, 80 and 90: all 232 pages matched at 96 DPI |

These measurements do not establish complete OOXML schema conformance, repair-free Microsoft Office opening or general visual fidelity. Current XLSX evidence: [XLSX validation](docs/xlsx-cells-images-validation.json). Loop evidence: [DOCX loop validation](docs/docx-loops-validation.json). Baseline evidence: [milestone validation](docs/milestone-1-validation.json), [Office results](docs/office-validation.json).

## Roadmap

1. **Implemented:** bounded conversion failures, automatic LibreOffice recycling and Docker recovery tests. Continue investigating the upstream layout stall without changing the reference documents to hide it.
2. **Implemented:** DOCX repeating rows, paragraph blocks, nested scopes and per-item images. Expand real-world template coverage and advanced cloning adapters.
3. **Implemented:** XLSX typed cells and images. Next: row repetition with formula/range rewriting and independently verified calculated results.
4. Add durable jobs, tenant authentication, template versions and object storage.
5. Expand into PPTX, OpenDocument, charts, barcodes and advanced PDF operations.

See the [corpus roadmap](docs/corpus-roadmap.json), [stack decisions](docs/stack-proposal.json) and [core-engine research](docs/core-engine-research.json) for the longer-term plan.
