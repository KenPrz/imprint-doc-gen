# imprint-doc-gen

[![Document corpus](https://github.com/KenPrz/imprint-doc-gen/actions/workflows/corpus.yml/badge.svg)](https://github.com/KenPrz/imprint-doc-gen/actions/workflows/corpus.yml)

Independent document generation in TypeScript and Node.js. Render a DOCX template with JSON data and image attachments using the SDK, CLI or authenticated HTTP API. Convert the result to PDF through a private LibreOffice service.

The template compiler and renderer are implemented in this repository. Carbone and Docxtemplater are not runtime dependencies. This is an early service foundation, with **DOCX levels 1–100 passing**; it is not feature-complete or template-syntax compatible with either product.

## Current capabilities

- Scalar placeholders such as `{{customer.name}}`, including tags split across formatted text runs.
- Text replacement in document bodies, tables, hyperlinks, content controls, headers and footers.
- Replacement of existing pictures using alternative text such as `{{image:logo}}` and PNG/JPEG attachments.
- Independent image placements even when pictures share their original media; existing drawing dimensions are preserved.
- Repeated table rows and paragraph blocks, nested object scopes, parent lookups and per-item images.
- Reusable compiled templates, JSON-only data and explicit errors for missing or unsupported bindings.
- Authenticated DOCX/PDF API with bounded worker execution and request concurrency.

**Not implemented yet:** XLSX rendering, multiline/tabbed replacement values, inline image insertion, SVG, strict OOXML, tenant accounts, durable jobs, template storage or billing. `null` becomes empty text; `false` and `0` remain meaningful values.

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

See the [corpus guide](docs/corpus-guide.txt) for the full dialect, including planned XLSX semantics.

## SDK and CLI

For local development without the API container, use Node.js 24:

```bash
npm ci --ignore-scripts
npm run build
npm run render -- template.docx data.json output.docx attachments.json
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

Compiled templates can be reused for concurrent renders. Compilation copies input buffers; the template language does not execute JavaScript. The SDK and CLI produce DOCX; the HTTP endpoint also supports PDF conversion.

## HTTP API

`GET /health` requires no authentication. Submit rendering requests to `POST /v1/render` with these headers:

```http
Authorization: Bearer <IMPRINT_API_KEY>
Content-Type: application/json
```

```json
{
  "template": "<base64 DOCX bytes>",
  "data": {"customer": {"name": "Example Company"}},
  "attachments": {"logo": "<base64 PNG or JPEG bytes>"},
  "output": "docx"
}
```

`output` is `docx` (default) or `pdf`; successful responses contain the binary file. The API admits two concurrent jobs and returns `503` with `Retry-After` when full.

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
| Automated tests | 63 passed |
| Fixture validation | 200 cases / 400 document packages passed |
| DOCX renderer | All 100 levels passed |
| XLSX renderer | 100 levels unsupported |
| LibreOffice fixture conversion | 385 converted, 3 failed, 12 not run |
| Converter recovery | Repeated timeout → successful conversion, including the production API |
| Selected DOCX PDF comparison | Levels 1, 20, 50, 60, 70, 80 and 90: all 232 pages matched at 96 DPI |

These measurements do not establish complete OOXML schema conformance, repair-free Microsoft Office opening or general visual fidelity. Current loop evidence: [DOCX loop validation](docs/docx-loops-validation.json). Baseline evidence: [milestone validation](docs/milestone-1-validation.json), [Office results](docs/office-validation.json).

## Roadmap

1. **Implemented:** bounded conversion failures, automatic LibreOffice recycling and Docker recovery tests. Continue investigating the upstream layout stall without changing the reference documents to hide it.
2. **Implemented:** DOCX repeating rows, paragraph blocks, nested scopes and per-item images. Expand real-world template coverage and advanced cloning adapters.
3. Implement XLSX typed cells and images, then row repetition and formula/range rewriting.
4. Add durable jobs, tenant authentication, template versions and object storage.
5. Expand into PPTX, OpenDocument, charts, barcodes and advanced PDF operations.

See the [corpus roadmap](docs/corpus-roadmap.json), [stack decisions](docs/stack-proposal.json) and [core-engine research](docs/core-engine-research.json) for the longer-term plan.
