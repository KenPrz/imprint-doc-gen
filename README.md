# imprint-doc-gen

[![Document corpus](https://github.com/KenPrz/imprint-doc-gen/actions/workflows/corpus.yml/badge.svg)](https://github.com/KenPrz/imprint-doc-gen/actions/workflows/corpus.yml)

Independent document generation in TypeScript and Node.js. Render a DOCX or XLSX template with JSON data and image attachments using the SDK, CLI or authenticated HTTP API. Run it as one Docker image with bundled PDF conversion, persistent templates and durable jobs, or use the separate development services.

The template compiler and renderer are implemented in this repository. Carbone and Docxtemplater are not runtime dependencies. This is an early service foundation, with **DOCX and XLSX levels 1–100 passing**; it is not feature-complete. A tested Carbone syntax adapter renders the **30 original seeded PMS DOCX reports without editing their templates**; see the [compatibility guide](docs/carbone-compatibility.txt) for supported syntax and limits.

## Current capabilities

- PMS Carbone DOCX keys such as `{d.property_name}`, `[i]`/`[i+1]` row and nested block repetition, and signature-image bindings.
- Scalar placeholders such as `{{customer.name}}`, including tags split across formatted text runs.
- Text replacement in document bodies, tables, hyperlinks, content controls, headers and footers.
- Replacement of existing pictures using alternative text such as `{{image:logo}}` and PNG/JPEG attachments.
- Independent image placements even when pictures share their original media; existing drawing dimensions are preserved.
- Repeated table rows and paragraph blocks, nested object scopes, parent lookups and per-item images.
- Reusable compiled templates, JSON-only data and explicit errors for missing or unsupported bindings.
- XLSX shared/rich strings, typed cells, repeated rows, formula/range updates and anchored picture replacement.
- Authenticated DOCX/XLSX/PDF API with bounded worker execution and request concurrency.
- Tenant API keys, immutable template versions, publish/rollback, and durable jobs with restart recovery.
- Single-container deployment with SQLite persistence and a working PMS invoice example.

**Not implemented yet:** nested/multi-row XLSX loops, a general Excel calculation engine, multiline/tabbed values in native DOCX `{{...}}` bindings, full Carbone formatter and XLSX syntax, inline image insertion, SVG, strict OOXML, customer login/SSO, PostgreSQL/S3 adapters or billing. `null` becomes empty text in DOCX and a blank cell for whole-cell XLSX bindings; `false` and `0` remain meaningful values.

## Quick start: one container

The release workflow publishes a **linux/amd64 preview** to GitHub Container Registry after all CI gates pass. The private package requires a GitHub login with package read access. `standalone-preview` follows the active standalone/compatibility feature branches; reviewed `main` builds publish `standalone`, and every release also receives a `sha-<commit>` tag.

```bash
docker login ghcr.io -u YOUR_GITHUB_USERNAME
docker pull ghcr.io/kenprz/imprint-doc-gen:standalone-preview
# Export IMPRINT_API_KEY from your secret manager: at least 24 characters.
docker run -d --name imprint-doc-gen --restart unless-stopped \
  --stop-timeout 120 --memory 2g --cpus 2 --pids-limit 384 \
  --read-only --tmpfs /tmp:rw,size=768m,mode=1777 \
  --cap-drop ALL --security-opt no-new-privileges:true \
  -p 127.0.0.1:8088:3000 -e IMPRINT_API_KEY \
  -v imprint-data:/data \
  ghcr.io/kenprz/imprint-doc-gen:standalone-preview
curl http://127.0.0.1:8088/ready
```

The API, queue, database and PDF converter run in that container. The named volume preserves templates, keys, jobs and results across container replacement. Deploy one instance per local volume. PostgreSQL/S3 and multiple-server operation remain future work.

To build locally, use `docker build -f Dockerfile.standalone -t imprint-doc-gen:standalone .`, then set `IMPRINT_IMAGE=imprint-doc-gen:standalone` when starting [`compose.standalone.yml`](compose.standalone.yml). The registry tag becomes available only when publication succeeds. Pin a tested image digest for production and back up the volume before upgrades.

Try the [PMS client](examples/pms/client.ts) and checked-in invoice templates with Node.js 24:

```bash
# Uses IMPRINT_API_KEY from your environment. Prefer a dedicated tenant key.
IMPRINT_URL=http://127.0.0.1:8088 node examples/pms/generate.ts pdf
IMPRINT_URL=http://127.0.0.1:8088 node examples/pms/generate.ts xlsx
```

The example uploads and publishes a template, submits booking data and a logo attachment, polls the durable job, and saves `artifacts/booking-invoice.<format>`. See the [standalone guide](docs/standalone-guide.txt) for tenant provisioning, endpoint contracts, retention, backups and limits.

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

Compose automatically combines [`compose.yml`](compose.yml) with [`compose.override.yml`](compose.override.yml). Development runs TypeScript in Node watch mode, mounts source/tests/scripts read-only and makes the corpus writable. The converter has no published host port. Service data persists in the `app-data` named volume; `docker compose down` preserves it.

## Production configuration: separate services

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

### Existing PMS / Carbone templates

Keep `{d.property_name}` and the existing `{d.rows[i].field}` / `{d.rows[i+1].field}` rows in your DOCX files. Nested task-sheet arrays and image alternative text `{d.guest_signature}` are supported. Missing values become blank, and empty arrays remove their repeat region.

The authenticated `POST /render/template?download=true` endpoint accepts the PMS request body (`template`, `data`, `convertTo`) and returns the generated file. Migration needs a service URL and Bearer key change; the 30 audited report templates stay unchanged. Custom uploaded templates still need auditing. Formatter chains and Carbone XLSX syntax remain unsupported. See [scope, evidence and migration details](docs/carbone-compatibility.txt).

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
- Native multi-row loop regions, inline text loops, repeated bookmarks/comments/notes, tracked changes, vertical merges, bound content controls and VML/OLE content require later adapters and fail explicitly. Basic content controls and DrawingML pictures get fresh identities when repeated.

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

For workbooks without row loops, existing formulas and cached results are preserved and full recalculation is requested on opening. Formula placeholders, automatic date conversion and bound phonetic strings require later adapters.

### XLSX row loops and formulas

Place `{{#each items}}` and `{{/each}}` in the **same worksheet row**. Every binding in that row uses its current item; `{{../title}}` reads the root object. Put ordinary Excel formulas in other cells of the row, for example:

| A40 | B40 | C40 | D40 | E40 |
| --- | --- | --- | --- | --- |
| `{{#each items}}{{name}}` | `{{quantity}}` | `{{price}}` | `=B40*C40` | `{{/each}}` |

The renderer repeats the row, keeps cell styles and row height, removes marker-only cells, and writes fresh formula results. Empty arrays remove the row. Multiple independent row loops can appear in a workbook; rows below each loop move by its expansion amount.

- Formula copies adjust relative rows while preserving `$` locks. Structural row insertion/deletion also moves absolute targets and cross-sheet references. An outside single-cell reference to a repeated row targets its first copy; range ends expand across all copies. A reference to a deleted cell becomes `#REF!`.
- Tables, filters, defined names, validation and conditional-formatting ranges follow row changes. Qualified workbook names and sheet-local A1 names are supported. Merged ranges below loops move; merges intersecting a loop are rejected.
- Existing pictures below a loop move with their cell anchors. Pictures intersecting a repeated row need a later cloning adapter and fail explicitly.
- In workbooks containing loops, all formulas must fit the bounded calculator: A1 cells/ranges, quoted sheet names, arithmetic, comparisons, text concatenation, and `SUM`, `MIN`, `MAX`, `AVERAGE`, `COUNT`, `IF`. Dependencies are recalculated across worksheets, and obsolete calculation chains are removed. Circular references fail explicitly. This is not a general Excel engine; advanced functions, named operands, structured references and shared/array/data-table formulas are rejected. Boolean concatenation and zero raised to a nonpositive power also return `UNSUPPORTED_FORMULA` instead of guessing at application-specific behavior.
- Nesting/multi-row loops, charts, pivots, comments, legacy drawings and unsupported extension metadata need later adapters. Loop workbooks require explicit ascending row coordinates. Limits include 10,000 total iterations, Excel row/column bounds, 8 MiB per part, 32 MiB per expanded package, 128 levels of formula/dependency depth and one million calculation operations.

The independent LibreOffice check in CI verifies cached results for representative supported expressions; see [XLSX loop validation](docs/xlsx-rows-validation.json).

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

`GET /health` and `GET /ready` require no authentication. The durable template/job endpoints are documented in the [standalone guide](docs/standalone-guide.txt). Submit rendering requests to `POST /v1/render` with these headers:

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

`output` can be `docx`, `xlsx` or `pdf`. Omit it to keep the detected template format. DOCX↔XLSX conversion is not supported; a mismatched format returns `422`. Successful responses contain the binary file with its matching content type and filename. The API admits two concurrent uploads/renders across synchronous requests and queued work, and returns `503` with `Retry-After` when synchronous capacity is full. Durable requests use `POST /v1/jobs` and wait in the queue.

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
| Automated tests | 169 passed |
| Original PMS Carbone templates | 30 unchanged DOCX files match independent snapshots; 9 extra array-length cases |
| Selected PMS PDF comparison | Five reports, seven pages: identical pixels at 96 DPI |
| Fixture validation | 200 cases / 400 document packages passed |
| DOCX renderer | All 100 levels passed |
| XLSX renderer | All 100 levels passed |
| LibreOffice fixture conversion | 385 converted, 3 failed, 12 not run |
| Converter recovery | Repeated timeout → successful conversion, including the production API |
| Selected XLSX loop PDF comparison | Levels 60, 80, 90 and 100: all 64 pages matched at 96 DPI |
| Independent formula calculation | 52 cached results match LibreOffice recalculation |
| Selected DOCX PDF comparison | Levels 1, 20, 50, 60, 70, 80 and 90: all 232 pages matched at 96 DPI |

These measurements do not establish complete OOXML schema conformance, repair-free Microsoft Office opening or general visual fidelity. Current XLSX evidence: [XLSX loop validation](docs/xlsx-rows-validation.json). Earlier cell/image evidence: [XLSX validation](docs/xlsx-cells-images-validation.json). Loop evidence: [DOCX loop validation](docs/docx-loops-validation.json). Baseline evidence: [milestone validation](docs/milestone-1-validation.json), [Office results](docs/office-validation.json).

## Roadmap

1. **Implemented:** bounded conversion failures, automatic LibreOffice recycling and Docker recovery tests. Continue investigating the upstream layout stall without changing the reference documents to hide it.
2. **Implemented:** DOCX repeating rows, paragraph blocks, nested scopes and per-item images. Expand real-world template coverage and advanced cloning adapters.
3. **Implemented:** XLSX typed cells, images, repeated rows, formula/range rewriting and bounded calculated results. Expand real-world fixtures, picture cloning and advanced formula adapters.
4. **Implemented for standalone:** one image, tenant keys, SQLite storage, immutable template versions, durable jobs and a PMS invoice example. Add a real PMS pilot, operational metrics and measured load limits next; PostgreSQL/S3 and separate workers follow when multiple-server deployment is needed.
5. **Implemented for seeded PMS reports:** Carbone DOCX syntax and the PMS one-shot API contract. Audit custom uploaded templates and pilot the URL/auth change next; see [compatibility evidence](docs/carbone-compatibility.json).
6. Expand into PPTX, OpenDocument, charts, barcodes and advanced PDF operations.

See the [corpus roadmap](docs/corpus-roadmap.json), [stack decisions](docs/stack-proposal.json) and [core-engine research](docs/core-engine-research.json) for the longer-term plan.
