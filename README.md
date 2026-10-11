# imprint-doc-gen

[![Document corpus](https://github.com/KenPrz/imprint-doc-gen/actions/workflows/corpus.yml/badge.svg)](https://github.com/KenPrz/imprint-doc-gen/actions/workflows/corpus.yml)

Independent document generation in TypeScript and Node.js. Render a DOCX or XLSX template with JSON data and image attachments using the SDK, CLI or authenticated HTTP API. Run it as one Docker image with bundled PDF conversion, persistent templates and durable jobs, or use the separate development services.

The template compiler and renderer are implemented in this repository. Carbone and Docxtemplater are not runtime dependencies. This is an early service foundation, with **DOCX and XLSX levels 1–100 passing**; it is not feature-complete. **Carbone-style syntax is the standard for DOCX and XLSX**, all 200 corpus templates, and the PMS examples. The 30 original seeded PMS DOCX reports also render unchanged. See the [syntax contract](docs/template-syntax.txt) and [PMS compatibility guide](docs/carbone-compatibility.txt) for the supported subset.

## Current capabilities

- Carbone-style `{d.customer.name}` placeholders in Word and Excel, including tags split across formatted text runs.
- `[i]`/`[i+1]` row repetition in both formats, nested DOCX blocks, and picture alternative text `{d.logo}`.
- Text replacement in document bodies, tables, hyperlinks, content controls, headers and footers.
- Replacement of existing pictures using alternative text such as `{d.logo}` and PNG/JPEG attachments.
- Independent image placements even when pictures share their original media; existing drawing dimensions are preserved.
- Repeated table rows and paragraph blocks, nested object scopes, parent lookups and per-item images.
- Reusable compiled templates, JSON-only data and explicit errors for missing or unsupported bindings.
- XLSX shared/rich strings, typed cells, repeated rows, formula/range updates and anchored picture replacement.
- Authenticated DOCX/XLSX/PDF API with bounded worker execution and request concurrency.
- Tenant API keys, immutable template versions, publish/rollback, and durable jobs with restart recovery.
- Single-container deployment with SQLite persistence and a working PMS invoice example.

**Not implemented yet:** nested/multi-row XLSX loops, a general Excel calculation engine, multiline/tabbed values in native DOCX `{{...}}` bindings, full Carbone formatter/filter syntax, inline image insertion, SVG, strict OOXML, customer login/SSO, PostgreSQL/S3 adapters or billing. `null` becomes empty text in DOCX and a blank cell for whole-cell XLSX bindings; `false` and `0` remain meaningful values.

## Quick start: one container

The release workflow publishes a **linux/amd64 preview** to GitHub Container Registry after all CI gates pass. The private package requires a GitHub login with package read access. Tested `main` builds publish `standalone`, and every release also receives a `sha-<commit>` tag.

```bash
docker login ghcr.io -u YOUR_GITHUB_USERNAME
docker pull ghcr.io/kenprz/imprint-doc-gen:standalone
# Export IMPRINT_API_KEY from your secret manager: at least 24 characters.
docker run -d --name imprint-doc-gen --restart unless-stopped \
  --stop-timeout 120 --memory 2g --cpus 2 --pids-limit 384 \
  --read-only --tmpfs /tmp:rw,size=768m,mode=1777 \
  --cap-drop ALL --security-opt no-new-privileges:true \
  -p 127.0.0.1:8088:3000 -e IMPRINT_API_KEY \
  -v imprint-data:/data \
  ghcr.io/kenprz/imprint-doc-gen:standalone
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

## Standard template syntax

Use **Carbone-style syntax** for every new template. No syntax flag or preprocessing is required; the SDK, CLI, synchronous endpoints and durable jobs use the same compiler.

| Purpose | Template |
| --- | --- |
| Root or nested field | `{d.customer.name}` |
| Repeated row field | `{d.items[i].name}` |
| Next-row marker | `{d.items[i+1].name}` |
| Nested DOCX field | `{d.groups[i].items[i].name}` |
| Picture alternative text | `{d.logo}` |

Missing/null scalars become blank; zero and false remain meaningful. Missing/null/empty arrays remove the repeated region. Replacements are escaped, support line breaks/tabs, and never execute JavaScript or recursively interpret data as tags. Loops accept arrays of objects.

### Pictures

Insert a placeholder picture and set its alternative-text description to `{d.logo}`. Supply PNG/JPEG bytes under the attachment key `logo`, or set `data.logo` to another attachment key or a PNG/JPEG data URL. Missing optional images preserve their placeholder. Geometry is retained and shared original media is replaced independently for each placement. Remote URLs and SVG are unsupported.

DOCX repeated pictures can use `{d.items[i].image}` with an attachment key or data URL on each item. XLSX pictures outside repeated rows support the same root binding syntax; pictures intersecting repeated rows still require a cloning adapter.

### DOCX repetition

Put `{d.items[i].name}` and other item fields in the row to repeat. Follow it with a marker row containing `{d.items[i+1].name}`. The marker row is discarded. Root fields such as `{d.customer.name}` retain their root meaning inside repeats.

Paragraph blocks use the same pattern: the first `[i]` paragraph begins the block, and its matching `[i+1]` paragraph ends it and is discarded. Nested arrays use full paths such as `{d.groups[i].items[i].name}`; parent fields use `{d.groups[i].name}`.

Sibling multi-row DOCX patterns are supported. Inline repeats, repeated bookmarks/comments/notes, tracked changes, vertical merges, bound controls and VML/OLE need later adapters. Basic controls and DrawingML pictures get fresh identities when cloned. Limits include 16 nested loops, 10,000 total iterations, 8 MiB per XML part and 32 MiB per expanded package.

### XLSX repetition, types and formulas

Use a pattern row followed immediately by its marker row:

| Cell | Template |
| --- | --- |
| A40 | `{d.items[i].name}` |
| B40 | `{d.items[i].quantity}` |
| C40 | `{d.items[i].price}` |
| D40 | `=B40*C40` (ordinary Excel formula) |
| A41 | `{d.items[i+1].name}` |

The renderer repeats row 40 and removes row 41. It updates following row coordinates, formulas, tables, filters, names, validation ranges, conditional formatting and picture anchors. Multiple independent repeats are supported. Nested, horizontal and multi-row XLSX patterns remain unsupported.

A whole-cell scalar preserves JSON numbers and booleans as native Excel values; null/missing becomes a blank cell with its style retained. Embedded bindings produce text. Strings beginning with `=` stay literal. Shared/rich strings, hidden worksheets, line breaks, tabs and Excel escape sequences are supported; strings are limited to 32,767 UTF-16 code units.

Formula copies adjust relative rows while preserving `$` locks; structural insertion and marker removal also move absolute and cross-sheet targets. A single-cell reference to the pattern targets its first output row. Ranges expand over the output rows; a reference to a deleted cell becomes `#REF!`.

Without repeats, existing formulas and caches are preserved and recalculation is requested. With repeats, Imprint calculates supported A1 formulas: arithmetic, comparisons, text concatenation, `SUM`, `MIN`, `MAX`, `AVERAGE`, `COUNT` and `IF`. This remains a bounded calculator; advanced functions, structured/named operands, shared/array/data-table formulas and cycles fail explicitly. Formula/dependency depth is limited to 128 with one million operations. Repeats also enforce Excel bounds, 10,000 total iterations and the package byte budgets.

Merged cells or pictures intersecting changed rows, charts, pivots, comments and unsupported extensions need later adapters. See [XLSX calculation evidence](docs/xlsx-rows-validation.json).

### Existing PMS templates and legacy bindings

The 30 seeded PMS DOCX templates keep their existing bytes and keys. `POST /render/template?download=true` accepts the existing `template`, `data`, `convertTo` request shape; migration needs a service URL and Bearer credentials. Custom uploaded templates still need auditing. See the [PMS migration guide](docs/carbone-compatibility.txt).

Earlier `{{name}}`, `{{#each items}}` and `{{image:logo}}` bindings remain supported for compatibility. Their strict missing-data and scope rules are unchanged. They are no longer used in the corpus or examples. Do not mix legacy and Carbone repeat markers in one Word story or worksheet.

Standardizing syntax does not imply complete Carbone feature parity. Formatter chains, filters/conditions, arbitrary iterators, translations and aliases require future adapters. Every future document-format adapter must use the same standard syntax; unsupported operations must fail explicitly.

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

## Synthetic PMS testing without a PMS pilot

Run the original 30 seeded PMS DOCX templates against fake guest, billing and housekeeping data, plus logo-bearing DOCX/XLSX invoice examples:

```bash
npm run pms:synthetic
# Optional: use a running Imprint service and inspect PDFs.
# Set IMPRINT_URL and IMPRINT_API_KEY in your environment first.
npm run pms:synthetic -- --api --pdf
```

The suite has **128 cases** covering normal data, empty lists, multilingual/long text and larger reports. It supplies synthetic PNG/JPEG signatures and a hotel logo. Open `artifacts/pms-synthetic/index.html` for documents, data, images, previews and results. PDF checks require Poppler tools; the default run performs offline document checks. See the [synthetic test guide](docs/synthetic-pms-guide.txt).

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
| Automated tests | 312 passed |
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
5. **Implemented:** Carbone-style syntax is standard across DOCX, XLSX, all 200 corpus templates and PMS examples; legacy syntax remains supported. The original 30 PMS reports stay unchanged. Audit custom uploads and pilot the URL/auth change next; formatter/filter support and advanced repeats follow real-template requirements.
6. Expand into PPTX, OpenDocument, charts, barcodes and advanced PDF operations.

See the [corpus roadmap](docs/corpus-roadmap.json), [stack decisions](docs/stack-proposal.json) and [core-engine research](docs/core-engine-research.json) for the longer-term plan.
