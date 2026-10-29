# imprint-doc-gen

[![Document corpus](https://github.com/KenPrz/imprint-doc-gen/actions/workflows/corpus.yml/badge.svg)](https://github.com/KenPrz/imprint-doc-gen/actions/workflows/corpus.yml)
[![Documentation](https://github.com/KenPrz/imprint-doc-gen/actions/workflows/pages.yml/badge.svg)](https://kenprz.github.io/imprint-doc-gen/)

Independent document generation in TypeScript and Node.js. Render **DOCX, XLSX, PPTX, ODT, ODS and ODP** templates with JSON and image attachments, or convert them to PDF. Run one image with bundled conversion and SQLite persistence, or use PostgreSQL/S3 with independent workers.

**[Read the documentation →](https://kenprz.github.io/imprint-doc-gen/)** · [API](https://kenprz.github.io/imprint-doc-gen/api.html) · [Template syntax](https://kenprz.github.io/imprint-doc-gen/syntax.html) · [Support matrix](https://kenprz.github.io/imprint-doc-gen/formats.html)

Version 1.0.0 provides bounded format support. The template compiler, repeat expansion and package adapters live in this repository. Third-party libraries provide ZIP/XML primitives, image transforms, PDF manipulation, barcodes and infrastructure clients. It is not a universal Office editor or a complete spreadsheet calculation engine.

## Capabilities

- Standard `{d.path}` fields, `[i]`/`[i+1]` repeats and picture metadata across all six formats.
- PNG/JPEG replacement, contain/cover, crop, rotation, orientation and per-item spreadsheet pictures.
- Locale-aware number/currency/date formatters, required fields, conditional values and array filters.
- DOCX nested repeats and note fields; XLSX multi-row blocks, contained merged cells and bounded formula calculation.
- Template inspection, field diagnostics and PDF preview through the SDK, CLI and API.
- PDF merge/split, text watermarks, page numbers, attachments and barcode/QR images.
- Tenant keys, immutable template versions, durable jobs, metrics and verified SQLite backup/restore.
- Optional PostgreSQL metadata/queue, S3 objects, independent workers and signed completion webhooks.

## Quick start: one container

The release workflow publishes a **linux/amd64 image** to GitHub Container Registry after all CI gates pass. The private package requires a GitHub login with package read access. Tested `main` builds publish `standalone`, and each build receives a `sha-<commit>` tag. Versioned releases additionally publish `v1.0.0`, `v1` and `latest` tags.

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

The API, queue, database and PDF converter run in that container. The named volume preserves templates, keys, jobs and results across container replacement. Deploy one instance per local volume. For independent workers, use the optional PostgreSQL/S3 deployment described in the documentation.

To build locally, use `docker build -f Dockerfile.standalone -t imprint-doc-gen:standalone .`, then set `IMPRINT_IMAGE=imprint-doc-gen:standalone` when starting [`compose.standalone.yml`](compose.standalone.yml). The registry tag becomes available only when publication succeeds. Pin a tested image digest for production and back up the volume before upgrades.

Try the [example client](examples/invoice/client.ts) and checked-in invoice templates with Node.js 24:

```bash
# Uses IMPRINT_API_KEY from your environment. Prefer a dedicated tenant key.
IMPRINT_URL=http://127.0.0.1:8088 node examples/invoice/generate.ts pdf
IMPRINT_URL=http://127.0.0.1:8088 node examples/invoice/generate.ts xlsx
```

The example uploads and publishes a template, submits invoice data and a logo attachment, polls the durable job, and saves `artifacts/example-invoice.<format>`. See the [standalone guide](docs/standalone-guide.txt) for tenant provisioning, endpoint contracts, retention, backups and limits.

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

**Bundled fonts:** Both PDF-capable images include a pinned, SHA-256-verified Google Fonts snapshot (3,913 files in 2,066 family directories), alongside the base image's fonts. The font files add about 2.3 GiB before image compression; the initial build downloads about 1.6 GiB and caches the archive. Conversion needs no font downloads at runtime. Original font files, available family licenses/notices, metadata, and shared license texts are retained.

This is not the complete Carbone cloud catalog: commercial fonts such as Aptos and Calibri are not included by default, and upstream font names/styles can differ. Each image contains an actual font inventory and a present/missing comparison against the pinned 6,361-entry Carbone list under `/usr/local/share/imprint/fonts/`. Add separately licensed fonts through [`docker/fonts/custom/`](docker/fonts/custom/README.md), then rebuild and recreate the container. See [font deployment and verification](docs/site/pages/deployment.html) for exact commands. Substituted fonts are never counted as installed coverage.

The validated default images match 4,771 reference face names and report 1,590 missing names. See [font validation evidence](docs/font-validation.json). LibreOffice loads the catalog before readiness so initialization does not consume the first document's request deadline. The font unit checks require Python 3.

**Known conversion limitation:** LibreOffice still stalls on DOCX stress reference level 93. Compose now recycles LibreOffice after every conversion attempt, including timeouts, so the following document can render without a manual restart. A timed-out request returns JSON with HTTP `504`; the failing document is not retried automatically or reported as converted. Recycling trades some throughput for predictable recovery. See [recovery evidence](docs/converter-recovery.json).

## Development and validation

```bash
npm ci --ignore-scripts
npm run typecheck
npm run build
npm test
npm run test:fonts
npm run corpus:render && npm run corpus:score
npm run test:formats
npm run test:synthetic
npm run docs:build
```

The original 100 DOCX and 100 XLSX levels retain independent expected outputs and scoring. Four additional 100-case format suites cover the implemented adapters. Set `GOTENBERG_URL` when running `test:formats` to compare 12 selected PDF conversions against independently generated expected documents; this requires Poppler. See [format limits](https://kenprz.github.io/imprint-doc-gen/formats.html), [operations](https://kenprz.github.io/imprint-doc-gen/operations.html) and [validation evidence](docs/roadmap-validation.json).

The documentation is static HTML/CSS with a small search index: no frontend framework or runtime server. `docs/site/pages/` contains the source, `npm run docs:build` validates links and writes `artifacts/docs/`, and the Pages workflow publishes it from main.

## Roadmap

The requested eight workstreams are implemented within the documented support contract. The completed feature expansion adds CSV/plain text, repeated classic notes/comments, explicit numbering restarts, passive SVG and inline DOCX pictures, ODS reference rewriting, more XLSX formulas, literal category charts, worker metrics aggregation and offline cluster migration. See [expanded feature contracts](https://kenprz.github.io/imprint-doc-gen/features.html). Remaining work includes representative application pilot acceptance, workbook-linked charts/pivots, threaded comments, broader formulas, full Office repair/schema validation and specialist formats. See the [release roadmap](https://kenprz.github.io/imprint-doc-gen/roadmap.html), [implementation tracking](docs/implementation-plan.json) and [longer-term corpus roadmap](docs/corpus-roadmap.json).

Current tests do not establish universal layout fidelity, PDF conformance or production capacity. The known LibreOffice DOCX stress-case timeout remains explicit and is covered by recovery checks.

## Document assurance (implementation branch)

The SDK, `POST /v1/assurance`, and `npm run assurance -- request.json evidence.json` expose typed input contracts, execution traces, customer regression suites, semantic/PDF visual comparisons, replay bundles, conservative migration, pinned text and rich DOCX components, bounded RPT artifact inspection, validated artifact handoff, interactive report snapshots and candidate-template verification. See [support contracts and examples](docs/site/pages/assurance.html), [implementation gates](docs/assurance-plan.json), and [independent research](docs/competitive-research.json). Version 1.0.0 is owner-approved with documented limitations: customer-specific visual acceptance, desktop Office repair-free validation, accessibility certification and full Crystal semantic migration remain unverified.

Run `npm test` (requires Poppler), `npm run test:assurance:browser` (requires `npx playwright install chromium`) and `npm run test:assurance:verapdf` (requires Docker; uses a pinned independent validator image). Generated evidence is in `artifacts/assurance/`.
