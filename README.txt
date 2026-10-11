IMPRINT-DOC-GEN

Independent TypeScript/Node.js document generation. Milestone 1 provides a DOCX
SDK, CLI and authenticated HTTP endpoint, with image attachment replacement and
PDF conversion. The template compiler and renderer are our own implementation.
Docxtemplater and Carbone are not runtime dependencies.

QUICK START — DOCKER DEVELOPMENT
  docker compose up --build -d
  docker compose logs -f app
  docker compose exec app npm run typecheck
  docker compose exec app npm test
  docker compose exec app node scripts/smoke.ts

compose.yml is the production base; compose.override.yml is loaded automatically
for development. Development mounts src/tests/scripts/docs read-only and corpus
read-write, runs TypeScript with Node watch mode, and uses the local development
API key "imprint-local-development-key" unless IMPRINT_API_KEY is set.
The endpoint is http://127.0.0.1:8088. GET /health needs no authentication.
The converter has no published host port and runs on an internal Docker network.

PRODUCTION CONFIGURATION
  cp .env.example .env
  # Set IMPRINT_API_KEY in .env to a generated secret of at least 24 characters.
  docker compose -f compose.yml up --build -d

The explicit -f omits the development override. Production contains compiled JS
and production dependencies, runs as a non-root user, has a read-only filesystem,
bounded temporary storage and resource limits, and refuses to start without a
key. Place a TLS reverse proxy in front of the loopback-bound API when deploying.
There is no externally deployed service from this repository setup.

Node and Gotenberg base images are pinned by digest. Dockerfile.gotenberg updates
the bundled pdfcpu from 0.15.0 to 0.16.1 using release asset SHA-256 verification.
The application enables Office-to-PDF conversion; it does not expose arbitrary
converter routes or download-from/webhook features. Only Linux amd64/arm64 have
patched converter build definitions; this workspace tests amd64.

SDK AND CLI
  npm ci --ignore-scripts
  npm run build
  npm run render -- template.docx data.json output.docx attachments.json

attachments.json maps keys to image paths relative to that JSON file:
  {"logo":"./logo.png"}

ES module SDK:
  import { compile, render } from './dist/index.js';
  const output = await render(templateBytes, data, {
    attachments: { logo: imageBytes }
  });
  const compiled = await compile(templateBytes);
  const anotherOutput = await compiled.render(otherData, { attachments });

Compiled templates can be reused for concurrent renders. Input buffers are
copied on compilation. Data is JSON-only; no template JavaScript is evaluated.

HTTP API
  POST /v1/render
  Authorization: Bearer <IMPRINT_API_KEY>
  Content-Type: application/json

  {
    "template": "<base64 DOCX bytes>",
    "data": {"customer":{"name":"Example"}},
    "attachments": {"logo":"<base64 PNG or JPEG bytes>"},
    "output": "docx"
  }

output can be "docx" or "pdf". The response is binary. The development API is
synchronous and admits two concurrent jobs. Each DOCX render runs in a separate
worker with a deadline; excess work gets 503 with Retry-After. Typed template
errors return 422, resource limits 413, timeouts 504, converter failures 502.

CURRENT SUPPORT
  - DOCX scalar placeholders, including tags fragmented across formatted runs.
  - Text in body, tables, hyperlinks, content controls, headers and footers.
  - Existing picture replacement through whole-picture alternative text bindings.
  - PNG/JPEG attachment bytes, bounded decoding and EXIF orientation correction.
  - Independent picture placements even when original media is shared.
  - Preservation of untouched XML bytes, package parts and drawing geometry.
  - Null -> empty text, false/zero preserved, missing fields produce errors.

CURRENT LIMITS
  - DOCX only; XLSX generation and structural loops are not implemented yet.
  - Multiline/tabbed replacement values, inline image insertion and SVG are not
    implemented. Unsupported templates fail explicitly instead of yielding a
    partial document. The source namespace must be transitional WordprocessingML;
    strict OOXML support is not claimed.
  - This is a single-service rendering foundation. Tenant accounts, durable jobs,
    database/object storage, quotas and billing remain future service milestones.
  - LibreOffice stalls on DOCX stress reference level 93; later conversions may
    time out until the converter is restarted (docker compose restart gotenberg).
    Converter recovery is a release gate before exposing the service externally.
  - It is not Carbone/Docxtemplater template-syntax compatible. The current dialect
    and semantics are documented in docs/corpus-guide.txt.

CORPUS AND VALIDATION
  Open corpus/index.html for 100 DOCX and 100 XLSX levels.
  npm run corpus:generate
  npm run corpus:validate
  npm run corpus:render
  npm run corpus:score
  docker compose exec app npm run corpus:office

The renderer reads templates/input/attachments, never reference documents.
The separate scorer reads expected documents and reports actual passes separately
from unsupported and unrun cases. Exit 2 means coverage is incomplete; exit 1
means a failure. Do not update reference outputs to conceal renderer regressions.
corpus:office converts representative template/reference files with LibreOffice.
Add -- --all to cover all 400 files. Conversion success is not visual approval.

See docs/milestone-1-validation.json for measured results, docs/corpus-guide.txt
for the corpus contract, and docs/corpus-roadmap.json for future formats/features.

REPOSITORY
  Intended private GitHub repository: KenPrz/imprint-doc-gen
  The local Git repository is the source of truth until GitHub creation/push has
  been confirmed. Check docs/milestone-1-validation.json for publication status.
