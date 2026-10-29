# Additional image fonts

Put font files (`.ttf`, `.otf`, or `.ttc`) and their license/notice files here,
optionally in family subdirectories. Both PDF-capable Dockerfiles copy this
directory into `/usr/local/share/fonts/imprint/custom` at build time. Rebuild and
recreate the converter (or standalone app) after changing the files.

Only include files you can use and distribute in the resulting image. The public
build bundles Google Fonts; it does not bundle Aptos, Calibri, Helvetica, Acumin
Pro, or other independently licensed fonts. Microsoft's freely downloadable
Aptos package permits installation on your devices but grants no distribution
rights. A font name in Carbone's cloud list is not a redistribution license.

Local files are Git-ignored but intentionally included in Docker's build context.
CI images contain only files actually provided to their build context. Never
commit private font files or font license keys. For an image with supplied Aptos
files, require the family to prevent accidentally building without it:

```sh
docker build -f Dockerfile.standalone \
  --build-arg 'REQUIRED_FONTS=Roboto,Lobster,Open Sans,Noto Sans,Aptos' \
  -t imprint-doc-gen:custom-fonts .
```

Inspect `/usr/local/share/imprint/fonts/carbone-coverage.json` in the image for
exact present/missing font-face names. A family being installed does not prove
every style, character, or document layout renders identically to Word.
