#!/usr/bin/env python3
"""Check actual DOCX-to-PDF font embedding through a running Gotenberg endpoint."""
import argparse
import hashlib
import io
import json
from pathlib import Path
import re
import tempfile
import urllib.request
import uuid
from xml.sax.saxutils import escape, quoteattr
import zipfile


def font_matches(family, name):
    normalize = lambda value: re.sub(r"[^a-z0-9]", "", value.casefold())
    styles = ["", "Regular", "Roman", "Book", "Medium", "Bold", "Italic", "Oblique",
              "BoldItalic", "BoldOblique", "RomanRegular", "RomanBold", "ItalicRegular", "ItalicBold"]
    return normalize(name) in {normalize(family + style) for style in styles}


def document(family):
    name = quoteattr(family)
    paragraphs = []
    for style in ["", "<w:b/>", "<w:i/>", "<w:b/><w:i/>"]:
        paragraphs.append(f'<w:p><w:r><w:rPr><w:rFonts w:ascii={name} w:hAnsi={name} '
                          f'w:eastAsia={name} w:cs={name}/><w:sz w:val="28"/>{style}</w:rPr>'
                          f'<w:t>{escape(family)} font preservation: AaBb 0123456789</w:t></w:r></w:p>')
    parts = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="word/document.xml"/></Relationships>',
        "word/document.xml": '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body>' + ''.join(paragraphs) + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>',
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as package:
        for name, value in parts.items():
            package.writestr(name, value)
    return buffer.getvalue()


def check(url, families, output):
    output.mkdir(parents=True, exist_ok=True)
    results = []
    for index, family in enumerate(families):
        docx = document(family)
        boundary = "imprint-font-" + uuid.uuid4().hex
        body = (f'--{boundary}\r\nContent-Disposition: form-data; name="files"; filename="font.docx"\r\n'
                'Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n'
                ).encode() + docx + f'\r\n--{boundary}--\r\n'.encode()
        request = urllib.request.Request(url.rstrip("/") + "/forms/libreoffice/convert", data=body,
                                         headers={"Content-Type": "multipart/form-data; boundary=" + boundary})
        with urllib.request.urlopen(request, timeout=90) as response:
            pdf = response.read()
        (output / f"font-{index}.docx").write_bytes(docx)
        (output / f"font-{index}.pdf").write_bytes(pdf)
        names = set()
        for raw in re.findall(rb"/BaseFont\s*/([^\s/<>]+)", pdf):
            decoded = re.sub(rb"#([0-9a-fA-F]{2})", lambda m: bytes([int(m[1], 16)]), raw).decode()
            names.add(decoded.split("+")[-1])
        passed = pdf.startswith(b"%PDF-") and bool(names) and all(
            font_matches(family, name) for name in names)
        passed = passed and len(re.findall(rb"/FontFile[23]?\b", pdf)) >= len(names)
        result = {"family": family, "embeddedFonts": sorted(names), "passed": passed,
                  "pdfSha256": hashlib.sha256(pdf).hexdigest()}
        results.append(result)
        print(json.dumps(result), flush=True)
    (output / "font-smoke.json").write_text(json.dumps(results, indent=2) + "\n")
    if not all(result["passed"] for result in results):
        raise SystemExit("Font preservation failed: requested family was substituted or not embedded")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--url", default="http://127.0.0.1:3000")
    parser.add_argument("--font", action="append", dest="fonts")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    families = args.fonts or ["Roboto", "Open Sans", "Lobster", "Ubuntu"]
    if args.output:
        check(args.url, families, args.output)
    else:
        with tempfile.TemporaryDirectory(prefix="imprint-font-smoke-") as temp:
            check(args.url, families, Path(temp))
