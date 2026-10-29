#!/usr/bin/env python3
"""Install pinned fonts at build time; audit installed faces without fallback matches."""
import argparse
import hashlib
import json
from pathlib import Path, PurePosixPath
import shutil
import subprocess
import tarfile
import tempfile
import urllib.request

HERE = Path(__file__).resolve().parent
ROOTS = {"ofl", "apache", "ufl"}
FONT_EXTENSIONS = {".ttf", ".otf", ".ttc"}


def digest(path):
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def download(source, cache):
    cache.mkdir(parents=True, exist_ok=True)
    archive = cache / (source["sha256"] + ".tar.gz")
    if archive.is_file() and digest(archive) == source["sha256"]:
        return archive
    # Only a complete, verified download becomes a reusable cache entry.
    with tempfile.NamedTemporaryFile(dir=cache, delete=False) as stream:
        temporary = Path(stream.name)
        try:
            print("Downloading pinned Google Fonts snapshot", flush=True)
            with urllib.request.urlopen(source["url"], timeout=120) as response:
                shutil.copyfileobj(response, stream, length=1024 * 1024)
            stream.close()
            if digest(temporary) != source["sha256"]:
                raise ValueError("Google Fonts archive SHA-256 mismatch")
            temporary.replace(archive)
        finally:
            temporary.unlink(missing_ok=True)
    return archive


def install_archive(archive, destination, source):
    if digest(archive) != source["sha256"]:
        raise ValueError("Google Fonts archive SHA-256 mismatch")
    destination.mkdir(parents=True, exist_ok=False)
    count, families = 0, set()
    try:
        with tarfile.open(archive, "r|gz") as package:
            for member in package:
                path = PurePosixPath(member.name)
                if path.is_absolute() or ".." in path.parts:
                    raise ValueError("Unsafe path in font archive")
                if not path.parts or path.parts[0] != "fonts-" + source["revision"]:
                    raise ValueError("Unexpected Google Fonts archive root")
                relative = PurePosixPath(*path.parts[1:])
                if len(relative.parts) < 3 or relative.parts[0] not in ROOTS:
                    continue
                is_font = relative.suffix.lower() in FONT_EXTENSIONS
                if not (is_font or relative.suffix.lower() == ".txt" or
                        relative.name in {"METADATA.pb", "upstream_info.md"}):
                    continue
                if not member.isfile():
                    raise ValueError("Font archive contains a non-regular selected file")
                target = destination.joinpath(*relative.parts)
                target.parent.mkdir(parents=True, exist_ok=True)
                with package.extractfile(member) as source_file, target.open("xb") as output:
                    shutil.copyfileobj(source_file, output)
                target.chmod(0o644)
                if is_font:
                    count += 1
                    families.add(tuple(relative.parts[:2]))
        if count != source["fontFiles"] or len(families) != source["families"]:
            raise ValueError(f"Unexpected font inventory: {count} files, {len(families)} families")
    except BaseException:
        shutil.rmtree(destination)
        raise
    return {"fontFiles": count, "families": len(families)}


def parse_fontconfig(text):
    records = []
    for line in text.splitlines():
        fields = line.split("\t")
        if len(fields) != 4 or not fields[0] or not fields[3]:
            raise ValueError("Invalid fontconfig inventory record")
        family, style, fullname, filename = fields
        records.append({"families": family.split(","), "styles": style.split(",") if style else [],
                        "fullnames": fullname.split(",") if fullname else [], "file": filename})
    if not records:
        raise ValueError("Fontconfig returned no installed fonts")
    return records


def normalize(value):
    return " ".join(value.split()).casefold()


def coverage(reference, records):
    available = set()
    for record in records:
        available.update(normalize(name) for name in record["fullnames"])
        # Fontconfig emits an additional continuous-variable record with no
        # style/fullname, as well as the named instances. Do not invent a face.
        if not record["styles"]:
            continue
        # Secondary family/style names are not a Cartesian product. E.g.
        # family=Example,Example Light style=Light,Regular is NOT Example Regular.
        family, style = record["families"][0], record["styles"][0]
        available.add(normalize(family + " " + style))
        if normalize(style) in {"regular", "normal", "roman", "book"}:
            for alias in record["families"]:
                available.add(normalize(alias))
                available.add(normalize(alias + " Regular"))
    present, missing = [], []
    for face in reference:
        (present if normalize(face) in available else missing).append(face)
    return {"total": len(reference), "presentCount": len(present), "missingCount": len(missing),
            "installedFiles": len({r["file"] for r in records}),
            "matchPolicy": "Installed full names or primary family/style names; no fallback aliases",
            "present": present, "missing": missing}


def check_requirements(report, records, required, strict):
    families = {normalize(f) for record in records for f in record["families"]}
    missing = [family for family in required if normalize(family) not in families]
    if missing:
        raise ValueError("Required font families are missing: " + ", ".join(missing))
    if strict and report["missingCount"]:
        raise ValueError(f'{report["missingCount"]} reference font faces are missing; see carbone-coverage.json')


def inventory(output, required, strict):
    lock = json.loads((HERE / "sources.lock.json").read_text())
    reference = HERE / "carbone-cloud-font-list.txt"
    if digest(reference) != lock["reference"]["sha256"]:
        raise ValueError("Carbone reference list SHA-256 mismatch")
    names = [line.strip() for line in reference.read_text().splitlines() if line.strip()]
    if len(names) != lock["reference"]["entries"]:
        raise ValueError("Carbone reference list count mismatch")
    result = subprocess.run(["fc-list", "--format", "%{family}\t%{style}\t%{fullname}\t%{file}\n"],
                            check=True, capture_output=True, text=True)
    records = sorted(parse_fontconfig(result.stdout), key=lambda r: (r["file"], r["fullnames"]))
    report = coverage(names, records)
    report["reference"] = lock["reference"]
    report["googleFonts"] = lock["googleFonts"]
    output.mkdir(parents=True, exist_ok=True)
    for name, value in [("installed-fonts.json", records), ("carbone-coverage.json", report),
                        ("sources.lock.json", lock)]:
        (output / name).write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n")
    print(f'Carbone reference: {report["presentCount"]}/{report["total"]} face names present; '
          f'{report["missingCount"]} missing; {report["installedFiles"]} installed font files')
    check_requirements(report, records, required, strict)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    install = commands.add_parser("install")
    install.add_argument("--archive", type=Path)
    install.add_argument("--cache", type=Path, default=Path("/var/cache/imprint-fonts"))
    install.add_argument("--destination", type=Path, required=True)
    audit = commands.add_parser("inventory")
    audit.add_argument("--output", type=Path, default=Path("/usr/local/share/imprint/fonts"))
    audit.add_argument("--require", default="Roboto,Lobster,Open Sans,Noto Sans",
                       help="Comma-separated family names; checks installed names, not substitutes")
    audit.add_argument("--strict", action="store_true", help="Fail if any Carbone reference face is absent")
    args = parser.parse_args()
    if args.command == "install":
        source = json.loads((HERE / "sources.lock.json").read_text())["googleFonts"]
        archive = args.archive or download(source, args.cache)
        print(json.dumps(install_archive(archive, args.destination, source)))
        # Preserve shared terms too: a few upstream families omit a local license
        # file. The original binaries retain their copyright/license metadata.
        shutil.copytree(HERE / "licenses", args.destination / "LICENSES")
    else:
        inventory(args.output, [x.strip() for x in args.require.split(",") if x.strip()], args.strict)


if __name__ == "__main__":
    try:
        main()
    except (ValueError, OSError, subprocess.CalledProcessError) as error:
        raise SystemExit(str(error)) from error
