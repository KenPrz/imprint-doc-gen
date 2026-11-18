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


def reference_names(lock):
    reference = HERE / "carbone-cloud-font-list.txt"
    if digest(reference) != lock["reference"]["sha256"]:
        raise ValueError("Carbone reference list SHA-256 mismatch")
    names = [line.strip() for line in reference.read_text().splitlines() if line.strip()]
    if len(names) != lock["reference"]["entries"]:
        raise ValueError("Carbone reference list count mismatch")
    return names


def installed_records():
    result = subprocess.run(["fc-list", "--format", "%{family}\t%{style}\t%{fullname}\t%{file}\n"],
                            check=True, capture_output=True, text=True)
    return sorted(parse_fontconfig(result.stdout), key=lambda r: (r["file"], r["fullnames"]))


# Static faces from installed variable fonts. Reference names are parsed as
# "<variable family> <style words>"; each word maps to an axis value.
WEIGHTS = {"hairline": 100, "thin": 100, "extralight": 200, "ultralight": 200, "light": 300, "regular": 400,
           "normal": 400, "book": 400, "medium": 500, "semibold": 600, "demibold": 600, "bold": 700,
           "extrabold": 800, "ultrabold": 800, "black": 900, "heavy": 900, "extrablack": 1000}
WIDTHS = {"ultracondensed": 50, "extracondensed": 62.5, "condensed": 75, "semicondensed": 87.5,
          "semiexpanded": 112.5, "expanded": 125, "extraexpanded": 150, "ultraexpanded": 200}
AXIS_WORDS = {"soft": {"SOFT": 50}, "supersoft": {"SOFT": 100}, "wonky": {"WONK": 1},
              "casual": {"CASL": 1}, "monospace": {"MONO": 1}, "linear": {"CASL": 0}}
SPACED = {"extra light": "extralight", "ultra light": "ultralight", "semi bold": "semibold",
          "demi bold": "demibold", "extra bold": "extrabold", "ultra bold": "ultrabold",
          "extra black": "extrablack", "ultra condensed": "ultracondensed", "extra condensed": "extracondensed",
          "semi condensed": "semicondensed", "semi expanded": "semiexpanded", "extra expanded": "extraexpanded",
          "ultra expanded": "ultraexpanded"}


def style_key(text):
    text = normalize(text)
    for spaced, joined in SPACED.items():
        text = text.replace(spaced, joined)
    return text


def parse_face(name, families):
    """Returns (family key, italic, axis coordinates) for a name built on a known family, else None."""
    words = style_key(name).split()
    for cut in range(len(words), 0, -1):
        family = " ".join(words[:cut])
        if family not in families:
            continue
        coords, italic = {}, False
        for word in words[cut:]:
            if word in WEIGHTS: coords["wght"] = WEIGHTS[word]
            elif word in WIDTHS: coords["wdth"] = WIDTHS[word]
            elif word in AXIS_WORDS: coords.update(AXIS_WORDS[word])
            elif word == "italic": italic = True
            elif word.endswith("pt") and word[:-2].isdigit(): coords["opsz"] = int(word[:-2])
            else: return None
        return family, italic, coords
    return None


def reserves_name(folder):
    """OFL fonts with a Reserved Font Name must not be modified under that name."""
    return any("with reserved font name" in path.read_text(errors="ignore").casefold()
               for path in folder.glob("*.txt"))


def face_names(name, italic, bold):
    words = name.split()
    weight = next((w for w in words if style_key(w) in WEIGHTS and style_key(w) != "regular"), None)
    family = " ".join(w for w in words if style_key(w) not in WEIGHTS and style_key(w) != "italic")
    style = " ".join(filter(None, [weight, "Italic" if italic else None])) or "Regular"
    legacy_family = family if weight in (None, "Bold") else f"{family} {weight}"
    legacy_style = " ".join(filter(None, ["Bold" if bold else None, "Italic" if italic else None])) or "Regular"
    postscript = "".join(c for c in family if c.isalnum() or c == "-") + "-" + "".join(c for c in style if c.isalnum())
    return {1: legacy_family, 2: legacy_style, 3: postscript, 4: name, 6: postscript, 16: family, 17: style}


def axis_location(axes, coords):
    """Full axis location, or None when out of range. No weight/width word means Regular/Normal, not the font default."""
    location = {tag: default for tag, (_, default, _) in axes.items()}
    for tag, value in (("wght", 400), ("wdth", 100)):
        if tag in axes and axes[tag][0] <= value <= axes[tag][2]:
            location[tag] = value
    for tag, value in coords.items():
        if tag not in axes or not axes[tag][0] <= value <= axes[tag][2]:
            return None
        location[tag] = value
    return location


def instantiate(job):
    from fontTools.ttLib import TTFont
    from fontTools.varLib.instancer import instantiateVariableFont
    path, faces, output = job
    source, made = TTFont(path), []
    axes = {axis.axisTag: (axis.minValue, axis.defaultValue, axis.maxValue) for axis in source["fvar"].axes}
    for name, italic, coords in faces:
        if italic and not source["OS/2"].fsSelection & 1:
            if "ital" in axes: coords = {**coords, "ital": 1}
            elif "slnt" in axes: coords = {**coords, "slnt": axes["slnt"][0]}
            else: continue
        location = axis_location(axes, coords)
        if location is None:
            continue
        font = instantiateVariableFont(source, location, inplace=False, updateFontNames=False)
        bold = location.get("wght") == 700
        names = face_names(name, italic, bold)
        table = font["name"]
        for name_id in (1, 2, 3, 4, 6, 16, 17, 21, 22, 25):
            table.removeNames(nameID=name_id)
        for name_id, value in names.items():
            table.setName(value, name_id, 3, 1, 0x409)
        # fsSelection bits: ITALIC=0, BOLD=5, REGULAR=6; usWeightClass/usWidthClass come from the instancer.
        selection = font["OS/2"].fsSelection & ~0b1100001
        font["OS/2"].fsSelection = selection | (1 if italic else 0) | (32 if bold else 0) | (0 if italic or bold else 64)
        font["head"].macStyle = (1 if bold else 0) | (2 if italic else 0)
        target = output / path.parent.name / (names[6] + path.suffix)
        target.parent.mkdir(parents=True, exist_ok=True)
        font.save(target)
        for license_file in path.parent.glob("*.txt"):
            shutil.copy(license_file, target.parent / license_file.name)
        made.append(name)
    return made


def generate(source, output):
    from concurrent.futures import ProcessPoolExecutor
    from fontTools.ttLib import TTFont
    missing = coverage(reference_names(json.loads((HERE / "sources.lock.json").read_text())), installed_records())["missing"]
    families = {}
    for path in sorted(source.rglob("*")):
        if path.suffix.lower() not in {".ttf", ".otf"} or reserves_name(path.parent):
            continue
        font = TTFont(path, lazy=True)
        if "fvar" in font:
            italic = bool(font["OS/2"].fsSelection & 1)
            family = font["name"].getDebugName(16) or font["name"].getDebugName(1)
            families.setdefault(style_key(family), {})[italic] = path
    jobs = {}
    for name in missing:
        parsed = parse_face(name, families)
        if parsed:
            family, italic, coords = parsed
            path = families[family].get(italic) or families[family].get(False)
            jobs.setdefault(path, []).append((name, italic, coords))
    output.mkdir(parents=True, exist_ok=False)
    # ponytail: one process per source font. The largest Noto family sets the floor (~3 min);
    # ~5 min on 4 cores. Split a font's faces across processes if builds need to be faster.
    with ProcessPoolExecutor() as pool:
        made = [name for names in pool.map(instantiate, [(p, f, output) for p, f in jobs.items()]) for name in names]
    print(f"Generated {len(made)} static faces from {len(jobs)} variable fonts; {len(missing) - len(made)} reference faces still missing")


def inventory(output, required, strict):
    lock = json.loads((HERE / "sources.lock.json").read_text())
    names = reference_names(lock)
    records = installed_records()
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
    instances = commands.add_parser("generate", help="Write static faces for missing reference names")
    instances.add_argument("--source", type=Path, required=True, help="Directory of installed variable fonts")
    instances.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.command == "generate":
        generate(args.source, args.output)
    elif args.command == "install":
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
