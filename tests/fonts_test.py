"""Build-time font checks; no network or installed font dependencies."""
import hashlib
import importlib.util
import io
from pathlib import Path
import tarfile
import tempfile
import unittest

MODULE = Path(__file__).resolve().parents[1] / "docker/fonts/manage.py"
spec = importlib.util.spec_from_file_location("imprint_fonts", MODULE)
fonts = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fonts)
smoke_spec = importlib.util.spec_from_file_location("font_smoke", MODULE.parents[2] / "scripts/font-smoke.py")
smoke = importlib.util.module_from_spec(smoke_spec)
smoke_spec.loader.exec_module(smoke)


class FontTests(unittest.TestCase):
    def test_pdf_smoke_does_not_accept_a_similarly_named_substitute(self):
        self.assertTrue(smoke.font_matches("Open Sans", "OpenSans-BoldItalic"))
        self.assertFalse(smoke.font_matches("Open Sans", "OpenSansCondensed-Regular"))
        self.assertFalse(smoke.font_matches("Ubuntu", "UbuntuMono-Regular"))
        self.assertFalse(smoke.font_matches("Aptos", "NotoSerif-Regular"))

    def archive(self, folder, entries):
        path = folder / "fonts.tar.gz"
        with tarfile.open(path, "w:gz") as archive:
            for name, content in entries.items():
                info = tarfile.TarInfo("fonts-test/" + name)
                info.size = len(content)
                archive.addfile(info, io.BytesIO(content))
        source = {"revision": "test", "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                  "fontFiles": 1, "families": 1}
        return path, source

    def test_installs_fonts_and_licenses_without_repository_extras(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            archive, source = self.archive(folder, {
                "ofl/example/static/Example.ttf": b"font bytes",
                "ofl/example/OFL.txt": b"license notice",
                "ofl/example/METADATA.pb": b'name: "Example"',
                "ofl/example/article/image.png": b"unused image",
                "axisregistry/test.ttf": b"test font, not a distributed family",
            })
            fonts.install_archive(archive, folder / "installed", source)
            self.assertEqual((folder / "installed/ofl/example/OFL.txt").read_bytes(), b"license notice")
            self.assertTrue((folder / "installed/ofl/example/static/Example.ttf").is_file())
            self.assertFalse((folder / "installed/axisregistry").exists())
            self.assertFalse((folder / "installed/ofl/example/article").exists())

    def test_corrupt_archive_fails_before_creating_destination(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            archive, source = self.archive(folder, {"ofl/example/Example.ttf": b"font"})
            source["sha256"] = "0" * 64
            with self.assertRaisesRegex(ValueError, "SHA-256"):
                fonts.install_archive(archive, folder / "installed", source)
            self.assertFalse((folder / "installed").exists())

    def test_traversal_and_missing_fonts_fail_and_remove_partial_output(self):
        for entries in [{"ofl/example/../../../escaped.ttf": b"escape"},
                        {"ofl/example/OFL.txt": b"only a license"}]:
            with self.subTest(entries=entries), tempfile.TemporaryDirectory() as temp:
                folder = Path(temp)
                archive, source = self.archive(folder, entries)
                with self.assertRaises(ValueError):
                    fonts.install_archive(archive, folder / "installed", source)
                self.assertFalse((folder / "escaped.ttf").exists())
                self.assertFalse((folder / "installed").exists())

    def test_inventory_never_counts_fontconfig_substitutions(self):
        records = fonts.parse_fontconfig(
            "Liberation Sans\tRegular\tLiberation Sans\t/fonts/liberation.ttf\n"
            "Example,Example Light\tLight,Regular\tExample Light\t/fonts/example.ttf\n"
            "Example\tBold\tExample Bold\t/fonts/example.ttf\n"
            "Example\tBold Italic\tExample Bold Italic\t/fonts/example-italic.ttf\n"
        )
        report = fonts.coverage(["Arial Regular", "Example Regular", "Example Light",
                                 "Example Bold", "Example Bold Italic", "Liberation Sans Regular"], records)
        self.assertEqual(report["missing"], ["Arial Regular", "Example Regular"])
        self.assertEqual(report["present"], ["Example Light", "Example Bold", "Example Bold Italic",
                                             "Liberation Sans Regular"])
        self.assertEqual(report["total"], 6)
        self.assertEqual(report["installedFiles"], 3)

    def test_regular_fullnames_and_variable_named_instances_are_detected(self):
        records = fonts.parse_fontconfig(
            "Lobster\tRegular\tLobster Regular\t/fonts/lobster.ttf\n"
            "Roboto\tBold\tRoboto Bold\t/fonts/roboto-variable.ttf\n"
            "Roboto\tLight\tRoboto Light\t/fonts/roboto-variable.ttf\n"
            "Roboto\t\t\t/fonts/roboto-variable.ttf\n"
        )
        report = fonts.coverage(["Lobster", "Lobster Regular", "Roboto Light", "Roboto Bold",
                                 "Roboto ExtraBold"], records)
        self.assertEqual(report["missing"], ["Roboto ExtraBold"])

    def test_strict_and_required_family_checks_fail_for_missing_fonts(self):
        records = fonts.parse_fontconfig("Lobster\tRegular\tLobster Regular\t/fonts/lobster.ttf\n")
        report = fonts.coverage(["Lobster Regular", "Aptos Regular"], records)
        fonts.check_requirements(report, records, ["Lobster"], strict=False)
        with self.assertRaisesRegex(ValueError, "Aptos"):
            fonts.check_requirements(report, records, ["Aptos"], strict=False)
        with self.assertRaisesRegex(ValueError, "1 reference"):
            fonts.check_requirements(report, records, [], strict=True)


if __name__ == "__main__":
    unittest.main()
