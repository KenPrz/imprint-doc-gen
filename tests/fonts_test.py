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

    def test_reference_names_map_to_variable_font_axes(self):
        families = {"fraunces": {}, "georama": {}, "noto sans thai": {}, "texturina": {}}
        self.assertEqual(fonts.parse_face("Fraunces Soft Wonky ExtraBold Italic", families),
                         ("fraunces", True, {"SOFT": 50, "WONK": 1, "wght": 800}))
        self.assertEqual(fonts.parse_face("Georama Ultra Condensed Semi Bold", families),
                         ("georama", False, {"wdth": 50, "wght": 600}))
        self.assertEqual(fonts.parse_face("Noto Sans Thai Condensed", families), ("noto sans thai", False, {"wdth": 75}))
        self.assertEqual(fonts.parse_face("Texturina 28pt Light", families), ("texturina", False, {"opsz": 28, "wght": 300}))
        self.assertIsNone(fonts.parse_face("Fraunces Caption Bold", families))
        self.assertIsNone(fonts.parse_face("Aptos Bold", families))

    def test_unnamed_weight_and_width_mean_regular_not_the_font_default(self):
        georama = {"wght": (100, 100, 900), "wdth": (62.5, 100, 150), "ital": (0, 0, 1)}
        self.assertEqual(fonts.axis_location(georama, {"wdth": 75, "ital": 1}), {"wght": 400, "wdth": 75, "ital": 1})
        fraunces = {"wght": (100, 900, 900), "SOFT": (0, 0, 100), "opsz": (9, 144, 144)}
        self.assertEqual(fonts.axis_location(fraunces, {"SOFT": 50}), {"wght": 400, "SOFT": 50, "opsz": 144})
        self.assertIsNone(fonts.axis_location(fraunces, {"wdth": 75}))
        self.assertIsNone(fonts.axis_location({"wght": (300, 300, 700)}, {"wght": 900}))

    def test_generated_names_keep_office_style_linking(self):
        names = fonts.face_names("Georama Condensed Black Italic", italic=True, bold=False)
        self.assertEqual((names[16], names[17], names[4]), ("Georama Condensed", "Black Italic", "Georama Condensed Black Italic"))
        self.assertEqual((names[1], names[2], names[6]), ("Georama Condensed Black", "Italic", "GeoramaCondensed-BlackItalic"))
        bold = fonts.face_names("Fraunces Soft Bold", italic=False, bold=True)
        self.assertEqual((bold[1], bold[2]), ("Fraunces Soft", "Bold"))

    def test_reserved_font_names_are_never_modified(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            (folder / "OFL.txt").write_text('Copyright 2010 The Lato Authors, with Reserved Font Name "Lato".\n'
                                            '"Reserved Font Name" refers to any names specified as such')
            self.assertTrue(fonts.reserves_name(folder))
            (folder / "OFL.txt").write_text('"Reserved Font Name" refers to any names specified as such')
            self.assertFalse(fonts.reserves_name(folder))


if __name__ == "__main__":
    unittest.main()
