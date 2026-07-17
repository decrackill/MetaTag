import tempfile
import unittest
from pathlib import Path

from PIL import Image, PngImagePlugin

from image_sync import RenameModel
from metatag_writer import read_existing_metadata, write_meta


class MetadataWriterTests(unittest.TestCase):
    def test_supported_formats_keep_readable_metadata(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            for suffix in (".jpg", ".png", ".tiff"):
                path = folder / f"image{suffix}"
                if suffix == ".png":
                    info = PngImagePlugin.PngInfo()
                    info.add_text("Keep", "yes")
                    Image.new("RGBA", (8, 8), (1, 2, 3, 4)).save(path, pnginfo=info)
                else:
                    Image.new("RGB", (8, 8), (1, 2, 3)).save(path)
                write_meta(str(path), {"Sitio": "A"})
                self.assertEqual(read_existing_metadata(str(path)), {"Sitio": "A"})

    def test_unsupported_format_is_rejected_without_changing_type(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "image.webp"
            Image.new("RGB", (8, 8)).save(path)
            with self.assertRaises(RuntimeError):
                write_meta(str(path), {"Sitio": "A"})
            with Image.open(path) as image:
                self.assertEqual(image.format, "WEBP")


class RenameModelTests(unittest.TestCase):
    def _model(self, folder, names):
        photos = []
        for index in range(len(names)):
            photo = Path(folder) / f"source_{index}.jpg"
            Image.new("RGB", (4, 4)).save(photo)
            photos.append(photo)
        model = RenameModel()
        model._photos = photos
        model._names = names
        return model

    def test_rename_rejects_path_traversal(self):
        with tempfile.TemporaryDirectory() as temporary:
            model = self._model(temporary, ["../outside"])
            result = []
            model.rename_all(lambda *_: None, lambda ok, errors: result.extend([ok, errors]))
            self.assertEqual(result[0], 0)
            self.assertTrue(result[1])
            self.assertTrue((Path(temporary) / "source_0.jpg").exists())

    def test_rename_rejects_existing_destination(self):
        with tempfile.TemporaryDirectory() as temporary:
            model = self._model(temporary, ["taken"])
            Image.new("RGB", (4, 4)).save(Path(temporary) / "taken.jpg")
            result = []
            model.rename_all(lambda *_: None, lambda ok, errors: result.extend([ok, errors]))
            self.assertEqual(result[0], 0)
            self.assertTrue(result[1])

    def test_rename_honors_keep_extension_option(self):
        with tempfile.TemporaryDirectory() as temporary:
            model = self._model(temporary, ["final.png"])
            result = []
            model.rename_all(lambda *_: None, lambda ok, errors: result.extend([ok, errors]), keep_ext=False)
            self.assertEqual(result[0], 1)
            self.assertTrue((Path(temporary) / "final.png").exists())


if __name__ == "__main__":
    unittest.main()
