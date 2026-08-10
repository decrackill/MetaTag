import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from metatag_v8 import MetaTagApp


class TestApp(MetaTagApp):
    def __init__(self):
        pass


def build_folder(files):
    tmp = tempfile.mkdtemp(prefix="metatag_test_")
    for name in files:
        with open(os.path.join(tmp, name), "wb") as f:
            f.write(b"fake")
    return tmp


class FindImageExTestCase(unittest.TestCase):
    def setUp(self):
        self.app = TestApp()
        self.app._img_cache = {}

    def test_ambiguous_cuando_dos_archivos_comparten_clave_id_suffix(self):
        # Escenario 0053: la clave (numero=0053, sufijo=F) es compartida por
        # dos archivos. Una variante del Excel que NO matchee por stem cae en
        # id-suffix y no debe elegirse uno arbitrariamente.
        folder = build_folder([
            "0053_EC_RS_372_F.jpg.JPG",
            "0053_EC_C7_XII_372_F.jpg.JPG",
        ])
        path, status, candidates = self.app._find_image_ex(
            "0053_OTRO_CAMPO_1_F.jpg", folder)
        self.assertIsNone(path)
        self.assertEqual(status, "ambiguous")
        self.assertEqual(len(candidates), 2)

    def test_id_suffix_unico_devuelve_ok(self):
        folder = build_folder([
            "0053_EC_RS_372_F.jpg.JPG",
            "0053_EC_C7_XII_372_R.jpg.JPG",
        ])
        # (0053, R) solo existe un archivo con esa clave → sin ambigüedad.
        path, status, candidates = self.app._find_image_ex(
            "0053_EC_C7_XII_372_R.jpg", folder)
        self.assertEqual(status, "ok")
        self.assertEqual(os.path.basename(path), "0053_EC_C7_XII_372_R.jpg.JPG")

    def test_normalize_resuelve_ceros_adicionales(self):
        folder = build_folder(["0001_UM_C4_UE18_0006_F.JPG"])
        path, status, _ = self.app._find_image_ex(
            "0001_UM_C4_UE18_00006_F.jpg", folder)
        self.assertEqual(status, "ok")
        self.assertEqual(os.path.basename(path), "0001_UM_C4_UE18_0006_F.JPG")

    def test_substring_ambiguo_con_dos_candidatos(self):
        folder = build_folder(["X_IMG_A.jpg", "X_IMG_B.jpg"])
        # 'x_img' está contenido en ambos stems; sin clave id-suffix.
        path, status, candidates = self.app._find_image_ex("X_IMG.jpg", folder)
        self.assertIsNone(path)
        self.assertEqual(status, "ambiguous")
        self.assertEqual(len(candidates), 2)

    def test_no_encontrada(self):
        folder = build_folder(["0001_UM_C4_UE18_0006_F.JPG"])
        path, status, candidates = self.app._find_image_ex(
            "9999_NOPE_X.jpg", folder)
        self.assertIsNone(path)
        self.assertEqual(status, "not_found")
        self.assertEqual(candidates, [])

    def test_find_image_delega_y_devuelve_path(self):
        folder = build_folder(["0053_EC_RS_372_F.jpg.JPG"])
        path = self.app._find_image("0053_EC_C7_XII_372_F.jpg", folder)
        self.assertEqual(os.path.basename(path), "0053_EC_RS_372_F.jpg.JPG")


class MatchRowsToFilesTestCase(unittest.TestCase):
    def setUp(self):
        self.app = TestApp()
        self.app._img_cache = {}

    def test_fila_ambigua_se_reporta_y_no_cae_en_no_encontradas(self):
        import pandas as pd

        folder = build_folder([
            "0053_EC_RS_372_F.jpg.JPG",
            "0053_EC_C7_XII_372_F.jpg.JPG",
        ])
        df = pd.DataFrame({"img": ["0053_OTRO_CAMPO_1_F.jpg"]})
        all_files = [os.path.join(folder, f)
                     for f in os.listdir(folder)]
        ordered, aprox, missing, used, ambiguas = self.app._match_rows_to_files(
            df, "img", all_files, {}, {}, folder)
        self.assertEqual(ordered, [])
        self.assertEqual(missing, [])
        self.assertEqual(len(ambiguas), 1)
        self.assertEqual(len(ambiguas[0][1]), 2)


if __name__ == "__main__":
    unittest.main()
