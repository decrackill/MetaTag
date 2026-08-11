"""Tests del motor de emparejamiento seguro puro (src/metatag_matching.py).

Verifican que el port reproduce EXACTAMENTE el algoritmo validado de
MetaTagApp._find_image_ex (sin Tkinter, sin PIL, sin umbrales difusos):
  - ambigüedad → nunca se elige un candidato arbitrario,
  - not_found → nunca se elige nada,
  - determinismo → mismos resultados y mismo orden de candidatos,
  - jerarquía de pasos con métodos etiquetables (para auditoría),
  - equivalencia con la línea base del dataset real (269 / 267 / 2 / 0,
    distribución 18 / 184 / 21 / 44).

Ejecutar:  .venv/bin/python -m unittest tests.test_metatag_matching -v
"""
import os
import re
import sys
import tempfile
import unittest
from collections import Counter
from pathlib import Path

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import metatag_matching as mm
from metatag_matching import ImageMatcher

SRC_FILE = Path(mm.__file__)


def build_folder(files):
    tmp = tempfile.mkdtemp(prefix="metatag_matching_")
    for name in files:
        with open(os.path.join(tmp, name), "wb") as f:
            f.write(b"fake")
    return tmp


class PurityTestCase(unittest.TestCase):
    def test_no_depende_de_tkinter_ni_de_pil_ni_de_difflib(self):
        import ast
        tree = ast.parse(SRC_FILE.read_text(encoding="utf-8"))
        imported = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                imported.update(alias.name.split(".")[0] for alias in node.names)
            elif isinstance(node, ast.ImportFrom):
                if node.module:
                    imported.add(node.module.split(".")[0])
        self.assertNotIn("tkinter", imported)
        self.assertNotIn("PIL", imported)
        self.assertNotIn("hashlib", imported)
        self.assertNotIn("difflib", imported)
        self.assertNotIn("unicodedata", imported)

    def test_importa_sin_tkinter_disponible(self):
        # El módulo debe poder importarse en un entorno sin Tkinter.
        import importlib
        saved = sys.modules.get("tkinter")
        try:
            importlib.reload(mm)
        finally:
            if saved is not None:
                sys.modules["tkinter"] = saved


class ImageMatcherTestCase(unittest.TestCase):
    def setUp(self):
        self.matcher = ImageMatcher()

    def test_ambiguo_cuando_dos_archivos_comparten_clave_id_suffix(self):
        # Escenario 0053: (numero=0053, sufijo=F) compartida por dos archivos.
        folder = build_folder([
            "0053_EC_RS_372_F.jpg.JPG",
            "0053_EC_C7_XII_372_F.jpg.JPG",
        ])
        path, status, candidates = self.matcher.find_image_ex(
            "0053_OTRO_CAMPO_1_F.jpg", folder)
        self.assertIsNone(path)
        self.assertEqual(status, "ambiguous")
        self.assertEqual(len(candidates), 2)
        self.assertEqual({Path(c).name for c in candidates},
                         {"0053_EC_RS_372_F.jpg.JPG",
                          "0053_EC_C7_XII_372_F.jpg.JPG"})

    def test_id_suffix_unico_devuelve_ok(self):
        folder = build_folder([
            "0053_EC_RS_372_F.jpg.JPG",
            "0053_EC_C7_XII_372_R.jpg.JPG",
        ])
        path, status, _ = self.matcher.find_image_ex(
            "0053_EC_C7_XII_372_R.jpg", folder)
        self.assertEqual(status, "ok")
        self.assertEqual(os.path.basename(path), "0053_EC_C7_XII_372_R.jpg.JPG")

    def test_normalize_resuelve_ceros_adicionales(self):
        folder = build_folder(["0001_UM_C4_UE18_0006_F.JPG"])
        path, status, _ = self.matcher.find_image_ex(
            "0001_UM_C4_UE18_00006_F.jpg", folder)
        self.assertEqual(status, "ok")
        self.assertEqual(os.path.basename(path), "0001_UM_C4_UE18_0006_F.JPG")

    def test_substring_ambiguo_con_dos_candidatos(self):
        folder = build_folder(["X_IMG_A.jpg", "X_IMG_B.jpg"])
        path, status, candidates = self.matcher.find_image_ex("X_IMG.jpg", folder)
        self.assertIsNone(path)
        self.assertEqual(status, "ambiguous")
        self.assertEqual(len(candidates), 2)

    def test_no_encontrada_ningun_paso_cuadra(self):
        folder = build_folder(["0001_UM_C4_UE18_0006_F.JPG"])
        path, status, candidates = self.matcher.find_image_ex(
            "9999_NOPE_X.jpg", folder)
        self.assertIsNone(path)
        self.assertEqual(status, "not_found")
        self.assertEqual(candidates, [])

    def test_nombre_vacio_nunca_empareja(self):
        path, status, _ = self.matcher.find_image_ex("   ", "/tmp")
        self.assertIsNone(path)
        self.assertEqual(status, "not_found")

    def test_find_image_delega_y_devuelve_path(self):
        folder = build_folder(["0053_EC_RS_372_F.jpg.JPG"])
        path = self.matcher.find_image("0053_EC_C7_XII_372_F.jpg", folder)
        self.assertEqual(os.path.basename(path), "0053_EC_RS_372_F.jpg.JPG")

    def test_resultados_deterministas_en_calls_repetidos(self):
        folder = build_folder([
            "0053_EC_RS_372_F.jpg.JPG",
            "0053_EC_C7_XII_372_F.jpg.JPG",
        ])
        r1 = self.matcher.find_image_ex("0053_OTRO_CAMPO_1_F.jpg", folder)
        r2 = self.matcher.find_image_ex("0053_OTRO_CAMPO_1_F.jpg", folder)
        self.assertEqual(r1, r2)
        self.assertEqual(r1[2], r2[2])  # mismo orden de candidatos

    def test_jerarquia_de_pasos_se_etiqueta(self):
        # Cada paso de la jerarquía debe poder atribuirse (para auditoría).
        direct = build_folder(["a.jpg"])
        self.assertEqual(self.matcher.find_image_ex_with_method("a.jpg", direct)[3], "direct")

        nombre = build_folder(["A.JPG"])
        self.assertEqual(
            self.matcher.find_image_ex_with_method("a.jpg", nombre)[3], "nombre-exacto")

        stem = build_folder(["A_B_C.JPG"])
        self.assertEqual(
            self.matcher.find_image_ex_with_method("a_b_c", stem)[3], "stem-exacto")

        clean = build_folder(["__A_B_C.JPG"])
        self.assertEqual(
            self.matcher.find_image_ex_with_method("#a_b_c#", clean)[3], "clean")

        norm = build_folder(["0001_UM_C4_UE18_0006_F.JPG"])
        self.assertEqual(
            self.matcher.find_image_ex_with_method(
                "0001_UM_C4_UE18_00006_F.jpg", norm)[3], "normalize")

        idsuf = build_folder(["0053_EC_RS_372_F.jpg.JPG"])
        self.assertEqual(
            self.matcher.find_image_ex_with_method(
                "0053_OTRO_CAMPO_1_F.jpg", idsuf)[3], "id-suffix")

        sub = build_folder(["IMG_A_B_C.jpg"])
        self.assertEqual(
            self.matcher.find_image_ex_with_method("img_a_b.jpg", sub)[3], "substring")

    def test_with_method_no_cambia_decision(self):
        folder = build_folder(["0001_UM_C4_UE18_0006_F.JPG"])
        plain = self.matcher.find_image_ex("0001_UM_C4_UE18_00006_F.jpg", folder)
        rich = self.matcher.find_image_ex_with_method(
            "0001_UM_C4_UE18_00006_F.jpg", folder)
        self.assertEqual(plain[:2], rich[:2])


# ── Equivalencia con la línea base del dataset real (269) ──────────────────

BASE = Path(__file__).resolve().parents[1]
FOLDER = BASE / "Finales 1 a 103"

MISSING_ESPERADAS = {"0053_EC_C7_XII_372_R.jpg", "0055_EC_C7_VI_146_P.jpg"}
HUERFANAS_ESPERADAS = ["0053_EC_RS_372_F.jpg.JPG", "0059_EC_RS_109_P.jpg.JPG"]
DISTRIBUCION_ESPERADA = {"nombre-exacto": 18, "stem-exacto": 184,
                         "normalize": 21, "id-suffix": 44}

try:
    import pandas as pd
    _XLSX = next(f for f in BASE.iterdir()
                 if f.suffix.lower() == ".xlsx" and not f.name.startswith("~$"))
except Exception:
    _XLSX = None

HAS_DATASET = FOLDER.is_dir() and _XLSX is not None


@unittest.skipUnless(HAS_DATASET, "dataset real (Finales 1 a 103 + xlsx) no presente")
class DatasetEquivalenceTestCase(unittest.TestCase):
    """El port debe reproducir IDÉNTICA la línea base validada de MetaTag."""

    @classmethod
    def setUpClass(cls):
        import metatag_v8
        cls.img_exts = metatag_v8.IMG_EXTS
        cls.app = metatag_v8.MetaTagApp.__new__(metatag_v8.MetaTagApp)
        cls.app._img_cache = {}
        df = pd.read_excel(_XLSX)
        ids = df["ID Imagen"].astype(str).str.strip()
        cls.ids = ids[(ids != "") & (ids != "nan")].drop_duplicates()

    def ref_find(self, name):
        """Réplica exacta del algoritmo original (como test_dataset_269.py)."""
        import metatag_v8
        name = name.strip()
        if not name:
            return None, None
        if not self.app._img_cache:
            self.app._img_cache = {
                self.app._full_stem(f.name).lower(): f
                for f in FOLDER.rglob("*")
                if f.is_file() and f.suffix.lower() in self.img_exts}
        p = FOLDER / name
        if p.exists():
            return str(p), "direct"
        name_lower = name.lower()
        name_stem = self.app._full_stem(name).lower()
        for _, fpath in self.app._img_cache.items():
            if fpath.name.lower() == name_lower:
                return str(fpath), "nombre-exacto"
        if name_stem in self.app._img_cache:
            return str(self.app._img_cache[name_stem]), "stem-exacto"
        name_clean = re.sub(r"^[#\s\-_]+|[#\s\-_]+$", "", name_stem)
        for sk, fpath in self.app._img_cache.items():
            if re.sub(r"^[#\s\-_]+|[#\s\-_]+$", "", sk) == name_clean:
                return str(fpath), "clean"
        name_norm = self.app._normalize_numbers(name_clean)
        for sk, fpath in self.app._img_cache.items():
            sk_clean = re.sub(r"^[#\s\-_]+|[#\s\-_]+$", "", sk)
            if self.app._normalize_numbers(sk_clean) == name_norm:
                return str(fpath), "normalize"
        ide = self.app._extract_id_suffix(name)
        if ide:
            for sk, fpath in self.app._img_cache.items():
                if self.app._extract_id_suffix(sk) == ide:
                    return str(fpath), "id-suffix"
        for sk, fpath in self.app._img_cache.items():
            if name_stem in sk or sk in name_stem:
                return str(fpath), "substring"
        return None, None

    def test_correspondencias_identicas_y_metodos_preservados(self):
        matcher = ImageMatcher()
        dist = Counter()
        missing = []
        for idv in self.ids:
            ref_path, method = self.ref_find(idv)
            path, status, _ = matcher.find_image_ex(idv, str(FOLDER))
            if status == "ok":
                self.assertIsNotNone(ref_path, idv)
                self.assertEqual(path, ref_path, idv)
                dist[method] += 1
            elif status == "ambiguous":
                self.fail(f"ambigüedad inesperada: {idv}")
            else:
                self.assertEqual(status, "not_found", idv)
                self.assertIsNone(ref_path, idv)
                missing.append(idv)
        self.assertEqual(set(missing), MISSING_ESPERADAS)
        self.assertEqual(dict(dist), DISTRIBUCION_ESPERADA)

    def test_no_reutiliza_archivos_y_huerfanas_esperadas(self):
        matcher = ImageMatcher()
        used = set()
        for idv in self.ids:
            p, status, _ = matcher.find_image_ex(idv, str(FOLDER))
            if status == "ok":
                self.assertNotIn(p, used, f"archivo reutilizado: {p}")
                used.add(p)
        allf = [f for f in FOLDER.iterdir()
                if f.is_file() and f.suffix.lower() in self.img_exts]
        orphans = sorted(f.name for f in allf if str(f) not in used)
        self.assertEqual(orphans, HUERFANAS_ESPERADAS)


if __name__ == "__main__":
    unittest.main()
