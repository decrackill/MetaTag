"""Integración: reproduce el matching ORIGINAL sobre el dataset real
(MetaTag_v8.9/Finales 1 a 103 + el .xlsx de la raíz) y verifica que
_find_image_ex NO cambia ninguna correspondencia ni reusa archivos.
"""
import os
import re
import sys
import unittest
from collections import Counter
from pathlib import Path

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from metatag_v8 import MetaTagApp, IMG_EXTS

BASE = Path(__file__).resolve().parents[1]
FOLDER = BASE / "Finales 1 a 103"
XLSX = next(
    f for f in BASE.iterdir()
    if f.suffix.lower() == ".xlsx" and not f.name.startswith("~$"))

MISSING_ESPERADAS = {"0053_EC_C7_XII_372_R.jpg", "0055_EC_C7_VI_146_P.jpg"}
HUERFANAS_ESPERADAS = ["0053_EC_RS_372_F.jpg.JPG", "0059_EC_RS_109_P.jpg.JPG"]
DISTRIBUCION_ESPERADA = {"nombre-exacto": 18, "stem-exacto": 184,
                         "normalize": 21, "id-suffix": 44}


class TestApp(MetaTagApp):
    def __init__(self):
        pass


class Dataset269TestCase(unittest.TestCase):
    def setUp(self):
        self.app = TestApp()
        self.app._img_cache = {}
        df = pd.read_excel(XLSX)
        ids = df["ID Imagen"].astype(str).str.strip()
        self.ids = ids[(ids != "") & (ids != "nan")].drop_duplicates()

    def ref_find(self, name):
        """Réplica EXACTA del algoritmo ORIGINAL de _find_image (antes del
        cambio de detección de ambigüedades). Devuelve (path, método)."""
        name = name.strip()
        if not name:
            return None, None
        if not self.app._img_cache:
            self.app._img_cache = {
                self.app._full_stem(f.name).lower(): f
                for f in FOLDER.rglob("*")
                if f.is_file() and f.suffix.lower() in IMG_EXTS}
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

    def test_correspondencias_identicas_al_original(self):
        for idv in self.ids:
            with self.subTest(id=idv):
                ref_path, _ = self.ref_find(idv)
                new_path, status, _ = self.app._find_image_ex(idv, str(FOLDER))
                if status == "ok":
                    self.assertIsNotNone(ref_path, idv)
                    self.assertEqual(new_path, ref_path, idv)
                else:
                    self.assertIsNone(ref_path, idv)
                    self.assertEqual(status, "not_found", idv)

    def test_distribucion_de_metodos_preservada_y_sin_ambiguas(self):
        dist = Counter()
        for idv in self.ids:
            ref_path, method = self.ref_find(idv)
            if not ref_path:
                continue  # 2 missing conocidas, validadas en test_missing...
            dist[method] += 1
            new_path, status, _ = self.app._find_image_ex(idv, str(FOLDER))
            self.assertEqual(status, "ok", idv)
            self.assertEqual(new_path, ref_path, idv)
        self.assertEqual(dict(dist), DISTRIBUCION_ESPERADA)

    def test_missing_son_exactamente_las_conocidas(self):
        missing = [idv for idv in self.ids
                   if self.ref_find(idv)[0] is None]
        self.assertEqual(set(missing), MISSING_ESPERADAS)

    def test_no_reutiliza_archivos_y_huerfanas_esperadas(self):
        used = set()
        for idv in self.ids:
            p, status, _ = self.app._find_image_ex(idv, str(FOLDER))
            if status == "ok":
                self.assertNotIn(p, used, f"archivo reutilizado: {p}")
                used.add(p)
        allf = [f for f in FOLDER.iterdir()
                if f.is_file() and f.suffix.lower() in IMG_EXTS]
        orphans = sorted(f.name for f in allf if str(f) not in used)
        self.assertEqual(orphans, HUERFANAS_ESPERADAS)


if __name__ == "__main__":
    unittest.main()
