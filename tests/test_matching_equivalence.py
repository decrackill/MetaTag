"""Equivalencia del matching con precompilación de regex (FASE 3B.1b).

La precompilación NO puede cambiar el algoritmo: misma jerarquía de pasos,
mismas prioridades, mismos resultados y mismo orden de candidatos. Esta
suite verifica que el código OPTIMIZADO reproduce EXACTAMENTE el resultado
del código ORIGINAL capturado en ``fixtures/matching_baseline.json``.

El fixture se generó ejecutando ``ImageMatcher`` y ``_find_image_ex`` con el
algoritmo original (regex inline) sobre el mismo corpus, ANTES del cambio.

Corpus cubierto:
  - dataset real (Finales 1 a 103 + Excel), si está presente;
  - casos sintéticos: doble extensión, marcador "(1)", cero-padding,
    bordes "#/_-", mayúsculas, sufijos F/R/P, candidatos múltiples,
    ambigüedad, inexistentes, cadena vacía y metacharacteres.
"""

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import metatag_matching as mm
from metatag_matching import ImageMatcher

BASE = Path(__file__).resolve().parents[1]
FIXTURE = Path(__file__).resolve().parent / "fixtures" / "matching_baseline.json"

# ── corpus sintético ────────────────────────────────────────────────────────
# Cada entrada: (archivos de la carpeta, [consultas]).
SYNTHETIC_CORPUS = [
    (["0053_EC_RS_372_F.jpg.JPG", "0053_EC_C7_XII_372_F.jpg.JPG"],
     ["0053_OTRO_CAMPO_1_F.jpg", "0053_OTRO_CAMPO_1_F.JPG"]),
    (["0053_EC_RS_372_F.jpg.JPG", "0053_EC_C7_XII_372_R.jpg.JPG"],
     ["0053_EC_C7_XII_372_R.jpg", "0053_EC_C7_XII_372_R"]),
    (["0001_UM_C4_UE18_0006_F.JPG"],
     ["0001_UM_C4_UE18_00006_F.jpg", "0001_UM_C4_UE18_00006_F.JPG",
      "9999_NOPE_X.jpg"]),
    (["X_IMG_A.jpg", "X_IMG_B.jpg"],
     ["X_IMG.jpg", "Y_IMG.jpg"]),
    (["A.JPG"], ["a.jpg", "a.JPG", "A"]),
    (["A_B_C.JPG"], ["a_b_c", "A_B_C", "a__b__c"]),
    (["__A_B_C.JPG"], ["#a_b_c#", "a_b_c"]),
    (["IMG_A_B_C.jpg"], ["img_a_b.jpg", "IMG_A_B_C.jpg"]),
    (["a.jpg (1).JPG"], ["a", "a.jpg", "a.jpg (1)"]),
    (["0001_UM_C4_IX_00034_P.jpg (1).JPG"],
     ["0001_UM_C4_IX_00034_P.jpg", "0001_UM_C4_IX_00034_P"]),
    (["0061_EC_C4_III_046.jpg.JPG"], ["0061_EC_C4_III_046", "0061_EC_C4_III_046.jpg"]),
    (["79_EC_PS_VI_250_R.jpg"], ["79_EC_PS_VI_250_R", "79_EC_PS_VI_250_R.jpg"]),
    (["0001_A_+_B_._C.jpg"], ["0001_A_+_B_._C", "0001_A_+_B_._C.jpg", "zzz_zzz"]),
    (["a.b.jpg"], ["a.b.jpg", "a.b", "a_b"]),
    ([], ["   ", "", "nada"]),
    (["IMG_0001_F.jpg"], ["img_0001_f.JPG", "IMG_0001_F.jpg", "IMG_0001_F"]),
    (["0001_UM_C4_UE18_0006_F.JPG", "0001_UM_C4_UE18_0007_F.JPG"],
     ["0001_UM_C4_UE18_00006_F", "0001_UM_C4_UE18_00007_R"]),
    (["C4_0001_F.jpg", "C4_0001_R.jpg"], ["C4_0001_F", "C4_0001_R", "C4_0001"]),
]


def normalize_result(res):
    """Convierte (path, status, candidates) a forma comparable por nombre."""
    path, status, candidates = res
    return {
        "status": status,
        "path": os.path.basename(path) if path else None,
        "candidates": sorted(os.path.basename(c) for c in candidates),
    }


def build_folder(files):
    tmp = tempfile.mkdtemp(prefix="metatag_baseline_")
    for name in files:
        with open(os.path.join(tmp, name), "wb") as f:
            f.write(b"fake")
    return tmp


def run_synthetic(matcher):
    """Ejecuta el corpus sintético y devuelve lista de resultados ordenada."""
    results = []
    for files, queries in SYNTHETIC_CORPUS:
        folder = build_folder(files)
        try:
            for q in queries:
                res = matcher.find_image_ex(q, folder)
                results.append([q, normalize_result(res)])
        finally:
            for name in files:
                try:
                    os.remove(os.path.join(folder, name))
                except OSError:
                    pass
            try:
                os.rmdir(folder)
            except OSError:
                pass
    results.sort(key=lambda item: item[0])
    return results


def real_corpus():
    """Devuelve [(id, folder)] con los IDs del dataset real si está presente."""
    folder = BASE / "Finales 1 a 103"
    try:
        import pandas as pd
        xlsx = next(f for f in BASE.iterdir()
                    if f.suffix.lower() == ".xlsx" and not f.name.startswith("~$"))
    except Exception:
        return None, None
    if not folder.is_dir():
        return None, None
    df = pd.read_excel(xlsx)
    ids = df["ID Imagen"].astype(str).str.strip()
    ids = ids[(ids != "") & (ids != "nan")].drop_duplicates()
    return list(ids), str(folder)


def run_real(matcher, ids, folder):
    results = []
    for q in ids:
        results.append([q, normalize_result(matcher.find_image_ex(q, folder))])
    results.sort(key=lambda item: item[0])
    return results


def load_fixture():
    with open(FIXTURE, encoding="utf-8") as f:
        return json.load(f)


@unittest.skipUnless(FIXTURE.exists(), "fixture de baseline no generado")
class MatchingEquivalenceTestCase(unittest.TestCase):
    def test_optimizado_sintetico_igual_a_original(self):
        matcher = ImageMatcher()
        baseline = load_fixture()["synthetic"]
        results = run_synthetic(matcher)
        self.assertEqual(results, baseline)

    @unittest.skipUnless(
        (BASE / "Finales 1 a 103").is_dir(), "dataset real no presente")
    def test_optimizado_dataset_real_igual_a_original(self):
        matcher = ImageMatcher()
        baseline = load_fixture()["real"]
        ids, folder = real_corpus()
        if ids is None:
            self.skipTest("dataset real no presente")
        results = run_real(matcher, ids, folder)
        self.assertEqual(results, baseline)


if __name__ == "__main__":
    unittest.main()
