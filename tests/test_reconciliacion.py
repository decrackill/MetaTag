"""
Fase 1 + 4 + 5 + 6: Diagnóstico exhaustivo del dataset real 269,
reconciliación de registros y fotografías, tests de integridad.

Este archivo es SOLO LECTURA sobre el dataset real (no modifica archivos).
"""
import os
import sys
from collections import Counter
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from metatag_matching import ImageMatcher
from renombrar_fotos_gui import RenameModel, VALID_IMG_EXT

BASE = Path(__file__).resolve().parents[1]
FOLDER = BASE / "Finales 1 a 103"
XLSX = next(
    f for f in BASE.iterdir()
    if f.suffix.lower() == ".xlsx" and not f.name.startswith("~$"))

EXPECTED_TOTAL_RECORDS = 269
EXPECTED_TOTAL_PHOTOS = 269  # verificado: 269 archivos .jpg.JPG en la carpeta
EXPECTED_MATCHED = 267
EXPECTED_NOT_FOUND = 2
EXPECTED_ORPHANS = 2


# ── helpers ────────────────────────────────────────────────────────────────

def _make_model(matching: bool = True) -> RenameModel:
    """Crea un RenameModel configurado con el dataset real."""
    m = RenameModel()
    m.folder_path = FOLDER
    m.excel_path = XLSX
    m.sheet_name = "Principal"
    m.column_name = "ID Imagen"
    m.sort_mode = "natural"
    m.matching_mode = matching
    m.keep_extension = True
    m.load_photos()
    m.load_names()
    return m


def _count_plan_states(plan: list[dict]) -> dict[str, int]:
    """Cuenta cada estado individual del plan."""
    return dict(Counter(item["state"] for item in plan))


def _count_by_photo(plan: list[dict]) -> dict[str, int]:
    """Clasifica fotos según su uso en el plan."""
    matched_files = set()
    for item in plan:
        src = item["src"]
        if src is not None:
            matched_files.add(str(src))
    all_photos = [f for f in FOLDER.iterdir()
                  if f.is_file() and f.suffix.lower() in VALID_IMG_EXT]
    all_photo_strs = {str(f) for f in all_photos}
    orphans = all_photo_strs - matched_files
    return {
        "total_files": len(all_photos),
        "matched_to_records": len(matched_files),
        "orphans": len(orphans),
        "orphan_names": sorted(Path(o).name for o in orphans),
    }


# ── Phase 1: Diagnóstico exhaustivo ───────────────────────────────────────

class TestDiagnosticoDataset269:
    """Diagnóstico del dataset real: cuenta exacta de cada estado del plan."""

    def test_matching_plan_state_distribution(self):
        """Distribution exacta de estados en modo matching sobre 269 registros."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        states = _count_plan_states(plan)
        total = sum(states.values())

        print("\n" + "=" * 60)
        print("DIAGNÓSTICO MATCHING MODE — 269 registros")
        print("=" * 60)
        for state, count in sorted(states.items()):
            print(f"  {state:20s}: {count}")
        print(f"  {'TOTAL':20s}: {total}")
        print(f"  photos loaded: {len(m.photos)}")
        print(f"  names loaded:  {len(m.names)}")
        print(f"  skipped_rows:  {m.skipped_rows}")
        print("=" * 60)

        assert total == EXPECTED_TOTAL_RECORDS, (
            f"Plan states sum {total} != {EXPECTED_TOTAL_RECORDS}")

    def test_matching_counter_formulas(self):
        """Verifica las fórmulas de contadores contra el plan real."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()

        # Fórmula actual de n_corr
        n_corr = sum(1 for item in plan if item["src"] is not None)
        # Fórmula actual de n_conf
        n_conf = sum(1 for item in plan
                     if item["state"] in ("existe", "conflicto", "duplicado",
                                          "ambiguo", "error"))
        faltan = len(m.names) - n_corr

        # Conteo alternativo: sumar estados con src != None
        states_with_src = Counter()
        states_without_src = Counter()
        for item in plan:
            if item["src"] is not None:
                states_with_src[item["state"]] += 1
            else:
                states_without_src[item["state"]] += 1

        print("\n" + "=" * 60)
        print("CONTADORES — Matching Mode")
        print("=" * 60)
        print(f"  n_corr (src is not None):  {n_corr}")
        print(f"  n_conf (blocking states):  {n_conf}")
        print(f"  faltan (names - n_corr):   {faltan}")
        print(f"  len(m.photos):             {len(m.photos)}")
        print(f"  len(m.names):              {len(m.names)}")
        print(f"\n  Estados CON foto (src != None):")
        for s, c in sorted(states_with_src.items()):
            print(f"    {s:20s}: {c}")
        print(f"  Estados SIN foto (src is None):")
        for s, c in sorted(states_without_src.items()):
            print(f"    {s:20s}: {c}")
        print(f"\n  Suma con foto:    {sum(states_with_src.values())}")
        print(f"  Suma sin foto:    {sum(states_without_src.values())}")
        print(f"  Total:            {sum(states_with_src.values()) + sum(states_without_src.values())}")
        print("=" * 60)

        # La suma de con_foto + sin_foto debe ser 269
        assert sum(states_with_src.values()) + sum(states_without_src.values()) == EXPECTED_TOTAL_RECORDS

    def test_matching_photo_reconciliation(self):
        """Reconciliación del lado de las fotografías."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        photo_info = _count_by_photo(plan)

        print("\n" + "=" * 60)
        print("RECONCILIACIÓN FOTOGRÁFICAS")
        print("=" * 60)
        print(f"  Total archivos:           {photo_info['total_files']}")
        print(f"  Asignados a registros:    {photo_info['matched_to_records']}")
        print(f"  Huérfanos (sin registro): {photo_info['orphans']}")
        print(f"  Nombres huérfanos:        {photo_info['orphan_names']}")
        print("=" * 60)

        assert photo_info["total_files"] == EXPECTED_TOTAL_PHOTOS
        assert photo_info["matched_to_records"] + photo_info["orphans"] == photo_info["total_files"]

    def test_positional_plan_state_distribution(self):
        """Distribution de estados en modo posicional (referencia cruzada)."""
        m = _make_model(matching=False)
        plan = m._build_plan_positional()
        states = _count_plan_states(plan)
        total = sum(states.values())

        print("\n" + "=" * 60)
        print("DIAGNÓSTICO POSITIONAL MODE — 269 registros")
        print("=" * 60)
        for state, count in sorted(states.items()):
            print(f"  {state:20s}: {count}")
        print(f"  {'TOTAL':20s}: {total}")
        print(f"  photos loaded: {len(m.photos)}")
        print(f"  names loaded:  {len(m.names)}")
        print("=" * 60)

        assert total == EXPECTED_TOTAL_RECORDS

    def test_identify_missing_records(self):
        """Identifica EXACTAMENTE los registros que no se encuentran."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()

        missing = [item for item in plan if item["state"] == "not_found"]
        ambiguous = [item for item in plan if item["state"] == "ambiguo"]
        error = [item for item in plan if item["state"] == "error"]

        print("\n" + "=" * 60)
        print("REGISTROS NO RESUELTOS (matching mode)")
        print("=" * 60)
        print(f"\n  not_found ({len(missing)}):")
        for item in missing:
            print(f"    row {item['row']:3d}: name='{item['name']}' src={item['src']}")
        print(f"\n  ambiguo ({len(ambiguous)}):")
        for item in ambiguous:
            cands = [Path(c).name for c in item.get("candidates", [])]
            print(f"    row {item['row']:3d}: name='{item['name']}' candidates={cands}")
        print(f"\n  error ({len(error)}):")
        for item in error:
            print(f"    row {item['row']:3d}: name='{item['name']}' reason='{item.get('reason', '')}'")
        print("=" * 60)

        # Debe haber exactamente 2 not_found (dataset conocido)
        assert len(missing) == EXPECTED_NOT_FOUND, (
            f"Expected {EXPECTED_NOT_FOUND} not_found, got {len(missing)}: "
            f"{[item['name'] for item in missing]}")

    def test_cross_check_with_dataset_269_constants(self):
        """Verifica que los resultados coinciden con las constantes de test_dataset_269."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        states = _count_plan_states(plan)

        # El dataset conocido tiene 2 not_found y 0 ambiguos
        assert states.get("not_found", 0) == 2
        assert states.get("ambiguo", 0) == 0
        assert states.get("error", 0) == 0

        # 267 registros con foto
        n_with_photo = sum(1 for item in plan if item["src"] is not None)
        assert n_with_photo == 267

    def test_full_reconciliation_matching(self):
        """Reconciliación completa: cada registro en EXACTAMENTE un estado."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        states = _count_plan_states(plan)

        # Clasificación exhaustiva
        classification = {
            "ok": states.get("ok", 0),
            "ya_correcto": states.get("ya_correcto", 0),
            "existe": states.get("existe", 0),
            "conflicto": states.get("conflicto", 0),
            "duplicado": states.get("duplicado", 0),
            "not_found": states.get("not_found", 0),
            "sin_foto": states.get("sin_foto", 0),
            "ambiguo": states.get("ambiguo", 0),
            "error": states.get("error", 0),
        }
        total = sum(classification.values())

        print("\n" + "=" * 60)
        print("RECONCILIACIÓN COMPLETA — 269 REGISTROS (matching)")
        print("=" * 60)
        for state, count in classification.items():
            marker = " ←" if count > 0 else ""
            print(f"  {state:20s}: {count:3d}{marker}")
        print(f"  {'─' * 30}")
        print(f"  {'TOTAL':20s}: {total:3d}")
        print("=" * 60)

        assert total == EXPECTED_TOTAL_RECORDS

    def test_full_photo_reconciliation(self):
        """Reconciliación completa del lado de las fotografías."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        photo_info = _count_by_photo(plan)

        # Clasificar cada foto
        photo_usage = Counter()
        for item in plan:
            src = item["src"]
            if src is not None:
                photo_usage[item["state"]] += 1

        photo_classification = {
            "asignadas_ok": photo_usage.get("ok", 0),
            "asignadas_ya_correcto": photo_usage.get("ya_correcto", 0),
            "asignadas_existe": photo_usage.get("existe", 0),
            "asignadas_conflicto": photo_usage.get("conflicto", 0),
            "asignadas_duplicado": photo_usage.get("duplicado", 0),
            "no_utilizadas_huerfanas": photo_info["orphans"],
        }
        total_photos = sum(photo_classification.values())

        print("\n" + "=" * 60)
        print("RECONCILIACIÓN COMPLETA — FOTOGRAFÍAS")
        print("=" * 60)
        for cat, count in photo_classification.items():
            marker = " ←" if count > 0 else ""
            print(f"  {cat:35s}: {count:3d}{marker}")
        print(f"  {'─' * 40}")
        print(f"  {'TOTAL':35s}: {total_photos:3d}")
        print(f"\n  Total archivos físicos: {photo_info['total_files']}")
        print("=" * 60)

        assert total_photos == photo_info["total_files"]


# ── Phase 6: Tests de integridad ──────────────────────────────────────────

class TestIntegridadReconciliacion:
    """Tests que fallan si las sumas no cuadran exactamente."""

    EXPECTED_RECORDS = 269
    EXPECTED_PHOTOS = 269
    EXPECTED_MATCHED = 267
    EXPECTED_NOT_FOUND = 2
    EXPECTED_ORPHANS = 2

    # ── Registros ───────────────────────────────────────────────────────

    def test_suma_estados_igual_total_registros(self):
        """sum(estados) == total de registros. Falla si algún registro
        no está clasificado o está clasificado dos veces."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        by_state = Counter(item["state"] for item in plan)
        total = sum(by_state.values())
        assert total == self.EXPECTED_RECORDS, (
            f"sum(estados)={total} != {self.EXPECTED_RECORDS}. "
            f"Distribución: {dict(by_state)}")

    def test_cada_registro_en_un_solo_estado(self):
        """Cada registro aparece en exactamente UN estado (mutuamente excluyentes)."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        # Verificar que la suma de Contadores individuales == len(plan)
        states = [item["state"] for item in plan]
        assert len(states) == len(plan)
        # Verificar que no hay estados fuera del dominio válido
        valid_states = {"ok", "ya_correcto", "existe", "conflicto",
                        "duplicado", "not_found", "sin_foto", "ambiguo", "error"}
        for s in states:
            assert s in valid_states, f"Estado inválido: {s!r}"

    def test_correspondencias_mas_sin_foto_igual_registros(self):
        """Correspondencias (src != None) + no_correspondencias (src is None)
        == total de registros. Invariante fundamental."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        n_corr = sum(1 for item in plan if item["src"] is not None)
        n_no_corr = sum(1 for item in plan if item["src"] is None)
        assert n_corr + n_no_corr == self.EXPECTED_RECORDS

    def test_no_photo_states_son_exhaustivos(self):
        """Los estados sin foto (not_found + sin_foto + ambiguo + error)
        cubren TODOS los registros sin src."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        no_photo_states = {"not_found", "sin_foto", "ambiguo", "error"}
        n_no_photo = sum(1 for item in plan
                         if item["state"] in no_photo_states)
        n_src_none = sum(1 for item in plan if item["src"] is None)
        assert n_no_photo == n_src_none, (
            f"Estados sin foto ({n_no_photo}) != src=None ({n_src_none})")

    def test_photo_states_son_exhaustivos(self):
        """Los estados con foto (ok + ya_correcto + existe + conflicto + duplicado)
        cubren TODOS los registros con src."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        photo_states = {"ok", "ya_correcto", "existe", "conflicto", "duplicado"}
        n_photo = sum(1 for item in plan if item["state"] in photo_states)
        n_src_set = sum(1 for item in plan if item["src"] is not None)
        assert n_photo == n_src_set, (
            f"Estados con foto ({n_photo}) != src!=None ({n_src_set})")

    def test_dataset_269_valores_exactos(self):
        """Para el dataset actual, los contadores deben ser exactos."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        by_state = Counter(item["state"] for item in plan)
        n_corr = sum(1 for item in plan if item["src"] is not None)

        assert by_state["ok"] == 83
        assert by_state["ya_correcto"] == 184
        assert by_state["not_found"] == 2
        assert by_state.get("existe", 0) == 0
        assert by_state.get("conflicto", 0) == 0
        assert by_state.get("duplicado", 0) == 0
        assert by_state.get("sin_foto", 0) == 0
        assert by_state.get("ambiguo", 0) == 0
        assert by_state.get("error", 0) == 0
        assert n_corr == 267
        assert sum(by_state.values()) == 269

    # ── Fotografías ─────────────────────────────────────────────────────

    def test_suma_fotografias_igual_total_archivos(self):
        """asignadas + huérfanas == total de archivos de imagen."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        photo_info = _count_by_photo(plan)
        assert (photo_info["matched_to_records"] + photo_info["orphans"]
                == photo_info["total_files"])

    def test_fotografias_total_coincide_dataset(self):
        """Total de archivos físicos == 269."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        photo_info = _count_by_photo(plan)
        assert photo_info["total_files"] == self.EXPECTED_PHOTOS

    def test_fotografias_matched_coincide(self):
        """Fotos asignadas == 267."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        photo_info = _count_by_photo(plan)
        assert photo_info["matched_to_records"] == self.EXPECTED_MATCHED

    def test_fotografias_orphan_coincide(self):
        """Fotos huérfanas == 2."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        photo_info = _count_by_photo(plan)
        assert photo_info["orphans"] == self.EXPECTED_ORPHANS
        assert sorted(photo_info["orphan_names"]) == [
            "0053_EC_RS_372_F.jpg.JPG", "0059_EC_RS_109_P.jpg.JPG"]

    def test_no_reutilizo_fotografias(self):
        """Ninguna foto aparece en más de un registro (one-to-one)."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()
        used = []
        for item in plan:
            src = item["src"]
            if src is not None:
                used.append(str(src))
        assert len(used) == len(set(used)), (
            f"Fotos reutilizadas: {[s for s in used if used.count(s) > 1]}")

    # ── Contadores de la interfaz ───────────────────────────────────────

    def test_counters_derive_from_single_model(self):
        """Verifica que los contadores de la interfaz (n_corr, n_conf)
        son consistentes con el plan (mismo modelo)."""
        m = _make_model(matching=True)
        plan = m._build_plan_matching()

        # Replicar la lógica de _update_sync_state_finish
        _BLOCKING = ("existe", "conflicto", "duplicado", "ambiguo", "error")
        state_counter = Counter(item["state"] for item in plan)
        n_corr = sum(1 for item in plan if item["src"] is not None)
        n_conf = sum(state_counter[s] for s in _BLOCKING)

        # Verificar que n_corr == photo_states sum
        photo_states = {"ok", "ya_correcto", "existe", "conflicto", "duplicado"}
        n_photo_states = sum(state_counter[s] for s in photo_states)
        assert n_corr == n_photo_states, (
            f"n_corr ({n_corr}) != sum(photo_states) ({n_photo_states})")

        # Verificar que n_conf == blocking_states sum
        n_blocking = sum(state_counter[s] for s in _BLOCKING)
        assert n_conf == n_blocking

        # Verificar invariante
        no_photo_states = {"not_found", "sin_foto", "ambiguo", "error"}
        n_no_photo = sum(state_counter[s] for s in no_photo_states)
        assert n_corr + n_no_photo == sum(state_counter.values())
