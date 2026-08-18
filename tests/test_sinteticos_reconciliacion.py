"""
Fase 7: Fixtures sintéticos que demuestran matemáticamente
el comportamiento del motor de plan sobre los 9 estados.

Cada caso crea archivos ficticios en un directorio temporal y verifica
que el plan produce exactamente los estados esperados.
"""
import os
import shutil
import sys
import tempfile
from collections import Counter
from pathlib import Path

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from renombrar_fotos_gui import RenameModel


@pytest.fixture
def tmp_dir():
    d = tempfile.mkdtemp()
    yield Path(d)
    shutil.rmtree(d, ignore_errors=True)


def _make_files(folder: Path, names: list[str]) -> None:
    """Crea archivos vacíos con contenido distinguishable."""
    for i, name in enumerate(names):
        (folder / name).write_bytes(f"content_{name}_{i}".encode())


def _plan_states(model: RenameModel) -> dict[str, int]:
    """Construye el plan y devuelve el conteo de estados."""
    plan = model._build_plan()
    return dict(Counter(item["state"] for item in plan))


# ── Caso 1: 3 registros → 3 fotos, nombre exacto ─────────────────────────

class TestCaso1PerfectMatch:
    """3 registros, 3 fotos: A→A, B→B, C→C. Todo ya_correcto."""

    def test_all_ya_correcto_matching(self, tmp_dir):
        _make_files(tmp_dir, ["A.jpg", "B.jpg", "C.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = True
        m.keep_extension = True
        m.load_photos()
        m._names = ["A", "B", "C"]

        plan = m._build_plan_matching()
        states = _plan_states(m)

        assert states == {"ya_correcto": 3}
        assert len(plan) == 3
        assert sum(1 for item in plan if item["src"] is not None) == 3
        assert sum(1 for item in plan
                   if item["state"] in ("existe", "conflicto", "duplicado",
                                        "ambiguo", "error")) == 0

    def test_all_ya_correcto_positional(self, tmp_dir):
        _make_files(tmp_dir, ["A.jpg", "B.jpg", "C.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["A", "B", "C"]

        plan = m._build_plan_positional()
        states = dict(Counter(item["state"] for item in plan))
        assert states == {"ya_correcto": 3}


# ── Caso 2: Swap resoluble A→B, B→A (posicional) ─────────────────────────

class TestCaso2Swap:
    """2 registros, 2 fotos en posicional: position 0的照片→name[0], position 1→name[1].
    Colisión interna resoluble con renombrado en dos fases.
    NOTA: en matching mode no hay swap porque cada nombre busca SU foto."""

    def test_swap_resuelto_posicional(self, tmp_dir):
        """Posicional: photos=[A.jpg, B.jpg], names=["B", "A"].
        A.jpg→B.jpg (colisión), B.jpg→A.jpg (colisión). Ciclo → resoluble."""
        _make_files(tmp_dir, ["A.jpg", "B.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "A"]

        plan = m._build_plan_positional()
        states = dict(Counter(item["state"] for item in plan))

        # Ciclo puro A→B, B→A: resuelto con rename en 2 fases → "ok"
        assert states.get("ok", 0) == 2
        assert states.get("conflicto", 0) == 0

    def test_swap_count_invariants(self, tmp_dir):
        _make_files(tmp_dir, ["A.jpg", "B.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "A"]

        plan = m._build_plan_positional()
        n_corr = sum(1 for item in plan if item["src"] is not None)
        assert n_corr == 2
        assert len(plan) == 2

    def test_matching_mode_no_swap(self, tmp_dir):
        """En matching mode, A y B buscan sus fotos → ambos ya_correcto (sin swap)."""
        _make_files(tmp_dir, ["A.jpg", "B.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = True
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "A"]

        plan = m._build_plan_matching()
        states = _plan_states(m)
        # Cada nombre encuentra SU foto → ya_correcto
        assert states == {"ya_correcto": 2}


# ── Caso 3: Ciclo resoluble A→B, B→C, C→A (posicional) ──────────────────

class TestCaso3Ciclo:
    """3 registros, 3 fotos en posicional: A.jpg→B, B.jpg→C, C.jpg→A.
    Ciclo puro, resoluble con renombrado en dos fases."""

    def test_ciclo_resuelto_posicional(self, tmp_dir):
        _make_files(tmp_dir, ["A.jpg", "B.jpg", "C.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "C", "A"]

        plan = m._build_plan_positional()
        states = dict(Counter(item["state"] for item in plan))

        # Ciclo puro: todos "ok" tras _enable_batch_swaps
        assert states.get("ok", 0) == 3
        assert states.get("conflicto", 0) == 0


# ── Caso 4: Cadena resoluble A→B, B→C, C→D, D→E (posicional) ────────────

class TestCaso4Cadena:
    """4 registros, 5 fotos en posicional: A.jpg→B, B.jpg→C, C.jpg→D, D.jpg→E.
    Cadena que termina en archivo libre (E.jpg no es destino de nadie)."""

    def test_cadena_resuelta_posicional(self, tmp_dir):
        """Cadena que termina en un archivo EXTERNO (E.jpg no es destino
        de nadie en el plan): D→E es 'existe', el resto queda 'conflicto'
        porque la cadena no es completamente resoluble."""
        _make_files(tmp_dir, ["A.jpg", "B.jpg", "C.jpg", "D.jpg", "E.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "C", "D", "E"]

        plan = m._build_plan_positional()
        states = dict(Counter(item["state"] for item in plan))

        # Cadena A→B→C→D→E: E es externo → D→E es "existe", A→B, B→C, C→D
        # quedan "conflicto" porque la cadena no se resuelve sin tocar E.
        assert states.get("existe", 0) == 1
        assert states.get("conflicto", 0) == 3

    def test_cadena_libre_resuelta(self, tmp_dir):
        """Cadena que termina en archivo LIBRE (no existe): A→B, B→C, C→D.
        D.jpg NO existe → la cadena se resuelve con rename en 2 fases."""
        _make_files(tmp_dir, ["A.jpg", "B.jpg", "C.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "C", "D"]

        plan = m._build_plan_positional()
        states = dict(Counter(item["state"] for item in plan))

        # D.jpg no existe → cadena libre → todos "ok"
        assert states.get("ok", 0) == 3
        assert states.get("conflicto", 0) == 0


# ── Caso 5: Destino externo existe ────────────────────────────────────────

class TestCaso5Existe:
    """Existe: la foto encontrada tiene un destino que ya existe como
    archivo externo (no pertenece al lote)."""

    def test_existe_posicional(self, tmp_dir):
        """Posicional: photos=[A.jpg, X.jpg], names=["X"].
        A.jpg→X.jpg: X.jpg existe en disco y NO es src de ningún plan item → existe."""
        _make_files(tmp_dir, ["A.jpg", "X.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["X"]

        plan = m._build_plan_positional()
        states = dict(Counter(item["state"] for item in plan))

        # X.jpg exists on disk, is not any plan item's src → "existe"
        assert states.get("existe", 0) == 1
        assert states.get("ok", 0) == 0

    def test_existe_counted_as_conflict(self, tmp_dir):
        _make_files(tmp_dir, ["A.jpg", "X.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["X"]

        plan = m._build_plan_positional()
        n_conf = sum(1 for item in plan
                     if item["state"] in ("existe", "conflicto", "duplicado",
                                          "ambiguo", "error"))
        assert n_conf == 1

    def test_existe_blocks_rename(self, tmp_dir):
        _make_files(tmp_dir, ["A.jpg", "X.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["X"]

        plan = m._build_plan_positional()
        blocked, _ = m.rename_blocked(plan)
        assert blocked is True


# ── Caso 6: Dos registros → misma fotografía (reuso) ─────────────────────

class TestCaso6Reuso:
    """1 foto A.jpg, 2 registros con el mismo nombre "A".
    Primera fila: ya_correcto. Segunda: duplicado (reuso)."""

    def test_reuso_duplicado(self, tmp_dir):
        _make_files(tmp_dir, ["A.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = True
        m.keep_extension = True
        m.load_photos()
        m._names = ["A", "A"]

        plan = m._build_plan_matching()
        states = _plan_states(m)

        assert states.get("ya_correcto", 0) == 1
        assert states.get("duplicado", 0) == 1
        assert sum(states.values()) == 2

    def test_reuso_both_have_src(self, tmp_dir):
        """Ambos registros tienen src != None (reuso cuenta como correspondencia)."""
        _make_files(tmp_dir, ["A.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = True
        m.keep_extension = True
        m.load_photos()
        m._names = ["A", "A"]

        plan = m._build_plan_matching()
        n_corr = sum(1 for item in plan if item["src"] is not None)
        assert n_corr == 2  # ambos tienen foto asociada


# ── Caso 7: Registro sin fotografía ───────────────────────────────────────

class TestCaso7SinFoto:
    """Sin fotos: registros no encontrados."""

    def test_not_found_matching(self, tmp_dir):
        _make_files(tmp_dir, [])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = True
        m.keep_extension = True
        m.load_photos()
        m._names = ["A"]

        plan = m._build_plan_matching()
        states = _plan_states(m)

        assert states.get("not_found", 0) == 1
        assert sum(1 for item in plan if item["src"] is None) == 1

    def test_sin_foto_positional(self, tmp_dir):
        """En posicional: hay más registros que fotos → sin_foto."""
        _make_files(tmp_dir, [])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["A", "B"]

        plan = m._build_plan_positional()
        states = dict(Counter(item["state"] for item in plan))

        assert states.get("sin_foto", 0) == 2

    def test_positional_partial_sin_foto(self, tmp_dir):
        """Posicional: 2 fotos, 3 registros → 1 sin_foto."""
        _make_files(tmp_dir, ["A.jpg", "B.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["X", "Y", "Z"]

        plan = m._build_plan_positional()
        states = dict(Counter(item["state"] for item in plan))

        # 2 fotos → position 0 y 1 tienen foto, position 2 no
        assert states.get("sin_foto", 0) == 1
        assert states.get("ok", 0) + states.get("ya_correcto", 0) + \
               states.get("conflicto", 0) + states.get("duplicado", 0) == 2


# ── Caso 8: Fotografía sin registro ───────────────────────────────────────

class TestCaso8FotoSinRegistro:
    """Foto sin registro: la foto queda huérfana (no aparece en el plan)."""

    def test_orphan_photo(self, tmp_dir):
        _make_files(tmp_dir, ["A.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = True
        m.keep_extension = True
        m.load_photos()
        m._names = []

        plan = m._build_plan_matching()
        assert len(plan) == 0
        assert len(m.photos) == 1

    def test_orphan_not_counted_in_plan(self, tmp_dir):
        _make_files(tmp_dir, ["A.jpg", "B.jpg"])
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = True
        m.keep_extension = True
        m.load_photos()
        m._names = ["A"]

        plan = m._build_plan_matching()
        states = _plan_states(m)

        assert states.get("ya_correcto", 0) == 1
        assert len(plan) == 1
        assert sum(1 for item in plan if item["src"] is not None) == 1


# ── Tests de integridad cruzada ───────────────────────────────────────────

class TestInvariantesSinteticos:
    """Invariante: sum(estados) == len(plan) == len(names) para cada fixture."""

    @pytest.mark.parametrize("names,photos,expected_total,mode", [
        (["A", "B", "C"], ["A.jpg", "B.jpg", "C.jpg"], 3, True),
        (["B", "A"], ["A.jpg", "B.jpg"], 2, False),       # positional swap
        (["B", "C", "A"], ["A.jpg", "B.jpg", "C.jpg"], 3, False),  # positional cycle
        (["B", "C", "D", "E"], ["A.jpg", "B.jpg", "C.jpg", "D.jpg", "E.jpg"], 4, False),
        (["X"], ["A.jpg", "X.jpg"], 1, False),             # existe
        (["A", "A"], ["A.jpg"], 2, True),                   # reuso
        (["A"], [], 1, True),                               # not_found
        ([], ["A.jpg"], 0, True),                           # sin registros
    ])
    def test_suma_estados(self, tmp_dir, names, photos, expected_total, mode):
        _make_files(tmp_dir, photos)
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = mode
        m.keep_extension = True
        m.load_photos()
        m._names = names

        plan = m._build_plan()
        by_state = Counter(item["state"] for item in plan)
        total = sum(by_state.values())
        assert total == expected_total
        assert len(plan) == expected_total

    @pytest.mark.parametrize("names,photos,n_corr_expected,mode", [
        (["A", "B", "C"], ["A.jpg", "B.jpg", "C.jpg"], 3, True),
        (["B", "A"], ["A.jpg", "B.jpg"], 2, False),
        (["B", "C", "A"], ["A.jpg", "B.jpg", "C.jpg"], 3, False),
        (["B", "C", "D", "E"], ["A.jpg", "B.jpg", "C.jpg", "D.jpg", "E.jpg"], 4, False),
        (["X"], ["A.jpg", "X.jpg"], 1, False),  # existe → src is not None
        (["A", "A"], ["A.jpg"], 2, True),        # reuso → ambos con src
        (["A"], [], 0, True),                    # not_found
        ([], ["A.jpg"], 0, True),                # sin registros
    ])
    def test_correspondencias_consistentes(self, tmp_dir, names, photos, n_corr_expected, mode):
        _make_files(tmp_dir, photos)
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = mode
        m.keep_extension = True
        m.load_photos()
        m._names = names

        plan = m._build_plan()
        n_corr = sum(1 for item in plan if item["src"] is not None)
        assert n_corr == n_corr_expected

    @pytest.mark.parametrize("names,photos,n_conf_expected,mode", [
        (["A", "B", "C"], ["A.jpg", "B.jpg", "C.jpg"], 0, True),
        (["B", "A"], ["A.jpg", "B.jpg"], 0, False),
        (["X"], ["A.jpg", "X.jpg"], 1, False),     # existe → n_conf=1
        (["A", "A"], ["A.jpg"], 1, True),           # duplicado → n_conf=1
        (["A"], [], 0, True),                        # not_found → no es conflicto
    ])
    def test_conflictos_consistentes(self, tmp_dir, names, photos, n_conf_expected, mode):
        _make_files(tmp_dir, photos)
        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = mode
        m.keep_extension = True
        m.load_photos()
        m._names = names

        plan = m._build_plan()
        _BLOCKING = ("existe", "conflicto", "duplicado", "ambiguo", "error")
        n_conf = sum(1 for item in plan if item["state"] in _BLOCKING)
        assert n_conf == n_conf_expected
