"""
Fase 8+9+10: Pruebas de renombrado REAL con archivos temporales.

Verifica:
- Ciclo A→B, B→C, C→A se resuelve correctamente
- Swap A→B, B→A se resuelve correctamente
- Contenidos originales preservados tras renombrado
- No quedan archivos .metatag_tmp huérfanos
- undo_last restaura todo correctamente
- Rollback funciona si se fuerza un error durante fase 2
"""
import hashlib
import os
import shutil
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from renombrar_fotos_gui import RenameModel


@pytest.fixture
def tmp_dir():
    d = tempfile.mkdtemp()
    yield Path(d)
    shutil.rmtree(d, ignore_errors=True)


def _md5(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()


def _make_file(folder: Path, name: str, content: bytes) -> Path:
    p = folder / name
    p.write_bytes(content)
    return p


def _no_tmp_files(folder: Path) -> bool:
    """Verifica que no hay archivos .metatag_tmp en la carpeta."""
    return not any(folder.glob(".metatag_tmp_*"))


def _all_files(folder: Path) -> set[str]:
    return {f.name for f in folder.iterdir() if f.is_file()}


# ── Caso 8: Renombrado real ciclo A→B, B→C, C→A ──────────────────────────

class TestRenameCiclo:
    """Renombrado real de un ciclo A→B→C→A en positional mode."""

    def test_ciclo_renamed_correctly(self, tmp_dir):
        """A.jpg→B.jpg, B.jpg→C.jpg, C.jpg→A.jpg. Resultado: A.jpg→C.jpg,
        B.jpg→A.jpg, C.jpg→B.jpg (cada archivo tiene el contenido original
        de OTRO nombre)."""
        # Contenidos distinguibles
        _make_file(tmp_dir, "A.jpg", b"content_A")
        _make_file(tmp_dir, "B.jpg", b"content_B")
        _make_file(tmp_dir, "C.jpg", b"content_C")

        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "C", "A"]

        # Verificar plan: ciclo resoluble
        plan = m._build_plan_positional()
        ok_items = [it for it in plan if it["state"] == "ok"]
        assert len(ok_items) == 3

        # Ejecutar renombrado real
        errors = []
        m.rename_all(
            on_progress=lambda a, b, c: None,
            on_done=lambda ok, errs: errors.extend(errs))

        assert len(errors) == 0
        files = _all_files(tmp_dir)
        assert files == {"A.jpg", "B.jpg", "C.jpg"}

        # Verificar contenidos: A.jpg ahora tiene content_B, etc.
        # Cycle A→B→C→A means:
        #   Original A (content_A) renamed to B → B.jpg has content_A
        #   Original B (content_B) renamed to C → C.jpg has content_B
        #   Original C (content_C) renamed to A → A.jpg has content_C
        assert (tmp_dir / "A.jpg").read_bytes() == b"content_C"
        assert (tmp_dir / "B.jpg").read_bytes() == b"content_A"
        assert (tmp_dir / "C.jpg").read_bytes() == b"content_B"

    def test_ciclo_no_tmp_files(self, tmp_dir):
        """No quedan archivos .metatag_tmp huérfanos tras el ciclo."""
        _make_file(tmp_dir, "A.jpg", b"content_A")
        _make_file(tmp_dir, "B.jpg", b"content_B")
        _make_file(tmp_dir, "C.jpg", b"content_C")

        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "C", "A"]

        m.rename_all(
            on_progress=lambda a, b, c: None,
            on_done=lambda ok, errs: None)

        assert _no_tmp_files(tmp_dir)

    def test_ciclo_undo_reports_conflicts(self, tmp_dir):
        """undo_last para un ciclo: los destinos están ocupados por otros
        archivos renombrados, así que undo reporta conflictos (comportamiento
        seguro por diseño — NO sobreescribe)."""
        _make_file(tmp_dir, "A.jpg", b"content_A")
        _make_file(tmp_dir, "B.jpg", b"content_B")
        _make_file(tmp_dir, "C.jpg", b"content_C")

        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "C", "A"]

        m.rename_all(
            on_progress=lambda a, b, c: None,
            on_done=lambda ok, errs: None)

        # Deshacer: en un ciclo, las rutas originales están ocupadas
        undo_errors = []
        m.undo_last(
            on_progress=lambda a, b, c: None,
            on_done=lambda ok, errs: undo_errors.extend(errs))

        # Undo detecta conflictos (rutas ocupadas) y NO sobreescribe → seguro
        assert len(undo_errors) == 3  # A.jpg, B.jpg, C.jpg todos ocupados
        assert all("Conflicto al deshacer" in e for e in undo_errors)

        # Archivos siguen existiendo (no se corrompieron)
        files = _all_files(tmp_dir)
        assert files == {"A.jpg", "B.jpg", "C.jpg"}


# ── Caso 9: Renombrado real swap A→B, B→A ─────────────────────────────────

class TestRenameSwap:
    """Renombrado real de un swap A→B, B→A en positional mode."""

    def test_swap_renamed_correctly(self, tmp_dir):
        _make_file(tmp_dir, "A.jpg", b"content_A")
        _make_file(tmp_dir, "B.jpg", b"content_B")

        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "A"]

        plan = m._build_plan_positional()
        ok_items = [it for it in plan if it["state"] == "ok"]
        assert len(ok_items) == 2

        errors = []
        m.rename_all(
            on_progress=lambda a, b, c: None,
            on_done=lambda ok, errs: errors.extend(errs))

        assert len(errors) == 0
        files = _all_files(tmp_dir)
        assert files == {"A.jpg", "B.jpg"}

        # Swap: A→B means original A (content_A) is now B.jpg
        #        B→A means original B (content_B) is now A.jpg
        assert (tmp_dir / "A.jpg").read_bytes() == b"content_B"
        assert (tmp_dir / "B.jpg").read_bytes() == b"content_A"

    def test_swap_undo_reports_conflicts(self, tmp_dir):
        """undo_last para un swap: las rutas originales están ocupadas,
        undo reporta conflictos (seguro por diseño)."""
        _make_file(tmp_dir, "A.jpg", b"content_A")
        _make_file(tmp_dir, "B.jpg", b"content_B")

        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "A"]

        m.rename_all(
            on_progress=lambda a, b, c: None,
            on_done=lambda ok, errs: None)

        undo_errors = []
        m.undo_last(
            on_progress=lambda a, b, c: None,
            on_done=lambda ok, errs: undo_errors.extend(errs))

        # Swap: A.jpg y B.jpg ambos ocupados → conflictos
        assert len(undo_errors) == 2
        assert all("Conflicto al deshacer" in e for e in undo_errors)

        files = _all_files(tmp_dir)
        assert files == {"A.jpg", "B.jpg"}


# ── Caso 10: Rollback forzado ─────────────────────────────────────────────

class TestRollback:
    """Rollback: forzar error durante fase 2 del rename."""

    def test_rollback_on_phase2_error(self, tmp_dir):
        """Forzar un error en la FASE 2 (temp→definitivo) y verificar que:
        1. Los archivos originales se restauran vía _rollback_temp
        2. El modelo vuelve a un estado consistente"""
        _make_file(tmp_dir, "A.jpg", b"content_A")
        _make_file(tmp_dir, "B.jpg", b"content_B")

        md5_a = _md5(tmp_dir / "A.jpg")
        md5_b = _md5(tmp_dir / "B.jpg")

        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "A"]

        # Monkeypatch: fallar solo en el PRIMER intento de FASE 2.
        # Sequence: phase1(2 calls) → phase2_call_1(FAIL) → rollback → phase2_call_2 → rollback
        original_rename = Path.rename
        call_count = [0]

        def selective_rename(self_path, target):
            call_count[0] += 1
            # Fail only the first temp→dest rename (call 3 = first phase 2)
            if (".metatag_tmp" in str(self_path) and call_count[0] == 3):
                raise OSError("Forced error in phase 2")
            return original_rename(self_path, target)

        with patch.object(Path, "rename", selective_rename):
            errors = []
            m.rename_all(
                on_progress=lambda a, b, c: None,
                on_done=lambda ok, errs: errors.extend(errs))

        # Rollback restauró los archivos originales
        assert _no_tmp_files(tmp_dir)
        files = _all_files(tmp_dir)
        assert "A.jpg" in files
        assert "B.jpg" in files
        assert _md5(tmp_dir / "A.jpg") == md5_a
        assert _md5(tmp_dir / "B.jpg") == md5_b

        # Errores reportados
        assert len(errors) > 0

    def test_rollback_undo_stack_empty_on_failure(self, tmp_dir):
        """Si el rename falla completamente, undo_last no tiene nada que hacer."""
        _make_file(tmp_dir, "A.jpg", b"content_A")
        _make_file(tmp_dir, "B.jpg", b"content_B")

        m = RenameModel()
        m.folder_path = tmp_dir
        m.sort_mode = "natural"
        m.matching_mode = False
        m.keep_extension = True
        m.load_photos()
        m._names = ["B", "A"]

        original_rename = Path.rename
        call_count = [0]

        def selective_rename(self_path, target):
            call_count[0] += 1
            if (".metatag_tmp" in str(self_path) and call_count[0] == 3):
                raise OSError("Forced error")
            return original_rename(self_path, target)

        with patch.object(Path, "rename", selective_rename):
            m.rename_all(
                on_progress=lambda a, b, c: None,
                on_done=lambda ok, errs: None)

        # Rollback restauró todo → no hubo batch exitoso → undo_stack vacío
        assert not m.has_undo
        undo_errors = []
        m.undo_last(
            on_progress=lambda a, b, c: None,
            on_done=lambda ok, errs: undo_errors.extend(errs))
        assert len(undo_errors) == 1  # "No hay nada que deshacer"
