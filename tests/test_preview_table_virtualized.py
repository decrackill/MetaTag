"""Tests de PreviewTable VIRTUALIZADO (FASE 3B.2).

Cubre la matriz exigida:
  - Tamaños de datos: 0, 1, <pool, =pool, pool+1, 269, 1000, 5000, 10000
  - Scroll: arriba / centro / abajo / rueda devuelve "break"
  - Filtro: 0 resultados, 1 resultado, tras scroll
  - Edición: visible, scroll-vuelta (commit), sanitizado de rutas
  - Estados: 7 estados, update_dup_states
  - Selección lógica: sobrevive scroll y filtro
  - Rebuild (tema): estado preservado y _rows truthy
  - No-crecimiento lineal de widgets

Ejecutar:
    env XMODIFIERS="@im=none" .venv/bin/python -m pytest tests/test_preview_table_virtualized.py -v
"""
import os
import sys
import tempfile
from pathlib import Path

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import renombrar_fotos_gui as mod
from renombrar_fotos_gui import PreviewTable

STATES = ["ok", "ya_correcto", "conflicto", "duplicado", "not_found", "ambiguo", "error"]
assert len(STATES) == 7


def make_pairs(n: int, states: bool = False, with_paths: bool = False):
    """Construye n pares (orig, new, photo_path, is_dup, state)."""
    pairs = []
    for i in range(n):
        orig = f"{i+1:04d}_UM_C4_XII_{i+1:04d}_F.jpg.JPG"
        new = f"{i+1:04d}_UM_C4_XII_{i+1:04d}_F.jpg"
        path = Path("/tmp/opencode/nonexistent.jpg") if with_paths else None
        pairs.append((orig, new, path, (i % 7) == 0, STATES[i % 7] if states else "ok"))
    return pairs


def _count_widgets(widget) -> int:
    n = 1
    for w in widget.winfo_children():
        n += _count_widgets(w)
    return n


@pytest.fixture
def root():
    import tkinter as tk
    r = tk.Tk()
    r.geometry("900x720+0+0")
    mod._init_fonts(r.winfo_screenwidth())
    yield r
    r.destroy()


@pytest.fixture
def preview(root):
    p = PreviewTable(root)
    p.pack(fill="both", expand=True)
    root.update()
    return p


def _pool_capacity(preview, root):
    """Capacidad real del pool (slots) con un render de sobra."""
    preview.render(make_pairs(100))
    root.update()
    return len(preview._rows)


# ── tamaño de datos ────────────────────────────────────────────────────────
class TestTamanos:
    def test_render_0_muestra_vacio(self, root, preview):
        preview.render([])
        root.update()
        assert preview._all_pairs == []
        assert preview._rows == []
        assert preview._lbl_empty.winfo_manager(), "el mensaje vacío debe estar visible"
        assert preview._cv is None, "sin datos no debe crearse el canvas"

    def test_render_1_fila(self, root, preview):
        preview.render(make_pairs(1))
        root.update()
        assert len(preview._all_pairs) == 1
        assert len(preview._rows) == 1
        assert preview._filtered == [0]
        slot = preview._rows[0]
        assert slot["pair_index"] == 0
        assert slot["orig_widget"].cget("text") == "0001_UM_C4_XII_0001_F.jpg.JPG"

    def test_render_menos_que_pool(self, root, preview):
        n = max(1, _pool_capacity(preview, root) - 5)
        preview.render(make_pairs(n))
        root.update()
        assert len(preview._rows) == n

    def test_render_igual_a_pool(self, root, preview):
        pool = _pool_capacity(preview, root)
        preview.render(make_pairs(pool))
        root.update()
        assert len(preview._rows) == pool
        assert all(s["index"] == k for k, s in enumerate(preview._rows))

    def test_render_pool_mas_1(self, root, preview):
        pool = _pool_capacity(preview, root)
        preview.render(make_pairs(pool + 1))
        root.update()
        assert len(preview._rows) == pool, "el pool no crece con el contenido"

    def test_render_269(self, root, preview):
        preview.render(make_pairs(269))
        root.update()
        assert len(preview._all_pairs) == 269
        assert len(preview._rows) <= 31
        assert preview._first == 0

    def test_render_1000(self, root, preview):
        preview.render(make_pairs(1000))
        root.update()
        assert len(preview._all_pairs) == 1000
        assert len(preview._rows) <= 31

    def test_render_5000_sin_badalloc(self, root, preview):
        preview.render(make_pairs(5000))
        root.update()
        assert len(preview._all_pairs) == 5000
        assert len(preview._rows) <= 31
        assert preview._cv is not None

    def test_render_10000_sin_badalloc(self, root, preview):
        preview.render(make_pairs(10000))
        root.update()
        assert len(preview._all_pairs) == 10000
        assert len(preview._rows) <= 31
        assert preview._first == 0


# ── scroll ─────────────────────────────────────────────────────────────────
class TestScroll:
    def _render(self, preview, root, n=1000):
        preview.render(make_pairs(n))
        root.update()
        root.update_idletasks()

    def test_scroll_arriba(self, root, preview):
        self._render(preview, root)
        preview._scroll_by(-10_000_000)
        root.update()
        assert preview._pixel_offset() == 0
        assert preview._first == 0

    def test_scroll_centro(self, root, preview):
        self._render(preview, root)
        total = len(preview._filtered) * PreviewTable.ROW_H
        view_h = preview._cv.winfo_height()
        middle = (total - view_h) // 2
        preview._scroll_by(middle)
        root.update()
        assert preview._first > 0
        assert preview._first < len(preview._filtered) - 1

    def test_scroll_abajo(self, root, preview):
        self._render(preview, root)
        preview._scroll_by(10_000_000)
        root.update()
        last = len(preview._filtered) - 1
        assert preview._first <= last
        assert preview._pixel_offset() > 0

    def test_scroll_rebasa_se_clampa(self, root, preview):
        self._render(preview, root)
        preview._scroll_by(10_000_000)
        root.update()
        bottom = preview._pixel_offset()
        preview._scroll_by(999_999_999)
        root.update()
        assert preview._pixel_offset() == bottom, "no debe exceder el final"

    def test_wheel_devuelve_break(self, root, preview):
        self._render(preview, root)
        from types import SimpleNamespace
        assert preview._on_wheel(SimpleNamespace(num=5, delta=None)) == "break"
        assert preview._on_wheel(SimpleNamespace(num=4, delta=None)) == "break"
        assert preview._on_wheel(SimpleNamespace(num=0, delta=120)) == "break"


# ── filtro ─────────────────────────────────────────────────────────────────
class TestFiltro:
    def test_filtro_0_resultados(self, root, preview):
        preview.render(make_pairs(269))
        root.update()
        preview._apply_filter("__NADA__")
        root.update()
        assert preview._filtered == []
        assert all(s["pair_index"] < 0 for s in preview._rows), "sin resultados no hay filas activas"

    def test_filtro_1_resultado(self, root, preview):
        preview.render(make_pairs(269))
        root.update()
        preview._apply_filter("0042")
        root.update()
        assert len(preview._filtered) == 1
        assert preview._all_pairs[preview._filtered[0]]["orig"].lower().find("0042") >= 0

    def test_filtro_tras_scroll(self, root, preview):
        preview.render(make_pairs(1000))
        root.update()
        preview._scroll_by(96 * 20)
        root.update()
        assert preview._first > 0
        preview._apply_filter("0950")
        root.update()
        assert len(preview._filtered) == 1
        assert preview._first == 0, "el viewport se reclampa tras filtrar"


# ── edición ────────────────────────────────────────────────────────────────
class TestEdicion:
    def test_edicion_visible_escribe_modelo(self, root, preview):
        preview.render(make_pairs(100))
        root.update()
        preview.set_edit_mode(True)
        root.update()
        slot = next(s for s in preview._rows if s["pair_index"] >= 0)
        pi = slot["pair_index"]
        assert slot["new_entry"].winfo_manager(), "en edición debe verse el entry"
        slot["new_var"].set("NUEVO_nombre.jpg")
        root.update()
        assert preview._all_pairs[pi]["new"] == "NUEVO_nombre.jpg"

    def test_edicion_scroll_vuelta_persiste(self, root, preview):
        edits = {}
        def cb(i, name, path):
            edits[i] = name
        p = PreviewTable(root, on_name_change=cb)
        p.pack(fill="both", expand=True)
        root.update()
        p.render(make_pairs(1000))
        root.update()
        p.set_edit_mode(True)
        root.update()
        slot = next(s for s in p._rows if s["pair_index"] >= 0)
        pi = slot["pair_index"]
        slot["new_var"].set("persistida.jpg")
        root.update()
        p._scroll_by(96 * 30)
        root.update()
        assert p._all_pairs[pi]["new"] == "persistida.jpg", "el valor debe persistir en el modelo"
        assert edits.get(pi) == "persistida.jpg", "el callback debe haberse notificado"

    def test_edicion_sanitiza_rutas(self, root, preview):
        preview.render(make_pairs(100))
        root.update()
        preview.set_edit_mode(True)
        root.update()
        slot = next(s for s in preview._rows if s["pair_index"] >= 0)
        pi = slot["pair_index"]
        slot["new_var"].set("carpeta/sub_nombre.jpg")
        root.update()
        assert preview._all_pairs[pi]["new"] == "sub_nombre.jpg"
        assert slot["new_var"].get() == "sub_nombre.jpg"

    def test_edit_mode_no_reconstruye(self, root, preview):
        preview.render(make_pairs(269))
        root.update()
        before = _count_widgets(preview)
        preview.set_edit_mode(True)
        root.update()
        preview.set_edit_mode(False)
        root.update()
        after = _count_widgets(preview)
        assert before == after, "set_edit_mode no debe crear/destruir widgets"


# ── estados ────────────────────────────────────────────────────────────────
class TestEstados:
    def test_siete_estados_visibles(self, root, preview):
        pairs = make_pairs(7, states=True)
        preview.render(pairs)
        root.update()
        for i, row in enumerate(preview._all_pairs):
            assert row["state"] == STATES[i]
        # cada slot muestra la etiqueta de estado correspondiente
        for slot in preview._rows:
            state = preview._all_pairs[slot["pair_index"]]["state"]
            assert slot["state_widget"].cget("text") == mod.STATE_LABELS.get(state, "")

    def test_update_dup_states_refresca_modelo(self, root, preview):
        preview.render(make_pairs(100))
        root.update()
        before = [r["state"] for r in preview._all_pairs]
        assert all(s == "ok" for s in before)
        pairs2 = make_pairs(100, states=True)
        preview.update_dup_states(pairs2)
        root.update()
        after = [r["state"] for r in preview._all_pairs]
        assert after == [STATES[i % 7] for i in range(100)]
        # el slot que muestra la fila 1 debe reflejar el nuevo color/estado
        slot = next(s for s in preview._rows if s["pair_index"] == 1)
        assert slot["state_widget"].cget("text") == mod.STATE_LABELS["ya_correcto"]


# ── selección / rebuild / widgets ──────────────────────────────────────────
class TestRobustez:
    def test_seleccion_sobrevive_scroll_y_filtro(self, root, preview):
        preview.render(make_pairs(1000))
        root.update()
        preview._set_row_selected(500, True)
        preview._scroll_by(96 * 5)
        root.update()
        assert 500 in preview._selected
        preview._apply_filter("0500")
        root.update()
        assert 500 in preview._selected
        preview._apply_filter("__NADA__")
        root.update()
        assert 500 in preview._selected, "la selección persiste aunque el filtro la oculte"

    def test_seleccion_desaparece_con_render(self, root, preview):
        preview.render(make_pairs(100))
        root.update()
        preview._set_row_selected(10, True)
        preview.render(make_pairs(5))
        root.update()
        assert 10 not in preview._selected
        assert preview._selected == set()

    def test_render_preserva_estado_tras_rebuild(self, root, preview):
        preview.render(make_pairs(1000))
        root.update()
        preview._set_row_selected(42, True)
        # rebuild de tema = re-render con los mismos datos
        preview.render(make_pairs(1000))
        root.update()
        assert 42 in preview._selected, "la selección sobrevive al rebuild"
        assert preview._rows, "tras el rebuild la tabla debe seguir materializada"

    def test_widgets_no_crecen_linealmente(self, root, preview):
        preview.render(make_pairs(1000))
        root.update()
        w1000 = _count_widgets(preview)
        preview.render(make_pairs(10000))
        root.update()
        w10000 = _count_widgets(preview)
        assert w10000 <= w1000 + 10, f"widgets deben ser planos: {w1000} -> {w10000}"

    def test_rebuild_tema_via_controller_preserva_rows(self, root):
        ctrl = mod.AppController()
        try:
            ctrl._last_pairs = [("a.jpg", "A.jpg", None, False, "ok")]
            target = next(t for t in mod.THEME_ORDER if t != mod.CURRENT_THEME)
            ctrl._apply_theme(target)
            assert ctrl._view._preview._rows, "la vista previa debe re-renderizarse"
        finally:
            ctrl._view.destroy()
