"""ExcelGrid: verificación de la optimización de redraw (Bloque 2).

Estrategia:
1. Equivalencia lógica (sin Tk): el enfoque "naive" (calcular
   _col_fully_selected dentro del bucle de celdas) y el "precalculado"
   (mapa col_sel_map calculado una sola vez por columna visible) deben
   producir exactamente las mismas decisiones de render, usando el método
   real ExcelGrid._col_fully_selected.
2. Selección: comportamiento real de _col_fully_selected ante distintos
   estados de selección (fila, columna, parcial, múltiple, ninguna).
3. Conteo de llamadas (Tk real): redraw() invoca _col_fully_selected como
   máximo una vez por columna visible (ya no una vez por celda).
4. Smoke test real de ExcelGrid (carga, redraw, toggle de columna,
   select_row, clear_selection, scroll) — se ejecuta SOLO si hay un display
   Tk funcional; en caso contrario se omite y queda pendiente de validación
   manual.
"""
import os
import sys
import tkinter as tk
import unittest
from unittest import mock

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
from metatag_widgets import ExcelGrid


def _tk_ok():
    try:
        root = tk.Tk()
        root.destroy()
        return True
    except Exception:
        return False


def make_df(nrows=4, ncols=3):
    return pd.DataFrame({f"C{i}": list(range(nrows)) for i in range(ncols)})


class FakeState:
    """Estado mínimo para evaluar la lógica real sin Tk."""

    def __init__(self, df, selected, hovered=None):
        self.df = df
        self.selected_cells = selected
        self.hovered_row = hovered


def _real_col_sel(state, ci):
    return ExcelGrid._col_fully_selected(state, ci)


def cell_style(ri, ci, state, col_sel):
    """Réplica exacta de la decisión de color de la celda en redraw()."""
    if (ri, ci) in state.selected_cells:
        return "sel"
    if col_sel:
        return "col"
    if ri == state.hovered_row:
        return "hover"
    return "row"


def render_naive(state, ncols, nrows):
    """Enfoque ANTES: col_sel recalculado en cada celda."""
    out = []
    for ci in range(ncols):
        out.append([_real_col_sel(state, ci)])
        for ri in range(nrows):
            out[ci].append(cell_style(ri, ci, state, _real_col_sel(state, ci)))
    return out


def render_map(state, ncols, nrows):
    """Enfoque DESPUÉS: col_sel_map precalculado una vez por columna."""
    col_sel_map = {ci: _real_col_sel(state, ci) for ci in range(ncols)}
    out = []
    for ci in range(ncols):
        out.append([col_sel_map[ci]])
        for ri in range(nrows):
            out[ci].append(cell_style(ri, ci, state, col_sel_map[ci]))
    return out


class LogicEquivalenceTestCase(unittest.TestCase):
    """1. Ambos enfoques producen exactamente los mismos estados por columna."""

    NROWS, NCOLS = 4, 3

    def _states(self):
        df = make_df(self.NROWS, self.NCOLS)
        all_cells = {(r, c) for r in range(self.NROWS) for c in range(self.NCOLS)}
        col0 = {(r, 0) for r in range(self.NROWS)}
        col2 = {(r, 2) for r in range(self.NROWS)}
        return [
            ("ninguna",        FakeState(df, set(), hovered=None)),
            ("fila 1",         FakeState(df, {(1, c) for c in range(self.NCOLS)}, hovered=None)),
            ("columna 0",      FakeState(df, col0, hovered=None)),
            ("parcial",        FakeState(df, {(0, 0), (1, 0), (2, 0)}, hovered=None)),
            ("multiple",       FakeState(df, col0 | col2 | {(0, 1)}, hovered=None)),
            ("con hover",      FakeState(df, {(0, 1), (1, 2)}, hovered=2)),
            ("todo",           FakeState(df, all_cells, hovered=None)),
            ("una celda",      FakeState(df, {(0, 1)}, hovered=None)),
        ]

    def test_equivalencia_naive_vs_mapa(self):
        for label, state in self._states():
            with self.subTest(estado=label):
                self.assertEqual(
                    render_naive(state, self.NCOLS, self.NROWS),
                    render_map(state, self.NCOLS, self.NROWS))

    def test_mapa_por_columna_es_el_estado_real(self):
        df = make_df(self.NROWS, self.NCOLS)
        state = FakeState(df, {(r, 0) for r in range(self.NROWS)})
        col_sel_map = {ci: _real_col_sel(state, ci) for ci in range(self.NCOLS)}
        self.assertTrue(col_sel_map[0])
        self.assertFalse(col_sel_map[1])
        self.assertFalse(col_sel_map[2])


class SelectionTestCase(unittest.TestCase):
    """2. _col_fully_selected real ante distintos estados."""

    NROWS, NCOLS = 4, 3

    def _grid(self, selected):
        return FakeState(make_df(self.NROWS, self.NCOLS), selected)

    def test_ninguna_seleccion(self):
        g = self._grid(set())
        self.assertEqual([_real_col_sel(g, c) for c in range(self.NCOLS)],
                         [False, False, False])

    def test_columna_seleccionada(self):
        g = self._grid({(r, 1) for r in range(self.NROWS)})
        self.assertEqual([_real_col_sel(g, c) for c in range(self.NCOLS)],
                         [False, True, False])

    def test_fila_seleccionada_no_es_columna_completa(self):
        # select_row selecciona una fila; ninguna columna queda totalmente
        # seleccionada salvo que el df tenga una sola fila.
        g = self._grid({(1, c) for c in range(self.NCOLS)})
        self.assertEqual([_real_col_sel(g, c) for c in range(self.NCOLS)],
                         [False, False, False])

    def test_seleccion_parcial_no_es_columna_completa(self):
        g = self._grid({(0, 0), (1, 0), (2, 0)})  # falta (3, 0)
        self.assertEqual([_real_col_sel(g, c) for c in range(self.NCOLS)],
                         [False, False, False])

    def test_seleccion_multiple(self):
        sel = {(r, 0) for r in range(self.NROWS)} | {(r, 2) for r in range(self.NROWS)}
        g = self._grid(sel)
        self.assertEqual([_real_col_sel(g, c) for c in range(self.NCOLS)],
                         [True, False, True])

    def test_todas_las_columnas_completas(self):
        sel = {(r, c) for r in range(self.NROWS) for c in range(self.NCOLS)}
        g = self._grid(sel)
        self.assertEqual([_real_col_sel(g, c) for c in range(self.NCOLS)],
                         [True, True, True])

    def test_df_vacio(self):
        g = FakeState(pd.DataFrame(), set())
        self.assertFalse(_real_col_sel(g, 0))


class _FakeApp:
    current_scale = 1.0


@unittest.skipUnless(_tk_ok(), "sin display Tk funcional — smoke test pendiente de validación manual")
class ExcelGridTkTestCase(unittest.TestCase):
    """3 y 4. Conteo de llamadas y smoke test real (requieren Tk)."""

    @classmethod
    def setUpClass(cls):
        cls.root = tk.Tk()
        cls.root.withdraw()

    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()

    def _grid(self, df):
        grid = ExcelGrid(self.root, app_ref=_FakeApp())
        grid.load(df)
        self.addCleanup(grid.destroy)
        return grid

    def test_col_fully_selected_una_vez_por_columna_visible(self):
        grid = self._grid(make_df(nrows=8, ncols=4))
        with mock.patch.object(grid, "_col_fully_selected",
                               wraps=grid._col_fully_selected) as m:
            grid.redraw()
            ncols = len(grid.df.columns)
            self.assertEqual(m.call_count, ncols)
            # ANTES (naive) habría sido (filas_visibles * ncols) + ncols de cabecera.
            naive_old = (len(grid.df) * ncols) + ncols
            self.assertLess(m.call_count, naive_old)

    def test_col_fully_selected_solo_columnas_visibles(self):
        grid = self._grid(make_df(nrows=6, ncols=3))
        grid.hidden_columns = {"C1"}
        with mock.patch.object(grid, "_col_fully_selected",
                               wraps=grid._col_fully_selected) as m:
            grid.redraw()
            self.assertEqual(m.call_count, 2)  # C0 y C2

    def test_smoke_redraw_y_selecciones(self):
        grid = self._grid(make_df(nrows=6, ncols=3))
        self.assertGreater(len(grid.canvas.find_all()), 0)

        grid._toggle_column(1)
        self.assertTrue(grid._col_fully_selected(1))
        grid.redraw()
        self.assertGreater(len(grid.canvas.find_all()), 0)

        grid.select_row(2)
        self.assertIn((2, 0), grid.selected_cells)
        grid.redraw()

        grid.clear_selection()
        self.assertEqual(len(grid.selected_cells), 0)

        grid.scroll_to_row(5)
        grid.redraw()
        self.assertGreater(len(grid.canvas.find_all()), 0)


if __name__ == "__main__":
    unittest.main()
