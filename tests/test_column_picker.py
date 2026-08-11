"""Smoke tests de las ventanas de selección de columnas.

Cubre los fixes de la Fase 4:
  - scroll por rueda/cursor sobre CANVAS, TEXTO, CHECKBUTTON y SCROLLBAR;
  - acciones rápidas "Todas / Ninguna / Invertir" y contador de selección;
  - redimensionado: el contenido se ancla a la izquierda sin overflow;
  - Aceptar devuelve la selección; Cancelar devuelve None.

Requisitos: display real con Tk (se omite en headless).
"""
import os
import sys
import time
import tkinter as tk
import unittest
from unittest import mock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import metatag_v8


def _tk_ok():
    try:
        r = tk.Tk()
        r.destroy()
        # No dejar una raíz huérfana como _default_root: si el picker crea
        # BooleanVar() sin master se ligaría al intérprete equivocado.
        tk._default_root = None
        return True
    except Exception:
        return False


import re

COLS_19 = [
    "ID", "FOTO", "PROCEDENCIA", "MORFOLOGIA", "TECNOLOGIA",
    "TRATAMIENTO", "PASTA", "DESGRASANTE", "COCCION", "COLOR",
    "DECORACION", "MEDIDAS", "PESO", "ANALISIS", "FECHA",
    "RESPONSABLE", "OBSERVACIONES", "ESTADO", "NIVEL",
]
IMG_COL = "FOTO"


def _win(app):
    wins = [w for w in app.winfo_children()
            if isinstance(w, tk.Toplevel) and w.winfo_exists()]
    return wins[-1]


def _all(app, cls):
    out = []

    def walk(w):
        for c in w.winfo_children():
            if isinstance(c, cls):
                out.append(c)
            walk(c)
    walk(app)
    return out


def _button(win, text):
    for b in _all(win, tk.Button):
        if str(b.cget("text")) == text:
            return b
    return None


def _cb_vars(app, win):
    return [(cb.cget("variable"), cb) for cb in _all(win, tk.Checkbutton)]


def _count_text(app, win):
    """Devuelve (seleccionadas, total) del label de contador ('18 / 19')."""
    for lb in _all(win, tk.Label):
        var = str(lb.cget("textvariable"))
        if var and var.startswith("PY_VAR"):
            m = re.match(r"^\s*(\d+)\s*/\s*(\d+)\s*$",
                         str(app.getvar(var)))
            if m:
                return int(m.group(1)), int(m.group(2))
    return None


def _click_when_ready(app, action, timeout_ms=3000):
    """Ejecuta action(win) en cuanto exista la ventana del picker."""
    import time as _t
    start = _t.time()

    def poll():
        try:
            wins = [w for w in app.winfo_children()
                    if isinstance(w, tk.Toplevel) and w.winfo_exists()]
            if wins:
                if action(wins[-1]):
                    return
        except tk.TclError:
            return
        if _t.time() - start > timeout_ms / 1000:
            return
        app.after(10, poll)
    app.after(10, poll)


def _click_button(action, text):
    def _do(win):
        b = _button(win, text)
        if b is None:
            return False
        action(b)
        return True
    return _do


def _pump_wait(app, win, timeout_ms=3000):
    """Bombea la cola hasta que la ventana se destruya o venza el timeout."""
    start = time.time()
    while win.winfo_exists() and (time.time() - start) < timeout_ms / 1000:
        app.update()
        time.sleep(0.01)


@unittest.skipUnless(_tk_ok(), "requiere display real con Tk")
class ColumnPickerBaseTestCase(unittest.TestCase):
    """Crea la app completa (patrón de test_responsive) y parchea wait_window
    para que el picker no bloquee el flujo de la prueba."""

    def setUp(self):
        p = mock.patch.object(metatag_v8.messagebox, "showwarning")
        self.m_mw = p.start()
        self.addCleanup(p.stop)
        p = mock.patch.object(metatag_v8.MetaTagApp, "_save_config",
                              lambda self: None)
        self.m_save = p.start()
        self.addCleanup(p.stop)
        p = mock.patch.object(metatag_v8.MetaTagApp, "_load_config_pre_build",
                              lambda self: None)
        self.m_pre = p.start()
        self.addCleanup(p.stop)
        p = mock.patch.object(metatag_v8.MetaTagApp, "_load_config_post_build",
                              lambda self: None)
        self.m_post = p.start()
        self.addCleanup(p.stop)
        self.app = metatag_v8.MetaTagApp()
        self.app.withdraw()
        self.app.update()

    def tearDown(self):
        try:
            self.app.destroy()
        except tk.TclError:
            pass

    def _open(self, cols, img_col="", method="batch"):
        result = [None]
        with mock.patch.object(metatag_v8.MetaTagApp, "wait_window",
                               lambda self_, w: self_.update()):
            if method == "batch":
                result[0] = self.app._batch_pick_columns(cols, img_col)
            else:
                result[0] = self.app._pick_sort_columns(cols)
        self.app.update()
        return _win(self.app), result[0]

    def _scroll(self, widget, count=5, down=True):
        btn = "<Button-5>" if down else "<Button-4>"
        for _ in range(count):
            widget.event_generate(btn, when="now")
            self.app.update()

    def _canvas(self, win):
        return _all(win, tk.Canvas)[0]


class ColumnPickerSmokeTestCase(ColumnPickerBaseTestCase):
    def test_abre_con_19_columnas(self):
        win, _ = self._open(COLS_19, IMG_COL)
        self.assertEqual(len(_all(win, tk.Checkbutton)), 19)
        self.assertEqual(_count_text(self.app, win), (18, 19))

    def test_todas_selecciona_todo(self):
        win, _ = self._open(COLS_19, IMG_COL)
        _button(win, "Todas").invoke()
        self.assertEqual(_count_text(self.app, win), (19, 19))

    def test_ninguna_limpia(self):
        win, _ = self._open(COLS_19, IMG_COL)
        _button(win, "Ninguna").invoke()
        self.assertEqual(_count_text(self.app, win), (0, 19))

    def test_invertir_preserva_columna_imagen(self):
        win, _ = self._open(COLS_19, IMG_COL)
        _button(win, "Todas").invoke()
        self.assertEqual(_count_text(self.app, win), (19, 19))
        _button(win, "Invertir").invoke()
        self.assertEqual(_count_text(self.app, win), (1, 19))

    def test_toggle_manual_actualiza_contador(self):
        win, _ = self._open(COLS_19, IMG_COL)
        cb = _all(win, tk.Checkbutton)[0]
        cb.deselect()
        self.assertEqual(_count_text(self.app, win), (17, 19))
        cb.select()
        self.assertEqual(_count_text(self.app, win), (18, 19))

    def test_aceptar_devuelve_seleccion(self):
        _click_when_ready(self.app,
                          _click_button(lambda b: b.invoke(), "Aceptar"))
        with mock.patch.object(metatag_v8.MetaTagApp, "wait_window",
                               lambda self_, w: _pump_wait(self_, w)):
            got = self.app._batch_pick_columns(COLS_19, IMG_COL)
        self.assertEqual(got, [c for c in COLS_19 if c != IMG_COL])

    def test_cancelar_devuelve_none(self):
        _click_when_ready(self.app,
                          _click_button(lambda b: b.invoke(), "Cancelar"))
        with mock.patch.object(metatag_v8.MetaTagApp, "wait_window",
                               lambda self_, w: _pump_wait(self_, w)):
            got = self.app._batch_pick_columns(COLS_19, IMG_COL)
        self.assertIsNone(got)

    def test_cancelar_limpia_traces(self):
        with mock.patch.object(metatag_v8.MetaTagApp, "wait_window",
                               lambda self_, w: self_.update()):
            self.app._batch_pick_columns(COLS_19, IMG_COL)
        self.app.update()
        win = _win(self.app)
        var_name = str(_all(win, tk.Checkbutton)[0].cget("variable"))
        _button(win, "Cancelar").invoke()
        self.app.update()
        info = self.app.tk.call("trace", "info", "variable", var_name)
        self.assertEqual(str(info).strip(), "",
                         "Cancelar debe limpiar las trazas write de Tk")


class ColumnPickerScrollTestCase(ColumnPickerBaseTestCase):
    """Scroll con ventana reducida para forzar desbordamiento vertical."""

    def setUp(self):
        super().setUp()
        self.win, _ = self._open(COLS_19, IMG_COL)
        self.win.geometry("460x360")
        self.app.update()

    def _top(self):
        return self._canvas(self.win).yview()[0]

    def test_scroll_cursor_sobre_canvas(self):
        self._scroll(self._canvas(self.win))
        self.assertGreater(self._top(), 0.0)

    def test_scroll_cursor_sobre_texto(self):
        self._scroll(_all(self.win, tk.Label)[0])
        self.assertGreater(self._top(), 0.0)

    def test_scroll_cursor_sobre_checkbutton(self):
        self._scroll(_all(self.win, tk.Checkbutton)[0])
        self.assertGreater(self._top(), 0.0)

    def test_scroll_cursor_sobre_scrollbar(self):
        self._scroll(_all(self.win, tk.Scrollbar)[0])
        self.assertGreater(self._top(), 0.0)

    def test_scrollbar_visible_y_anclada(self):
        vsb = _all(self.win, tk.Scrollbar)[0]
        self.assertIn("yview", str(vsb.cget("command")))
        self.assertEqual(str(vsb.winfo_manager()), "grid")

    def test_seleccion_sobrevive_al_desplazamiento(self):
        cbs = _all(self.win, tk.Checkbutton)
        for cb in cbs[::2]:
            cb.invoke()
        antes = [self.app.getvar(cb.cget("variable")) for cb in cbs]
        self._scroll(self._canvas(self.win))
        self.assertGreater(self._top(), 0.0)
        despues = [self.app.getvar(cb.cget("variable")) for cb in cbs]
        self.assertEqual(antes, despues)

    def test_redimensionar_crece_contenido_sin_overflow(self):
        self.win.geometry("900x720")
        self.app.update()
        canvas = self._canvas(self.win)
        self.assertGreater(canvas.winfo_width(), 400)
        left, right = canvas.xview()
        self.assertEqual((left, right), (0.0, 1.0))

    def test_volver_arriba_tras_bajar(self):
        self._scroll(self._canvas(self.win), down=True)
        self.assertGreater(self._top(), 0.0)
        self._scroll(self._canvas(self.win), count=10, down=False)
        self.assertAlmostEqual(self._top(), 0.0, places=2)


class ClampMatrixTestCase(unittest.TestCase):
    """Matriz de resoluciones: la ventana de selección siempre cabe en pantalla.

    Se simula el tamaño de pantalla parcheando winfo_screenwidth/height; se
    aplica la misma fórmula de dimensionado que usa el picker real.
    """

    RESOLUCIONES = [
        (640, 480), (800, 600), (1024, 768), (1280, 720),
        (1366, 768), (1536, 864), (1920, 1080), (2560, 1440),
    ]

    def setUp(self):
        self.app = tk.Tk()
        self.app.withdraw()
        self.addCleanup(self.app.destroy)

    def _screen(self, w, h):
        self.app.winfo_screenwidth = lambda: w
        self.app.winfo_screenheight = lambda: h

    @staticmethod
    def _intended(sw, sh, scale=1.0):
        w = min(int(480 * scale), max(int(380 * scale), int(sw * 0.45)))
        h = min(int(620 * scale), max(int(420 * scale), int(sh * 0.68)))
        return w, h

    def test_picker_cabe_en_pantalla(self):
        captura = {}
        self.app.geometry = lambda *a: captura.setdefault(
            "g", a[0]) if a else self.app.winfo_geometry()
        for sw, sh in self.RESOLUCIONES:
            self._screen(sw, sh)
            w, h = self._intended(sw, sh)
            metatag_v8._clamp_toplevel(self.app, None, w, h)
            m = re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)", captura["g"])
            W, H, X, Y = map(int, m.groups())
            tag = f"{sw}x{sh}"
            self.assertGreaterEqual(W, 24, tag)
            self.assertGreaterEqual(H, 24, tag)
            self.assertLessEqual(W, sw - 24, tag)
            self.assertLessEqual(H, sh - 24, tag)
            self.assertGreaterEqual(X, 0, tag)
            self.assertGreaterEqual(Y, 0, tag)
            self.assertLessEqual(X + W, sw, tag)
            self.assertLessEqual(Y + H, sh, tag)


class SortColumnsSmokeTestCase(ColumnPickerBaseTestCase):
    def test_abre_con_todas_deseleccionadas(self):
        win, _ = self._open(COLS_19, method="sort")
        cbs = _all(win, tk.Checkbutton)
        self.assertEqual(len(cbs), 19)
        self.assertTrue(all(self.app.getvar(str(cb.cget("variable"))) != "1"
                            for cb in cbs))

    def test_aceptar_devuelve_orden(self):
        def _selecciona_y_acepta(win):
            cbs = _all(win, tk.Checkbutton)
            if not cbs:
                return False
            cbs[0].select()
            b = _button(win, "Aceptar")
            if b is None:
                return False
            b.invoke()
            return True
        _click_when_ready(self.app, _selecciona_y_acepta)
        with mock.patch.object(metatag_v8.MetaTagApp, "wait_window",
                               lambda self_, w: _pump_wait(self_, w)):
            got = self.app._pick_sort_columns(COLS_19)
        self.assertEqual(got, [COLS_19[0]])


if __name__ == "__main__":
    unittest.main()
