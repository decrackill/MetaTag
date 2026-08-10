"""Bloque 4 — responsividad de la interfaz.

Verifica que:
  1. `_clamp_toplevel` centra las ventanas secundarias y las limita al
     área visible (nunca quedan fuera de pantalla), para cualquier
     combinación de resolución y tamaño solicitado (headless, sin Tk).
  2. `_place_labels_clean` reduce el tamaño de fuente cuando hay muchas
     categorías para evitar solapamiento de etiquetas (headless).
  3. Con display Tk + matplotlib: la ventana de estadísticas se
     redimensiona dentro de la pantalla, tiene scrollbar de insights y
     la EXPORTACIÓN produce un PNG del mismo tamaño sin importar el
     tamaño de la ventana (visualización y exportación desacopladas).

Rango documentado: >= 1024x768 soportado. 900x600 y 800x600 son pruebas
de degradación (best effort), no soporte oficial.
"""
import os
import re
import sys
import tempfile
import unittest
from unittest import mock

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import metatag_v8
from metatag_v8 import MetaTagApp, _clamp_toplevel

try:
    import matplotlib
    import metatag_graficas as mg
    MATPLOTLIB_OK = mg.MATPLOTLIB_OK
except Exception:
    MATPLOTLIB_OK = False

try:
    from PIL import Image
    PIL_OK = True
except Exception:
    PIL_OK = False


class _FakeWin:
    def __init__(self, sw, sh):
        self._sw, self._sh = sw, sh
        self.geometry_calls = []

    def winfo_screenwidth(self):
        return self._sw

    def winfo_screenheight(self):
        return self._sh

    def geometry(self, geo):
        self.geometry_calls.append(geo)


class _FakeParent:
    def __init__(self, rx, ry, w, h):
        self._rx, self._ry, self._w, self._h = rx, ry, w, h

    def winfo_rootx(self):
        return self._rx

    def winfo_rooty(self):
        return self._ry

    def winfo_width(self):
        return self._w

    def winfo_height(self):
        return self._h


def _parse_geo(geo):
    m = re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)", geo)
    return tuple(int(g) for g in m.groups())


class ClampToplevelTestCase(unittest.TestCase):
    """Matemática de _clamp_toplevel (headless)."""

    SCREENS = [(1024, 768), (1366, 768), (1920, 1080),
               (2560, 1440), (900, 600), (800, 600)]

    def test_nunca_fuera_de_pantalla(self):
        for sw, sh in self.SCREENS:
            for w, h in [(200, 150), (400, 300), (800, 600),
                         (1000, 700), (2000, 1500)]:
                win = _FakeWin(sw, sh)
                _clamp_toplevel(win, None, w, h)
                gw, gh, gx, gy = _parse_geo(win.geometry_calls[-1])
                self.assertGreaterEqual(gx, 0, f"{sw}x{sh} size={w}x{h}")
                self.assertGreaterEqual(gy, 0, f"{sw}x{sh} size={w}x{h}")
                self.assertLessEqual(gx + gw, sw, f"{sw}x{sh} size={w}x{h}")
                self.assertLessEqual(gy + gh, sh, f"{sw}x{sh} size={w}x{h}")
                self.assertLessEqual(gw, sw, f"{sw}x{sh} size={w}x{h}")
                self.assertLessEqual(gh, sh, f"{sw}x{sh} size={w}x{h}")

    def test_ventana_grande_se_limita_a_pantalla(self):
        win = _FakeWin(1024, 768)
        _clamp_toplevel(win, None, 2000, 1500)
        gw, gh, _, _ = _parse_geo(win.geometry_calls[-1])
        self.assertEqual(gw, 1024 - 24)
        self.assertEqual(gh, 768 - 24)

    def test_centra_sobre_padre(self):
        win = _FakeWin(1024, 768)
        parent = _FakeParent(100, 100, 800, 600)  # centro = (500, 400)
        _clamp_toplevel(win, parent, 400, 300)
        gw, gh, gx, gy = _parse_geo(win.geometry_calls[-1])
        self.assertEqual((gx, gy), (300, 250))

    def test_padre_cerca_del_borde_no_sale_de_pantalla(self):
        win = _FakeWin(1024, 768)
        # El centro del padre está bajo: la ventana 800x700 no cabría debajo.
        parent = _FakeParent(0, 450, 1024, 400)  # centro = (512, 650)
        _clamp_toplevel(win, parent, 800, 700)
        gw, gh, gx, gy = _parse_geo(win.geometry_calls[-1])
        self.assertGreaterEqual(gx, 0)
        self.assertLessEqual(gx + gw, 1024)
        self.assertLessEqual(gy + gh, 768)


@unittest.skipUnless(MATPLOTLIB_OK, "sin matplotlib")
class PlaceLabelsTestCase(unittest.TestCase):
    """Fuente adaptativa de _place_labels_clean con muchas categorías."""

    class FakeWedge:
        def __init__(self, t1, t2):
            self.theta1, self.theta2 = t1, t2

    class FakeAx:
        def __init__(self):
            self.annots = []

        def annotate(self, *a, **kw):
            self.annots.append((a, kw))

    def _run(self, n):
        ax = self.FakeAx()
        wedges = [self.FakeWedge(10, 170)] * n
        items = [(1.0, f"cat {i}", 0.0, i, 1.0) for i in range(n)]
        mg._place_labels_clean(items, "left", wedges, ax,
                               ["#000000"] * n, "#111111", "#222222", "#EEEEEE")
        return ax

    def test_muchas_categorias_fuente_pequena(self):
        ax = self._run(20)
        self.assertTrue(ax.annots)
        for _, kw in ax.annots:
            self.assertEqual(kw["fontsize"], 6)

    def test_pocas_categorias_fuente_grande(self):
        ax = self._run(3)
        self.assertTrue(ax.annots)
        for _, kw in ax.annots:
            self.assertEqual(kw["fontsize"], 9)


def _tk_ok():
    try:
        import tkinter as tk
        root = tk.Tk()
        root.destroy()
        return True
    except Exception:
        return False


def _find_widget(widget, pred):
    if pred(widget):
        return widget
    for child in widget.winfo_children():
        found = _find_widget(child, pred)
        if found:
            return found
    return None


def _recursive_geometry(win):
    m = re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)", win.geometry())
    if not m:
        return None
    w, h, x, y = (int(g) for g in m.groups())
    return w, h, x, y


@unittest.skipUnless(_tk_ok() and MATPLOTLIB_OK and PIL_OK,
                     "sin display Tk / matplotlib / PIL — smoke pendiente")
class ResponsiveSmokeTestCase(unittest.TestCase):
    """Pruebas con UI real (requieren display)."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="metatag_responsive_")

    @classmethod
    def tearDownClass(cls):
        import shutil
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def _make_stats(self):
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        self.addCleanup(root.destroy)
        df = pd.DataFrame({
            "img": [f"im{i:02d}.jpg" for i in range(30)],
            "sitio": (["A"] * 12 + ["B"] * 8 + ["C"] * 6 + ["D"] * 4),
        })
        img_var = tk.StringVar(value="img")
        mg.show_stats(root, df, img_var, 1.0,
                      dict(metatag_v8.C), metatag_v8.FONTS)
        root.update()
        wins = [w for w in root.winfo_children() if isinstance(w, tk.Toplevel)]
        self.assertEqual(len(wins), 1)
        return root, wins[0]

    def _make_app(self):
        import tkinter as tk
        mb = mock.MagicMock()
        patchers = [
            mock.patch.object(metatag_v8, "messagebox", mb),
            mock.patch.object(MetaTagApp, "_save_config", lambda self: None),
            mock.patch.object(MetaTagApp, "_load_config_pre_build", lambda self: None),
            mock.patch.object(MetaTagApp, "_load_config_post_build", lambda self: None),
        ]
        for p in patchers:
            p.start()
        self.addCleanup(lambda: [p.stop() for p in patchers])
        app = MetaTagApp()
        app.withdraw()
        self.addCleanup(app.destroy)
        return app

    def test_show_stats_resize_dentro_de_pantalla(self):
        root, win = self._make_stats()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        for tw, th in [(900, 600), (1024, 768), (1600, 1000)]:
            win.geometry(f"{tw}x{th}")
            root.update_idletasks()
            geo = _recursive_geometry(win)
            self.assertIsNotNone(geo)
            w, h, x, y = geo
            self.assertGreaterEqual(x, 0, f"target={tw}x{th}")
            self.assertGreaterEqual(y, 0, f"target={tw}x{th}")
            self.assertLessEqual(x + w, sw, f"target={tw}x{th}")
            self.assertLessEqual(y + h, sh, f"target={tw}x{th}")

    def test_show_stats_minsize_escalado(self):
        root, win = self._make_stats()
        self.assertEqual(win.minsize(), (800, 500))

    def test_show_stats_tiene_scrollbar_de_insights(self):
        import tkinter as tk
        from tkinter import ttk
        root, win = self._make_stats()
        sb = _find_widget(win, lambda w: isinstance(w, ttk.Scrollbar))
        self.assertIsNotNone(sb)
        self.assertEqual(str(sb.cget("orient")), "vertical")

    def test_export_independiente_del_tamano_de_ventana(self):
        import tkinter as tk
        root, win = self._make_stats()
        btn = _find_widget(
            win, lambda w: isinstance(w, tk.Button)
            and "Exportar" in str(w.cget("text")))
        self.assertIsNotNone(btn)

        path1 = os.path.join(self.tmp, "export_pequeno.png")
        path2 = os.path.join(self.tmp, "export_grande.png")

        def _invoke(path):
            with mock.patch.object(mg, "_native_file_save",
                                   return_value=path), \
                 mock.patch.object(mg.messagebox, "showinfo",
                                   lambda *a, **k: None):
                btn.invoke()
            root.update_idletasks()

        win.geometry("900x600")
        root.update_idletasks()
        _invoke(path1)
        with Image.open(path1) as img:
            size1 = img.size

        win.geometry("1700x1100")
        root.update_idletasks()
        _invoke(path2)
        with Image.open(path2) as img:
            size2 = img.size

        self.assertEqual(size1, size2,
                         "la exportación depende del tamaño de ventana")

    def test_ventana_principal_resize_dentro_de_pantalla(self):
        app = self._make_app()
        sw, sh = app.winfo_screenwidth(), app.winfo_screenheight()
        for tw, th in [(1024, 768), (1366, 768), (900, 600), (800, 600)]:
            app.geometry(f"{tw}x{th}")
            app.update_idletasks()
            w, h = app.winfo_width(), app.winfo_height()
            self.assertLessEqual(w, sw, f"target={tw}x{th}")
            self.assertLessEqual(h, sh, f"target={tw}x{th}")
            self.assertGreaterEqual(w, 860, f"minsize ancho target={tw}x{th}")
            self.assertGreaterEqual(h, 520, f"minsize alto target={tw}x{th}")

    def test_ventanas_secundarias_dentro_de_pantalla(self):
        import tkinter as tk
        app = self._make_app()
        app._show_shortcuts()
        app.update_idletasks()
        sw, sh = app.winfo_screenwidth(), app.winfo_screenheight()
        wins = [w for w in app.winfo_children() if isinstance(w, tk.Toplevel)]
        self.assertTrue(wins)
        for win in wins:
            geo = _recursive_geometry(win)
            if geo is None:
                continue
            w, h, x, y = geo
            self.assertGreaterEqual(x, 0)
            self.assertGreaterEqual(y, 0)
            self.assertLessEqual(x + w, sw)
            self.assertLessEqual(y + h, sh)


if __name__ == "__main__":
    unittest.main()
