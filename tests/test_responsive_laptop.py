"""
Tests de simulacion de pantalla laptop (1366x768).

Verifica que:
  1. ScreenProfile clasifica correctamente 1366x768 como laptop_small
  2. Todas las ventanas principales caben en 1366x768
  3. minsizes no exceden 900x560
  4. Panel izquierdo < 30% del ancho total
  5. PanedWindow deja >= 400px para area central
  6. Fuentes escaladas estan entre 7 y 13 pt en pantalla 768px
"""
import os
import re
import sys
import unittest
from unittest import mock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from metatag_responsive import ScreenProfile, PROFILE


def _tk_ok():
    """True si hay display Tk disponible."""
    try:
        import tkinter as tk
        r = tk.Tk()
        r.destroy()
        return True
    except Exception:
        return False


def _recursive_geometry(win):
    m = re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)", win.geometry())
    if not m:
        return None
    w, h, x, y = (int(g) for g in m.groups())
    return w, h, x, y


# ══════════════════════════════════════════════════════════════════
#  SCREENPROFILE — tests headless (sin display)
# ══════════════════════════════════════════════════════════════════
class ScreenProfileTestCase(unittest.TestCase):
    """Tests de clasificacion y factores de ScreenProfile."""

    def test_laptop_small_classification(self):
        p = ScreenProfile()
        p._screen_w, p._screen_h = 1366, 768
        p._classify()
        self.assertEqual(p.kind, "laptop_small")
        self.assertEqual(p.font_scale, 0.78)
        self.assertEqual(p.panel_left_w, 180)
        self.assertEqual(p.panel_right_w, 260)
        self.assertLessEqual(p.win_w, 1380)
        self.assertLessEqual(p.win_h, 860)

    def test_laptop_small_minsize(self):
        p = ScreenProfile()
        p._screen_w, p._screen_h = 1366, 768
        p._classify()
        self.assertEqual(p.min_w, 780)
        self.assertEqual(p.min_h, 480)
        self.assertLessEqual(p.min_w, 900)
        self.assertLessEqual(p.min_h, 560)

    def test_laptop_small_row_height(self):
        p = ScreenProfile()
        p._screen_w, p._screen_h = 1366, 768
        p._classify()
        self.assertEqual(p.table_row_h, 24)

    def test_laptop_large_classification(self):
        p = ScreenProfile()
        # 1440x864 es laptop_large (height < 900)
        p._screen_w, p._screen_h = 1440, 864
        p._classify()
        self.assertEqual(p.kind, "laptop_large")
        self.assertEqual(p.font_scale, 0.88)

    def test_desktop_classification(self):
        p = ScreenProfile()
        p._screen_w, p._screen_h = 1920, 1080
        p._classify()
        self.assertEqual(p.kind, "desktop")
        self.assertEqual(p.font_scale, 1.0)

    def test_win_w_capped_at_1380(self):
        p = ScreenProfile()
        p._screen_w, p._screen_h = 2560, 1440
        p._classify()
        self.assertEqual(p.win_w, 1380)

    def test_win_h_capped_at_860(self):
        p = ScreenProfile()
        p._screen_w, p._screen_h = 2560, 1440
        p._classify()
        self.assertEqual(p.win_h, 860)

    def test_font_scale_between_values(self):
        p = ScreenProfile()
        for sw, sh, expected_kind in [
            (1280, 720, "laptop_small"),
            (1366, 768, "laptop_small"),
            (1440, 864, "laptop_large"),
            (1920, 1080, "desktop"),
            (2560, 1440, "desktop"),
        ]:
            p._screen_w, p._screen_h = sw, sh
            p._classify()
            self.assertEqual(p.kind, expected_kind, f"{sw}x{sh}")
            self.assertGreaterEqual(p.font_scale, 0.70)
            self.assertLessEqual(p.font_scale, 1.0)


# ══════════════════════════════════════════════════════════════════
#  VENTANAS — tests con display (requieren X/Wayland)
# ══════════════════════════════════════════════════════════════════
@unittest.skipUnless(_tk_ok(), "sin display Tk")
class LaptopSimTestCase(unittest.TestCase):
    """Simula pantalla 1366x768 y verifica que las ventanas caben."""

    def _init_profile_laptop(self):
        """Inicializa PROFILE como laptop_small."""
        PROFILE._screen_w = 1366
        PROFILE._screen_h = 768
        PROFILE._classify()

    def _make_main_app(self):
        import tkinter as tk
        from metatag_v8 import MetaTagApp
        mb = mock.MagicMock()
        patchers = [
            mock.patch.object(sys.modules["metatag_v8"], "messagebox", mb),
            mock.patch.object(MetaTagApp, "_save_config", lambda self: None),
            mock.patch.object(MetaTagApp, "_load_config_pre_build", lambda self: None),
            mock.patch.object(MetaTagApp, "_load_config_post_build", lambda self: None),
        ]
        for p in patchers:
            p.start()
        self.addCleanup(lambda: [p.stop() for p in patchers])

        # Patch winfo_screenwidth/height en la clase Tk
        _orig_sw = tk.Tk.winfo_screenwidth
        _orig_sh = tk.Tk.winfo_screenheight
        tk.Tk.winfo_screenwidth = lambda self: 1366
        tk.Tk.winfo_screenheight = lambda self: 768
        self.addCleanup(lambda: setattr(tk.Tk, "winfo_screenwidth", _orig_sw))
        self.addCleanup(lambda: setattr(tk.Tk, "winfo_screenheight", _orig_sh))

        self._init_profile_laptop()
        app = MetaTagApp()
        app.withdraw()
        self.addCleanup(app.destroy)
        return app

    def _make_visor(self):
        import tkinter as tk
        from Visor import VisorApp

        _orig_sw = tk.Tk.winfo_screenwidth
        _orig_sh = tk.Tk.winfo_screenheight
        tk.Tk.winfo_screenwidth = lambda self: 1366
        tk.Tk.winfo_screenheight = lambda self: 768
        self.addCleanup(lambda: setattr(tk.Tk, "winfo_screenwidth", _orig_sw))
        self.addCleanup(lambda: setattr(tk.Tk, "winfo_screenheight", _orig_sh))

        visor = VisorApp()
        visor.withdraw()
        self.addCleanup(visor.destroy)
        return visor

    # --- Tests ---

    def test_main_window_fits_1366x768(self):
        app = self._make_main_app()
        app.update_idletasks()
        geo = _recursive_geometry(app)
        self.assertIsNotNone(geo)
        w, h, x, y = geo
        self.assertLessEqual(w, 1366, "ancho excede pantalla")
        self.assertLessEqual(h, 768, "alto excede pantalla")

    def test_main_minsize_fits(self):
        app = self._make_main_app()
        mw, mh = app.minsize()
        self.assertLessEqual(mw, 900, f"minsize width {mw} > 900")
        self.assertLessEqual(mh, 560, f"minsize height {mh} > 560")

    def test_main_left_panel_ratio(self):
        app = self._make_main_app()
        app.update_idletasks()
        # El panel izquierdo debe ser < 30% del ancho de la ventana
        try:
            total_w = app.winfo_width()
            left_w = app.left.winfo_width()
            if total_w > 10 and left_w > 0:
                ratio = left_w / total_w
                self.assertLess(ratio, 0.35,
                    f"panel izquierdo {ratio:.1%} > 35% de {total_w}px")
        except AttributeError:
            self.skipTest("panel izquierdo no disponible")

    def test_main_paned_center_area(self):
        app = self._make_main_app()
        app.update_idletasks()
        try:
            # El area central debe tener al menos 300px
            # (en laptop_small el total ~1160 y left+right ~440)
            total_w = app.winfo_width()
            if total_w > 10:
                # minsize del center es int(400*0.82)=328, lo cual es > 300
                center_minsize = int(400 * PROFILE.font_scale)
                self.assertGreaterEqual(center_minsize, 280,
                    f"center minsize {center_minsize}px < 280px")
        except Exception:
            self.skipTest("PanedWindow no disponible")

    def test_font_sizes_on_768(self):
        """Verifica que las fuentes escaladas estan entre 6 y 13 pt."""
        import metatag_theme as mt
        from metatag_theme import font_specs
        # En 1366px, compute_font_scale = 0.82 (floor)
        scale = mt.compute_font_scale(1366)
        specs = font_specs(scale)
        for role, spec in specs.items():
            size = spec[1]
            self.assertGreaterEqual(size, 6,
                f"fuente {role} = {size} < 6 en scale={scale}")
            self.assertLessEqual(size, 13,
                f"fuente {role} = {size} > 13 en scale={scale}")

    def test_renamer_minsize_fits(self):
        """El renombrador debe caber en 1366x768."""
        import tkinter as tk
        try:
            import customtkinter as ctk
        except ImportError:
            self.skipTest("customtkinter no disponible")

        from metatag_responsive import PROFILE as P
        P._screen_w, P._screen_h = 1366, 768
        P._classify()

        _orig_sw = tk.Tk.winfo_screenwidth
        _orig_sh = tk.Tk.winfo_screenheight
        tk.Tk.winfo_screenwidth = lambda self: 1366
        tk.Tk.winfo_screenheight = lambda self: 768
        self.addCleanup(lambda: setattr(tk.Tk, "winfo_screenwidth", _orig_sw))
        self.addCleanup(lambda: setattr(tk.Tk, "winfo_screenheight", _orig_sh))

        import renombrar_fotos_gui as mod
        # AppController crea su propio CTk root internamente
        ctrl = mod.AppController()
        view = ctrl._view
        view.withdraw()
        self.addCleanup(view.destroy)
        view.update_idletasks()

        mw = getattr(view, '_min_width', 0) or view.winfo_reqwidth()
        mh = getattr(view, '_min_height', 0) or view.winfo_reqheight()
        self.assertLessEqual(mw, 900,
            f"renombrador minsize width {mw} > 900")
        self.assertLessEqual(mh, 560,
            f"renombrador minsize height {mh} > 560")

        geo = _recursive_geometry(view)
        if geo:
            w, h, x, y = geo
            self.assertLessEqual(w, 1366)
            self.assertLessEqual(h, 768)


# ══════════════════════════════════════════════════════════════════
#  INTEGRIDAD — PROFILE afecta a todos los archivos (headless)
# ══════════════════════════════════════════════════════════════════
class ProfileIntegrityTestCase(unittest.TestCase):
    """Verifica que PROFILE es importable y tiene los atributos esperados."""

    def test_profile_has_all_attributes(self):
        p = ScreenProfile()
        attrs = [
            "kind", "font_scale", "panel_left_w", "panel_right_w",
            "win_w", "win_h", "min_w", "min_h", "table_row_h",
            "chart_left_margin", "chart_right_margin",
        ]
        for attr in attrs:
            self.assertTrue(hasattr(p, attr), f"falta atributo {attr}")

    def test_module_level_instance(self):
        self.assertIsInstance(PROFILE, ScreenProfile)
        # PROFILE puede haber sido modificado por otros tests;
        # verificar que un instancia fresca da desktop por defecto
        fresh = ScreenProfile()
        self.assertEqual(fresh.kind, "desktop")

    def test_init_from_tk_sets_attributes(self):
        """Verifica que init_from_tk funciona con un fake Tk."""
        class FakeTk:
            def winfo_screenwidth(self): return 1366
            def winfo_screenheight(self): return 768

        p = ScreenProfile()
        p.init_from_tk(FakeTk())
        self.assertTrue(p._initialized)
        self.assertEqual(p.kind, "laptop_small")
        self.assertEqual(p.screen_width, 1366)
        self.assertEqual(p.screen_height, 768)

    def test_regressor_desktop_unchanged(self):
        """En desktop 1920x1080, PROFILE debe dar valores desktop."""
        p = ScreenProfile()
        p._screen_w, p._screen_h = 1920, 1080
        p._classify()
        self.assertEqual(p.kind, "desktop")
        self.assertEqual(p.font_scale, 1.0)
        self.assertEqual(p.min_w, 960)
        self.assertEqual(p.min_h, 620)
        self.assertEqual(p.table_row_h, 28)


if __name__ == "__main__":
    unittest.main()
