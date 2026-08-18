"""
metatag_responsive.py — Detección de pantalla y factores de escalado.

Clasifica la pantalla en categorías (laptop_small, laptop_large, desktop)
y computa dimensiones de layout consistentes para toda la aplicación.

Uso:
    from metatag_responsive import PROFILE
    # Después de crear tk.Tk():
    PROFILE.init_from_tk(root)
    # Ahora PROFILE.win_w, PROFILE.min_w, etc. están disponibles.
"""

from __future__ import annotations


class ScreenProfile:
    """Detecta resolución de pantalla y computa factores de escalado."""

    def __init__(self) -> None:
        self._initialized = False
        self._screen_w = 1920
        self._screen_h = 1080
        # Valores por defecto (desktop) mientras no se inicializa
        self.kind = "desktop"
        self.font_scale = 1.0
        self.panel_left_w = 235
        self.panel_right_w = 330
        self.win_w = 1380
        self.win_h = 860
        self.min_w = 960
        self.min_h = 620
        self.table_row_h = 28
        self.chart_left_margin = 0.18
        self.chart_right_margin = 0.56

    def init_from_tk(self, tk_root) -> None:
        """Inicializa el perfil detectando la pantalla real.

        Debe llamarse después de crear tk.Tk() o tk.Toplevel().
        """
        try:
            self._screen_w = tk_root.winfo_screenwidth()
            self._screen_h = tk_root.winfo_screenheight()
        except Exception:
            pass
        self._classify()
        self._initialized = True

    def _classify(self) -> None:
        sw, sh = self._screen_w, self._screen_h

        if sw < 1400 and sh < 820:
            # Portatil pequeno (1366x768, 1280x720, etc.)
            self.kind = "laptop_small"
            self.font_scale = 0.78
            self.panel_left_w = 180
            self.panel_right_w = 260
            self.min_w = 780
            self.min_h = 480
            self.table_row_h = 24
            self.chart_left_margin = 0.12
            self.chart_right_margin = 0.62
        elif sw < 1600 and sh < 900:
            # Portatil grande (1440x900, 1536x864, etc.)
            self.kind = "laptop_large"
            self.font_scale = 0.88
            self.panel_left_w = 210
            self.panel_right_w = 295
            self.min_w = 860
            self.min_h = 540
            self.table_row_h = 26
            self.chart_left_margin = 0.15
            self.chart_right_margin = 0.60
        else:
            # Escritorio (1920x1080 o mayor)
            self.kind = "desktop"
            self.font_scale = 1.0
            self.panel_left_w = 235
            self.panel_right_w = 330
            self.min_w = 960
            self.min_h = 620
            self.table_row_h = 28
            self.chart_left_margin = 0.18
            self.chart_right_margin = 0.56

        self.win_w = min(int(sw * 0.95), 1380)
        self.win_h = min(int(sh * 0.92), 860)

    @property
    def screen_width(self) -> int:
        return self._screen_w

    @property
    def screen_height(self) -> int:
        return self._screen_h


# Instancia unica del modulo — se inicializa lazy despues de tk.Tk()
PROFILE = ScreenProfile()
