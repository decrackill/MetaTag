"""
metatag_theme.py — Tokens técnicos de tema y tipografía de MetaTag v8.9.

FUENTE DE VERDAD TÉCNICA de los temas y del motor de escalado de fuentes.

Este módulo contiene SOLO datos y lógica pura: no importa tkinter ni
customtkinter, por lo que puede ser consumido por cualquier parte del
proyecto (MetaTag con tk/ttk y el Renombrador con customtkinter) sin
acoplamiento a un toolkit concreto.

* THEMES / THEME_ORDER / DEFAULT_THEME / THEME_ICONS son los valores
  canónicos reales que usa MetaTag (extraídos de metatag_v8.py).
* El motor de fuentes replica exactamente la fórmula de escalado de
  MetaTag (mismo divisor de referencia, mismo rango de escala y la misma
  fórmula `max(floor, int(base * scale))`).
* Los adaptadores traducen los tokens canónicos al esquema que espera
  cada toolkit SIN inventar colores: los valores canónicos se pasan tal
  cual y las claves que no existen en el esquema canónico se derivan de
  forma determinista a partir de él.

Autor: Deivis
"""
from __future__ import annotations

# ══════════════════════════════════════════════════════════════════
#  TEMAS CANÓNICOS  (valores reales de MetaTag — metatag_v8.py)
# ══════════════════════════════════════════════════════════════════
THEMES: dict[str, dict[str, object]] = {
    "Arqueológico (Oscuro Refinado)": {
        "bg": "#121212", "surface": "#1E1E1E", "card": "#1A1A1A", "panel": "#2D2D30", "border": "#3E3E42", "border_light": "#252526",
        "accent": "#A67C52", "accent_hover": "#D4A574", "accent_light": "#7A4F2D", "accent_pale": "#3D1F0A", "header_bg": "#2D2D30",
        "header_fg": "#E0E0E0", "row_even": "#1E1E1E", "row_odd": "#1A1A1A", "sel_bg": "#3D1F0A", "sel_fg": "#E0E0E0", "col_sel": "#1A1A1A",
        "text": "#E0E0E0", "text2": "#AAAAAA", "text3": "#707070", "ok": "#4EC9B0", "err": "#F44747", "warn": "#CB4B16", "grid_line": "#3E3E42",
        "btn_ghost_bg": "#1A1A1A", "chart_colors": ["#A67C52", "#D4A574", "#7A4F2D", "#3D1F0A", "#B1A28F", "#5C3518"]
    },
    "Noche Total": {
        "bg": "#0A0A0A", "surface": "#111111", "card": "#0A0A0A", "panel": "#141414", "border": "#2A2A2A", "border_light": "#1E1E1E",
        "accent": "#BB86FC", "accent_hover": "#D0A8FF", "accent_light": "#6200EA", "accent_pale": "#1A0A2E", "header_bg": "#1A1A1A",
        "header_fg": "#E8E8E8", "row_even": "#111111", "row_odd": "#161616", "sel_bg": "#3700B3", "sel_fg": "#FFFFFF", "col_sel": "#1A0A2E",
        "text": "#E8E8E8", "text2": "#AAAAAA", "text3": "#666666", "ok": "#03DAC6", "err": "#CF6679", "warn": "#FF9800", "grid_line": "#1E1E1E",
        "btn_ghost_bg": "#1E1E1E", "chart_colors": ["#BB86FC", "#6200EA", "#03DAC6", "#CF6679", "#018786", "#FF9800"]
    },
    "Carbón": {
        "bg": "#1E1E1E", "surface": "#252526", "card": "#1E1E1E", "panel": "#252526", "border": "#3E3E42", "border_light": "#2D2D30",
        "accent": "#569CD6", "accent_hover": "#79B8FF", "accent_light": "#264F78", "accent_pale": "#1E3A5F", "header_bg": "#007ACC",
        "header_fg": "#FFFFFF", "row_even": "#252526", "row_odd": "#2D2D30", "sel_bg": "#264F78", "sel_fg": "#FFFFFF", "col_sel": "#1E3A5F",
        "text": "#D4D4D4", "text2": "#9CDCFE", "text3": "#6A9955", "ok": "#4EC9B0", "err": "#F44747", "warn": "#D97706", "grid_line": "#3E3E42",
        "btn_ghost_bg": "#2D2D30", "chart_colors": ["#569CD6", "#007ACC", "#4EC9B0", "#F44747", "#CE9178", "#9CDCFE"]
    },
}

THEME_ORDER: list[str] = list(THEMES)

DEFAULT_THEME: str = "Arqueológico (Oscuro Refinado)"

THEME_ICONS: dict[str, str] = {
    "Arqueológico (Oscuro Refinado)": "🏺",
    "Noche Total": "🌑",
    "Carbón": "⬛",
}

# Texto sobre fondo accent. MetaTag lo usa así, en mayúsculas (metatag_v8.py).
ACCENT_TEXT: str = "#FFF5E8"

# ══════════════════════════════════════════════════════════════════
#  MOTOR DE FUENTES DINÁMICO
# ══════════════════════════════════════════════════════════════════
SCALE_REFERENCE_WIDTH: int = 1920
SCALE_RANGE: tuple[float, float] = (0.82, 1.35)


def compute_font_scale(screen_width: int) -> float:
    """Factor de escala idéntico al de MetaTag: clamp(sw/1920, 0.82, 1.35)."""
    return max(SCALE_RANGE[0], min(SCALE_RANGE[1], screen_width / SCALE_REFERENCE_WIDTH))


def scaled_size(base: int, scale: float, floor: int = 6) -> int:
    """Misma fórmula que MetaTag: `max(floor, int(base * scale))`."""
    return max(floor, int(base * scale))


# Descripción exacta de las fuentes de MetaTag (estructura idéntica a
# `set_font_scale` de metatag_v8.py: los roles sin peso devuelven tuplas
# de 2 elementos y los roles en negrita tuplas de 3).
_FONT_BASE: dict[str, tuple[str, int, str | None, int]] = {
    "TITLE":   ("Georgia",  15, "bold", 8),
    "H2":      ("Georgia",  11, "bold", 8),
    "LABEL":   ("Segoe UI",  9, None,   7),
    "LABEL_B": ("Segoe UI",  9, "bold", 7),
    "BODY":    ("Segoe UI", 10, None,   8),
    "MONO":    ("Consolas",  9, None,   7),
    "CELL":    ("Segoe UI",  9, None,   7),
    "HEAD":    ("Segoe UI",  9, "bold", 7),
    "TINY":    ("Segoe UI",  8, None,   6),
}


def font_specs(scale: float) -> dict[str, tuple[str, int] | tuple[str, int, str]]:
    """Genera el diccionario FONTS con el mismo algoritmo que MetaTag."""
    out: dict[str, tuple[str, int] | tuple[str, int, str]] = {}
    for role, (family, size, weight, floor) in _FONT_BASE.items():
        spec: tuple[str, int] | tuple[str, int, str] = (family, max(floor, int(size * scale)))
        if weight:
            spec = (spec[0], spec[1], weight)
        out[role] = spec
    return out


def fit_to_screen(sw: int, sh: int, minsize: int, ideal: int, padding: int = 30) -> int:
    """Tamaño que mejor se adapta a la pantalla sin pasar del ideal."""
    scale = min(1.0, max(minsize / ideal, (sw - padding) / ideal, (sh - padding) / ideal))
    return max(minsize, int(ideal * scale))


# ══════════════════════════════════════════════════════════════════
#  UTILIDADES DE COLOR
# ══════════════════════════════════════════════════════════════════
def _channels(hex_color: str) -> tuple[int, int, int]:
    return (int(hex_color[1:3], 16), int(hex_color[3:5], 16), int(hex_color[5:7], 16))


def mix(c1: str, c2: str, t: float) -> str:
    """Interpola linealmente c1→c2 (t=0→1). Devuelve #RRGGBB en mayúsculas."""
    a, b = _channels(c1), _channels(c2)
    rgb = tuple(round(x + (y - x) * t) for x, y in zip(a, b))
    return "#%02X%02X%02X" % rgb


def relative_luminance(hex_color: str) -> float:
    """Luminancia relativa aproximada (no lineal) de un color hex."""
    def _lin(c: int) -> float:
        c /= 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = _channels(hex_color)
    return 0.2126 * _lin(r) + 0.7152 * _lin(g) + 0.0722 * _lin(b)


# ══════════════════════════════════════════════════════════════════
#  ADAPTADORES
# ══════════════════════════════════════════════════════════════════
class TkThemeAdapter:
    """Adaptador Tk/ttk: los tokens canónicos se consumen sin transformación.

    MetaTag define su esquema y este adaptador lo devuelve literalmente,
    de modo que no hay ninguna capa intermedia que pueda alterar la imagen.
    """

    names = THEME_ORDER
    default = DEFAULT_THEME

    @classmethod
    def palette(cls, theme: str) -> dict[str, object]:
        if theme not in cls.names:
            theme = cls.default
        return dict(THEMES[theme])


class CustomTkinterThemeAdapter:
    """Adaptador customtkinter: traduce los tokens canónicos de MetaTag al
    esquema que consume el Renombrador.

    Nada se inventa:
    * las claves que existen en el esquema canónico se pasan tal cual;
    * las claves propias del esquema del Renombrador que NO existen en el
      esquema canónico se derivan de forma determinista (comentado en cada
      una); los colores de estado reemplazan los literales CSS
      ("green"/"red"/"yellow") y los fondos de estado reemplazan los
      literales antiguos por tintes derivados de los semánticos reales.
    """

    names = THEME_ORDER
    default = DEFAULT_THEME

    # Estados del plan de renombrado → token semántico canónico.
    _STATE_TOKEN = {
        "ya_correcto": "ok",
        "conflicto":   "err",
        "duplicado":   "err",
        "not_found":   "err",
        "ambiguo":     "warn",
        "error":       "err",
    }

    # Proporción del tinte para los fondos de estado (derivado, no inventado).
    _STATE_TINT = 0.14

    @classmethod
    def palette(cls, theme: str) -> dict[str, object]:
        if theme not in cls.names:
            theme = cls.default
        t = dict(THEMES[theme])

        # Derivaciones (valores que el esquema del Renombrador espera y que
        # no existen como tales en el esquema canónico de MetaTag).
        surface2 = t["btn_ghost_bg"]
        dup_bg = mix(str(t["surface"]), str(t["err"]), cls._STATE_TINT)
        t.update({
            # Mapeos 1:1 del esquema canónico.
            "subtext": t["text2"],
            "overlay": t["text3"],
            "accent2": t["accent_hover"],
            "green":   t["ok"],
            "red":     t["err"],
            "yellow":  t["warn"],
            # Derivaciones por mezcla determinista.
            "surface2":   surface2,
            "surface3":   mix(surface2, str(t["border"]), 0.5),
            "dup_bg":     dup_bg,
            "accent_text": ACCENT_TEXT,
        })
        state_bg: dict[str, str] = {}
        state_fg: dict[str, str] = {}
        for state, token in cls._STATE_TOKEN.items():
            state_fg[state] = str(t[token])
            state_bg[state] = dup_bg if token == "err" else mix(str(t["surface"]), str(t[token]), cls._STATE_TINT)
        t["state_bg"] = state_bg
        t["state_fg"] = state_fg
        return t
