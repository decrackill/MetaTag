"""Tests de regresión del módulo central de temas ``metatag_theme``.

Verifican que:
  - Los 3 temas canónicos de MetaTag existen y contienen todas las claves.
  - Los valores canónicos son EXACTAMENTE los reales de ``metatag_v8.py``
    (snapshot de la fuente de verdad; nada inventado).
  - ``TkThemeAdapter`` devuelve los tokens canónicos sin transformación.
  - ``CustomTkinterThemeAdapter`` traduce el esquema canónico al esquema
    del Renombrador de forma determinista y 1:1 (solo deriva lo que el
    esquema canónico no define, con reglas comentadas).
  - El motor de fuentes es idéntico al de ``metatag_v8.set_font_scale``.
  - ``fit_to_screen`` devuelve ventanas que caben en pantalla.
  - La migración de ``metatag_v8`` no introduce ninguna diferencia visual.
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import metatag_theme as mt


CANONICAL_KEYS = {
    "bg", "surface", "card", "panel", "border", "border_light",
    "accent", "accent_hover", "accent_light", "accent_pale", "header_bg",
    "header_fg", "row_even", "row_odd", "sel_bg", "sel_fg", "col_sel",
    "text", "text2", "text3", "ok", "err", "warn", "grid_line",
    "btn_ghost_bg", "chart_colors",
}

# Valores reales de metatag_v8.py:213-235 (fuente de verdad).
EXPECTED_CANONICAL = {
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


def _is_hex(value) -> bool:
    return isinstance(value, str) and value.startswith("#") and len(value) == 7


class TestCanonicalThemes:
    def test_theme_order_and_count(self):
        assert mt.THEME_ORDER == list(mt.THEMES)
        assert mt.THEME_ORDER == [
            "Arqueológico (Oscuro Refinado)", "Noche Total", "Carbón",
        ]
        assert len(mt.THEMES) == 3

    @pytest.mark.parametrize("name", mt.THEME_ORDER)
    def test_all_themes_have_full_key_set(self, name):
        assert CANONICAL_KEYS <= set(mt.THEMES[name])
        for key in CANONICAL_KEYS - {"chart_colors"}:
            assert _is_hex(mt.THEMES[name][key]), (
                f"{name}.{key} no es un color hex válido: {mt.THEMES[name][key]!r}"
            )
        cc = mt.THEMES[name]["chart_colors"]
        assert isinstance(cc, list) and len(cc) == 6
        assert all(_is_hex(c) for c in cc)

    @pytest.mark.parametrize("name", mt.THEME_ORDER)
    def test_historical_values_unchanged(self, name):
        assert mt.THEMES[name] == EXPECTED_CANONICAL[name]

    def test_default_theme(self):
        assert mt.DEFAULT_THEME == "Arqueológico (Oscuro Refinado)"
        assert mt.DEFAULT_THEME in mt.THEMES

    def test_theme_icons(self):
        assert set(mt.THEME_ICONS) == set(mt.THEMES)
        assert all(mt.THEME_ICONS[n] for n in mt.THEMES)

    def test_accent_text_is_metatag_constant(self):
        assert mt.ACCENT_TEXT == "#FFF5E8"


class TestTkAdapter:
    def test_palette_is_canonical_passthrough(self):
        for name in mt.THEME_ORDER:
            assert mt.TkThemeAdapter.palette(name) == mt.THEMES[name]

    def test_unknown_theme_falls_back_to_default(self):
        assert mt.TkThemeAdapter.palette("no-existe") == mt.THEMES[mt.DEFAULT_THEME]

    def test_metadata(self):
        assert mt.TkThemeAdapter.names == mt.THEME_ORDER
        assert mt.TkThemeAdapter.default == mt.DEFAULT_THEME


class TestCustomTkinterAdapter:
    # Esquema completo que el Renombrador consume.
    RENOMBRADOR_KEYS = {
        "bg", "surface", "surface2", "surface3", "accent", "accent2",
        "green", "red", "yellow", "text", "subtext", "border", "overlay",
    }
    EXTRA_KEYS = {"accent_text", "dup_bg", "state_bg", "state_fg"}

    @pytest.mark.parametrize("name", mt.THEME_ORDER)
    def test_complete_schema(self, name):
        p = mt.CustomTkinterThemeAdapter.palette(name)
        assert self.RENOMBRADOR_KEYS <= set(p)
        assert self.EXTRA_KEYS <= set(p)
        assert set(p["state_bg"]) == set(p["state_fg"])
        assert set(p["state_bg"]) == {"ya_correcto", "conflicto", "duplicado",
                                      "not_found", "ambiguo", "error"}

    @pytest.mark.parametrize("name", mt.THEME_ORDER)
    def test_deterministic(self, name):
        a = mt.CustomTkinterThemeAdapter.palette(name)
        b = mt.CustomTkinterThemeAdapter.palette(name)
        assert a == b

    def test_unknown_theme_falls_back_to_default(self):
        assert mt.CustomTkinterThemeAdapter.palette("no-existe") == \
            mt.CustomTkinterThemeAdapter.palette(mt.DEFAULT_THEME)

    @pytest.mark.parametrize("name", mt.THEME_ORDER)
    def test_canonical_mapping_is_11(self, name):
        p = mt.CustomTkinterThemeAdapter.palette(name)
        t = mt.THEMES[name]
        assert p["subtext"] == t["text2"]
        assert p["overlay"] == t["text3"]
        assert p["accent2"] == t["accent_hover"]
        assert p["green"] == t["ok"]
        assert p["red"] == t["err"]
        assert p["yellow"] == t["warn"]
        assert p["surface2"] == t["btn_ghost_bg"]
        assert p["surface"] == t["surface"]
        assert p["bg"] == t["bg"]
        assert p["accent"] == t["accent"]
        assert p["text"] == t["text"]
        assert p["border"] == t["border"]

    @pytest.mark.parametrize("name", mt.THEME_ORDER)
    def test_derived_values(self, name):
        p = mt.CustomTkinterThemeAdapter.palette(name)
        t = mt.THEMES[name]
        assert p["dup_bg"] == mt.mix(str(t["surface"]), str(t["err"]),
                                     mt.CustomTkinterThemeAdapter._STATE_TINT)
        assert p["surface3"] == mt.mix(str(p["surface2"]), str(t["border"]), 0.5)
        assert p["accent_text"] == mt.ACCENT_TEXT
        # Los fondos de estado "error" usan el mismo tinte de duplicado.
        for state in ("conflicto", "duplicado", "not_found", "error"):
            assert p["state_bg"][state] == p["dup_bg"]
        assert p["state_bg"]["ya_correcto"] == mt.mix(
            str(t["surface"]), str(t["ok"]), mt.CustomTkinterThemeAdapter._STATE_TINT)
        assert p["state_bg"]["ambiguo"] == mt.mix(
            str(t["surface"]), str(t["warn"]), mt.CustomTkinterThemeAdapter._STATE_TINT)

    @pytest.mark.parametrize("name", mt.THEME_ORDER)
    def test_state_fg_follows_semantics(self, name):
        p = mt.CustomTkinterThemeAdapter.palette(name)
        t = mt.THEMES[name]
        assert p["state_fg"]["ya_correcto"] == t["ok"]
        for state in ("conflicto", "duplicado", "not_found", "error"):
            assert p["state_fg"][state] == t["err"]
        assert p["state_fg"]["ambiguo"] == t["warn"]

    @pytest.mark.parametrize("name", mt.THEME_ORDER)
    def test_all_values_are_valid_hex(self, name):
        p = mt.CustomTkinterThemeAdapter.palette(name)
        for key, value in p.items():
            if key in ("state_bg", "state_fg", "chart_colors"):
                continue
            assert _is_hex(value), f"{key} no es un color hex: {value!r}"
        for key, value in p["state_bg"].items():
            assert _is_hex(value), f"state_bg.{key} no es un color hex: {value!r}"
        for key, value in p["state_fg"].items():
            assert _is_hex(value), f"state_fg.{key} no es un color hex: {value!r}"


class TestFontScale:
    def test_reference_and_range(self):
        assert mt.SCALE_REFERENCE_WIDTH == 1920
        assert mt.SCALE_RANGE == (0.82, 1.35)

    @pytest.mark.parametrize(
        "width,expected",
        [
            (1920, 1.0),    # referencia → escala 1.0
            (1560, 0.82),   # 1560/1920=0.8125 → clamp al mínimo 0.82
            (1000, 0.82),   # por debajo del mínimo → clamp
            (800, 0.82),    # límite inferior del clamp
            (3000, 1.35),   # por encima del máximo → clamp
            (2592, 1.35),   # límite superior exacto del clamp
        ],
    )
    def test_parity_with_metatag_v8_formula(self, width, expected):
        # La fórmula exacta de metatag_v8: max(0.82, min(1.35, sw / 1920)).
        assert mt.compute_font_scale(width) == expected

    def test_scaled_size_uses_metatag_formula(self):
        # max(floor, int(base * scale)) — idéntico a set_font_scale.
        assert mt.scaled_size(15, 1.0, 8) == 15
        assert mt.scaled_size(15, 0.82, 8) == max(8, int(15 * 0.82))
        assert mt.scaled_size(9, 0.82, 7) == 7
        assert mt.scaled_size(18, 1.35, 10) == int(18 * 1.35)

    def test_font_specs_parity_scale_1(self):
        # Mismo output que set_font_scale(1.0) de metatag_v8.
        expected = {
            "TITLE":   ("Georgia", 15, "bold"),
            "H2":      ("Georgia", 11, "bold"),
            "LABEL":   ("Segoe UI", 9),
            "LABEL_B": ("Segoe UI", 9, "bold"),
            "BODY":    ("Segoe UI", 10),
            "MONO":    ("Consolas", 9),
            "CELL":    ("Segoe UI", 9),
            "HEAD":    ("Segoe UI", 9, "bold"),
            "TINY":    ("Segoe UI", 8),
        }
        assert mt.font_specs(1.0) == expected

    def test_font_specs_parity_scale_090(self):
        expected = {
            "TITLE":   ("Georgia", 13, "bold"),
            "H2":      ("Georgia", 9, "bold"),
            "LABEL":   ("Segoe UI", 8),
            "LABEL_B": ("Segoe UI", 8, "bold"),
            "BODY":    ("Segoe UI", 9),
            "MONO":    ("Consolas", 8),
            "CELL":    ("Segoe UI", 8),
            "HEAD":    ("Segoe UI", 8, "bold"),
            "TINY":    ("Segoe UI", 7),
        }
        assert mt.font_specs(0.9) == expected

    def test_font_specs_respect_floors(self):
        specs = mt.font_specs(0.82)
        assert specs["TITLE"][1] == max(8, int(15 * 0.82))
        assert specs["TINY"][1] == max(6, int(8 * 0.82))
        assert specs["HEAD"][1] == max(7, int(9 * 0.82))


class TestColorHelpers:
    def test_mix_endpoints(self):
        assert mt.mix("#000000", "#FFFFFF", 0.0) == "#000000"
        assert mt.mix("#000000", "#FFFFFF", 1.0) == "#FFFFFF"

    def test_mix_midpoint(self):
        assert mt.mix("#000000", "#FFFFFF", 0.5) == "#808080"

    def test_relative_luminance(self):
        assert mt.relative_luminance("#000000") == 0.0
        assert mt.relative_luminance("#FFFFFF") == pytest.approx(1.0, abs=0.01)
        assert mt.relative_luminance("#808080") < 0.3


class TestFitToScreen:
    def test_returns_ideal_when_screen_has_room(self):
        # 1920×1080 da espacio de sobra para el ideal 1200 → 1200.
        assert mt.fit_to_screen(1920, 1080, 800, 1200) == 1200

    def test_scales_down_to_fit_small_screen(self):
        # 1366×768 obliga a reducir por debajo del ideal (pero >= minsize).
        size = mt.fit_to_screen(1366, 768, 800, 1200)
        assert 800 <= size <= 1200

    def test_reduces_more_on_tiny_screen(self):
        size = mt.fit_to_screen(1024, 768, 600, 1200)
        assert 600 <= size < 1200

    def test_never_below_minsize(self):
        assert mt.fit_to_screen(800, 600, 800, 1200) >= 800
        assert mt.fit_to_screen(100, 100, 800, 1200) == 800


class TestMetatagV8Parity:
    """Migración de metatag_v8: debe consumir los mismos tokens sin cambios."""

    def test_v8_uses_same_themes(self):
        import metatag_v8 as v8
        assert v8.THEMES == mt.THEMES
        assert v8.CURRENT_THEME == mt.DEFAULT_THEME
        assert v8.C == mt.THEMES[mt.DEFAULT_THEME]

    def test_v8_font_engine_identical(self):
        import metatag_v8 as v8
        for scale in (1.0, 0.9, 0.82, 1.35, 0.96):
            v8.set_font_scale(scale)
            assert v8.FONTS == mt.font_specs(scale), f"escala {scale}"

    def test_v8_theme_icons_identical(self):
        import metatag_v8 as v8
        assert v8.THEME_ICONS == mt.THEME_ICONS
