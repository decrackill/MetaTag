"""Tests del layout con CTkScrollableFrame y scroll aislado.

Verifica:
  - CTkScrollableFrame con scroll principal
  - PreviewTable dentro del scrollable frame
  - PreviewTable height adaptativa (crece con contenido)
  - Scroll de PreviewTable aislado (devuelve "break")
  - _on_frame_configure maneja resize correctamente
  - Pool se reconstruye cuando es necesario

Ejecutar:
    env XMODIFIERS="@im=none" .venv/bin/python -m pytest tests/test_scroll_layout.py -v
"""
import os
import sys
import tkinter as tk

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import renombrar_fotos_gui as mod
from renombrar_fotos_gui import MainView, PreviewTable, AppController
from metatag_responsive import PROFILE


def make_pairs(n: int):
    pairs = []
    for i in range(n):
        orig = f"{i+1:04d}_FOTO_{i+1:04d}.jpg"
        new = f"{i+1:04d}_FOTO_{i+1:04d}_v2.jpg"
        pairs.append((orig, new, None, False, "ok"))
    return pairs


@pytest.fixture
def root():
    r = tk.Tk()
    r.geometry("1200x800+0+0")
    mod._init_fonts(r.winfo_screenwidth())
    r.update_idletasks()
    yield r
    r.destroy()


@pytest.fixture
def ctrl(root):
    c = AppController()
    yield c
    try:
        c._view.destroy()
    except Exception:
        pass


# ── Layout structure ──────────────────────────────────────────────
class TestLayoutStructure:
    def test_scroll_frame_exists(self, ctrl):
        view = ctrl._view
        assert hasattr(view, "_scroll"), "MainView debe tener _scroll (CTkScrollableFrame)"

    def test_scroll_is_scrollable(self, ctrl):
        import customtkinter as ctk
        view = ctrl._view
        assert isinstance(view._scroll, ctk.CTkScrollableFrame)

    def test_preview_in_scroll(self, ctrl):
        view = ctrl._view
        preview_str = str(view._preview)
        scroll_children = [str(w) for w in view._scroll.winfo_children()]
        assert any(preview_str in s for s in scroll_children), \
            "PreviewTable debe estar dentro de _scroll"

    def test_all_sections_in_scroll(self, ctrl):
        view = ctrl._view
        scroll_children = [str(w) for w in view._scroll.winfo_children()]
        assert len(scroll_children) >= 10, \
            f"Scroll debe contener todas las secciones, tiene {len(scroll_children)}"


# ── PreviewTable height ──────────────────────────────────────────
class TestPreviewTableHeight:
    def test_height_from_screen(self, root):
        pt = PreviewTable(root)
        expected = max(PreviewTable._MIN_H,
                       min(PreviewTable._MAX_H,
                           int(root.winfo_screenheight() * 0.45)))
        assert pt._adaptive_height == expected
        pt.destroy()

    def test_height_grows_with_content(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        initial = pt._adaptive_height
        pt.render(make_pairs(100))
        root.update()
        assert pt._adaptive_height > initial, \
            f"Height should grow: {pt._adaptive_height} <= {initial}"
        pt.destroy()

    def test_height_capped_at_max(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        pt.render(make_pairs(10000))
        root.update()
        assert pt._adaptive_height == PreviewTable._MAX_H
        pt.destroy()

    def test_height_configure_updates(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        event = tk.Event()
        event.widget = pt
        event.width = 800
        event.height = 500
        pt._on_frame_configure(event)
        assert pt._adaptive_height == 500
        pt.destroy()

    def test_height_ignores_other_widget(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        original = pt._adaptive_height
        event = tk.Event()
        event.widget = root
        event.height = 500
        pt._on_frame_configure(event)
        assert pt._adaptive_height == original
        pt.destroy()


# ── Scroll isolation ──────────────────────────────────────────────
class TestScrollIsolation:
    def test_on_wheel_returns_break(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        pt.render(make_pairs(50))
        root.update()
        event = tk.Event()
        event.num = 5
        event.delta = 0
        result = pt._on_wheel(event)
        assert result == "break"
        pt.destroy()

    def test_on_wheel_button4(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        pt.render(make_pairs(50))
        root.update()
        event = tk.Event()
        event.num = 4
        event.delta = 0
        result = pt._on_wheel(event)
        assert result == "break"
        pt.destroy()

    def test_on_wheel_mousewheel(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        pt.render(make_pairs(50))
        root.update()
        event = tk.Event()
        event.num = 0
        event.delta = 120
        result = pt._on_wheel(event)
        assert result == "break"
        pt.destroy()

    def test_canvas_has_local_bindings(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        pt.render(make_pairs(10))
        root.update()
        assert pt._cv is not None
        cv_bindings = pt._cv.bind()
        for event in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            assert event in str(cv_bindings), \
                f"Canvas must have local binding for {event}"


# ── Pool and viewport ─────────────────────────────────────────────
class TestPoolAndViewport:
    def test_scrollregion_consistent_after_render(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        pt.render(make_pairs(100))
        root.update()
        sr = pt._cv.cget("scrollregion")
        parts = [int(x) for x in sr.split()]
        assert len(parts) == 4
        expected_h = 100 * pt.ROW_H
        assert parts[3] == expected_h, \
            f"scrollregion height={parts[3]}, expected={expected_h}"
        pt.destroy()

    def test_scrollregion_after_resize(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        pt.render(make_pairs(100))
        root.update()
        event = tk.Event()
        event.widget = pt
        event.width = 800
        event.height = 300
        pt._on_frame_configure(event)
        root.update()
        sr = pt._cv.cget("scrollregion")
        parts = [int(x) for x in sr.split()]
        assert parts[3] == 100 * pt.ROW_H
        pt.destroy()

    def test_viewport_clamped_after_scroll(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        pt.render(make_pairs(20))
        root.update()
        pt._scroll_by(10 ** 9)
        root.update()
        offset = pt._pixel_offset()
        total = 20 * pt.ROW_H
        vh = pt._view_height()
        max_px = max(0, total - vh)
        assert offset <= max_px + pt.ROW_H
        pt.destroy()

    def test_pool_no_rebuild_on_small_change(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        pt.render(make_pairs(50))
        root.update()
        old_pool = len(pt._rows)
        view_h = pt._view_height()
        event = tk.Event()
        event.widget = pt
        event.width = 800
        event.height = view_h + 1
        pt._on_frame_configure(event)
        root.update()
        new_pool = len(pt._rows)
        assert new_pool == old_pool
        pt.destroy()


# ── Stress tests ──────────────────────────────────────────────────
class TestScrollStress:
    def test_many_renders_no_growth(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        pool_sizes = []
        for i in range(10):
            n = 20 + (i * 10)
            pt.render(make_pairs(n))
            root.update_idletasks()
            pool_sizes.append(len(pt._rows))
        max_pool = max(pool_sizes)
        assert max_pool < 100
        pt.destroy()

    def test_scroll_render_cycle(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        pt.render(make_pairs(100))
        root.update_idletasks()
        pt._scroll_by(500)
        root.update_idletasks()
        pt.render(make_pairs(50))
        root.update_idletasks()
        pt._scroll_by(-200)
        root.update_idletasks()
        assert len(pt._rows) > 0
        pt.destroy()

    def test_resize_during_scroll(self, root):
        pt = PreviewTable(root)
        pt.pack(fill="both", expand=True)
        root.update()
        pt.render(make_pairs(100))
        root.update_idletasks()
        pt._scroll_by(500)
        root.update_idletasks()
        event = tk.Event()
        event.widget = pt
        event.width = 600
        event.height = 150
        pt._on_frame_configure(event)
        root.update_idletasks()
        pt._scroll_by(-200)
        root.update_idletasks()
        assert len(pt._rows) > 0
        pt.destroy()
