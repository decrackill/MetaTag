"""
MetaTag v8.9 — Escritor de metadatos para cerámica arqueológica
Novedades v8.9: Lote por Orden (posición Excel→foto), ExcelGrid con viewport culling
para 300+ filas sin lag, barra de progreso en ventana dedicada.
Dependencias: pip install pandas openpyxl pillow piexif matplotlib
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os, json, re, threading, shutil, sys, subprocess
from typing import Union, Optional

if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd: ctypes.windll.user32.ShowWindow(hwnd, 0)
    except Exception: pass

from pathlib import Path

try:
    from PIL import Image, ImageTk, ImageOps
    import piexif, piexif.helper
    PIL_OK = True
except ImportError:
    PIL_OK = False

try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
    import numpy as np
    MATPLOTLIB_OK = True
except ImportError:
    MATPLOTLIB_OK = False

# ══════════════════════════════════════════════════════════════════
#  TEMAS DE COLOR Y MOTOR DE FUENTES DINÁMICO
# ══════════════════════════════════════════════════════════════════
THEMES = {
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

CURRENT_THEME = "Arqueológico (Oscuro Refinado)"
C = dict(THEMES[CURRENT_THEME])

FONTS = {}
def set_font_scale(scale):
    FONTS["TITLE"]   = ("Georgia",  max(8, int(15 * scale)), "bold")
    FONTS["H2"]      = ("Georgia",  max(8, int(11 * scale)), "bold")
    FONTS["LABEL"]   = ("Segoe UI", max(7, int(9  * scale)))
    FONTS["LABEL_B"] = ("Segoe UI", max(7, int(9  * scale)), "bold")
    FONTS["BODY"]    = ("Segoe UI", max(8, int(10 * scale)))
    FONTS["MONO"]    = ("Consolas", max(7, int(9  * scale)))
    FONTS["CELL"]    = ("Segoe UI", max(7, int(9  * scale)))
    FONTS["HEAD"]    = ("Segoe UI", max(7, int(9  * scale)), "bold")
    FONTS["TINY"]    = ("Segoe UI", max(6, int(8  * scale)))

set_font_scale(1.0)

IMG_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"}
META_GROUPS = {
    "Ubicacion":   ["Sitio", "Corte", "Cuadrante", "Unidad", "Nivel", "Profundidad Cm"],
    "Descripcion": ["Vista", "Parte", "Perfil", "Labio"],
    "Tecnica":     ["Tratamiento", "Tecnica", "Motivo"],
    "Notas":       ["Observaciones", "Excluido"],
}
META_GROUP_ORDER = ["Ubicacion", "Descripcion", "Tecnica", "Notas"]


# ══════════════════════════════════════════════════════════════════
#  EXPLORADOR DE ARCHIVOS NATIVO DEL SISTEMA
# ══════════════════════════════════════════════════════════════════
def _native_file_open(title="Seleccionar archivo", filetypes=None):
    if sys.platform == "linux":
        cmd = ["zenity", "--file-selection", "--title", title]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            if r.returncode == 0 and r.stdout.strip():
                return r.stdout.strip().split("\n")[0]
            if r.returncode == 1:
                return None
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
        cmd = ["kdialog", "--getopenfilename", ".", "--title", title]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            if r.returncode == 0 and r.stdout.strip():
                return r.stdout.strip().split("\n")[0]
            if r.returncode == 1:
                return None
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
    return filedialog.askopenfilename(title=title, filetypes=filetypes)


def _native_folder_open(title="Seleccionar carpeta"):
    if sys.platform == "linux":
        cmd = ["zenity", "--file-selection", "--directory", "--title", title]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            if r.returncode == 0 and r.stdout.strip():
                return r.stdout.strip().split("\n")[0]
            if r.returncode == 1:
                return None
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
        cmd = ["kdialog", "--getexistingdirectory", ".", "--title", title]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            if r.returncode == 0 and r.stdout.strip():
                return r.stdout.strip().split("\n")[0]
            if r.returncode == 1:
                return None
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
    return filedialog.askdirectory(title=title)


# ══════════════════════════════════════════════════════════════════
#  TABLA DE CELDAS INDIVIDUALES — OPTIMIZADA PARA 300+ FILAS
#  (viewport culling: solo pinta las filas visibles en pantalla)
# ══════════════════════════════════════════════════════════════════
class ExcelGrid(tk.Frame):
    def __init__(self, master, on_selection_change=None, on_row_click=None, app_ref=None, **kw):
        super().__init__(master, bg=C["surface"], **kw)
        self.on_selection_change = on_selection_change
        self.on_row_click        = on_row_click
        self.app_ref             = app_ref

        self.df = None  # type: Optional[pd.DataFrame]
        self.col_widths = []  # type: list
        self.selected_cells = set()
        self.hovered_row = None  # type: Optional[int]
        self.hidden_columns: set     = set()
        self._redraw_pending         = False   # evita redraws dobles

        self.unhide_bar = tk.Frame(self, bg=C["accent_pale"])
        self.unhide_lbl = tk.Label(self.unhide_bar, text="", bg=C["accent_pale"],
                                   fg=C["accent"], font=FONTS["TINY"], anchor="w")
        self.unhide_lbl.pack(side="left", padx=8, pady=2)
        self.unhide_btn = tk.Button(self.unhide_bar, text="Mostrar todas",
                                    bg=C["accent"], fg="#FFF5E8", font=FONTS["TINY"],
                                    relief="flat", bd=0, cursor="hand2",
                                    command=self._show_all_hidden)
        self.unhide_btn.pack(side="right", padx=8, pady=2)

        self.canvas = tk.Canvas(self, bg=C["surface"], highlightthickness=0, cursor="arrow")
        self.vsb = ttk.Scrollbar(self, orient="vertical",   command=self._on_vscroll)
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self._on_hscroll)
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        self.vsb.pack(side="right",  fill="y")
        self.hsb.pack(side="bottom", fill="x")
        self.canvas.pack(fill="both", expand=True)

        self.canvas.bind("<ButtonRelease-1>", self._on_click)
        self.canvas.bind("<Button-3>",        self._on_right_click)
        self.canvas.bind("<Motion>",          self._on_motion)
        self.canvas.bind("<Leave>",           self._on_leave)
        self.canvas.bind("<MouseWheel>",      self._on_wheel)
        self.canvas.bind("<Button-4>",        self._on_wheel)
        self.canvas.bind("<Button-5>",        self._on_wheel)
        self.canvas.bind("<Shift-Button-4>",  self._on_shift_wheel)
        self.canvas.bind("<Shift-Button-5>",  self._on_shift_wheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_shift_wheel)
        self.canvas.bind("<Configure>",       self._on_canvas_resize)

    # ── Alturas dinámicas ──
    @property
    def ROW_H(self):
        scale = getattr(self.app_ref, "current_scale", 1.0) if self.app_ref else 1.0
        return int(26 * scale)
    @property
    def HDR_H(self):
        scale = getattr(self.app_ref, "current_scale", 1.0) if self.app_ref else 1.0
        return int(28 * scale)

    # ── Scroll sincronizado + repintado ──
    def _on_vscroll(self, *args):
        self.canvas.yview(*args)
        self._schedule_redraw()

    def _on_hscroll(self, *args):
        self.canvas.xview(*args)
        self._schedule_redraw()

    def _on_wheel(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-3, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(3, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self._schedule_redraw()

    def _on_shift_wheel(self, event):
        if event.num == 4:
            self.canvas.xview_scroll(-3, "units")
        elif event.num == 5:
            self.canvas.xview_scroll(3, "units")
        else:
            self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
        self._schedule_redraw()

    def _on_canvas_resize(self, event):
        self._schedule_redraw()

    def _schedule_redraw(self):
        """Evita redraws múltiples en el mismo frame."""
        if not self._redraw_pending:
            self._redraw_pending = True
            self.canvas.after_idle(self._deferred_redraw)

    def _deferred_redraw(self):
        self._redraw_pending = False
        self.redraw()

    # ── Carga de datos ──
    def load(self, df: pd.DataFrame):
        self.df = df
        self.selected_cells.clear()
        self._calc_col_widths()
        self.canvas.yview_moveto(0)
        self.redraw()

    def _calc_col_widths(self):
        if self.df is None: return
        self.col_widths = []
        scale = getattr(self.app_ref, "current_scale", 1.0) if self.app_ref else 1.0
        min_cw, max_cw, pad = int(80*scale), int(250*scale), int(8*scale)
        for col in self.df.columns:
            max_len = len(str(col))
            for val in self.df[col].head(50):
                try:
                    max_len = max(max_len, len(str(val)))
                except Exception:
                    pass
            w = max(min_cw, min(max_cw, int(max_len * 7 * scale) + pad * 2))
            self.col_widths.append(w)

    @staticmethod
    def _sanitize_cell(val, max_len=200):
        try:
            if pd.isna(val):
                return ""
        except (ValueError, TypeError):
            pass
        try:
            s = str(val).strip()
        except Exception:
            return "N/A"
        if s.lower() in ("nan", "none", "nat"):
            return ""
        if len(s) > max_len:
            s = s[:max_len] + "..."
        return s

    # ── RENDER OPTIMIZADO — solo filas en viewport ──
    def redraw(self):
        if self.df is None: return
        self.canvas.configure(bg=C["surface"])
        self.canvas.delete("all")

        cols  = list(self.df.columns)
        nrows = len(self.df)
        pad   = int(8 * getattr(self.app_ref, "current_scale", 1.0))

        vis_cols = [(ci, col, cw) for ci, (col, cw) in enumerate(zip(cols, self.col_widths))
                    if col not in self.hidden_columns]
        total_w = sum(cw for _, _, cw in vis_cols) + 1
        total_h = self.HDR_H + nrows * self.ROW_H + 1
        self.canvas.configure(scrollregion=(0, 0, total_w, total_h))

        self._update_unhide_bar()

        # ── Calcular rango de filas visibles ──
        try:
            y1_frac, y2_frac = self.canvas.yview()
        except Exception:
            y1_frac, y2_frac = 0.0, 1.0

        vis_top = y1_frac * total_h
        vis_bot = y2_frac * total_h

        first_vis = max(0, int((vis_top - self.HDR_H) / self.ROW_H))
        last_vis  = min(nrows - 1, int((vis_bot - self.HDR_H) / self.ROW_H) + 1)

        # ── Cabecera (siempre visible) ──
        x = 0
        for ci, col, cw in vis_cols:
            is_col_sel = self._col_fully_selected(ci)
            bg = C["sel_bg"]    if is_col_sel else C["header_bg"]
            fg = C["sel_fg"]    if is_col_sel else C["header_fg"]
            self.canvas.create_rectangle(x, 0, x + cw, self.HDR_H,
                                         fill=bg, outline=C["grid_line"], width=1)
            self.canvas.create_text(x + pad, self.HDR_H // 2,
                                    text=col, anchor="w", font=FONTS["HEAD"], fill=fg)
            x += cw

        # ── Solo filas dentro del viewport ──
        for ri in range(first_vis, last_vis + 1):
            if ri >= nrows: break
            try:
                row = self.df.iloc[ri]
            except Exception:
                continue
            y   = self.HDR_H + ri * self.ROW_H
            row_bg = C["row_even"] if ri % 2 == 0 else C["row_odd"]
            x = 0
            for ci, col, cw in vis_cols:
                is_sel  = (ri, ci) in self.selected_cells
                col_sel = self._col_fully_selected(ci)

                if is_sel:              bg, fg = C["sel_bg"],    C["sel_fg"]
                elif col_sel:           bg, fg = C["col_sel"],   C["text"]
                elif ri == self.hovered_row: bg, fg = C["accent_pale"], C["text"]
                else:                   bg, fg = row_bg,         C["text"]

                self.canvas.create_rectangle(x, y, x + cw, y + self.ROW_H,
                                             fill=bg, outline=C["grid_line"], width=1)
                try:
                    cell_text = self._sanitize_cell(row[col])
                except Exception:
                    cell_text = "ERROR"
                self.canvas.create_text(x + pad, y + self.ROW_H // 2,
                                        text=cell_text, anchor="w",
                                        font=FONTS["CELL"], fill=fg)
                x += cw

    def _col_fully_selected(self, ci: int) -> bool:
        if self.df is None or len(self.df) == 0: return False
        return all((r, ci) in self.selected_cells for r in range(len(self.df)))

    def _hit(self, cx, cy):
        if self.df is None: return None, None
        cols = list(self.df.columns)
        x, ci = 0, None
        for i, cw in enumerate(self.col_widths):
            if cols[i] in self.hidden_columns: continue
            if x <= cx < x + cw: ci = i; break
            x += cw
        if ci is None: return None, None
        if cy < self.HDR_H: return None, ci
        ri = int((cy - self.HDR_H) // self.ROW_H)
        if ri < 0 or ri >= len(self.df): return None, ci
        return ri, ci

    def _on_click(self, event):
        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        ri, ci = self._hit(cx, cy)
        if ci is None: return
        if ri is None: self._toggle_column(ci)
        else:
            key = (ri, ci)
            if key in self.selected_cells: self.selected_cells.discard(key)
            else: self.selected_cells.add(key)
            if self.on_row_click: self.on_row_click(ri)
        self.redraw()
        if self.on_selection_change: self.on_selection_change()

    def _toggle_column(self, ci: int):
        nrows = len(self.df)
        if self._col_fully_selected(ci):
            for r in range(nrows): self.selected_cells.discard((r, ci))
        else:
            for r in range(nrows): self.selected_cells.add((r, ci))

    def _on_motion(self, event):
        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        ri, _ = self._hit(cx, cy)
        if ri != self.hovered_row:
            self.hovered_row = ri
            self._schedule_redraw()

    def _on_leave(self, event):
        self.hovered_row = None
        self._schedule_redraw()

    def _on_right_click(self, event):
        if self.df is None: return
        self._dismiss_col_menu()
        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        _, ci = self._hit(cx, cy)
        if ci is None: return
        col_name = self.df.columns[ci]
        self._col_menu = tk.Menu(self, tearoff=0, bg=C["surface"], fg=C["text"],
                                 font=FONTS["LABEL"], relief="flat")
        if col_name in self.hidden_columns:
            self._col_menu.add_command(label=f"Mostrar columna '{col_name}'",
                                       command=lambda c=col_name: self._toggle_col_visibility(c))
        else:
            self._col_menu.add_command(label=f"Ocultar columna '{col_name}'",
                                       command=lambda c=col_name: self._toggle_col_visibility(c))
        if self.hidden_columns:
            self._col_menu.add_separator()
            self._col_menu.add_command(label=f"Mostrar todas ({len(self.hidden_columns)} ocultas)",
                                       command=self._show_all_hidden)
        self._col_menu.post(event.x_root, event.y_root)
        root = self.winfo_toplevel()
        root.bind("<Button-1>", self._dismiss_col_menu, add=True)
        root.bind("<Button-3>", self._dismiss_col_menu, add=True)

    def _dismiss_col_menu(self, event=None):
        if hasattr(self, "_col_menu") and self._col_menu:
            try: self._col_menu.destroy()
            except Exception: pass
            self._col_menu = None
        try:
            root = self.winfo_toplevel()
            root.unbind("<Button-1>")
            root.unbind("<Button-3>")
        except Exception: pass

    def _toggle_col_visibility(self, col_name: str):
        if col_name in self.hidden_columns:
            self.hidden_columns.discard(col_name)
        else:
            self.hidden_columns.add(col_name)
            to_discard = {k for k in self.selected_cells if self.df.columns[k[1]] == col_name}
            self.selected_cells -= to_discard
        self.redraw()
        if self.on_selection_change: self.on_selection_change()

    def _show_all_hidden(self):
        self.hidden_columns.clear()
        self.redraw()
        if self.on_selection_change: self.on_selection_change()

    def _update_unhide_bar(self):
        if self.hidden_columns:
            names = ", ".join(sorted(self.hidden_columns))
            self.unhide_lbl.configure(text=f"Ocultas: {names}")
            self.unhide_bar.pack(fill="x", before=self.canvas)
        else:
            self.unhide_bar.pack_forget()

    def select_row(self, ri: int):
        if self.df is None: return
        cols = list(self.df.columns)
        for ci in range(len(self.col_widths)):
            if cols[ci] not in self.hidden_columns:
                self.selected_cells.add((ri, ci))
        self.redraw()
        if self.on_selection_change: self.on_selection_change()

    def clear_selection(self):
        self.selected_cells.clear()
        self.redraw()
        if self.on_selection_change: self.on_selection_change()

    def scroll_to_row(self, ri: int):
        if self.df is None: return
        canvas_h = self.canvas.winfo_height()
        if canvas_h <= 1: canvas_h = 400
        total_h = self.HDR_H + len(self.df) * self.ROW_H
        y_row   = self.HDR_H + ri * self.ROW_H
        if total_h <= canvas_h:
            self.canvas.yview_moveto(0)
        else:
            target = (y_row - canvas_h // 3) / total_h
            self.canvas.yview_moveto(max(0, min(target, 1.0)))
        self._schedule_redraw()

    def _get_val(self, val):
        if pd.isna(val): return ""
        s = str(val).strip()
        if s.lower() in ("nan", "none"): return ""
        return s

    def get_selected_metadata(self, img_col_idx):
        if self.df is None: return {}
        cols   = list(self.df.columns)
        result = {}
        omit_empty = True
        if self.app_ref:
            omit_var = getattr(self.app_ref, 'omit_empty_var', None)
            omit_empty = omit_var.get() if omit_var else True

        temp_result = {}
        for (ri, ci) in self.selected_cells:
            if ci == img_col_idx: continue
            if cols[ci] in self.hidden_columns: continue
            val = self._get_val(self.df.iloc[ri, ci])
            if omit_empty and not val: continue
            temp_result.setdefault(ri, {})[cols[ci]] = val

        if self.app_ref:
            locked = getattr(self.app_ref, 'locked_columns', None)
            if locked:
                for ri in temp_result.keys():
                    for col in locked:
                        if col in cols and col not in self.hidden_columns:
                            ci  = cols.index(col)
                            val = self._get_val(self.df.iloc[ri, ci])
                            if omit_empty and not val: continue
                            temp_result[ri][col] = val

        for ri, row_data in temp_result.items():
            result[ri] = {col: row_data[col] for col in cols
                          if col in row_data and col not in self.hidden_columns}
        return result

    def get_row_metadata(self, ri, img_col_idx):
        if self.df is None: return {}
        cols       = list(self.df.columns)
        omit_empty = True
        if self.app_ref:
            omit_var = getattr(self.app_ref, 'omit_empty_var', None)
            omit_empty = omit_var.get() if omit_var else True
        temp_meta  = {}

        for ci in range(len(cols)):
            if ci == img_col_idx: continue
            if cols[ci] in self.hidden_columns: continue
            if (ri, ci) in self.selected_cells:
                val = self._get_val(self.df.iloc[ri, ci])
                if not omit_empty or val: temp_meta[cols[ci]] = val

        if self.app_ref:
            locked = getattr(self.app_ref, 'locked_columns', None)
            if locked:
                for col in locked:
                    if col in cols and col not in self.hidden_columns:
                        ci  = cols.index(col)
                        val = self._get_val(self.df.iloc[ri, ci])
                        if not omit_empty or val: temp_meta[col] = val

        return {col: temp_meta[col] for col in cols if col in temp_meta}


# ══════════════════════════════════════════════════════════════════
#  EXPLORADOR DE IMÁGENES
# ══════════════════════════════════════════════════════════════════
class ImageBrowser(tk.Frame):
    def __init__(self, master, on_select=None, **kw):
        super().__init__(master, bg=C["panel"], **kw)
        self.on_select = on_select
        self.folder    = None
        self.img_files = []

        top = tk.Frame(self, bg=C["panel"])
        top.pack(fill="x", padx=8, pady=(8, 4))
        tk.Label(top, text="Explorador de imágenes", bg=C["panel"],
                 fg=C["header_bg"], font=FONTS["H2"]).pack(side="left")

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._filter())
        sf = tk.Frame(self, bg=C["panel"])
        sf.pack(fill="x", padx=8, pady=(0, 4))
        tk.Label(sf, text="🔍", bg=C["panel"], fg=C["text3"], font=FONTS["LABEL"]).pack(side="left")
        self.search_entry = tk.Entry(sf, textvariable=self.search_var,
                                     bg=C["surface"], fg=C["text"],
                                     relief="solid", bd=1, font=FONTS["LABEL"])
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(4, 0))

        lf = tk.Frame(self, bg=C["panel"])
        lf.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        vsb = ttk.Scrollbar(lf, orient="vertical")
        self.listbox = tk.Listbox(lf, bg=C["surface"], fg=C["text"],
                                  selectbackground=C["sel_bg"], selectforeground=C["sel_fg"],
                                  font=FONTS["CELL"], relief="flat", borderwidth=0,
                                  activestyle="none", yscrollcommand=vsb.set)
        vsb.configure(command=self.listbox.yview)
        vsb.pack(side="right", fill="y")
        self.listbox.pack(fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        self.info_lbl = tk.Label(self, text="Sin carpeta seleccionada",
                                 bg=C["panel"], fg=C["text3"], font=FONTS["TINY"])
        self.info_lbl.pack(anchor="w", padx=8, pady=(0, 4))

    def load_folder(self, folder: str):
        self.folder    = folder
        self.img_files = sorted(
            [f for f in Path(folder).iterdir() if f.suffix.lower() in IMG_EXTS],
            key=lambda p: p.name.lower()
        )
        self._excel_count = None   # reset: ya no hay separación Excel/huérfanas
        self._filter()
        self.info_lbl.configure(text=f"{len(self.img_files)} imágenes en carpeta")

    def _filter(self):
        q = self.search_var.get().lower()
        self.listbox.delete(0, "end")
        self._filtered = [f for f in self.img_files if q in f.name.lower()]
        orphans = getattr(self, "_orphan_files", set())
        for i, f in enumerate(self._filtered):
            if f.name in orphans:
                self.listbox.insert("end", f"  {f.name}  [sin fila]")
                self.listbox.itemconfigure(i, foreground="#888888")
            else:
                self.listbox.insert("end", f.name)

    def _on_select(self, event):
        sel = self.listbox.curselection()
        if not sel: return
        path = self._filtered[sel[0]]
        if self.on_select: self.on_select(str(path))

    def highlight(self, filename: str):
        for i, f in enumerate(getattr(self, "_filtered", [])):
            if f.name == filename:
                self.listbox.selection_clear(0, "end")
                self.listbox.selection_set(i)
                self.listbox.see(i)
                break


# ══════════════════════════════════════════════════════════════════
#  APP PRINCIPAL
# ══════════════════════════════════════════════════════════════════
class MetaTagApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MetaTag v8.9 — Edición Arqueológica Avanzada")

        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        # current_scale antes quedaba fijo en 1.0 sin importar la pantalla,
        # así que en un laptop pequeño (ej. 1366×768) la UI se veía grande/
        # apretada, y en un monitor grande (ej. 2560×1440+) se veía chica.
        # Se calcula contra una resolución de referencia de escritorio
        # (1920×1080) y se limita entre 0.82 y 1.35 para no deformar la UI
        # en extremos (netbooks muy chicos o monitores 4K muy grandes).
        _ref_w = 1920
        self.current_scale = max(0.82, min(1.35, sw / _ref_w))
        set_font_scale(self.current_scale)

        win_w = int(sw * 0.85)
        win_h = int(sh * 0.85)
        self.geometry(f"{win_w}x{win_h}+{(sw-win_w)//2}+{(sh-win_h)//2}")
        self.minsize(860, 520)
        self.configure(bg=C["bg"])

        self.original_df    = None
        self.df             = None
        self.csv_path_var   = tk.StringVar()
        self.img_folder_var = tk.StringVar()
        self.img_col_var    = tk.StringVar()
        self.theme_var      = tk.StringVar(value=CURRENT_THEME)
        self.process_mode   = tk.StringVar(value="Inteligente")
        self.current_img    = None
        self.current_row    = None
        self.progress_var   = tk.DoubleVar()
        self.status_var     = tk.StringVar(value="Carga un archivo Excel o CSV para comenzar.")
        self._img_cache     = {}
        self.locked_columns = set()
        self.omit_empty_var          = tk.BooleanVar(value=True)
        self.meta_mode_organized     = tk.BooleanVar(value=False)
        self.meta_mode_organized.trace_add("write", lambda *_: self._update_meta_preview())

        if getattr(sys, "frozen", False): base = Path(sys.executable).parent
        else: base = Path(__file__).parent
        self.output_base   = base
        self.output_folder = base / "Metadatos_Escritos"

        self._build_styles()
        self._load_config_pre_build()
        self._build_ui()
        self._load_config_post_build()
        self._toggle_mode()

        self.bind("<Up>",          self._nav_up)
        self.bind("<Down>",        self._nav_down)
        self.bind("<Control-o>",   lambda e: self._browse_csv())
        self.bind("<Control-O>",   lambda e: self._browse_csv())
        self.bind("<Control-b>",   lambda e: self._browse_folder())
        self.bind("<Control-B>",   lambda e: self._browse_folder())
        self.bind("<Control-e>",   lambda e: self._start_processing())
        self.bind("<Control-E>",   lambda e: self._start_processing())
        self.bind("<Control-i>",   lambda e: self._inject_manual())
        self.bind("<Control-I>",   lambda e: self._inject_manual())
        self.bind("<Control-g>",   lambda e: self._show_stats())
        self.bind("<Control-G>",   lambda e: self._show_stats())
        self.bind("<Control-l>",   lambda e: self._clear_selection())
        self.bind("<Control-L>",   lambda e: self._clear_selection())
        self.bind("<Control-f>",   self._focus_search)
        self.bind("<Control-F>",   self._focus_search)
        self.bind("<Control-r>",   lambda e: self._sync_images_to_excel())
        self.bind("<Control-R>",   lambda e: self._sync_images_to_excel())
        self.bind("<Control-Shift-b>",   lambda e: self._batch_write_by_order())
        self.bind("<Control-Shift-B>",   lambda e: self._batch_write_by_order())
        self.bind("<Control-q>",   lambda e: self._on_close())
        self.bind("<Control-Q>",   lambda e: self._on_close())
        self.bind("<Escape>",      self._on_escape)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        self._save_config()
        self.destroy()

    def _focus_search(self, event=None):
        if hasattr(self, "filter_entry"):
            self.filter_entry.focus_set()
            self.filter_entry.select_range(0, "end")

    def _on_escape(self, event=None):
        if hasattr(self, "loupe_window") and self.loupe_window.winfo_exists():
            self.loupe_window.destroy()
            return
        for w in self.winfo_toplevel().winfo_children():
            if isinstance(w, tk.Toplevel) and w.winfo_exists():
                w.destroy()
                return
        self._clear_selection()

    def _show_shortcuts(self):
        win = tk.Toplevel(self)
        win.title("Atajos de teclado")
        win.configure(bg=C["bg"])
        win.geometry(f"{int(420*self.current_scale)}x{int(520*self.current_scale)}")
        win.resizable(False, False)
        win.attributes("-topmost", True)

        hdr = tk.Frame(win, bg=C["header_bg"])
        hdr.pack(fill="x")
        tk.Label(hdr, text="  ⌨  Atajos de teclado", bg=C["header_bg"],
                 fg=C["header_fg"], font=FONTS["H2"]).pack(side="left", pady=10, padx=8)

        canvas = tk.Canvas(win, bg=C["surface"], highlightthickness=0)
        vsb = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=C["surface"])
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        vsb.pack(side="right", fill="y", padx=(0, 8), pady=8)

        shortcuts = [
            ("Navegación", [
                ("↑ / ↓", "Navegar entre filas"),
            ]),
            ("Archivos", [
                ("Ctrl + O", "Abrir archivo Excel/CSV"),
                ("Ctrl + B", "Abrir carpeta de imágenes"),
            ]),
            ("Edición", [
                ("Ctrl + E", "Escribir metadatos (Modo Inteligente)"),
                ("Ctrl + I", "Inyectar en foto actual (Modo Libre)"),
                ("Ctrl + L", "Limpiar selección"),
            ]),
            ("Herramientas", [
                ("Ctrl + G", "Ver estadísticas / gráficos"),
                ("Ctrl + R", "Reordenar imágenes según Excel"),
                ("Ctrl + Shift + B", "Lote por orden (Excel→Fotos)"),
            ]),
            ("Búsqueda", [
                ("Ctrl + F", "Enfocar campo de búsqueda"),
            ]),
            ("General", [
                ("Escape", "Cerrar ventana / limpiar selección"),
                ("Ctrl + Q", "Salir del programa"),
            ]),
        ]

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def _on_mousewheel_linux(event):
            if event.num == 4: canvas.yview_scroll(-1, "units")
            elif event.num == 5: canvas.yview_scroll(1, "units")
        def _bind_scroll_recursive(widget):
            widget.bind("<MouseWheel>", _on_mousewheel, add=True)
            widget.bind("<Button-4>", _on_mousewheel_linux, add=True)
            widget.bind("<Button-5>", _on_mousewheel_linux, add=True)
            for child in widget.winfo_children():
                _bind_scroll_recursive(child)

        for group_name, items in shortcuts:
            tk.Label(inner, text=group_name, bg=C["surface"], fg=C["accent"],
                     font=FONTS["LABEL_B"], anchor="w").pack(fill="x", padx=10, pady=(10, 2))
            for key, desc in items:
                row = tk.Frame(inner, bg=C["surface"])
                row.pack(fill="x", padx=10, pady=1)
                tk.Label(row, text=key, bg=C["accent_pale"], fg=C["accent"],
                         font=FONTS["MONO"], width=18, anchor="w",
                         padx=6, pady=3).pack(side="left")
                tk.Label(row, text=desc, bg=C["surface"], fg=C["text"],
                         font=FONTS["LABEL"], anchor="w").pack(side="left", padx=6)

        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win.after(100, lambda: _bind_scroll_recursive(inner))

    # ─────────────────────────────────────────────────────────────
    #  NAVEGACIÓN CON TECLADO
    # ─────────────────────────────────────────────────────────────
    def _nav_up(self, event):
        if isinstance(event.widget, tk.Entry): return
        if self.grid.df is None or len(self.grid.df) == 0: return
        if self.current_row is None: self.current_row = 0
        elif self.current_row > 0: self.current_row -= 1
        self._auto_select_nav(fast=True)

    def _nav_down(self, event):
        if isinstance(event.widget, tk.Entry): return
        if self.grid.df is None or len(self.grid.df) == 0: return
        if self.current_row is None: self.current_row = 0
        elif self.current_row < len(self.grid.df) - 1:
            prev = self.current_row
            self.current_row += 1
            self._apply_autocompletion(prev, self.current_row)
        self._auto_select_nav(fast=True)

    def _auto_select_nav(self, fast=False):
        self.grid.clear_selection()
        self.grid.select_row(self.current_row)
        self.grid.scroll_to_row(self.current_row)
        if fast:
            self._update_meta_preview()
            if hasattr(self, "_nav_img_timer"): self.after_cancel(self._nav_img_timer)
            self._nav_img_timer = self.after(250, self._nav_load_image)
        else:
            self._on_row_click(self.current_row)

    def _nav_load_image(self):
        if self.current_row is None or self.grid.df is None: return
        folder = self.img_folder_var.get()
        if not folder: return
        row = self.grid.df.iloc[self.current_row]
        for col in self.grid.df.columns:
            img_name = str(row[col]).strip()
            if img_name and img_name.lower() not in ("nan", "none", ""):
                img_path = self._find_image(img_name, folder)
                if img_path:
                    self._load_image(img_path)
                    self.browser.highlight(Path(img_path).name)
                    break

    def _apply_autocompletion(self, prev_row_idx: int, new_row_idx: int):
        if not self.locked_columns or self.df is None: return
        cambios = False
        for col in self.locked_columns:
            if col in self.df.columns:
                prev_val = self.df.at[prev_row_idx, col]
                self.df.at[new_row_idx, col] = prev_val
                if self.original_df is not None:
                    idx_real = self.df.index[new_row_idx]
                    self.original_df.at[idx_real, col] = prev_val
                cambios = True
        if cambios: self.grid.redraw()

    # ─────────────────────────────────────────────────────────────
    #  ESTADÍSTICAS / GRÁFICOS
    # ─────────────────────────────────────────────────────────────
    def _show_stats(self):
        if self.df is None or len(self.df) == 0:
            return messagebox.showinfo("Sin datos", "Carga un archivo de datos primero.")
        if not MATPLOTLIB_OK:
            return messagebox.showerror("Falta Librería", "pip install matplotlib")

        plt.rcParams['lines.antialiased']  = True
        plt.rcParams['patch.antialiased']  = True
        plt.rcParams['path.simplify']      = True

        S_BG           = C["bg"]
        S_CARD         = C["surface"]
        S_BORDER       = C["border"]
        S_ACCENT       = C["accent"]
        S_ACCENT_LIGHT = C.get("accent_light", C["accent"])
        S_TEXT         = C["text"]
        S_TEXT_MUTE    = C["text2"]
        S_CHART_COLORS = C.get("chart_colors",
                                ["#A67C52","#D4A574","#7A4F2D","#3D1F0A","#B1A28F","#5C3518"])

        win = tk.Toplevel(self)
        win.title("📊 Análisis Cuantitativo Arqueológico (v8.9)")
        win.geometry(f"{int(1100*self.current_scale)}x{int(720*self.current_scale)}")
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        ww = min(int(1200 * self.current_scale), sw - 40)
        wh = min(int(720  * self.current_scale), sh - 60)
        win.geometry(f"{ww}x{wh}")
        win.minsize(800, 500)
        win.configure(bg=S_BG)

        img_c  = self.img_col_var.get()
        validas = [c for c in self.df.columns if 1 < self.df[c].nunique() < 40 and c != img_c]
        if not validas:
            return messagebox.showinfo("Sin datos", "No hay categorías repetitivas para graficar.")

        top_frame = tk.Frame(win, bg=S_BG, pady=15, padx=20)
        top_frame.pack(fill="x")

        # ── Selector custom con popup ──
        def make_selector(parent, label_text, icon, options, default, on_change):
            frame = tk.Frame(parent, bg=S_BG)
            tk.Label(frame, text=f"{icon} {label_text}", bg=S_BG, fg=S_TEXT_MUTE,
                     font=FONTS["TINY"]).pack(anchor="w")

            current = tk.StringVar(value=default)

            btn_frame = tk.Frame(frame, bg=S_BORDER, padx=1, pady=1)
            btn_frame.pack(fill="x", pady=(4, 0))

            inner = tk.Frame(btn_frame, bg=C["surface"], cursor="hand2")
            inner.pack(fill="x")

            display = tk.Label(inner, text=default, bg=C["surface"], fg=S_TEXT,
                               font=FONTS["BODY"], anchor="w", padx=12, pady=7)
            display.pack(side="left", fill="x", expand=True)

            arrow = tk.Label(inner, text="▾", bg=C["surface"], fg=C["accent"],
                             font=FONTS["LABEL_B"], padx=8)
            arrow.pack(side="right")

            popup_ref = [None]

            def toggle(e=None):
                if popup_ref[0] and popup_ref[0].winfo_exists():
                    popup_ref[0].destroy()
                    popup_ref[0] = None
                    return

                m = tk.Toplevel(win)
                popup_ref[0] = m
                m.withdraw()
                m.overrideredirect(True)
                win.update_idletasks()
                x = btn_frame.winfo_rootx()
                y = btn_frame.winfo_rooty() + btn_frame.winfo_height() + 2
                min_w  = btn_frame.winfo_width()
                char_w = max((len(opt) for opt in options), default=10)
                popup_w = max(min_w, char_w * 9 + 32)
                popup_h = len(options) * 34 + 8
                screen_w = win.winfo_screenwidth()
                if x + popup_w > screen_w - 10:
                    x = screen_w - popup_w - 10
                m.geometry(f"{popup_w}x{popup_h}+{x}+{y}")
                m.configure(bg=S_BORDER)

                container = tk.Frame(m, bg=C["surface"])
                container.pack(fill="both", expand=True, padx=1, pady=1)

                for opt in options:
                    is_sel = (opt == current.get())
                    bg = S_ACCENT if is_sel else C["surface"]
                    fg = S_BG if is_sel else S_TEXT

                    row = tk.Frame(container, bg=bg, cursor="hand2")
                    row.pack(fill="x", padx=2, pady=1)

                    lbl = tk.Label(row, text=f"  {opt}", bg=bg, fg=fg,
                                   font=FONTS["BODY"], anchor="w", padx=8, pady=4)
                    lbl.pack(fill="x")

                    def select(val=opt, win_ref=m):
                        current.set(val)
                        display.configure(text=val)
                        on_change(val)
                        win_ref.destroy()
                        popup_ref[0] = None

                    for w in (row, lbl):
                        w.bind("<Button-1>", lambda e, _s=select: _s())
                        w.bind("<Enter>", lambda e, r=row, l=lbl, v=opt:
                               (r.configure(bg=S_ACCENT_LIGHT), l.configure(bg=S_ACCENT_LIGHT, fg=S_TEXT)))
                        w.bind("<Leave>", lambda e, r=row, l=lbl, v=opt:
                               (r.configure(bg=S_ACCENT if v == current.get() else C["surface"]),
                                l.configure(bg=S_ACCENT if v == current.get() else C["surface"],
                                            fg=S_BG if v == current.get() else S_TEXT)))

                m.deiconify()

            def _close_popup(e=None):
                if popup_ref[0] and popup_ref[0].winfo_exists():
                    popup_ref[0].destroy()
                    popup_ref[0] = None

            win.bind("<FocusOut>",  _close_popup, add="+")
            win.bind("<Unmap>",     _close_popup, add="+")
            win.bind("<Configure>", _close_popup, add="+")

            for w in (inner, display, arrow):
                w.bind("<Button-1>", toggle)

            return frame, current

        # ── Selectores ──
        def on_var_change(val):
            update_chart()

        def on_style_change(val):
            update_chart()

        CHART_OPTIONS = [
            "🍩 Dona HD",
            "🥧 Pastel Profesional",
            "📊 Barras Horizontal",
            "📈 Barras Vertical",
            "🎯 Lollipop",
        ]

        selector_var_frame, combo_var = make_selector(
            top_frame, "Variable a analizar", "📊", validas, validas[0], on_var_change)
        selector_style_frame, chart_type_var = make_selector(
            top_frame, "Estilo del Gráfico", "🎨",
            CHART_OPTIONS, "🍩 Dona HD", on_style_change)

        selector_var_frame.pack(side="left")
        selector_style_frame.pack(side="left", padx=(30, 0))

        def export_chart():
            # Carpeta inicial adaptada al sistema operativo
            if sys.platform == "win32":
                default_dir = os.path.join(os.path.expanduser("~"), "Pictures")
            elif sys.platform == "darwin":
                default_dir = os.path.join(os.path.expanduser("~"), "Pictures")
            else:  # Linux
                default_dir = os.path.join(os.path.expanduser("~"), "Imágenes")
                if not os.path.isdir(default_dir):
                    default_dir = os.path.expanduser("~")
            if not os.path.isdir(default_dir):
                default_dir = os.path.expanduser("~")

            path = filedialog.asksaveasfilename(
                title="Exportar gráfica",
                initialdir=default_dir,
                defaultextension=".png",
                filetypes=[("Imagen PNG (alta calidad)", "*.png"),
                           ("PDF vectorial", "*.pdf"),
                           ("Imagen JPEG", "*.jpg")],
                initialfile=f"grafica_{combo_var.get()}")
            if not path:
                return

            # En Linux, algunos diálogos GTK no respetan defaultextension.
            # Si el usuario no escribió extensión, se la agregamos manualmente.
            if not os.path.splitext(path)[1]:
                path += ".png"
            try:
                fig.savefig(path, dpi=300, facecolor=fig.get_facecolor(),
                            bbox_inches="tight", pad_inches=0.3)
                messagebox.showinfo("Exportado", f"Gráfica guardada en:\n{path}")
            except Exception as e:
                messagebox.showerror("Error al exportar", str(e))

        export_btn_frame = tk.Frame(top_frame, bg=S_BG)
        export_btn_frame.pack(side="right", anchor="s")
        tk.Button(export_btn_frame, text="⬇ Exportar gráfica", command=export_chart,
                  bg=C["accent"], fg=S_BG, font=FONTS["LABEL_B"], relief="flat",
                  cursor="hand2", padx=14, pady=6).pack()

        body_paned = tk.PanedWindow(win, orient="horizontal",
                                    bg=S_BORDER, sashwidth=4, bd=0)
        body_paned.pack(fill="both", expand=True, padx=20, pady=15)

        info_frame = tk.Frame(body_paned, bg=S_CARD,
                              highlightbackground=S_BORDER, highlightthickness=1)
        body_paned.add(info_frame, minsize=int(280*self.current_scale))
        info_frame.pack_propagate(False)

        header_insights = tk.Frame(info_frame, bg=C["header_bg"], pady=12)
        header_insights.pack(fill="x")
        tk.Label(header_insights, text="💡 Insights Arqueológicos",
                 bg=C["header_bg"], fg=C["header_fg"], font=FONTS["H2"]).pack()
        insight_text = tk.Text(info_frame, bg=S_CARD, fg=S_TEXT, font=FONTS["BODY"],
                               wrap="word", relief="flat", padx=18, pady=18)
        insight_text.pack(fill="both", expand=True)
        mode_footer = tk.Frame(info_frame, bg=C["accent_pale"],
                               highlightbackground=S_BORDER, highlightthickness=1)
        mode_footer.pack(fill="x", side="bottom")
        tk.Label(mode_footer,
                 text="👆 Modo Exploración Visual Activo:\nHaz clic en una barra o sector para filtrar automáticamente tu tabla.",
                 bg=C["accent_pale"], fg=S_ACCENT, font=FONTS["LABEL_B"],
                 justify="center", wraplength=int(250*self.current_scale), pady=15).pack()

        chart_frame = tk.Frame(body_paned, bg=S_BG)
        body_paned.add(chart_frame, minsize=int(500*self.current_scale))
        try:
            _dpi_screen = win.winfo_fpixels("1i")
        except Exception:
            _dpi_screen = 96
        _fig_dpi = min(160, max(100, int(_dpi_screen * 1.25)))
        fig = Figure(figsize=(8, 6), dpi=_fig_dpi, facecolor=S_BG)
        canvas_widget = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas_widget.get_tk_widget().pack(fill="both", expand=True)

        def update_chart(*args):
            col   = combo_var.get()
            ctype = chart_type_var.get()
            if not isinstance(col, str) or col not in self.df.columns:
                return
            data   = self.df[col].replace("", pd.NA).dropna()
            counts = data.value_counts().sort_values(ascending=False)
            fig.clear()
            fig.patch.set_facecolor(S_BG)
            total_count = len(data)
            n_cats = len(counts)
            _small_items_for_insight = []
            LABEL_THRESHOLD_PCT = 1.0

            if total_count == 0:
                ax = fig.add_subplot(111)
                ax.set_facecolor(S_BG)
                ax.text(0.5, 0.5, "No hay datos suficientes.",
                        ha="center", va="center", color=S_TEXT, fontsize=12)

            # ── PIE / DONA ────────────────────────────────────────────
            elif "Dona" in ctype or "Pastel" in ctype:
                ax = fig.add_axes([0.22, 0.08, 0.56, 0.84])
                ax.set_facecolor(S_BG)
                counts_sorted = counts.sort_values(ascending=True)
                wedge_w = 0.45 if "Dona" in ctype else 1.0
                colors_cycle = (S_CHART_COLORS * (n_cats // len(S_CHART_COLORS) + 1))

                pie_result = ax.pie(
                    counts_sorted,
                    labels=None,
                    autopct=None,
                    colors=colors_cycle[:n_cats],
                    wedgeprops=dict(width=wedge_w, edgecolor=S_BG, linewidth=1.4,
                                    antialiased=True),
                    startangle=90)
                wedges = pie_result[0]

                for i, p in enumerate(wedges):
                    pct = counts_sorted.values[i] / total_count * 100
                    p._metatag_valor = f"{int(counts_sorted.values[i])} piezas ({pct:.1f}%)"
                    p._metatag_label = counts_sorted.index[i]

                LABEL_THRESHOLD_PCT = 1.0

                left_items  = []
                right_items = []
                small_items = []

                for i, p in enumerate(wedges):
                    ang = (p.theta2 - p.theta1) / 2.0 + p.theta1
                    if ang % 90 == 0:
                        ang += 0.5
                    pct = counts_sorted.values[i] / total_count * 100
                    nombre = counts_sorted.index[i]
                    n_abs  = int(counts_sorted.values[i])

                    if pct < LABEL_THRESHOLD_PCT:
                        small_items.append((nombre, n_abs, pct))
                        continue

                    ys_ang = np.sin(np.deg2rad(ang))
                    xs_ang = np.cos(np.deg2rad(ang))
                    label  = f"{nombre}  {pct:.1f}%"
                    if xs_ang < 0:
                        left_items.append((ys_ang, label, xs_ang, i, pct))
                    else:
                        right_items.append((ys_ang, label, xs_ang, i, pct))

                def place_labels_clean(items, side):
                    if not items:
                        return
                    ordered = sorted(items, key=lambda x: -x[0])
                    n = len(ordered)

                    y_top  =  1.15
                    y_bot  = -1.15
                    step   = (y_top - y_bot) / max(n, 1)
                    total_h = step * (n - 1)
                    y_start = (y_top + y_bot) / 2.0 + total_h / 2.0

                    ys_placed = [y_start - k * step for k in range(n)]

                    x_col = -1.80 if side == "left" else 1.80
                    ha    = "right" if side == "left" else "left"
                    bbox  = dict(boxstyle="round,pad=0.28",
                                 fc=S_CARD, ec=S_BORDER, lw=0.7, alpha=0.90)

                    for k, (y_orig, label, xs_ang, idx_w, pct) in enumerate(ordered):
                        y_lbl = ys_placed[k]
                        ang_mid = ((wedges[idx_w].theta2 - wedges[idx_w].theta1)
                                   / 2.0 + wedges[idx_w].theta1)
                        r_arrow = 1.02
                        xa = r_arrow * np.cos(np.deg2rad(ang_mid))
                        ya = r_arrow * np.sin(np.deg2rad(ang_mid))
                        ax.annotate(
                            label,
                            xy=(xa, ya),
                            xytext=(x_col, y_lbl),
                            horizontalalignment=ha,
                            fontsize=8,
                            color=S_TEXT,
                            arrowprops=dict(
                                arrowstyle="-",
                                color=S_TEXT_MUTE,
                                lw=1.1,
                                alpha=0.55,
                                shrinkA=0, shrinkB=4,
                                connectionstyle="arc3,rad=0.15",
                                capstyle="round"),
                            bbox=bbox,
                            va="center",
                            zorder=5,
                            annotation_clip=False)

                place_labels_clean(left_items,  "left")
                place_labels_clean(right_items, "right")

                tk_chart_title.configure(text=f"Distribución de {col}")

                if wedge_w < 1.0:
                    top_val = counts_sorted.index[-1]
                    top_pct = counts_sorted.max() / total_count * 100
                    ax.text(0, 0.10, f"TOTAL\n{total_count}\npiezas",
                            ha="center", va="center",
                            color=S_TEXT, fontsize=13, fontweight="bold")
                    ax.text(0, -0.30, f"Dominante:\n{top_val} ({top_pct:.1f}%)",
                            ha="center", va="center",
                            color=S_TEXT_MUTE, fontsize=8, fontstyle="italic")

                _small_items_for_insight = small_items

            # ── BARRAS HORIZONTAL ─────────────────────────────────────
            elif "Horizontal" in ctype:
                ax = fig.add_subplot(111)
                ax.set_facecolor(S_BG)
                counts_asc = counts.sort_values(ascending=True)
                bar_h = max(0.35, min(0.65, 4.0 / max(n_cats, 1)))
                colors_list = (S_CHART_COLORS * (n_cats // len(S_CHART_COLORS) + 1))
                bars = ax.barh(range(n_cats), counts_asc.values,
                               color=colors_list[:n_cats],
                               edgecolor=S_BG, linewidth=1.2, height=bar_h)
                ax.set_yticks(range(n_cats))
                ax.set_yticklabels(counts_asc.index, color=S_TEXT,
                                   fontsize=max(7, min(10, 120 // max(n_cats, 1))))
                ax.set_xlabel("Cantidad de piezas", color=S_TEXT_MUTE, fontsize=9)
                ax.tick_params(colors=S_TEXT, length=0)
                for sp in ("top", "right", "left"):
                    ax.spines[sp].set_visible(False)
                ax.spines["bottom"].set_color(S_BORDER)
                ax.xaxis.grid(True, linestyle=(0, (4, 4)), alpha=0.22,
                               linewidth=0.6, color=S_BORDER)
                ax.set_axisbelow(True)
                max_val = counts_asc.max()
                for bar, val in zip(bars, counts_asc.values):
                    pct = val / total_count * 100
                    ax.text(bar.get_width() + max_val * 0.015,
                            bar.get_y() + bar.get_height() / 2,
                            f"{val}  ({pct:.1f}%)",
                            va="center", ha="left",
                            color=S_TEXT,
                            fontsize=max(7, min(9, 100 // max(n_cats, 1))),
                            fontweight="bold")
                ax.set_xlim(0, max_val * 1.32)
                for bar, val in zip(bars, counts_asc.values):
                    pct = val / total_count * 100
                    bar._metatag_valor = f"{val} piezas ({pct:.1f}%)"
                    bar._metatag_label = counts_asc.index[list(counts_asc.values).index(val)]
                tk_chart_title.configure(text=f"Frecuencia: {col}")

            # ── BARRAS VERTICAL ───────────────────────────────────────
            elif "Vertical" in ctype:
                ax = fig.add_subplot(111)
                ax.set_facecolor(S_BG)
                colors_list = (S_CHART_COLORS * (n_cats // len(S_CHART_COLORS) + 1))
                bars = ax.bar(range(n_cats), counts.values,
                              color=colors_list[:n_cats],
                              edgecolor=S_BG, linewidth=1.2, width=0.65)
                rot = 30 if n_cats <= 8 else 45
                fs  = max(7, min(9, 80 // max(n_cats, 1)))
                ax.set_xticks(range(n_cats))
                ax.set_xticklabels(counts.index, color=S_TEXT,
                                   fontsize=fs, rotation=rot, ha="right")
                ax.set_ylabel("Cantidad de piezas", color=S_TEXT_MUTE, fontsize=9)
                ax.tick_params(colors=S_TEXT, length=0)
                for sp in ("top", "right"):
                    ax.spines[sp].set_visible(False)
                ax.spines["left"].set_color(S_BORDER)
                ax.spines["bottom"].set_color(S_BORDER)
                ax.yaxis.grid(True, linestyle=(0, (4, 4)), alpha=0.22,
                               linewidth=0.6, color=S_BORDER)
                ax.set_axisbelow(True)
                max_val = counts.max()
                for bar, val in zip(bars, counts.values):
                    pct = val / total_count * 100
                    ax.text(bar.get_x() + bar.get_width() / 2,
                            bar.get_height() + max_val * 0.012,
                            f"{val} ({pct:.1f}%)",
                            ha="center", va="bottom",
                            color=S_TEXT,
                            fontsize=max(6, min(8, 80 // max(n_cats, 1))),
                            fontweight="bold")
                ax.set_ylim(0, max_val * 1.28)
                for bar, val in zip(bars, counts.values):
                    pct = val / total_count * 100
                    bar._metatag_valor = f"{val} piezas ({pct:.1f}%)"
                    bar._metatag_label = counts.index[list(counts.values).index(val)]
                fig.subplots_adjust(bottom=0.18)
                tk_chart_title.configure(text=f"Frecuencia: {col}")

            # ── LOLLIPOP ──────────────────────────────────────────────
            elif "Lollipop" in ctype:
                ax = fig.add_subplot(111)
                ax.set_facecolor(S_BG)
                counts_asc = counts.sort_values(ascending=True)
                y_pos   = list(range(n_cats))
                max_val = counts_asc.max()
                colors_list = (S_CHART_COLORS * (n_cats // len(S_CHART_COLORS) + 1))
                for i, (y, val) in enumerate(zip(y_pos, counts_asc.values)):
                    color = colors_list[i]
                    ax.plot([0, val], [y, y], color=color, linewidth=1.8,
                            alpha=0.55, solid_capstyle="round")
                    ax.scatter(val, y, color=color, s=90, zorder=5,
                               edgecolors=S_BG, linewidth=1.2)
                    pct = val / total_count * 100
                    x_offset = max_val * 0.02
                    fs = max(7, min(9, 100 // max(n_cats, 1)))
                    ax.text(val + x_offset, y,
                            f"{val}  ({pct:.1f}%)",
                            va="center", ha="left",
                            color=S_TEXT, fontsize=fs)
                ax.set_yticks(y_pos)
                ax.set_yticklabels(counts_asc.index, color=S_TEXT,
                                   fontsize=max(7, min(10, 120 // max(n_cats, 1))))
                ax.tick_params(colors=S_TEXT, length=0)
                ax.set_xlabel("Cantidad de piezas", color=S_TEXT_MUTE, fontsize=9)
                for sp in ("top", "right", "left"):
                    ax.spines[sp].set_visible(False)
                ax.spines["bottom"].set_color(S_BORDER)
                ax.xaxis.grid(True, linestyle=(0, (4, 4)), alpha=0.22,
                               linewidth=0.6, color=S_BORDER)
                ax.set_axisbelow(True)
                ax.set_xlim(-max_val * 0.02, max_val * 1.30)
                tk_chart_title.configure(text=f"Distribución: {col}")

            try:
                if "Dona" in ctype or "Pastel" in ctype:
                    fig.subplots_adjust(left=0.02, right=0.98, top=0.96, bottom=0.04)
                else:
                    fig.subplots_adjust(left=0.08, right=0.95, top=0.95, bottom=0.10)
            except Exception:
                fig.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)

            canvas_widget.draw()
            canvas_widget.get_tk_widget().update_idletasks()

            # ── Panel de insights ─────────────────────────────────────
            try:
                insight_text.configure(state="normal")
                insight_text.delete("1.0", "end")
                if total_count > 0:
                    counts_desc = counts.sort_values(ascending=False)
                    top_val     = counts_desc.index[0]
                    top_count_v = int(counts_desc.iloc[0])
                    low_val     = counts_desc.index[-1]
                    low_count_v = int(counts_desc.iloc[-1])
                    pct_top     = top_count_v / total_count * 100
                    pct_low     = low_count_v / total_count * 100
                    diversidad  = len(counts_desc)

                    insight_text.tag_configure("h",
                        font=FONTS["LABEL_B"], foreground=S_ACCENT)
                    insight_text.tag_configure("li",
                        font=FONTS["BODY"], spacing1=4)
                    insight_text.tag_configure("small_h",
                        font=FONTS["LABEL_B"], foreground=S_TEXT_MUTE)
                    insight_text.tag_configure("small_row",
                        font=FONTS["BODY"], spacing1=3,
                        foreground=S_TEXT_MUTE)

                    insight_text.insert("end",
                        f"ANÁLISIS DE '{col.upper()}'\n\n", "h")
                    insight_text.insert("end",
                        f"Se analizaron {total_count} elementos con esta "
                        f"característica registrada.\n\n")
                    insight_text.insert("end",
                        "🥇 Categoría Dominante:\n", "li")
                    insight_text.insert("end",
                        f" ▸ '{top_val}' lidera con {top_count_v} piezas.\n", "li")
                    insight_text.insert("end",
                        f" ▸ Representa el {pct_top:.1f}% de la muestra.\n\n", "li")
                    if diversidad > 1:
                        insight_text.insert("end",
                            "📉 Categoría Minoritaria:\n", "li")
                        insight_text.insert("end",
                            f" ▸ '{low_val}' es la menos frecuente, "
                            f"con {low_count_v} apariciones ({pct_low:.1f}%).\n\n", "li")
                    insight_text.insert("end",
                        "⚖️ Diversidad Registrada:\n", "li")
                    insight_text.insert("end",
                        f" ▸ Existen {diversidad} variaciones únicas "
                        f"documentadas.\n", "li")

                    small = _small_items_for_insight
                    if small:
                        insight_text.insert("end",
                            f"\n\n🔎 Categorías menores "
                            f"(< {LABEL_THRESHOLD_PCT:.0f}%, sin etiqueta "
                            f"en el gráfico):\n", "small_h")
                        for nombre, n_abs, pct in sorted(small,
                                                         key=lambda x: -x[2]):
                            insight_text.insert("end",
                                f" ▸ {nombre}: "
                                f"{n_abs} piezas ({pct:.1f}%)\n",
                                "small_row")

                else:
                    insight_text.insert("end",
                        "Carga datos para visualizar el análisis.")
                insight_text.configure(state="disabled")
            except Exception:
                pass

        tk_chart_title = tk.Label(chart_frame, text="", bg=S_BG, fg=S_TEXT,
                                   font=("Georgia", int(13*self.current_scale), "bold"), pady=10)
        tk_chart_title.pack(fill="x")

        # ── Tooltip interactivo al pasar el mouse sobre barras/sectores ──
        tooltip_annot = None

        def on_hover(event):
            nonlocal tooltip_annot
            if event.inaxes is None:
                if tooltip_annot is not None:
                    tooltip_annot.set_visible(False)
                    canvas_widget.draw_idle()
                return
            ax = event.inaxes
            found = False
            for patch in ax.patches:
                contains, _ = patch.contains(event)
                if contains:
                    found = True
                    valor = getattr(patch, "_metatag_valor", None)
                    etiqueta = getattr(patch, "_metatag_label", None)
                    if valor is None:
                        continue
                    if tooltip_annot is None or tooltip_annot.axes != ax:
                        if tooltip_annot is not None:
                            tooltip_annot.remove()
                        tooltip_annot = ax.annotate(
                            "", xy=(0, 0), xytext=(15, 15), textcoords="offset points",
                            bbox=dict(boxstyle="round,pad=0.4", fc=S_ACCENT, ec="none", alpha=0.95),
                            color=S_BG, fontsize=9, fontweight="bold", zorder=10)
                    tooltip_annot.xy = (event.xdata, event.ydata)
                    tooltip_annot.set_text(f"{etiqueta}\n{valor}")
                    tooltip_annot.set_visible(True)
                    canvas_widget.draw_idle()
                    break
            if not found and tooltip_annot is not None:
                tooltip_annot.set_visible(False)
                canvas_widget.draw_idle()

        canvas_widget.mpl_connect("motion_notify_event", on_hover)

        update_chart()
        combo_var.trace_add("write", update_chart)
        chart_type_var.trace_add("write", update_chart)

    # ─────────────────────────────────────────────────────────────
    #  SELECTOR DE TEMA
    # ─────────────────────────────────────────────────────────────
    def _show_theme_menu(self):
        THEME_ICONS = {
            "Arqueológico (Oscuro Refinado)": "🏺",
            "Noche Total": "🌑", "Carbón": "⬛"
        }
        popup = tk.Toplevel(self)
        popup.overrideredirect(True)
        popup.configure(bg=C["border"])
        popup.resizable(False, False)

        btn = self._theme_btn
        self.update_idletasks()
        bx, by = btn.winfo_rootx(), btn.winfo_rooty() + btn.winfo_height() + 2
        popup.geometry(f"+{bx}+{by}")

        outer = tk.Frame(popup, bg=C["border"], padx=1, pady=1)
        outer.pack(fill="both", expand=True)
        inner = tk.Frame(outer, bg=C["surface"])
        inner.pack(fill="both", expand=True)

        hdr = tk.Frame(inner, bg=C["header_bg"])
        hdr.pack(fill="x")
        tk.Label(hdr, text="🎨  Elige un tema", bg=C["header_bg"], fg=C["header_fg"],
                 font=FONTS["LABEL_B"], padx=12, pady=6, anchor="w").pack(fill="x")
        tk.Frame(inner, bg=C["border"], height=1).pack(fill="x")

        def _select(name):
            self.theme_var.set(name)
            popup.destroy()
            self._apply_rebuild()

        for name in THEMES:
            icon       = THEME_ICONS.get(name, "🎨")
            is_current = (name == CURRENT_THEME)
            swatch_clr = THEMES[name]["accent"]
            bg_row, fg_row = ((C["accent_pale"], C["accent"]) if is_current
                              else (C["surface"], C["text"]))
            row = tk.Frame(inner, bg=bg_row, cursor="hand2")
            row.pack(fill="x")
            tk.Frame(row, bg=swatch_clr, width=5).pack(side="left", fill="y")
            mark = "✔ " if is_current else "   "
            lbl = tk.Label(row, text=f"{mark}{icon}  {name}", bg=bg_row, fg=fg_row,
                           font=FONTS["LABEL"], anchor="w", padx=10, pady=7)
            lbl.pack(side="left", fill="x", expand=True)

            def _bind(r, l, n, bg_r, fg_r):
                def _enter(_): r.configure(bg=C["accent_pale"]); l.configure(bg=C["accent_pale"], fg=C["accent"])
                def _leave(_): r.configure(bg=bg_r); l.configure(bg=bg_r, fg=fg_r)
                def _click(_): _select(n)
                for w in (r, l):
                    w.bind("<Enter>", _enter); w.bind("<Leave>", _leave)
                    w.bind("<Button-1>", _click)
            _bind(row, lbl, name, bg_row, fg_row)

        def _on_focus_out(e):
            if popup.winfo_exists():
                popup.after(150, lambda: popup.destroy() if popup.winfo_exists() else None)
        popup.bind("<FocusOut>", _on_focus_out)
        popup.focus_set()

    def _apply_rebuild(self):
        global C, CURRENT_THEME
        CURRENT_THEME = self.theme_var.get()
        C.update(THEMES[CURRENT_THEME])
        self.configure(bg=C["bg"])
        saved_df, saved_folder = self.df, self.img_folder_var.get()
        for widget in self.winfo_children(): widget.destroy()
        self._build_styles()
        self._build_ui()
        self._toggle_mode()
        if saved_df is not None:
            self.df = saved_df
            self.grid.load(self.df)
            self._populate_img_col(self.df)
            self._update_filter_columns()
            self.row_lbl.configure(text=f"{len(self.df)} filas · {len(self.df.columns)} columnas")
            self.status_var.set(f"Archivo cargado — {len(self.df)} filas")
        if saved_folder:
            self.img_folder_var.set(saved_folder)
            self.browser.load_folder(saved_folder)
        self._save_config()

    # ─────────────────────────────────────────────────────────────
    #  CONSTRUCCIÓN DE UI
    # ─────────────────────────────────────────────────────────────
    def _build_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", background=C["bg"], foreground=C["text"], font=FONTS["BODY"])
        s.configure("TFrame",       background=C["bg"])
        s.configure("Panel.TFrame", background=C["panel"])
        s.configure("TLabel",       background=C["bg"], foreground=C["text"])
        s.configure("TScrollbar",   background=C["border"], troughcolor=C["surface"],
                    arrowcolor=C["surface"], borderwidth=0, relief="flat",
                    width=int(12*self.current_scale), arrowsize=0)
        s.map("TScrollbar", background=[("active", C["accent_light"]),
                                        ("pressed", C["accent"]), ("!active", C["border"])])
        s.configure("TProgressbar", troughcolor=C["border_light"], background=C["accent"],
                    borderwidth=0, thickness=int(4*self.current_scale))
        s.configure("TCombobox", fieldbackground=C["surface"], foreground=C["text"],
                    selectbackground=C["accent_light"], selectforeground=C["text"], font=FONTS["BODY"])
        s.map("TCombobox", fieldbackground=[("readonly", C["surface"])])

    def _build_ui(self):
        self._build_topbar()
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")
        body = tk.Frame(self, bg=C["bg"])
        body.pack(fill="both", expand=True)

        self.paned_window = tk.PanedWindow(body, orient="horizontal",
                                           bg=C["border"], sashwidth=int(4*self.current_scale), bd=0)
        self.paned_window.pack(fill="both", expand=True)

        self.left = tk.Frame(self.paned_window, bg=C["panel"])
        left_canvas = tk.Canvas(self.left, bg=C["panel"], highlightthickness=0, bd=0)
        left_vsb = ttk.Scrollbar(self.left, orient="vertical", command=left_canvas.yview)
        left_inner = tk.Frame(left_canvas, bg=C["panel"])
        left_inner.bind("<Configure>", lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        left_canvas.create_window((0, 0), window=left_inner, anchor="nw")
        left_canvas.configure(yscrollcommand=left_vsb.set)
        left_canvas.pack(side="left", fill="both", expand=True)
        left_vsb.pack(side="right", fill="y")

        s = ttk.Style()
        s.configure("Hidden.Vertical.TScrollbar", troughcolor=C["panel"],
                     background=C["panel"], bordercolor=C["panel"],
                     arrowcolor=C["panel"], relief="flat")
        left_vsb.configure(style="Hidden.Vertical.TScrollbar")

        def _on_left_mousewheel(event):
            left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def _on_left_mousewheel_linux(event):
            if event.num == 4:
                left_canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                left_canvas.yview_scroll(1, "units")
        def _bind_scroll_recursive(widget):
            widget.bind("<MouseWheel>", _on_left_mousewheel)
            widget.bind("<Button-4>", _on_left_mousewheel_linux)
            widget.bind("<Button-5>", _on_left_mousewheel_linux)
            for child in widget.winfo_children():
                _bind_scroll_recursive(child)

        self._build_control_panel(left_inner)
        _bind_scroll_recursive(left_inner)
        self.paned_window.add(self.left, minsize=int(220*self.current_scale))

        center = tk.Frame(self.paned_window, bg=C["bg"])
        self._build_center(center)
        self.paned_window.add(center, minsize=int(400*self.current_scale))

        right = tk.Frame(self.paned_window, bg=C["panel"])
        self._build_right(right)
        self.paned_window.add(right, minsize=int(260*self.current_scale))

        self._build_statusbar()
        self.update_idletasks()
        total_w = self.winfo_width()
        if total_w > 10:
            self.paned_window.paneconfigure(self.left,   width=int(total_w * 0.20))
            self.paned_window.paneconfigure(center,      width=int(total_w * 0.55))

    def _build_topbar(self):
        bar = tk.Frame(self, bg=C["header_bg"], height=int(52*self.current_scale))
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="⬡  MetaTag v8.9", bg=C["header_bg"],
                 fg="#FFF5E8", font=FONTS["TITLE"]).pack(side="left", padx=20, pady=int(10*self.current_scale))

        dep_ok = PIL_OK and MATPLOTLIB_OK
        self.dep_lbl = tk.Label(bar, bg=C["header_bg"], font=FONTS["LABEL"],
                                text="✔ Todo instalado" if dep_ok else "⚠ Faltan dependencias",
                                fg=C["accent_light"] if dep_ok else "#FF9966")
        self.dep_lbl.pack(side="right", padx=20)

        THEME_ICONS = {"Arqueológico (Oscuro Refinado)": "🏺", "Noche Total": "🌑", "Carbón": "⬛"}
        icon = THEME_ICONS.get(CURRENT_THEME, "🎨")
        theme_wrap = tk.Frame(bar, bg=C["accent_hover"], padx=1, pady=1)
        theme_wrap.pack(side="right", padx=(0, 14), pady=int(10*self.current_scale))
        self._theme_btn = tk.Button(theme_wrap, text=f"{icon}  {CURRENT_THEME}  ▾",
                                    bg=C["accent"], fg=C["header_fg"], font=FONTS["LABEL_B"],
                                    relief="flat", bd=0, padx=14, pady=0, cursor="hand2",
                                    activebackground=C["accent_hover"], activeforeground="#FFF5E8",
                                    command=self._show_theme_menu)
        self._theme_btn.pack(ipady=int(6*self.current_scale))

    def _build_control_panel(self, parent):
        def section(txt):
            f = tk.Frame(parent, bg=C["panel"])
            f.pack(fill="x", padx=10, pady=(int(14*self.current_scale), 2))
            tk.Label(f, text=txt, bg=C["panel"], fg=C["text3"], font=FONTS["TINY"]).pack(side="left")
            tk.Frame(parent, bg=C["border_light"], height=1).pack(fill="x", padx=10, pady=(2, 4))

        def browse_row(par, var, cmd):
            f = tk.Frame(par, bg=C["panel"])
            f.pack(fill="x", padx=10, pady=(0, int(6*self.current_scale)))
            e = tk.Entry(f, textvariable=var, bg=C["surface"], fg=C["text"],
                         relief="solid", bd=1, font=FONTS["LABEL"])
            e.pack(side="left", fill="x", expand=True)
            tk.Button(f, text="…", bg=C["btn_ghost_bg"], fg=C["accent"],
                      font=FONTS["TINY"], relief="flat", bd=0, padx=6,
                      cursor="hand2", command=cmd).pack(side="left", padx=(4, 0))

        section("ARCHIVO DE DATOS")
        tk.Label(parent, text="Excel / CSV", bg=C["panel"],
                 fg=C["text2"], font=FONTS["LABEL"]).pack(anchor="w", padx=10)
        browse_row(parent, self.csv_path_var, self._browse_csv)
        self._btn(parent, "⟳  Cargar archivo", self._load_file, primary=True)

        section("CARPETA DE IMÁGENES")
        browse_row(parent, self.img_folder_var, self._browse_folder)
        self._btn(parent, "🗁  Abrir carpeta", self._open_folder, primary=False)

        section("SELECCIÓN DE DATOS")
        self.img_col_cb = ttk.Combobox(parent, textvariable=self.img_col_var,
                                       state="readonly", font=FONTS["LABEL"])
        self.img_col_cb.pack(fill="x", padx=10, pady=(0, 4))
        self.btn_omit_empty = tk.Button(parent, text="[ ✓ ] Omitir celdas vacías",
                                        bg=C["sel_bg"], fg=C["sel_fg"], font=FONTS["TINY"],
                                        relief="flat", bd=0, cursor="hand2", anchor="w",
                                        padx=8, pady=4, command=self._toggle_omit_empty)
        self.btn_omit_empty.pack(fill="x", padx=10, pady=(8, 0))
        tk.Frame(parent, bg=C["panel"], height=5).pack(fill="x")
        self._btn(parent, "↔  Seleccionar fila activa", self._select_active_row, primary=False)
        self._btn(parent, "✕  Limpiar selección",       self._clear_selection,   primary=False)

        section("HERRAMIENTAS AVANZADAS")
        self._btn(parent, "📊 Ver Estadísticas (Gráficos)",  self._show_stats,           primary=False)
        self._btn(parent, "🗂 Lote por Orden (Excel→Fotos)", self._batch_write_by_order, primary=False)

        section("SINCRONIZACIÓN DE ORDEN")
        self._btn(parent, "🔄 Reordenar imágenes según Excel",  self._sync_images_to_excel, primary=False)

        section("VERIFICACIÓN DE INTEGRIDAD")
        self._btn(parent, "🔍 Verificar imágenes originales",   self._verify_source_images, primary=False)

        section("⚙️ MODO DE TRABAJO")
        self.mode_frame = tk.Frame(parent, bg=C["btn_ghost_bg"], padx=2, pady=2)
        self.mode_frame.pack(fill="x", padx=10, pady=(0, 4))
        self.btn_mode_smart = tk.Button(self.mode_frame, text="🧠 Inteligente",
                                        font=FONTS["LABEL_B"], relief="flat", bd=0,
                                        cursor="hand2", pady=4,
                                        command=lambda: self._set_mode("Inteligente"))
        self.btn_mode_smart.pack(side="left", fill="x", expand=True, padx=1)
        self.btn_mode_free = tk.Button(self.mode_frame, text="⚡ Libre",
                                       font=FONTS["LABEL"], relief="flat", bd=0,
                                       cursor="hand2", pady=4,
                                       command=lambda: self._set_mode("Libre"))
        self.btn_mode_free.pack(side="left", fill="x", expand=True, padx=1)

        section("CARPETA DE SALIDA")
        rel = (self.output_folder.relative_to(self.output_base)
               if self.output_folder.is_relative_to(self.output_base)
               else self.output_folder)
        self.out_lbl = tk.Label(parent, text=f"/{rel}", bg=C["panel"], fg=C["accent"],
                                font=FONTS["TINY"], wraplength=int(200*self.current_scale),
                                justify="left")
        self.out_lbl.pack(anchor="w", padx=10, pady=(0, 4))
        self._btn(parent, "📂  Abrir carpeta de salida", self._open_output_folder, primary=False)

        section("ATAJOS DE TECLADO")
        self._btn(parent, "⌨  Ver atajos disponibles", self._show_shortcuts, primary=False)

        tk.Frame(parent, bg=C["border_light"], height=1).pack(fill="x", padx=10, pady=10)

        self.action_frame = tk.Frame(parent, bg=C["panel"])
        self.action_frame.pack(fill="x", padx=10)
        self._write_btn  = tk.Button(self.action_frame, text="▶  Escribir Metadatos",
                                     command=self._start_processing, bg=C["accent"],
                                     fg="#FFF5E8", font=FONTS["LABEL_B"], relief="flat",
                                     bd=0, cursor="hand2", pady=int(9*self.current_scale),
                                     activebackground=C["accent_hover"], activeforeground="#FFF5E8")
        self._inject_btn = tk.Button(self.action_frame, text="📌 Inyectar en Foto Actual",
                                     command=self._inject_manual, bg=C["warn"],
                                     fg="#FFFFFF", font=FONTS["LABEL_B"], relief="flat",
                                     bd=0, cursor="hand2", pady=int(9*self.current_scale),
                                     activebackground=C["accent_hover"], activeforeground="#FFFFFF")
        self.sel_info = tk.Label(parent, text="Sin selección", bg=C["panel"],
                                 fg=C["text3"], font=FONTS["TINY"],
                                 wraplength=int(200*self.current_scale), justify="left")
        self.sel_info.pack(anchor="w", padx=10, pady=(6, 0))

    def _btn(self, parent, text, cmd, primary=True):
        bg, fg = (C["accent"], "#FFF5E8") if primary else (C["btn_ghost_bg"], C["accent"])
        b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                      font=FONTS["LABEL_B"] if primary else FONTS["LABEL"],
                      relief="flat", bd=0, cursor="hand2",
                      pady=int(6*self.current_scale),
                      activebackground=C["accent_hover"], activeforeground="#FFF5E8")
        b.pack(fill="x", padx=10, pady=(0, 4))
        return b

    def _toggle_omit_empty(self):
        self.omit_empty_var.set(not self.omit_empty_var.get())
        if self.omit_empty_var.get():
            self.btn_omit_empty.configure(text="[ ✓ ] Omitir celdas vacías",
                                          bg=C["sel_bg"], fg=C["sel_fg"])
        else:
            self.btn_omit_empty.configure(text="[   ] Incluir celdas vacías",
                                          bg=C["btn_ghost_bg"], fg=C["text"])
        self._update_meta_preview()
        self._on_sel_change()

    def _set_mode(self, mode):
        self.process_mode.set(mode)
        self._toggle_mode()

    def _toggle_mode(self):
        mode = self.process_mode.get()
        if hasattr(self, "btn_mode_smart"):
            if mode == "Inteligente":
                self.btn_mode_smart.configure(bg=C["accent"],        fg="#FFF5E8", font=FONTS["LABEL_B"])
                self.btn_mode_free.configure( bg=C["btn_ghost_bg"],  fg=C["text"], font=FONTS["LABEL"])
            else:
                self.btn_mode_smart.configure(bg=C["btn_ghost_bg"],  fg=C["text"], font=FONTS["LABEL"])
                self.btn_mode_free.configure( bg=C["warn"],           fg="#FFFFFF", font=FONTS["LABEL_B"])
        if mode == "Inteligente":
            self._inject_btn.pack_forget()
            self._write_btn.pack(fill="x", pady=(0, 4))
            self.sel_info.configure(
                text="Modo Inteligente: Los nombres de archivo deben coincidir con la tabla.")
            if self.current_row is not None: self._on_row_click(self.current_row)
        else:
            self._write_btn.pack_forget()
            self._inject_btn.pack(fill="x", pady=(0, 4))
            self.sel_info.configure(
                text="Modo Libre: Se inyectarán los datos de la fila seleccionada en la foto que estés viendo.")
        self._save_config()

    def _build_center(self, parent):
        top = tk.Frame(parent, bg=C["bg"])
        top.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(top, text="Tabla de datos", bg=C["bg"],
                 fg=C["header_bg"], font=FONTS["H2"]).pack(side="left")
        self.sel_badge = tk.Label(top, text="", bg=C["accent_pale"], fg=C["accent"],
                                  font=FONTS["TINY"], padx=8, pady=2)
        self.sel_badge.pack(side="left", padx=10)
        self.row_lbl = tk.Label(top, text="", bg=C["bg"], fg=C["text3"], font=FONTS["TINY"])
        self.row_lbl.pack(side="right")

        sf = tk.Frame(parent, bg=C["bg"])
        sf.pack(fill="x", padx=12, pady=(0, 8))
        tk.Label(sf, text="🔍 Buscar en:", bg=C["bg"], fg=C["text3"], font=FONTS["LABEL"]).pack(side="left")
        self.filter_col_var = tk.StringVar(value="Todas las columnas")
        self.filter_col_cb  = ttk.Combobox(sf, textvariable=self.filter_col_var,
                                            state="readonly", font=FONTS["LABEL"], width=18)
        self.filter_col_cb.pack(side="left", padx=(4, 8))
        self.filter_col_cb.bind("<<ComboboxSelected>>", self._on_filter_type)
        self.filter_var   = tk.StringVar()
        self.filter_var.trace_add("write", self._on_filter_type)
        self.filter_entry = tk.Entry(sf, textvariable=self.filter_var,
                                     bg=C["surface"], fg=C["text"], relief="solid", bd=1, font=FONTS["LABEL"])
        self.filter_entry.pack(side="left", fill="x", expand=True)

        self.grid = ExcelGrid(parent, on_selection_change=self._on_sel_change,
                              on_row_click=self._on_row_click, app_ref=self)
        self.grid.pack(fill="both", expand=True, padx=12, pady=(0, 4))

        tk.Frame(parent, bg=C["border_light"], height=1).pack(fill="x", padx=12, pady=(2, 4))
        tk.Label(parent, text="Registro de actividad", bg=C["bg"],
                 fg=C["text3"], font=FONTS["TINY"]).pack(anchor="w", padx=12)
        lf = tk.Frame(parent, bg=C["bg"])
        lf.pack(fill="x", padx=12, pady=(2, 8))
        self.log = tk.Text(lf, height=5, bg=C["surface"], fg=C["text"],
                           font=FONTS["MONO"], relief="flat", state="disabled",
                           wrap="word", borderwidth=0)
        lsb = ttk.Scrollbar(lf, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=lsb.set)
        lsb.pack(side="right", fill="y")
        self.log.pack(fill="x")
        self.log.tag_configure("ok",   foreground=C["ok"])
        self.log.tag_configure("err",  foreground=C["err"])
        self.log.tag_configure("info", foreground=C["warn"])
        self.log.tag_configure("head", foreground=C["header_bg"], font=FONTS["LABEL_B"])

    def _build_right(self, parent):
        self.browser = ImageBrowser(parent, on_select=self._on_img_select)
        self.browser.pack(fill="x", padx=0, pady=0)
        tk.Frame(parent, bg=C["border"], height=1).pack(fill="x", pady=4)

        vhdr = tk.Frame(parent, bg=C["panel"])
        vhdr.pack(fill="x", padx=8, pady=(8, 4))
        tk.Label(vhdr, text="Vista previa", bg=C["panel"],
                 fg=C["text"], font=FONTS["H2"]).pack(side="left")

        def launch_visor():
            visor_path = None
            for f in self.output_base.iterdir():
                if f.name.lower() == "visor.py":
                    visor_path = str(f)
                    break
            if not visor_path:
                return messagebox.showerror("Error",
                    f"No se encontró el archivo visor.py en:\n{self.output_base}")
            self.withdraw()
            proc = subprocess.Popen([sys.executable, visor_path, "STANDALONE", CURRENT_THEME])

            def _check_visor():
                if proc.poll() is not None:
                    self.deiconify()
                    self.lift()
                    self.focus_force()
                else:
                    self.after(500, _check_visor)
            self.after(500, _check_visor)

        self._btn_visor = tk.Button(vhdr, text="👁 Visor Pro", bg=C["surface"],
                                    fg=C["accent"], font=FONTS["LABEL_B"], relief="flat",
                                    bd=0, cursor="hand2", padx=10, pady=2,
                                    activebackground=C["accent"], activeforeground="#FFFFFF",
                                    command=launch_visor)
        self._btn_visor.pack(side="right", padx=(4, 0))
        self._btn_open = tk.Button(vhdr, text="📂", bg=C["btn_ghost_bg"], fg=C["text2"],
                                   font=FONTS["TINY"], relief="flat", bd=0, cursor="hand2",
                                   padx=6, pady=2, command=self._browse_single_image)
        self._btn_open.pack(side="right")

        self.img_name_lbl = tk.Label(parent, text="Sin imagen seleccionada",
                                     bg=C["panel"], fg=C["accent"], font=FONTS["TINY"],
                                     anchor="w", wraplength=int(310*self.current_scale), justify="left")
        self.img_name_lbl.pack(fill="x", padx=8, pady=(2, 0))
        self.img_hint_lbl = tk.Label(parent, text="Rueda del mouse para zoom · Clic y arrastra para mover",
                                     bg=C["panel"], fg=C["text3"], font=FONTS["TINY"], anchor="w")
        self.img_hint_lbl.pack(fill="x", padx=8, pady=(0, 4))

        self.img_canvas = tk.Canvas(parent, bg=C["surface"], highlightthickness=1,
                                    highlightbackground=C["border_light"], cursor="crosshair")
        self.img_canvas.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        self.img_canvas.create_text(int(155*self.current_scale), int(120*self.current_scale),
                                    text="Sin imagen", fill=C["text3"], font=FONTS["LABEL"],
                                    justify="center", tags="ph")
        self.current_img_tk = None
        self._preview_zoom    = 1.0
        self._preview_pan_x   = 0
        self._preview_pan_y   = 0
        self._preview_drag_start = None
        self._preview_pil     = None
        self.img_canvas.bind("<Configure>",   self._on_preview_resize)
        self.img_canvas.bind("<MouseWheel>",  self._on_preview_wheel)
        self.img_canvas.bind("<Button-4>",    self._on_preview_wheel)
        self.img_canvas.bind("<Button-5>",    self._on_preview_wheel)
        self.img_canvas.bind("<ButtonPress-1>",   self._on_preview_drag_start)
        self.img_canvas.bind("<B1-Motion>",       self._on_preview_drag)
        self.img_canvas.bind("<ButtonRelease-1>", self._on_preview_drag_end)
        self.img_canvas.bind("<Double-Button-1>", lambda e: self._reset_preview_zoom())

        tk.Frame(parent, bg=C["border_light"], height=1).pack(fill="x", padx=8, pady=4)
        tk.Label(parent, text="Metadatos a escribir:", bg=C["panel"],
                 fg=C["text2"], font=FONTS["TINY"]).pack(anchor="w", padx=8)
        mode_row = tk.Frame(parent, bg=C["panel"])
        mode_row.pack(fill="x", padx=10, pady=(4, 0))
        tk.Label(mode_row, text="Formato:", bg=C["panel"], fg=C["text3"], font=FONTS["TINY"]).pack(side="left")
        tk.Checkbutton(mode_row, text="Organizado", variable=self.meta_mode_organized,
                       bg=C["panel"], fg=C["text2"], selectcolor=C["surface"],
                       activebackground=C["panel"], font=FONTS["TINY"], cursor="hand2").pack(side="right")
        self.meta_txt = tk.Text(parent, height=7, bg=C["surface"], fg=C["text2"],
                                font=FONTS["MONO"], relief="flat", state="disabled",
                                wrap="word", borderwidth=0)
        self.meta_txt.pack(fill="x", padx=8, pady=(4, 8))
        self.meta_txt.tag_configure("head", foreground=C["header_bg"], font=FONTS["LABEL_B"])

    def _build_statusbar(self):
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x", side="bottom")
        bar = tk.Frame(self, bg=C["card"], height=int(28*self.current_scale))
        bar.pack(fill="x", side="bottom")
        ttk.Progressbar(bar, variable=self.progress_var, maximum=100,
                        style="TProgressbar",
                        length=int(160*self.current_scale)).pack(side="right", padx=12, pady=6)
        tk.Label(bar, textvariable=self.status_var, bg=C["card"],
                 fg=C["text3"], font=FONTS["TINY"]).pack(side="left", padx=12, pady=6)

    # ─────────────────────────────────────────────────────────────
    #  LUPA HD
    # ─────────────────────────────────────────────────────────────
    def _on_preview_wheel(self, event):
        if not hasattr(self, "_preview_pil") or self._preview_pil is None:
            return
        if event.num == 4 or (hasattr(event, "delta") and event.delta > 0):
            factor = 1.12
        else:
            factor = 1 / 1.12
        new_zoom = max(0.2, min(self._preview_zoom * factor, 12.0))
        self._preview_zoom = new_zoom
        self._schedule_preview_redraw()

    def _schedule_preview_redraw(self):
        # Si el usuario mueve la rueda muy rápido, Tk genera muchos eventos
        # en ráfaga. Antes cada uno disparaba su propio resize+redraw, y
        # como no daba tiempo a terminar uno antes de que llegara el
        # siguiente, se iban encolando y la UI se sentía "trabada".
        # Con after_idle solo programamos UN redraw por ciclo de eventos:
        # si llegan más eventos antes de que se ejecute, simplemente
        # actualizan _preview_zoom/_preview_pan y el redraw pendiente
        # toma el valor más reciente cuando por fin corre.
        if not getattr(self, "_preview_redraw_pending", False):
            self._preview_redraw_pending = True
            self.after_idle(self._flush_preview_redraw)

    def _flush_preview_redraw(self):
        self._preview_redraw_pending = False
        self._redraw_preview_zoomed(fast=True)
        if hasattr(self, "_preview_hq_timer"):
            self.after_cancel(self._preview_hq_timer)
        self._preview_hq_timer = self.after(120, lambda: self._redraw_preview_zoomed(fast=False))

    def _on_preview_drag_start(self, event):
        self._preview_drag_start = (event.x, event.y)
        self.img_canvas.configure(cursor="fleur")

    def _on_preview_drag(self, event):
        if self._preview_drag_start is None:
            return
        dx = event.x - self._preview_drag_start[0]
        dy = event.y - self._preview_drag_start[1]
        self._preview_pan_x += dx
        self._preview_pan_y += dy
        self._preview_drag_start = (event.x, event.y)
        self._schedule_preview_redraw()

    def _on_preview_drag_end(self, event):
        self._preview_drag_start = None
        self.img_canvas.configure(cursor="crosshair")

    def _reset_preview_zoom(self):
        self._preview_zoom  = 1.0
        self._preview_pan_x = 0
        self._preview_pan_y = 0
        self._redraw_preview_zoomed()

    def _redraw_preview_zoomed(self, fast=False):
        if not hasattr(self, "_preview_pil") or self._preview_pil is None:
            return
        cw = max(self.img_canvas.winfo_width(),  int(310 * self.current_scale))
        ch = max(self.img_canvas.winfo_height(), int(260 * self.current_scale))
        iw, ih = self._preview_pil.size
        scale = min(cw / iw, ch / ih) * self._preview_zoom
        nw = max(1, int(iw * scale))
        nh = max(1, int(ih * scale))

        # Durante interacción (rueda/arrastre) usamos un resample barato
        # (BILINEAR) sobre una miniatura pre-reducida (_preview_pil_fast,
        # ~1400px de lado mayor) en vez de la imagen original a full
        # resolución. Redimensionar una foto de 4000×3000 en cada frame
        # es lo que más pesaba; partir de una miniatura lo vuelve casi
        # instantáneo. Solo al terminar el gesto (debounce de 120ms) se
        # usa la imagen completa con LANCZOS para nitidez real.
        cache = getattr(self, "_preview_resize_cache", None)
        if cache is not None and cache[0] == (nw, nh) and cache[1] == fast:
            resized = cache[2]
        else:
            if fast and getattr(self, "_preview_pil_fast", None) is not None:
                resample = Image.BILINEAR
                resized = self._preview_pil_fast.resize((nw, nh), resample)
            else:
                resample = Image.BILINEAR if fast else Image.LANCZOS
                resized = self._preview_pil.resize((nw, nh), resample)
            self._preview_resize_cache = ((nw, nh), fast, resized)

        cx = cw // 2 + self._preview_pan_x
        cy = ch // 2 + self._preview_pan_y
        self.current_img_tk = ImageTk.PhotoImage(resized)
        self.img_canvas.delete("all")
        self.img_canvas.create_image(cx, cy, anchor="center", image=self.current_img_tk)
        if abs(self._preview_zoom - 1.0) > 0.05:
            self.img_canvas.create_text(
                cw - 8, ch - 8, anchor="se",
                text=f"×{self._preview_zoom:.1f}",
                fill=C["accent"], font=FONTS["TINY"])

    def _on_preview_resize(self, event):
        if hasattr(self, "_preview_resize_timer"): self.after_cancel(self._preview_resize_timer)
        self._preview_resize_timer = self.after(150, self._redraw_preview)

    def _redraw_preview(self):
        self._redraw_preview_zoomed()

    def _show_loupe_window(self, event=None):
        if not self.current_img or not os.path.exists(self.current_img): return
        if hasattr(self, "loupe_window") and self.loupe_window.winfo_exists():
            self.loupe_window.lift(); self.loupe_window.focus_force()
            self._load_loupe_image(self.current_img); return

        self.loupe_window = tk.Toplevel(self)
        self.loupe_window.geometry(f"{int(1000*self.current_scale)}x{int(700*self.current_scale)}")
        self.loupe_window.configure(bg=C["surface"])
        self.loupe_canvas = tk.Canvas(self.loupe_window, bg=C["surface"],
                                      highlightthickness=0, cursor="fleur")
        self.loupe_canvas.pack(fill="both", expand=True)
        self.loupe_window.bind("<Configure>", self._on_loupe_resize)
        self.loupe_canvas.bind("<ButtonPress-1>", self._on_loupe_press)
        self.loupe_canvas.bind("<B1-Motion>",     self._on_loupe_drag)
        self.loupe_canvas.bind("<MouseWheel>",    self._on_loupe_zoom)
        self.loupe_canvas.bind("<Button-4>",      lambda e: self._on_loupe_zoom_linux(e, 1.15))
        self.loupe_canvas.bind("<Button-5>",      lambda e: self._on_loupe_zoom_linux(e, 1/1.15))
        self._loupe_orig_img = None
        self._loupe_scale    = 1.0
        self._loupe_x = self._loupe_y = 0
        self.after(50, lambda: self._load_loupe_image(self.current_img))

    def _on_loupe_press(self, e): self._drag_start_x, self._drag_start_y = e.x, e.y
    def _on_loupe_drag(self, e):
        self._loupe_x += e.x - self._drag_start_x
        self._loupe_y += e.y - self._drag_start_y
        self._drag_start_x, self._drag_start_y = e.x, e.y
        self._render_loupe_img(fast=True)
        if hasattr(self, "_loupe_hq_timer"): self.after_cancel(self._loupe_hq_timer)
        self._loupe_hq_timer = self.after(200, lambda: self._render_loupe_img(fast=False))
    def _on_loupe_resize(self, e):
        if e.widget == self.loupe_window:
            if hasattr(self, "_loupe_resize_timer"): self.after_cancel(self._loupe_resize_timer)
            self._loupe_resize_timer = self.after(200, lambda: self._render_loupe_img(fast=False))
    def _on_loupe_zoom(self, e):
        if self._loupe_orig_img is None: return
        factor    = 1.15 if e.delta > 0 else (1/1.15)
        new_scale = max(0.05, min(self._loupe_scale * factor, 15.0))
        actual_f  = new_scale / self._loupe_scale
        self._loupe_scale = new_scale
        self._loupe_x = e.x - (e.x - self._loupe_x) * actual_f
        self._loupe_y = e.y - (e.y - self._loupe_y) * actual_f
        self._render_loupe_img(fast=True)
        if hasattr(self, "_loupe_hq_timer"): self.after_cancel(self._loupe_hq_timer)
        self._loupe_hq_timer = self.after(200, lambda: self._render_loupe_img(fast=False))

    def _on_loupe_zoom_linux(self, e, factor):
        if self._loupe_orig_img is None: return
        new_scale = max(0.05, min(self._loupe_scale * factor, 15.0))
        actual_f  = new_scale / self._loupe_scale
        self._loupe_scale = new_scale
        self._loupe_x = e.x - (e.x - self._loupe_x) * actual_f
        self._loupe_y = e.y - (e.y - self._loupe_y) * actual_f
        self._render_loupe_img(fast=True)
        if hasattr(self, "_loupe_hq_timer"): self.after_cancel(self._loupe_hq_timer)
        self._loupe_hq_timer = self.after(200, lambda: self._render_loupe_img(fast=False))

    def _load_loupe_image(self, path):
        if not hasattr(self, "loupe_window") or not self.loupe_window.winfo_exists(): return
        self.loupe_window.title(f"Lupa HD — {Path(path).name}")
        try:
            img = Image.open(path)
            self._loupe_orig_img = ImageOps.exif_transpose(img)
            w_win = self.loupe_window.winfo_width()  or 1000
            h_win = self.loupe_window.winfo_height() or 700
            self._loupe_scale = min(w_win / self._loupe_orig_img.size[0],
                                    h_win / self._loupe_orig_img.size[1]) * 0.95
            self._loupe_x = (w_win - int(self._loupe_orig_img.size[0] * self._loupe_scale)) // 2
            self._loupe_y = (h_win - int(self._loupe_orig_img.size[1] * self._loupe_scale)) // 2
            self._render_loupe_img(fast=False)
        except Exception as e:
            self.loupe_canvas.delete("all")
            self.loupe_canvas.create_text(500, 350, text=f"Error: {e}", fill=C["err"])

    def _render_loupe_img(self, fast=False):
        if self._loupe_orig_img is None: return
        nw = max(1, int(self._loupe_orig_img.size[0] * self._loupe_scale))
        nh = max(1, int(self._loupe_orig_img.size[1] * self._loupe_scale))
        filtro  = Image.BILINEAR if fast else Image.LANCZOS
        resized = self._loupe_orig_img.resize((nw, nh), filtro)
        self.loupe_tk_img = ImageTk.PhotoImage(resized)
        self.loupe_canvas.delete("all")
        self.loupe_canvas.create_image(self._loupe_x, self._loupe_y,
                                       anchor="nw", image=self.loupe_tk_img)

    # ─────────────────────────────────────────────────────────────
    #  DIÁLOGOS DE ARCHIVOS
    # ─────────────────────────────────────────────────────────────
    def _browse_csv(self):
        p = _native_file_open(
            title="Seleccionar archivo de datos",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")])
        if p:
            self.csv_path_var.set(p)
            if not self.img_folder_var.get():
                self.img_folder_var.set(str(Path(p).parent))

    def _browse_folder(self):
        f = _native_folder_open(title="Seleccionar carpeta de imágenes")
        if f:
            self.img_folder_var.set(f)
            self.browser.load_folder(f)
            self._img_cache = {self._full_stem(fp.name).lower(): fp
                               for fp in Path(f).rglob("*")
                               if fp.is_file() and fp.suffix.lower() in IMG_EXTS}
            self._save_config()

    def _open_folder(self):
        f = _native_folder_open(title="Seleccionar carpeta de imágenes")
        if f:
            self.img_folder_var.set(f)
            self.browser.load_folder(f)
            self._img_cache = {self._full_stem(fp.name).lower(): fp
                               for fp in Path(f).rglob("*")
                               if fp.is_file() and fp.suffix.lower() in IMG_EXTS}
            self._save_config()

    def _browse_single_image(self):
        p = _native_file_open(
            title="Seleccionar imagen",
            filetypes=[("Imágenes", "*.jpg *.jpeg *.png *.tif *.tiff *.bmp *.webp"),
                       ("Todos", "*.*")])
        if p: self._load_image(p)

    # ─────────────────────────────────────────────────────────────
    #  CARGA DE EXCEL / CSV
    # ─────────────────────────────────────────────────────────────
    def _load_file(self):
        p = self.csv_path_var.get().strip()
        if not p or not os.path.exists(p):
            return messagebox.showerror("Error", "Selecciona un archivo válido.")
        try:
            ext = Path(p).suffix.lower()
            df  = (pd.read_csv(p, dtype=str, keep_default_na=False, na_values=[])
                   if ext == ".csv"
                   else pd.read_excel(p, dtype=str, keep_default_na=False, na_values=[]))
            df = df.map(lambda v: "" if (
                v is None or
                (isinstance(v, float) and str(v).lower() == "nan") or
                str(v).strip().lower() in ("nat", "none")
            ) else str(v).strip())
            df.columns = [str(c).strip() for c in df.columns]
            self.original_df = df
            self.df          = df.copy()
            self.grid.load(self.df)
            self._populate_img_col(self.df)
            self._update_filter_columns()
            self.row_lbl.configure(text=f"{len(df)} filas · {len(df.columns)} columnas")
            self._log(f"✔ Archivo cargado: {Path(p).name}  ({len(df)} filas)\n", "head")
            self.status_var.set(f"Archivo cargado — {len(df)} filas")
            self.title(f"MetaTag v8.9  —  {Path(p).name}  ({len(df)} registros)")
            self._save_config()
        except Exception as e:
            messagebox.showerror("Error al cargar", str(e))

    def _populate_img_col(self, df):
        cols = list(df.columns)
        self.img_col_cb["values"] = cols
        for g in ["id", "imagen", "image", "file", "archivo", "nombre", "name", "foto", "photo"]:
            m = [c for c in cols if g in c.lower()]
            if m: return self.img_col_var.set(m[0])
        self.img_col_var.set(cols[0] if cols else "")

    # ─────────────────────────────────────────────────────────────
    #  SELECCIÓN Y FILTROS
    # ─────────────────────────────────────────────────────────────
    def _update_filter_columns(self):
        if self.original_df is not None:
            self.filter_col_cb["values"] = ["Todas las columnas"] + list(self.original_df.columns)

    def _on_filter_type(self, *args):
        if hasattr(self, "_filter_timer"): self.after_cancel(self._filter_timer)
        self._filter_timer = self.after(300, self._apply_filter)

    def _apply_filter(self):
        if self.original_df is None: return
        query, col_b = self.filter_var.get().strip(), self.filter_col_var.get()
        if not query:
            self.df = self.original_df.copy()
        else:
            if col_b == "Todas las columnas":
                mask = self.original_df.apply(
                    lambda row: row.astype(str).str.contains(query, case=False, na=False).any(),
                    axis=1)
            else:
                mask = self.original_df[col_b].astype(str).str.contains(
                    query, case=False, na=False)
            self.df = self.original_df[mask].reset_index(drop=True)
        self.grid.load(self.df)
        self.row_lbl.configure(
            text=(f"{len(self.df)} filas" if not query
                  else f"Filtrado: {len(self.df)} / {len(self.original_df)} filas"))
        self._clear_selection()
        self._update_meta_preview()

    def _on_sel_change(self):
        if hasattr(self, "_sel_debounce"): self.after_cancel(self._sel_debounce)
        n = len(self.grid.selected_cells)
        if n == 0:
            self.sel_badge.configure(text="")
            self.status_var.set("Listo")
        else:
            cols_sel, rows_sel = set(), set()
            for (r, c) in self.grid.selected_cells:
                rows_sel.add(r)
                if self.grid.df is not None: cols_sel.add(self.grid.df.columns[c])
            self.sel_badge.configure(
                text=f"{n} celda(s) · {len(cols_sel)} col(s) · {len(rows_sel)} fila(s)")
            if self.grid.df is not None:
                cols_list    = list(self.grid.df.columns)
                img_col_idx  = self._img_col_idx()
                empty_found  = []
                for (ri, ci) in self.grid.selected_cells:
                    if ci == img_col_idx: continue
                    val = self.grid._get_val(self.grid.df.iloc[ri, ci])
                    if not val or val.lower() in ("nan", "none", ""):
                        empty_found.append(f"{cols_list[ci]} (fila {ri+1})")
                if empty_found and self.omit_empty_var.get():
                    self.status_var.set(
                        f"⚠ {len(empty_found)} celda(s) vacía(s) seleccionada(s) — se omitirán")
                else:
                    self.status_var.set("Listo")
        self._sel_debounce = self.after(80, self._update_meta_preview)

    def _on_row_click(self, ri: int):
        self.current_row = ri
        if self.process_mode.get() == "Inteligente" and self.grid.df is not None:
            folder = self.img_folder_var.get()
            if folder:
                row = self.grid.df.iloc[ri]
                for col in self.grid.df.columns:
                    img_name = str(row[col]).strip()
                    if img_name and img_name.lower() not in ("nan", "none", ""):
                        img_path = self._find_image(img_name, folder)
                        if img_path:
                            self._load_image(img_path)
                            self.browser.highlight(Path(img_path).name)
                            break
        self._update_meta_preview()


    def _get_stem_index(self):
        df = self.grid.df
        if df is None:
            return {}, {}
        img_col = self.img_col_var.get()
        cache = getattr(self, "_stem_index_cache", None)
        if cache is not None and cache[0] is df and cache[1] == img_col:
            return cache[2], cache[3]

        col_index = {}
        if img_col and img_col in df.columns:
            for ri in range(len(df)):
                val = str(df.iloc[ri][img_col]).strip()
                if val and val.lower() not in ("nan", "none", ""):
                    col_index.setdefault(self._full_stem(val).lower(), ri)

        all_index = {}
        for ri in range(len(df)):
            row = df.iloc[ri]
            for col in df.columns:
                val = str(row[col]).strip()
                if val and val.lower() not in ("nan", "none", ""):
                    all_index.setdefault(self._full_stem(val).lower(), ri)

        self._stem_index_cache = (df, img_col, col_index, all_index)
        return col_index, all_index

    def _on_img_select(self, path: str):
        self._load_image(path)
        if self.process_mode.get() != "Inteligente" or self.grid.df is None:
            return

        img_stem = self._full_stem(Path(path).name).lower()
        col_index, all_index = self._get_stem_index()
        ri = col_index.get(img_stem)
        if ri is None:
            ri = all_index.get(img_stem)
        if ri is not None:
            self.current_row = ri
            self.grid.clear_selection()
            self.grid.select_row(ri)
            self.grid.scroll_to_row(ri)
            self._update_meta_preview()

    def _select_active_row(self):
        if self.current_row is None:
            return messagebox.showinfo("Sin fila activa", "Haz clic en una fila.")
        self.grid.select_row(self.current_row)

    def _clear_selection(self): self.grid.clear_selection()

    def _load_image(self, path: str, update_loupe=True):
        self.current_img = path
        self.img_name_lbl.configure(text=f"📄 {Path(path).name}")
        if update_loupe and hasattr(self, "loupe_window") and self.loupe_window.winfo_exists():
            self._load_loupe_image(path)
        if not PIL_OK: return
        try:
            img = ImageOps.exif_transpose(Image.open(path))
            self._preview_pil   = img.copy()
            self._preview_zoom  = 1.0
            self._preview_pan_x = 0
            self._preview_pan_y = 0
            self._preview_resize_cache = None
            # Miniatura pre-reducida usada solo para el redibujo "rápido"
            # (rueda/arrastre). Evita redimensionar la imagen a full
            # resolución en cada evento de la rueda del mouse.
            _fw, _fh = self._preview_pil.size
            _fmax = 1400
            if max(_fw, _fh) > _fmax:
                _fscale = _fmax / max(_fw, _fh)
                self._preview_pil_fast = self._preview_pil.resize(
                    (max(1, int(_fw * _fscale)), max(1, int(_fh * _fscale))), Image.BILINEAR)
            else:
                self._preview_pil_fast = self._preview_pil
            self._redraw_preview_zoomed(fast=True)
            if hasattr(self, "_preview_hq_timer"):
                self.after_cancel(self._preview_hq_timer)
            self._preview_hq_timer = self.after(80, lambda: self._redraw_preview_zoomed(fast=False))
        except Exception as e:
            self.img_canvas.delete("all")
            self.img_canvas.create_text(
                int(155*self.current_scale), int(120*self.current_scale),
                text=f"No se pudo abrir:\n{e}", fill=C["err"],
                font=FONTS["LABEL"], justify="center")

    def _update_meta_preview(self):
        if self.current_row is None or self.grid.df is None: return
        meta = self.grid.get_row_metadata(self.current_row, self._img_col_idx())
        self.meta_txt.configure(state="normal")
        self.meta_txt.delete("1.0", "end")
        if not meta:
            self.meta_txt.insert("end", "(ningún campo seleccionado)")
            self.meta_txt.configure(state="disabled")
            return
        if self.meta_mode_organized.get():
            remaining = dict(meta)
            for group in META_GROUP_ORDER:
                items = {k: v for k, v in meta.items() if k in META_GROUPS[group]}
                if not items: continue
                self.meta_txt.insert("end", f"▸ {group}\n", "head")
                for k, v in items.items():
                    lock_icon = " [🔒]" if (self.locked_columns and k in self.locked_columns) else ""
                    self.meta_txt.insert("end", f"  {k}: {v}{lock_icon}\n")
                self.meta_txt.insert("end", "\n")
                for k in items: remaining.pop(k, None)
            if remaining:
                self.meta_txt.insert("end", "▸ Otros\n", "head")
                for k, v in remaining.items():
                    lock_icon = " [🔒]" if (self.locked_columns and k in self.locked_columns) else ""
                    self.meta_txt.insert("end", f"  {k}: {v}{lock_icon}\n")
        else:
            self.meta_txt.insert("end", "▸ Orden Original (Excel)\n", "head")
            for k, v in meta.items():
                lock_icon = " [🔒]" if (self.locked_columns and k in self.locked_columns) else ""
                self.meta_txt.insert("end", f"  {k}: {v}{lock_icon}\n")
        self.meta_txt.configure(state="disabled")

    def _img_col_idx(self) -> int:
        if self.grid.df is None: return 0
        col, cols = self.img_col_var.get(), list(self.grid.df.columns)
        return cols.index(col) if col in cols else 0

    # ─────────────────────────────────────────────────────────────
    #  SINCRONIZACIÓN DE ORDEN
    # ─────────────────────────────────────────────────────────────
    def _sync_excel_to_images(self):
        if self.grid.df is None:
            return messagebox.showwarning("Sin datos", "Carga un archivo Excel / CSV.")
        folder = self.img_folder_var.get()
        if not folder or not os.path.isdir(folder):
            return messagebox.showwarning("Sin carpeta", "Selecciona la carpeta de imágenes.")
        img_files = sorted(Path(folder).iterdir(),
                          key=lambda p: p.name.lower())
        img_files = [f for f in img_files if f.is_file() and f.suffix.lower() in IMG_EXTS]
        if not img_files:
            return messagebox.showwarning("Sin imágenes", "La carpeta no contiene imágenes compatibles.")

        img_col = self.img_col_var.get()
        if not img_col or img_col not in self.grid.df.columns:
            return messagebox.showwarning("Sin columna de imagen",
                "Selecciona la columna de imagen en la herramienta principal.")

        def get_clean_name(filename):
            name = filename
            while '.' in name:
                name = name.rsplit('.', 1)[0]
            return name.lower().strip()

        file_order = {}
        for i, p in enumerate(img_files):
            file_order[get_clean_name(p.name)] = i

        def _sort_key(row):
            val = str(row[img_col]).strip()
            if not val or val.lower() in ("nan", "none", ""):
                return 9999
            clean = get_clean_name(val)
            if clean in file_order:
                return file_order[clean]
            for fk, fi in file_order.items():
                if clean.startswith(fk) or fk.startswith(clean):
                    return fi
            return 9999

        new_df = self.grid.df.copy()
        new_df["_sort_key"] = new_df.apply(_sort_key, axis=1)
        new_df = new_df.sort_values("_sort_key").drop(columns=["_sort_key"]).reset_index(drop=True)
        self.grid.df = new_df
        self.grid.load(self.grid.df)
        self.status_var.set(f"✓ Excel reordenado según imágenes ({len(self.grid.df)} filas)")
        self._log(f"✓ Excel reordenado según orden de imágenes\n", "ok")

    def _sync_images_to_excel(self):
        """
        Hace que las imágenes en el explorador sigan EXACTAMENTE el orden
        de filas que ya tiene el Excel cargado — sin reordenar ni una sola
        fila del Excel. Usa matching EXACTO primero (nombre de archivo
        idéntico o stem idéntico) para evitar que la búsqueda tolerante
        haga que dos filas distintas colisionen en la misma imagen.
        Solo si el matching exacto falla, intenta la búsqueda tolerante
        de _find_image, marcando esos casos como "aproximados" en el log.
        """
        if self.grid.df is None:
            return messagebox.showwarning("Sin datos", "Carga un archivo Excel / CSV.")
        folder = self.img_folder_var.get()
        if not folder or not os.path.isdir(folder):
            return messagebox.showwarning("Sin carpeta", "Selecciona la carpeta de imágenes.")

        img_col = self.img_col_var.get()
        if not img_col or img_col not in self.grid.df.columns:
            return messagebox.showwarning("Sin columna de imagen",
                "Selecciona la columna de imagen en la herramienta principal.")

        df_actual = self.grid.df

        # ── Índice de archivos en disco, por nombre exacto y por stem ──
        all_files = [f for f in Path(folder).iterdir()
                     if f.is_file() and f.suffix.lower() in IMG_EXTS]
        by_exact_name = {f.name.lower(): f for f in all_files}
        by_stem: dict = {}
        for f in all_files:
            key = self._safe_stem(f.name).lower()
            by_stem.setdefault(key, []).append(f)

        ordered_files     = []
        used_files        = set()     # rutas ya asignadas a una fila (evita colisiones)
        no_encontradas    = []
        matches_aprox     = []        # filas que solo encontraron match por tolerancia

        for ri in range(len(df_actual)):
            val = str(df_actual.iloc[ri][img_col]).strip()
            if not val or val.lower() in ("nan", "none", ""):
                continue

            found_path = None

            # 1) Coincidencia EXACTA por nombre completo
            cand = by_exact_name.get(val.lower())
            if cand and cand not in used_files:
                found_path = cand

            # 2) Coincidencia EXACTA por stem (sin extensión de imagen)
            if found_path is None:
                stem = self._safe_stem(val).lower()
                candidates = by_stem.get(stem, [])
                for cand in candidates:
                    if cand not in used_files:
                        found_path = cand
                        break

            # 3) Solo si lo anterior falló, usar la búsqueda tolerante
            if found_path is None:
                approx = self._find_image(val, folder)
                if approx:
                    approx_path = Path(approx)
                    if approx_path not in used_files:
                        found_path = approx_path
                        matches_aprox.append((val, found_path.name))

            if found_path:
                ordered_files.append(found_path)
                used_files.add(found_path)
            else:
                no_encontradas.append(val)

        if not ordered_files:
            return messagebox.showwarning("Sin coincidencias",
                "No se encontró ninguna imagen que coincida con la columna seleccionada.")

        # ── Imágenes huérfanas: existen en disco pero NUNCA fueron asignadas ──
        huerfanas = sorted(
            [f for f in all_files if f not in used_files],
            key=lambda p: p.name.lower())

        # Construir razón explicativa por cada huérfana
        orphan_reasons = {}
        no_encontradas_set = {self._full_stem(v).lower() for v in no_encontradas}
        for f in huerfanas:
            stem_f = self._full_stem(f.name).lower()
            if stem_f in no_encontradas_set:
                orphan_reasons[f.name] = "el Excel tiene esta fila pero el nombre no coincidió exactamente"
            else:
                orphan_reasons[f.name] = "no existe ninguna fila en el Excel con este nombre"

        self._sync_excel_count  = len(ordered_files)
        self._sync_orphan_count = len(huerfanas)

        ordered_files.extend(huerfanas)

        self.browser.img_files     = ordered_files
        self.browser._excel_count  = self._sync_excel_count
        self.browser._orphan_files = {f.name for f in huerfanas}
        self.browser._orphan_reasons = orphan_reasons
        self.browser._filter()

        msg = (f"{len(ordered_files)} imágenes  "
               f"({self._sync_excel_count} del Excel"
               + (f" + {self._sync_orphan_count} sin fila" if huerfanas else "")
               + ")")
        self.browser.info_lbl.configure(text=msg)
        self.status_var.set("✓ Imágenes reordenadas según el orden actual del Excel")

        self._log(
            f"✓ Imágenes reordenadas para coincidir, fila por fila, con el "
            f"orden ACTUAL del Excel (sin reordenar el Excel).\n"
            f"  • {self._sync_excel_count} imágenes con fila en el Excel "
            f"({self._sync_excel_count - len(matches_aprox)} exactas, "
            f"{len(matches_aprox)} por similitud)\n", "ok")

        if matches_aprox:
            self._log(
                f"  ⚠ {len(matches_aprox)} fila(s) NO tuvieron nombre de archivo "
                f"idéntico — se emparejaron por similitud. Verifica que sean "
                f"correctas:\n", "warn")
            for excel_val, file_name in matches_aprox[:10]:
                self._log(f"     '{excel_val}'  →  {file_name}\n", "warn")
            if len(matches_aprox) > 10:
                self._log(f"     … y {len(matches_aprox)-10} más.\n", "warn")

        if huerfanas:
            self._log(
                f"  ⚠ {len(huerfanas)} imagen(es) marcadas [sin fila]:\n", "warn")
            for f in huerfanas[:12]:
                razon = orphan_reasons.get(f.name, "razón desconocida")
                self._log(f"     • {f.name}\n       → {razon}\n", "warn")
            if len(huerfanas) > 12:
                self._log(f"     … y {len(huerfanas)-12} más.\n", "warn")
        if no_encontradas:
            self._log(
                f"  ⚠ {len(no_encontradas)} valores de '{img_col}' en el Excel "
                f"no tuvieron imagen correspondiente en disco: "
                f"{', '.join(no_encontradas[:8])}"
                f"{' …' if len(no_encontradas) > 8 else ''}\n", "warn")

    def _pick_sort_columns(self, all_cols):  # -> Optional[list]
        S = C
        sc = self.current_scale
        win = tk.Toplevel(self)
        win.title("Columnas de ordenamiento")
        win.configure(bg=S["bg"])
        win.geometry(f"{int(400*sc)}x{int(380*sc)}")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        win.grid_rowconfigure(2, weight=1)
        win.grid_columnconfigure(0, weight=1)

        result = [None]
        col_vars = {}
        def on_ok():
            result[0] = [c for c, v in col_vars.items() if v.get()]
            win.destroy()
        def on_cancel():
            win.destroy()

        hdr = tk.Frame(win, bg=S["header_bg"])
        hdr.grid(row=0, column=0, sticky="ew")
        tk.Label(hdr, text="  Seleccionar columnas de orden", bg=S["header_bg"],
                 fg=S["header_fg"], font=FONTS["H2"]).pack(side="left", pady=10, padx=8)

        tk.Label(win, text="Elige las columnas por las que se ordenarán las imágenes (en el orden que selecciones)",
                 bg=S["bg"], fg=S["text3"], font=FONTS["BODY"],
                 wraplength=int(360*sc)).grid(row=1, column=0, pady=(10, 4), padx=14)

        list_frame = tk.Frame(win, bg=S["surface"], highlightbackground=S["border"],
                              highlightthickness=1)
        list_frame.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 4))

        canvas = tk.Canvas(list_frame, bg=S["surface"], highlightthickness=0)
        canvas.pack_propagate(False)
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=S["surface"])
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def _sync_scroll(e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _sync_scroll)
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def _on_mousewheel_linux(event):
            if event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")
        for w in (canvas, inner, list_frame):
            w.bind("<MouseWheel>", _on_mousewheel)
            w.bind("<Button-4>", _on_mousewheel_linux)
            w.bind("<Button-5>", _on_mousewheel_linux)

        def _bind_all_children(widget):
            widget.bind("<MouseWheel>", _on_mousewheel)
            widget.bind("<Button-4>",   _on_mousewheel_linux)
            widget.bind("<Button-5>",   _on_mousewheel_linux)

        def _update_vsb(*_):
            lo, hi = canvas.yview()
            if lo <= 0.0 and hi >= 1.0:
                vsb.pack_forget()
            else:
                if not vsb.winfo_ismapped():
                    vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.bind("<Configure>", lambda e: (inner.update_idletasks(), _update_vsb()))
        inner.bind("<Configure>",  lambda e: _update_vsb())
        for idx, col in enumerate(all_cols):
            row_bg = S["row_even"] if idx % 2 == 0 else S["row_odd"]
            var = tk.BooleanVar(value=False)
            col_vars[col] = var

            row = tk.Frame(inner, bg=row_bg, cursor="hand2")
            row.pack(fill="x")

            indicator = tk.Frame(row, bg=S["accent"], width=4)
            indicator.pack(side="left", fill="y")

            cb = tk.Checkbutton(row, text=f"  {col}", variable=var,
                                bg=row_bg, fg=S["text"],
                                selectcolor=S["surface"],
                                activebackground=row_bg,
                                font=FONTS["BODY"], anchor="w",
                                cursor="hand2")
            cb.pack(side="left", fill="x", expand=True, padx=4, pady=5)

            def _enter(e, r=row, ind=indicator):
                r.configure(bg=S["accent_pale"])
                ind.configure(bg=S["accent_hover"])
            def _leave(e, r=row, ind=indicator, bg=row_bg):
                r.configure(bg=bg)
                ind.configure(bg=S["accent"])
            for w in (row, cb):
                w.bind("<Enter>", _enter)
                w.bind("<Leave>", _leave)

        sep = tk.Frame(win, bg=S["border"], height=1)
        sep.grid(row=3, column=0, sticky="ew")
        btn_frame = tk.Frame(win, bg=S["bg"])
        btn_frame.grid(row=4, column=0, sticky="ew", padx=14, pady=10)
        tk.Button(btn_frame, text="Cancelar", bg=S["btn_ghost_bg"], fg=S["text"],
                  font=FONTS["LABEL"], relief="flat", cursor="hand2",
                  command=on_cancel).pack(side="right", ipady=4)
        tk.Button(btn_frame, text="Aceptar", bg=S["accent"], fg="#FFF5E8",
                  font=FONTS["LABEL_B"], relief="flat", cursor="hand2",
                  activebackground=S["accent_hover"],
                  command=on_ok).pack(side="right", padx=(0, 8), ipady=4)

        self.wait_window(win)
        return result[0]

    # ─────────────────────────────────────────────────────────────
    #  LOTE POR ORDEN (NUEVA FEATURE v8.9)
    #  Fila 1 del Excel → Foto 1 (alfabética), etc.
    #  Optimizado para 300 fotos: hilo independiente + actualizaciones
    #  de barra cada 5 imágenes.
    # ─────────────────────────────────────────────────────────────
    def _batch_write_by_order(self):
        """Abre diálogos para seleccionar Excel y carpeta, y escribe metadatos por posición."""
        if not PIL_OK:
            return messagebox.showerror("Error", "Instala: pip install pillow piexif")

        # ── Opción A: usar el Excel ya cargado ──
        use_loaded = False
        if self.df is not None and len(self.df) > 0:
            use_loaded = messagebox.askyesno(
                "¿Usar tabla cargada?",
                f"Ya tienes un archivo cargado con {len(self.df)} filas.\n\n"
                "¿Usar esos datos para el lote por orden?\n\n"
                "• Sí → usa la tabla actual\n"
                "• No → selecciona un Excel nuevo")

        if use_loaded:
            batch_df = self.df.copy()
        else:
            excel_path = _native_file_open(
                title="Seleccionar Excel con metadatos",
                filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("CSV", "*.csv"), ("Todos", "*.*")])
            if not excel_path: return
            try:
                ext      = Path(excel_path).suffix.lower()
                batch_df = (pd.read_csv(excel_path, dtype=str, keep_default_na=False, na_values=[])
                            if ext == ".csv"
                            else pd.read_excel(excel_path, dtype=str, keep_default_na=False, na_values=[]))
                batch_df = batch_df.map(lambda v: "" if (
                    v is None or
                    (isinstance(v, float) and str(v).lower() == "nan") or
                    str(v).strip().lower() in ("nat", "none")
                ) else str(v).strip())
                batch_df.columns = [str(c).strip() for c in batch_df.columns]
            except Exception as e:
                return messagebox.showerror("Error al leer Excel", str(e))

        if len(batch_df) == 0:
            return messagebox.showwarning("Excel vacío", "El archivo no contiene filas de datos.")

        # ── Seleccionar carpeta de fotos ──
        folder = self.img_folder_var.get()
        if not folder or not os.path.isdir(folder):
            folder = _native_folder_open(title="Carpeta con las fotos")
            if not folder: return
            self.img_folder_var.set(folder)

        # ── SELECCIÓN DE COLUMNAS ──
        img_col  = self.img_col_var.get()
        all_cols = list(batch_df.columns)
        meta_cols = self._batch_pick_columns(all_cols, img_col)
        if meta_cols is None:
            return

        # ── ORDEN DE FOTOS ──
        sort_mode = self._batch_pick_sort()
        if sort_mode is None:
            return

        img_files = self._sort_images(folder, sort_mode)
        if not img_files:
            return messagebox.showwarning("Sin imágenes",
                "La carpeta no contiene imágenes compatibles.")

        total = min(len(img_files), len(batch_df))

        if len(img_files) != len(batch_df):
            if not messagebox.askyesno(
                "Conteo diferente",
                f"Imágenes en carpeta: {len(img_files)}\n"
                f"Filas en Excel:      {len(batch_df)}\n\n"
                f"Se procesarán las primeras {total} parejas. ¿Continuar?"
            ): return

        omit_empty = self.omit_empty_var.get()
        organizado = self.meta_mode_organized.get()

        # ── REANUDAR LOTE ──
        progress_file = self.output_folder / "_batch_progress.json"
        start_idx = 0
        ok_count  = 0
        errors    = []
        if progress_file.exists():
            try:
                prev = json.loads(progress_file.read_text(encoding="utf-8"))
                if prev.get("total") == total and prev.get("excel") == str(self.csv_path_var.get()):
                    done_count = prev.get("done", 0)
                    if done_count > 0 and done_count < total:
                        if messagebox.askyesno(
                            "Lote anterior encontrado",
                            f"Se procesaron {done_count}/{total} imágenes anteriormente.\n\n"
                            f"¿Continuar desde la imagen {done_count + 1}?"
                        ):
                            start_idx = done_count
                            ok_count  = prev.get("ok", 0)
                            errors    = prev.get("errors", [])
            except Exception:
                pass

        self.output_folder.mkdir(parents=True, exist_ok=True)

        # ── Ventana de progreso ──
        remaining = total - start_idx
        prog_win = tk.Toplevel(self)
        prog_win.title("Lote por Orden — Escribiendo…")
        prog_win.configure(bg=C["panel"])
        prog_win.geometry("480x200")
        prog_win.resizable(False, False)
        prog_win.grab_set()

        tk.Label(prog_win, text="Escribiendo metadatos en lote por orden",
                 bg=C["panel"], fg=C["accent"], font=FONTS["H2"]).pack(pady=(18, 4))

        lbl_file = tk.Label(prog_win, text="Iniciando…",
                            bg=C["panel"], fg=C["text2"], font=FONTS["LABEL"])
        lbl_file.pack()

        lbl_count = tk.Label(prog_win, text=f"{start_idx} / {total}",
                             bg=C["panel"], fg=C["text3"], font=FONTS["TINY"])
        lbl_count.pack()

        bar_outer = tk.Frame(prog_win, bg=C["border"], height=10)
        bar_outer.pack(fill="x", padx=24, pady=12)
        bar_fill  = tk.Frame(bar_outer, bg=C["accent"], height=10)
        bar_fill.place(x=0, y=0, relheight=1, width=0)

        prog_win.update_idletasks()
        bar_w = bar_outer.winfo_width() or 432

        def update_ui(done: int, filename: str):
            frac = done / total
            bar_fill.place(width=int(bar_w * frac))
            lbl_file.configure(text=filename)
            lbl_count.configure(text=f"{done} / {total}")
            prog_win.update_idletasks()

        cancelled = threading.Event()

        def on_cancel():
            cancelled.set()
            prog_win.destroy()

        btn_cancel = tk.Button(prog_win, text="Cancelar", bg=C["btn_ghost_bg"],
                               fg=C["text"], font=FONTS["LABEL"], relief="flat",
                               cursor="hand2", command=on_cancel)
        btn_cancel.pack(pady=(0, 8))

        def worker():
            nonlocal ok_count, errors

            for i in range(start_idx, total):
                if cancelled.is_set():
                    break
                img_path = img_files[i]
                row      = batch_df.iloc[i]

                meta = {}
                for col in meta_cols:
                    val = str(row[col]).strip()
                    if val.lower() in ("nan", "none", ""): val = ""
                    if omit_empty and not val: continue
                    meta[col] = val

                for col in self.locked_columns:
                    if col in batch_df.columns and col not in meta:
                        val = str(row[col]).strip()
                        if val.lower() in ("nan", "none", ""): val = ""
                        if omit_empty and not val: continue
                        meta[col] = val

                if not meta:
                    if (i + 1) % 5 == 0 or (i + 1) == total:
                        self.after(0, update_ui, i + 1, img_path.name)
                    continue

                out_path = self.output_folder / img_path.name
                try:
                    divergencias = self._check_metadata_divergence(str(img_path), meta)
                    if divergencias and (i + 1) % 5 == 0:
                        self._log(
                            f"  ⚠ {img_path.name}: divergencia detectada y corregida\n", "warn")
                    shutil.copy2(str(img_path), str(out_path))
                    self._write_meta(str(out_path), meta, organizado)
                    ok_count += 1
                except Exception as e:
                    errors.append(f"[{img_path.name}]: {e}")

                if (i + 1) % 5 == 0 or (i + 1) == total:
                    self.after(0, update_ui, i + 1, img_path.name)

                # ── Guardar progreso cada 10 fotos ──
                if (i + 1) % 10 == 0:
                    progress_data = {
                        "total": total, "done": i + 1, "ok": ok_count,
                        "errors": errors, "excel": str(self.csv_path_var.get()),
                        "sort": sort_mode, "meta_cols": meta_cols,
                    }
                    progress_file.write_text(
                        json.dumps(progress_data, ensure_ascii=False, indent=2),
                        encoding="utf-8")

            # Guardar progreso final
            progress_data = {
                "total": total, "done": total, "ok": ok_count,
                "errors": errors, "excel": str(self.csv_path_var.get()),
                "sort": sort_mode, "meta_cols": meta_cols,
            }
            progress_file.write_text(
                json.dumps(progress_data, ensure_ascii=False, indent=2),
                encoding="utf-8")

            def finish():
                if cancelled.is_set():
                    self._log(f"\n⏹ Lote cancelado: {ok_count}/{total} procesadas.\n", "info")
                    self.status_var.set(f"⏹ Lote cancelado: {ok_count}/{total}")
                else:
                    try: progress_file.unlink()
                    except Exception: pass
                    summary = (f"✅ {ok_count} de {total} imágenes procesadas correctamente.\n"
                               f"📁 Guardadas en: {self.output_folder}")
                    if errors:
                        summary += f"\n\n⚠ {len(errors)} error(es):\n"
                        summary += "\n".join(errors[:8])
                        if len(errors) > 8:
                            summary += f"\n… y {len(errors)-8} más."
                    messagebox.showinfo("Lote por Orden — Completado", summary)
                    self._log(f"\n✔ Lote por orden: {ok_count}/{total} procesadas. "
                              f"{len(errors)} errores.\n", "head")
                    self.status_var.set(f"✔ Lote por orden: {ok_count}/{total} listas")
                if self.current_img:
                    cur_name = Path(self.current_img).name
                    processed_names = {img_files[j].name for j in range(total)}
                    if cur_name in processed_names:
                        out = self.output_folder / cur_name
                        if out.exists(): self._load_image(str(out), update_loupe=False)
                prog_win.destroy()

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    # ── Selección de columnas para lote ──
    def _batch_pick_columns(self, all_cols, img_col=""):  # -> Optional[list]
        S = C
        sc = self.current_scale
        win = tk.Toplevel(self)
        win.title("Seleccionar columnas de metadatos")
        win.configure(bg=S["bg"])
        win.geometry(f"{int(400*sc)}x{int(420*sc)}")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        win.grid_rowconfigure(2, weight=1)
        win.grid_columnconfigure(0, weight=1)

        col_vars = {}
        result = [None]
        def on_ok():
            result[0] = [c for c, v in col_vars.items() if v.get()]
            win.destroy()
        def on_cancel():
            win.destroy()
        def select_all():
            for c, v in col_vars.items(): v.set(True)
            _update_count()
        def deselect_all():
            for c, v in col_vars.items(): v.set(False)
            _update_count()
        def invert_sel():
            for c, v in col_vars.items():
                if c != img_col: v.set(not v.get())
            _update_count()

        hdr = tk.Frame(win, bg=S["header_bg"])
        hdr.grid(row=0, column=0, sticky="ew")
        tk.Label(hdr, text="  Columnas de metadatos", bg=S["header_bg"],
                 fg=S["header_fg"], font=FONTS["H2"]).pack(side="left", pady=10, padx=8)

        top_info = tk.Frame(win, bg=S["bg"])
        top_info.grid(row=1, column=0, sticky="ew", padx=14)
        tk.Label(top_info, text="Selecciona las columnas que se escribirán en los metadatos",
                 bg=S["bg"], fg=S["text3"], font=FONTS["TINY"],
                 wraplength=int(360*sc)).pack(pady=(12, 4))
        sel_count_var = tk.StringVar(value=f"{len([c for c in all_cols if c != img_col])} / {len(all_cols)}")
        tk.Label(top_info, textvariable=sel_count_var, bg=S["bg"], fg=S["accent"],
                 font=FONTS["LABEL_B"]).pack(pady=(0, 6))
        actions = tk.Frame(top_info, bg=S["bg"])
        actions.pack()
        tk.Button(actions, text="Todas", bg=S["btn_ghost_bg"], fg=S["accent"],
                  font=FONTS["TINY"], relief="flat", bd=0, cursor="hand2",
                  command=select_all).pack(side="left", padx=(0, 4))
        tk.Button(actions, text="Ninguna", bg=S["btn_ghost_bg"], fg=S["text3"],
                  font=FONTS["TINY"], relief="flat", bd=0, cursor="hand2",
                  command=deselect_all).pack(side="left", padx=(0, 4))
        tk.Button(actions, text="Invertir", bg=S["btn_ghost_bg"], fg=S["text3"],
                  font=FONTS["TINY"], relief="flat", bd=0, cursor="hand2",
                  command=invert_sel).pack(side="left")

        list_frame = tk.Frame(win, bg=S["surface"], highlightbackground=S["border"],
                              highlightthickness=1)
        list_frame.grid(row=2, column=0, sticky="nsew", padx=14, pady=(4, 4))

        canvas = tk.Canvas(list_frame, bg=S["surface"], highlightthickness=0)
        canvas.pack_propagate(False)
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=S["surface"])
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def _on_mousewheel_linux(event):
            if event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")
        for w in (canvas, inner, list_frame):
            w.bind("<MouseWheel>", _on_mousewheel)
            w.bind("<Button-4>", _on_mousewheel_linux)
            w.bind("<Button-5>", _on_mousewheel_linux)

        def _bind_all_children(widget):
            widget.bind("<MouseWheel>", _on_mousewheel)
            widget.bind("<Button-4>",   _on_mousewheel_linux)
            widget.bind("<Button-5>",   _on_mousewheel_linux)

        def _update_vsb(*_):
            lo, hi = canvas.yview()
            if lo <= 0.0 and hi >= 1.0:
                vsb.pack_forget()
            else:
                if not vsb.winfo_ismapped():
                    vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.bind("<Configure>", lambda e: (inner.update_idletasks(), _update_vsb()))
        inner.bind("<Configure>",  lambda e: _update_vsb())

        def _update_count():
            n = sum(1 for v in col_vars.values() if v.get())
            sel_count_var.set(f"{n} / {len(all_cols)}")

        for idx, col in enumerate(all_cols):
            is_img = (col == img_col)
            var = tk.BooleanVar(value=(not is_img))
            col_vars[col] = var

            row_bg = S["row_even"] if idx % 2 == 0 else S["row_odd"]
            row = tk.Frame(inner, bg=row_bg)
            row.pack(fill="x")

            var.trace_add("write", lambda *_: _update_count())

            if is_img:
                indicator = tk.Frame(row, bg=S["text3"], width=4)
                indicator.pack(side="left", fill="y")
                cb = tk.Checkbutton(row, text=f"  {col}", variable=var,
                                    bg=row_bg, fg=S["text3"],
                                    selectcolor=S["surface"],
                                    activebackground=row_bg,
                                    font=FONTS["BODY"], anchor="w",
                                    cursor="hand2", disabledforeground=S["text3"])
                cb.pack(side="left", fill="x", expand=True, padx=4, pady=5)
                tk.Label(row, text="imagen", bg=row_bg, fg=S["text3"],
                         font=FONTS["TINY"]).pack(side="right", padx=8)
            else:
                indicator = tk.Frame(row, bg=S["accent"], width=4)
                indicator.pack(side="left", fill="y")
                cb = tk.Checkbutton(row, text=f"  {col}", variable=var,
                                    bg=row_bg, fg=S["text"],
                                    selectcolor=S["surface"],
                                    activebackground=row_bg,
                                    font=FONTS["BODY"], anchor="w",
                                    cursor="hand2")
                cb.pack(side="left", fill="x", expand=True, padx=4, pady=5)

                def _enter(e, r=row, ind=indicator):
                    r.configure(bg=S["accent_pale"])
                    ind.configure(bg=S["accent_hover"])
                def _leave(e, r=row, ind=indicator, bg=row_bg):
                    r.configure(bg=bg)
                    ind.configure(bg=S["accent"])
                row.bind("<Enter>", _enter)
                row.bind("<Leave>", _leave)
                cb.bind("<Enter>", _enter)
                cb.bind("<Leave>", _leave)
            _bind_all_children(row)

        def _sync_scroll(e=None):
            inner.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _sync_scroll)
        canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win.after(100, _sync_scroll)

        sep = tk.Frame(win, bg=S["border"], height=1)
        sep.grid(row=3, column=0, sticky="ew")

        btn_frame = tk.Frame(win, bg=S["bg"])
        btn_frame.grid(row=4, column=0, sticky="ew", padx=14, pady=10)
        tk.Button(btn_frame, text="Cancelar", bg=S["btn_ghost_bg"], fg=S["text"],
                  font=FONTS["LABEL"], relief="flat", cursor="hand2",
                  command=on_cancel).pack(side="right", ipady=4)
        tk.Button(btn_frame, text="Aceptar", bg=S["accent"], fg="#FFF5E8",
                  font=FONTS["LABEL_B"], relief="flat", cursor="hand2",
                  activebackground=S["accent_hover"],
                  command=on_ok).pack(side="right", padx=(0, 8), ipady=4)

        self.wait_window(win)
        return result[0]

    # ── Selección de orden de fotos ──
    def _batch_pick_sort(self):  # -> Optional[str]
        S = C
        sc = self.current_scale
        win = tk.Toplevel(self)
        win.title("Orden de las fotos")
        win.configure(bg=S["bg"])
        win.geometry(f"{int(420*sc)}x{int(320*sc)}")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        win.grid_rowconfigure(2, weight=1)
        win.grid_columnconfigure(0, weight=1)

        sort_var = tk.StringVar(value="alfabetico")
        options = [
            ("orden_excel",   "Orden del Excel",              "Respeta el orden de las filas tal como están"),
            ("alfabetico",    "Alfabético (A → Z)",           "Ordena por nombre de archivo"),
            ("fecha_mod",     "Fecha de modificación ↑",      "Más antigua primero"),
            ("fecha_mod_inv", "Fecha de modificación ↓",      "Más reciente primero"),
            ("fecha_exif",    "Fecha EXIF de la foto",        "Usa la fecha grabada en la foto"),
            ("numeral",       "Por número en el nombre",      "1, 2, 10, 20… en vez de 1, 10, 2, 20"),
        ]

        result = [None]
        def on_ok():
            result[0] = sort_var.get()
            win.destroy()
        def on_cancel():
            win.destroy()

        hdr = tk.Frame(win, bg=S["header_bg"])
        hdr.grid(row=0, column=0, sticky="ew")
        tk.Label(hdr, text="  Orden de las fotos", bg=S["header_bg"],
                 fg=S["header_fg"], font=FONTS["H2"]).pack(side="left", pady=10, padx=8)

        tk.Label(win, text="¿Cómo quieres ordenar las imágenes antes de emparejarlas?",
                 bg=S["bg"], fg=S["text3"], font=FONTS["BODY"],
                 wraplength=int(380*sc)).grid(row=1, column=0, pady=(14, 8), padx=14)

        list_frame = tk.Frame(win, bg=S["surface"], highlightbackground=S["border"],
                              highlightthickness=1)
        list_frame.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 4))

        canvas = tk.Canvas(list_frame, bg=S["surface"], highlightthickness=0)
        canvas.pack_propagate(False)
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=S["surface"])
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)

        def _sync_scroll(e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _sync_scroll)
        canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def _on_mousewheel_linux(event):
            if event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")
        for w in (canvas, inner, list_frame):
            w.bind("<MouseWheel>", _on_mousewheel)
            w.bind("<Button-4>", _on_mousewheel_linux)
            w.bind("<Button-5>", _on_mousewheel_linux)

        row_refs = []
        for idx, (val, label, desc) in enumerate(options):
            row_bg = S["row_even"] if idx % 2 == 0 else S["row_odd"]
            row = tk.Frame(inner, bg=row_bg, cursor="hand2")
            row.pack(fill="x")

            indicator = tk.Frame(row, bg=S["accent"], width=5)
            indicator.pack(side="left", fill="y")

            mark_lbl = tk.Label(row, text="  ", bg=row_bg, fg=S["accent"],
                                font=FONTS["BODY"], width=2, anchor="center")
            mark_lbl.pack(side="left", padx=(4, 0))

            text_frame = tk.Frame(row, bg=row_bg, cursor="hand2")
            text_frame.pack(side="left", fill="x", expand=True, pady=6)
            tk.Label(text_frame, text=label, bg=row_bg, fg=S["text"],
                     font=FONTS["BODY"], anchor="w").pack(anchor="w")
            tk.Label(text_frame, text=desc, bg=row_bg, fg=S["text3"],
                     font=FONTS["TINY"], anchor="w").pack(anchor="w")

            row_refs.append((val, row, indicator, mark_lbl, row_bg))

            def _enter(e, r=row, ind=indicator):
                r.configure(bg=S["accent_pale"])
                ind.configure(bg=S["accent_hover"])
            def _leave(e, r=row, ind=indicator, bg=row_bg, v=val):
                sel = sort_var.get()
                if v != sel:
                    r.configure(bg=bg)
                    ind.configure(bg=S["accent"])
            def _click(e, v=val):
                sort_var.set(v)
            for w in (row, text_frame, mark_lbl):
                w.bind("<Enter>", _enter)
                w.bind("<Leave>", _leave)
                w.bind("<Button-1>", _click)
            for lbl in text_frame.winfo_children():
                lbl.bind("<MouseWheel>",  _on_mousewheel)
                lbl.bind("<Button-4>",    _on_mousewheel_linux)
                lbl.bind("<Button-5>",    _on_mousewheel_linux)
                lbl.bind("<Button-1>",    _click)

        def _highlight_selected(*_):
            sel = sort_var.get()
            for val, row, ind, mark, bg in row_refs:
                if val == sel:
                    row.configure(bg=S["accent_pale"])
                    ind.configure(bg=S["accent_hover"])
                    mark.configure(text="●", fg=S["accent"])
                else:
                    row.configure(bg=bg)
                    ind.configure(bg=S["accent"])
                    mark.configure(text="  ", fg=S["accent"])
        sort_var.trace_add("write", _highlight_selected)
        _highlight_selected()

        sep = tk.Frame(win, bg=S["border"], height=1)
        sep.grid(row=3, column=0, sticky="ew")
        btn_frame = tk.Frame(win, bg=S["bg"])
        btn_frame.grid(row=4, column=0, sticky="ew", padx=14, pady=10)
        tk.Button(btn_frame, text="Cancelar", bg=S["btn_ghost_bg"], fg=S["text"],
                  font=FONTS["LABEL"], relief="flat", cursor="hand2",
                  command=on_cancel).pack(side="right", ipady=4)
        tk.Button(btn_frame, text="Aceptar", bg=S["accent"], fg="#FFF5E8",
                  font=FONTS["LABEL_B"], relief="flat", cursor="hand2",
                  activebackground=S["accent_hover"],
                  command=on_ok).pack(side="right", padx=(0, 8), ipady=4)

        self.wait_window(win)
        return result[0]

    def _sort_images(self, folder: str, mode: str) -> list:
        files = [p for p in Path(folder).iterdir()
                 if p.is_file() and p.suffix.lower() in IMG_EXTS]

        if mode in ("alfabetico", "orden_excel"):
            return sorted(files, key=lambda p: p.name.lower())
        elif mode == "fecha_mod":
            return sorted(files, key=lambda p: p.stat().st_mtime)
        elif mode == "fecha_mod_inv":
            return sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)
        elif mode == "fecha_exif":
            def exif_date(p):
                try:
                    if not PIL_OK: return 0
                    img = Image.open(str(p))
                    exif = img._getexif()
                    if exif and 36867 in exif:
                        return exif[36867]
                except Exception:
                    pass
                return ""
            return sorted(files, key=exif_date)
        elif mode == "numeral":
            def extract_num(p):
                nums = re.findall(r'\d+', p.stem)
                return int(nums[0]) if nums else 0
            return sorted(files, key=extract_num)
        return sorted(files, key=lambda p: p.name.lower())

    # ─────────────────────────────────────────────────────────────
    #  PROCESAMIENTO NORMAL (MODO INTELIGENTE)
    # ─────────────────────────────────────────────────────────────
    def _start_processing(self):
        if self.grid.df is None:
            return messagebox.showwarning("Sin datos", "Carga un archivo Excel / CSV.")
        folder = self.img_folder_var.get()
        if not folder or not os.path.isdir(folder):
            return messagebox.showwarning("Sin carpeta", "Selecciona la carpeta de imágenes.")
        if not self.grid.selected_cells:
            return messagebox.showwarning("Sin selección", "Selecciona celdas en la tabla.")

        img_col_idx  = self._img_col_idx()
        meta_by_row  = self.grid.get_selected_metadata(img_col_idx)
        total_sel    = sum(1 for (r, c) in self.grid.selected_cells if c != img_col_idx)
        total_valid  = sum(len(m) for m in meta_by_row.values())

        if total_valid == 0:
            return messagebox.showwarning("Sin datos",
                "Las celdas seleccionadas están vacías o son omitidas.")

        self.config(cursor="watch")
        self._write_btn.configure(text="⏳  Procesando…", state="disabled")
        self.output_folder.mkdir(parents=True, exist_ok=True)
        self._log(f"\n📁 Salida Lote: {self.output_folder}\n", "info")
        self.progress_var.set(0)
        self.status_var.set("Procesando Lote…")
        self._save_config()
        threading.Thread(target=self._process_all, args=(folder, meta_by_row), daemon=True).start()

    def _find_img_name_in_row(self, ri: int) -> str:
        row = self.grid.df.iloc[ri]
        return self._find_img_name_in_row_data(row)

    def _find_img_name_in_row_data(self, row) -> str:
        for col in row.index:
            val = str(row[col]).strip()
            if val and val.lower() not in ("nan", "none", ""):
                if Path(val).suffix.lower() in IMG_EXTS:
                    return val
        return ""

    def _process_all(self, folder: str, meta_by_row: dict):
        rows_to_process    = sorted(meta_by_row.keys())
        total, ok, err     = len(rows_to_process) or 1, 0, 0
        organizado         = self.meta_mode_organized.get()

        for i, ri in enumerate(rows_to_process):
            img_name = self._find_img_name_in_row(ri)
            meta     = meta_by_row[ri]
            self.after(0, lambda i=i, t=total: self.status_var.set(f"Procesando {i+1}/{t}…"))

            if not img_name:
                self._update_progress((i + 1) / total * 100); continue

            img_path = self._find_image(img_name, folder)
            if not img_path:
                self._log(f"  ✗ No encontrada: {img_name}\n", "err")
                err += 1; self._update_progress((i + 1) / total * 100); continue

            out_path = self.output_folder / Path(img_path).name
            try:
                divergencias = self._check_metadata_divergence(img_path, meta)
                if divergencias:
                    self._log(
                        f"  ⚠ {Path(img_path).name}: metadatos previos NO coinciden "
                        f"con el Excel, se sobreescribirán:\n"
                        f"     {' | '.join(divergencias)}\n", "warn")

                shutil.copy2(img_path, out_path)
                self._write_meta(str(out_path), meta, organizado)
                self._log(
                    f"  ✔ {Path(img_path).name}\n"
                    f"     → {' | '.join(f'{k}: {v}' for k,v in meta.items())}\n", "ok")
                ok += 1
            except Exception as e:
                self._log(f"  ✗ {img_name}: {e}\n", "err")
                err += 1
            self._update_progress((i + 1) / total * 100)

        empty_cnt  = max(0, sum(1 for (r, c) in self.grid.selected_cells
                                if c != self._img_col_idx())
                         - sum(len(m) for m in meta_by_row.values()))
        empty_note = f" · {empty_cnt} omitidas" if empty_cnt > 0 and self.omit_empty_var.get() else ""
        self._log(f"\n── Completado: {ok} escritas · {err} errores{empty_note} ──\n", "head")
        self.after(0, lambda: (
            self.config(cursor=""),
            self._write_btn.configure(text="▶  Escribir Metadatos", state="normal"),
            self.status_var.set(f"✔ {ok} guardadas · {err} errores{empty_note}")
        ))

    def _inject_manual(self):
        if self.current_row is None or self.grid.df is None:
            return messagebox.showwarning("Falta fila", "Selecciona una fila.")
        if not self.current_img or not os.path.exists(self.current_img):
            return messagebox.showwarning("Falta imagen", "Selecciona una imagen.")

        img_col_idx = self._img_col_idx()
        meta        = self.grid.get_row_metadata(self.current_row, img_col_idx)
        organizado  = self.meta_mode_organized.get()

        self.config(cursor="watch")
        self._inject_btn.configure(text="⏳ Inyectando...", state="disabled")
        self.output_folder.mkdir(parents=True, exist_ok=True)
        out_path = self.output_folder / Path(self.current_img).name

        def task():
            try:
                divergencias = self._check_metadata_divergence(self.current_img, meta)
                if divergencias:
                    self._log(
                        f"  ⚠ {Path(self.current_img).name}: metadatos previos NO "
                        f"coinciden con el Excel, se sobreescribirán:\n"
                        f"     {' | '.join(divergencias)}\n", "warn")
                shutil.copy2(self.current_img, out_path)
                self._write_meta(str(out_path), meta, organizado)
                img_name = Path(self.current_img).name
                self._log(f"📌 Inyección Manual Exitosa:\n"
                          f"   ✔ Fila {self.current_row+1} ➔ [{img_name}]\n", "ok")
                self.after(0, lambda: self.status_var.set(f"✔ Inyección en {img_name} completada"))
            except Exception as e:
                self._log(f"  ✗ Error inyectando en {Path(self.current_img).name}: {e}\n", "err")
                self.after(0, lambda: self.status_var.set("✗ Error en inyección"))
            finally:
                self.after(0, lambda: (
                    self.config(cursor=""),
                    self._inject_btn.configure(text="📌 Inyectar en Foto Actual", state="normal")
                ))
        threading.Thread(target=task, daemon=True).start()

    # ─────────────────────────────────────────────────────────────
    #  BÚSQUEDA INTELIGENTE DE IMAGEN
    # ─────────────────────────────────────────────────────────────
    _IMG_EXTS_LOWER = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".webp"}

    def _safe_stem(self, s: str) -> str:
        p = Path(s)
        changed = True
        while changed:
            changed = False
            if p.suffix.lower() in self._IMG_EXTS_LOWER:
                p = p.with_suffix("")
                changed = True
                continue
            # Despoja marcadores de duplicado tipo "nombre (1)", "nombre(2)"
            # que quedaban pegados al stem impidiendo el emparejamiento
            # (ej: "0006_UM_C4_IX_00034_P.jpg (1).JPG").
            m = re.match(r"^(.*?)\s*\(\d+\)$", p.name)
            if m and m.group(1):
                p = p.with_name(m.group(1))
                changed = True
        return p.name

    def _full_stem(self, s: str) -> str:
        return self._safe_stem(s)

    def _normalize_numbers(self, s: str) -> str:
        return re.sub(r"\d+", lambda m: str(int(m.group())), s)

    def _extract_id_suffix(self, s):  # -> Optional[tuple]
        """
        Extrae (numero_pieza, sufijo_vista) de un nombre, sin importar
        cuántos campos haya en medio, NI si el string trae extensión
        de archivo incluida (incluso doble, ej. '.jpg.JPG').
        '0001_UM_C4_UE18_00006_F.jpg' -> ('1', 'F')
        '79_EC_PS_VI_250_R'           -> ('79', 'R')
        '0061_EC_C4_III_046'          -> ('61', '')   <- sin sufijo de vista
        """
        s = s.strip()
        changed = True
        while changed:
            changed = False
            for ext in IMG_EXTS:
                if s.lower().endswith(ext):
                    s = s[: -len(ext)]
                    changed = True
            m_dup = re.match(r"^(.*?)\s*\(\d+\)$", s)
            if m_dup and m_dup.group(1):
                s = m_dup.group(1)
                changed = True
        s = s.lstrip("#").strip("_-").upper()
        m_num = re.match(r"^0*(\d+)", s)
        if not m_num:
            return None
        numero = m_num.group(1)
        m_suf = re.search(r"[FRP]$", s)
        sufijo = m_suf.group(0) if m_suf else ""
        return (numero, sufijo)

    def _find_image(self, name, folder):  # -> Optional[str]
        name = name.strip()
        if not name: return None
        folder_path = Path(folder)
        if not self._img_cache:
            self._img_cache = {self._full_stem(f.name).lower(): f
                               for f in folder_path.rglob("*")
                               if f.is_file() and f.suffix.lower() in IMG_EXTS}

        p = folder_path / name
        if p.exists(): return str(p)

        name_lower = name.lower()
        name_stem  = self._full_stem(name).lower()

        for _, fpath in self._img_cache.items():
            if fpath.name.lower() == name_lower: return str(fpath)
        if name_stem in self._img_cache: return str(self._img_cache[name_stem])

        name_clean = re.sub(r"^[#\s\-_]+|[#\s\-_]+$", "", name_stem)
        for stem_key, fpath in self._img_cache.items():
            if re.sub(r"^[#\s\-_]+|[#\s\-_]+$", "", stem_key) == name_clean:
                return str(fpath)

        name_normalized = self._normalize_numbers(name_clean)
        for stem_key, fpath in self._img_cache.items():
            stem_clean = re.sub(r"^[#\s\-_]+|[#\s\-_]+$", "", stem_key)
            if self._normalize_numbers(stem_clean) == name_normalized:
                return str(fpath)

        # ── ÚLTIMO RECURSO Y MÁS ROBUSTO: comparar solo número + sufijo ──
        # Ignora completamente cuántos campos tiene cada estructura de nombre.
        id_excel = self._extract_id_suffix(name)
        if id_excel:
            for stem_key, fpath in self._img_cache.items():
                id_archivo = self._extract_id_suffix(stem_key)
                if id_archivo and id_archivo == id_excel:
                    return str(fpath)

        for stem_key, fpath in self._img_cache.items():
            if name_stem in stem_key or stem_key in name_stem:
                return str(fpath)

        return None

    # ─────────────────────────────────────────────────────────────
    #  ESCRITURA DE METADATOS
    # ─────────────────────────────────────────────────────────────
    def _formatear_metadatos(self, meta: dict, organizado: bool = True) -> str:
        if not organizado:
            return "\n".join(f"{k}: {v}" for k, v in meta.items())
        partes, restantes = dict(meta), []
        for grupo in META_GROUP_ORDER:
            items = {k: v for k, v in meta.items() if k in META_GROUPS[grupo]}
            if not items: continue
            restantes.append(f"[{grupo}]")
            for k, v in items.items(): restantes.append(f"  {k}: {v}")
            for k in items: partes.pop(k, None)
        if partes:
            restantes.append("[Otros]")
            for k, v in partes.items(): restantes.append(f"  {k}: {v}")
        return "\n".join(restantes)

    def _write_meta(self, path: str, meta: dict, organizado: bool = False):
        """Escribe el diccionario `meta` en los metadatos del archivo en `path`."""
        if not PIL_OK: raise RuntimeError("Pillow / piexif no instalados.")
        ext = Path(path).suffix.lower()
        if ext in (".jpg", ".jpeg"):    self._write_jpeg(path, meta, organizado)
        elif ext == ".png":             self._write_png(path,  meta, organizado)
        elif ext in (".tif", ".tiff"):  self._write_tiff(path, meta, organizado)
        else:
            try:    self._write_jpeg(path, meta, organizado)
            except: raise RuntimeError(f"Formato no soportado: {ext}")

    def _read_existing_metadata(self, path: str) -> dict:
        """
        Lee el JSON de metadatos ya escrito por MetaTag dentro de una imagen
        (si existe). Devuelve {} si no hay datos previos o no se puede leer.
        """
        if not PIL_OK:
            return {}
        try:
            ext = Path(path).suffix.lower()
            with Image.open(path) as img:
                info = img.info
                if ext in (".jpg", ".jpeg") and "exif" in info:
                    ed = piexif.load(info["exif"])
                    uc = ed.get("Exif", {}).get(piexif.ExifIFD.UserComment)
                    if uc:
                        return json.loads(piexif.helper.UserComment.load(uc))
                elif ext == ".png":
                    c = getattr(img, "text", {}).get("Comment", "")
                    if c:
                        return json.loads(c)
        except Exception:
            pass
        return {}

    def _check_metadata_divergence(self, path, expected_meta):  # -> list
        """
        Compara los metadatos ya escritos en la imagen contra los valores
        que el Excel dice que DEBERÍAN estar ahí (expected_meta).
        Devuelve la lista de campos que divergen, para loguearlos antes
        de sobreescribir. El Excel siempre se trata como fuente de verdad.
        """
        existing = self._read_existing_metadata(path)
        if not existing:
            return []
        diffs = []
        for k, v_new in expected_meta.items():
            v_old = str(existing.get(k, "")).strip()
            if v_old and v_old != str(v_new).strip():
                diffs.append(f"{k}: '{v_old}' → '{v_new}'")
        return diffs

    def _verify_source_images(self):
        """
        Revisa la CARPETA ORIGINAL de imágenes (la que se usa para escribir
        metadatos nuevos) — NO la carpeta Metadatos_Escritos/. Para cada
        imagen detecta si la sección donde MetaTag escribe sus datos
        (UserComment JSON en JPEG, o el chunk 'Comment' en PNG) ya está
        ocupada por datos previos. Si encuentra datos previos, ofrece
        limpiarlos para que la próxima escritura no quede mezclada con
        información vieja.
        """
        folder = self.img_folder_var.get()
        if not folder or not os.path.isdir(folder):
            return messagebox.showwarning("Sin carpeta",
                "Selecciona la carpeta de imágenes originales primero.")
        if not PIL_OK:
            return messagebox.showerror("Error", "Instala: pip install pillow piexif")

        all_files = [f for f in Path(folder).iterdir()
                     if f.is_file() and f.suffix.lower() in IMG_EXTS]
        if not all_files:
            return messagebox.showinfo("Carpeta vacía",
                "No hay imágenes compatibles en esta carpeta.")

        total = len(all_files)
        self._log(f"\n🔍 Verificando metadatos previos en {total} "
                  f"imágenes de la carpeta original...\n", "info")

        prog_win = tk.Toplevel(self)
        prog_win.title("Verificando imágenes originales…")
        prog_win.configure(bg=C["panel"])
        prog_win.geometry("480x180")
        prog_win.resizable(False, False)
        prog_win.grab_set()
        tk.Label(prog_win, text="Revisando metadatos previos",
                 bg=C["panel"], fg=C["accent"], font=FONTS["H2"]).pack(pady=(18, 4))
        lbl_file = tk.Label(prog_win, text="Iniciando…",
                            bg=C["panel"], fg=C["text2"], font=FONTS["LABEL"])
        lbl_file.pack()
        lbl_count = tk.Label(prog_win, text=f"0 / {total}",
                             bg=C["panel"], fg=C["text3"], font=FONTS["TINY"])
        lbl_count.pack()
        bar_outer = tk.Frame(prog_win, bg=C["border"], height=10)
        bar_outer.pack(fill="x", padx=24, pady=12)
        bar_fill = tk.Frame(bar_outer, bg=C["accent"], height=10)
        bar_fill.place(x=0, y=0, relheight=1, width=0)
        prog_win.update_idletasks()
        bar_w = bar_outer.winfo_width() or 432

        def update_ui(done, filename):
            bar_fill.place(width=int(bar_w * (done / total)))
            lbl_file.configure(text=filename)
            lbl_count.configure(text=f"{done} / {total}")

        def worker():
            ocupadas = []
            for i, f in enumerate(all_files):
                existing = self._read_existing_metadata(str(f))
                if existing:
                    ocupadas.append((f, existing))
                if (i + 1) % 3 == 0 or (i + 1) == total:
                    self.after(0, update_ui, i + 1, f.name)
            self.after(0, finish, ocupadas)

        def finish(ocupadas):
            prog_win.destroy()
            if not ocupadas:
                self._log(
                    f"✔ Verificación completa: {len(all_files)} imágenes revisadas, "
                    f"NINGUNA tiene metadatos previos de MetaTag. Todas las "
                    f"secciones están limpias.\n", "ok")
                messagebox.showinfo("Verificación completa",
                    f"Se revisaron {len(all_files)} imágenes en:\n{folder}\n\n"
                    f"✔ Ninguna tiene metadatos previos de MetaTag.\n"
                    f"Puedes escribir los nuevos datos sin riesgo de mezclas.")
                return

            self._log(
                f"⚠ Verificación completa: {len(all_files)} revisadas, "
                f"{len(ocupadas)} YA TIENEN metadatos de MetaTag escritos:\n", "warn")
            for f, meta in ocupadas[:15]:
                preview = ", ".join(f"{k}={v}" for k, v in list(meta.items())[:3])
                self._log(f"   • {f.name}: {preview}{' …' if len(meta) > 3 else ''}\n", "warn")
            if len(ocupadas) > 15:
                self._log(f"   … y {len(ocupadas)-15} más.\n", "warn")

            respuesta = messagebox.askyesno(
                "Metadatos previos encontrados",
                f"Se revisaron {len(all_files)} imágenes en la carpeta original.\n\n"
                f"⚠ {len(ocupadas)} YA TIENEN metadatos de MetaTag escritos "
                f"(la sección donde se guardan los datos no está vacía).\n\n"
                f"Si escribes los nuevos metadatos ahora, estos se MEZCLARÁN "
                f"con los datos viejos.\n\n"
                f"¿Quieres limpiar esos {len(ocupadas)} archivos ahora mismo "
                f"(borrar solo la sección de metadatos de MetaTag, sin tocar "
                f"el resto de la imagen) antes de continuar?")
            if respuesta:
                self._clear_metatag_sections([f for f, _ in ocupadas])

        threading.Thread(target=worker, daemon=True).start()

    def _clear_metatag_sections(self, files: list):
        """
        Limpia ÚNICAMENTE la sección de metadatos que MetaTag usa
        (ImageDescription, UserComment/Comment con el JSON, XPComment,
        XPKeywords) — sin tocar ninguna otra propiedad de la imagen.
        """
        total = len(files)
        prog_win = tk.Toplevel(self)
        prog_win.title("Limpiando metadatos previos…")
        prog_win.configure(bg=C["panel"])
        prog_win.geometry("480x180")
        prog_win.resizable(False, False)
        prog_win.grab_set()
        tk.Label(prog_win, text="Limpiando secciones de metadatos MetaTag",
                 bg=C["panel"], fg=C["accent"], font=FONTS["H2"]).pack(pady=(18, 4))
        lbl_file = tk.Label(prog_win, text="Iniciando…",
                            bg=C["panel"], fg=C["text2"], font=FONTS["LABEL"])
        lbl_file.pack()
        lbl_count = tk.Label(prog_win, text=f"0 / {total}",
                             bg=C["panel"], fg=C["text3"], font=FONTS["TINY"])
        lbl_count.pack()
        bar_outer = tk.Frame(prog_win, bg=C["border"], height=10)
        bar_outer.pack(fill="x", padx=24, pady=12)
        bar_fill = tk.Frame(bar_outer, bg=C["accent"], height=10)
        bar_fill.place(x=0, y=0, relheight=1, width=0)
        prog_win.update_idletasks()
        bar_w = bar_outer.winfo_width() or 432

        def update_ui(done, filename):
            bar_fill.place(width=int(bar_w * (done / total)))
            lbl_file.configure(text=filename)
            lbl_count.configure(text=f"{done} / {total}")

        def worker():
            ok = err = 0
            for i, f in enumerate(files):
                try:
                    ext = f.suffix.lower()
                    if ext in (".jpg", ".jpeg"):
                        img = Image.open(str(f))
                        try:
                            exif = piexif.load(img.info.get("exif", b""))
                        except Exception:
                            exif = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}}
                        exif["0th"].pop(piexif.ImageIFD.ImageDescription, None)
                        exif["0th"].pop(40092, None)
                        exif["0th"].pop(40094, None)
                        exif["Exif"].pop(piexif.ExifIFD.UserComment, None)
                        img.save(str(f), "jpeg", exif=piexif.dump(exif), quality=95)

                    elif ext == ".png":
                        from PIL import PngImagePlugin
                        img = Image.open(str(f))
                        new_info = PngImagePlugin.PngInfo()
                        old_text = dict(getattr(img, "text", {}))
                        for k in ("Description", "Comment"):
                            old_text.pop(k, None)
                        for k, v in old_text.items():
                            new_info.add_text(k, v)
                        img.save(str(f), "PNG", pnginfo=new_info)

                    elif ext in (".tif", ".tiff"):
                        img = Image.open(str(f))
                        try:
                            exif = piexif.load(img.info.get("exif", b""))
                        except Exception:
                            exif = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}}
                        exif["0th"].pop(piexif.ImageIFD.ImageDescription, None)
                        img.save(str(f), exif=piexif.dump(exif))

                    ok += 1
                except Exception as e:
                    self._log(f"  ✗ Error limpiando {f.name}: {e}\n", "err")
                    err += 1
                self.after(0, update_ui, i + 1, f.name)

            def finish():
                prog_win.destroy()
                self._log(f"\n🧹 Limpieza de secciones MetaTag: {ok} limpiadas · "
                          f"{err} errores.\n", "head")
                self.status_var.set(f"✔ {ok} imágenes limpiadas, listas para nuevos metadatos")
                messagebox.showinfo("Limpieza completa",
                    f"✔ {ok} imágenes quedaron con la sección de metadatos vacía.\n"
                    f"{f'✗ {err} con errores.' if err else ''}\n\n"
                    f"Ya puedes escribir los metadatos nuevos sin riesgo de mezclas.")

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _write_jpeg(self, path: str, meta: dict, organizado: bool = True):
        img              = Image.open(path)
        texto_organizado = self._formatear_metadatos(meta, organizado)
        as_json          = json.dumps(meta, ensure_ascii=False)
        keywords         = ";".join(v for v in meta.values() if v.strip())
        try:    exif = piexif.load(img.info.get("exif", b""))
        except: exif = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}}
        exif["0th"][piexif.ImageIFD.ImageDescription] = texto_organizado.encode("utf-8")
        exif["Exif"][piexif.ExifIFD.UserComment] = piexif.helper.UserComment.dump(
            as_json, encoding="unicode")
        exif["0th"][40092] = (texto_organizado + "\x00").encode("utf-16-le")
        exif["0th"][40094] = (keywords         + "\x00").encode("utf-16-le")
        img.save(path, "jpeg", exif=piexif.dump(exif), quality=95)

    def _write_png(self, path: str, meta: dict, organizado: bool = True):
        from PIL import PngImagePlugin
        img  = Image.open(path)
        info = PngImagePlugin.PngInfo()
        texto_organizado = self._formatear_metadatos(meta, organizado)
        for k, v in meta.items(): info.add_text(str(k), str(v))
        info.add_text("Description", texto_organizado)
        info.add_text("Comment",     json.dumps(meta, ensure_ascii=False))
        img.save(path, "PNG", pnginfo=info)

    def _write_tiff(self, path: str, meta: dict, organizado: bool = True):
        img  = Image.open(path)
        texto_organizado = self._formatear_metadatos(meta, organizado)
        try:    exif = piexif.load(img.info.get("exif", b""))
        except: exif = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}}
        exif["0th"][piexif.ImageIFD.ImageDescription] = texto_organizado.encode("utf-8")
        img.save(path, exif=piexif.dump(exif))

    # ─────────────────────────────────────────────────────────────
    #  LOG Y PROGRESO
    # ─────────────────────────────────────────────────────────────
    def _log(self, msg: str, tag: str = ""):
        self.after(0, lambda m=msg, t=tag: self._log_safe(m, t))

    def _log_safe(self, msg: str, tag: str = ""):
        self.log.configure(state="normal")
        self.log.insert("end", msg, tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _update_progress(self, val: float):
        self.after(0, lambda: self.progress_var.set(val))

    # ─────────────────────────────────────────────────────────────
    #  CONFIGURACIÓN PERSISTENTE
    # ─────────────────────────────────────────────────────────────
    def _open_output_folder(self):
        self.output_folder.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform == "win32":
                os.startfile(str(self.output_folder))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(self.output_folder)])
            else:
                subprocess.Popen(["xdg-open", str(self.output_folder)])
        except Exception:
            messagebox.showinfo("Carpeta de salida", str(self.output_folder))

    def _config_path(self) -> Path:
        return self.output_base / "metatag_config.json"

    def _save_config(self):
        try:
            cfg = {
                "csv_path":    self.csv_path_var.get(),
                "img_folder":  self.img_folder_var.get(),
                "theme":       CURRENT_THEME,
                "process_mode": self.process_mode.get(),
            }
            try:
                self.update_idletasks()
                total_w = self.winfo_width()
                if total_w > 0:
                    left_w = self.left.winfo_width()
                    cfg["left_ratio"] = round(left_w / total_w, 3)
            except Exception: pass
            self._config_path().write_text(
                json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception: pass

    def _load_config_pre_build(self):
        try:
            cfg_file = self._config_path()
            if not cfg_file.exists(): return
            cfg = json.loads(cfg_file.read_text(encoding="utf-8"))
            if cfg.get("theme") and cfg["theme"] in THEMES:
                global C, CURRENT_THEME
                CURRENT_THEME = cfg["theme"]
                C.update(THEMES[CURRENT_THEME])
                self.configure(bg=C["bg"])
                self.theme_var.set(CURRENT_THEME)
        except Exception: pass

    def _load_config_post_build(self):
        try:
            cfg_file = self._config_path()
            if not cfg_file.exists(): return
            cfg = json.loads(cfg_file.read_text(encoding="utf-8"))
            if cfg.get("process_mode"):
                self.process_mode.set(cfg["process_mode"])
            if cfg.get("csv_path") and Path(cfg["csv_path"]).exists():
                self.csv_path_var.set(cfg["csv_path"])
            if cfg.get("img_folder") and Path(cfg["img_folder"]).is_dir():
                self.img_folder_var.set(cfg["img_folder"])
                self.browser.load_folder(cfg["img_folder"])
                self._img_cache = {
                    self._full_stem(f.name).lower(): f
                    for f in Path(cfg["img_folder"]).rglob("*")
                    if f.is_file() and f.suffix.lower() in IMG_EXTS
                }
            if cfg.get("left_ratio"):
                def _apply_sashes():
                    try:
                        self.update_idletasks()
                        total_w = self.winfo_width()
                        if total_w > 0:
                            left_w = int(total_w * cfg["left_ratio"])
                            panes = self.paned_window.panes()
                            if panes:
                                self.paned_window.paneconfigure(panes[0], width=left_w)
                    except Exception: pass
                self.after(500, _apply_sashes)
        except Exception: pass


# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = MetaTagApp()
    app.mainloop()