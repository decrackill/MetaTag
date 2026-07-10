"""
MetaTag Widgets — Componentes de UI reutilizables
ExcelGrid optimizada para 300+ filas con viewport culling.
"""

import tkinter as tk
from tkinter import ttk
import pandas as pd
from typing import Optional


class ExcelGrid(tk.Frame):
    def __init__(self, master, on_selection_change=None, on_row_click=None, app_ref=None,
                 theme_colors=None, fonts=None, **kw):
        colors = theme_colors or {}
        super().__init__(master, bg=colors.get("surface", "#1E1E1E"), **kw)
        self.on_selection_change = on_selection_change
        self.on_row_click        = on_row_click
        self.app_ref             = app_ref
        self.C = colors
        self.FONTS = fonts or {}

        self.df = None  # type: Optional[pd.DataFrame]
        self.col_widths = []  # type: list
        self.selected_cells = set()
        self.hovered_row = None  # type: Optional[int]
        self.hidden_columns: set     = set()
        self._redraw_pending         = False

        self.unhide_bar = tk.Frame(self, bg=self.C.get("accent_pale", "#3D1F0A"))
        self.unhide_lbl = tk.Label(self.unhide_bar, text="", bg=self.C.get("accent_pale", "#3D1F0A"),
                                   fg=self.C.get("accent", "#A67C52"),
                                   font=self.FONTS.get("TINY", ("Segoe UI", 8)), anchor="w")
        self.unhide_lbl.pack(side="left", padx=8, pady=2)
        self.unhide_btn = tk.Button(self.unhide_bar, text="Mostrar todas",
                                    bg=self.C.get("accent", "#A67C52"), fg="#FFF5E8",
                                    font=self.FONTS.get("TINY", ("Segoe UI", 8)),
                                    relief="flat", bd=0, cursor="hand2",
                                    command=self._show_all_hidden)
        self.unhide_btn.pack(side="right", padx=8, pady=2)

        self.canvas = tk.Canvas(self, bg=self.C.get("surface", "#1E1E1E"),
                                highlightthickness=0, cursor="arrow")
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

    @property
    def ROW_H(self):
        scale = getattr(self.app_ref, "current_scale", 1.0) if self.app_ref else 1.0
        return int(26 * scale)
    @property
    def HDR_H(self):
        scale = getattr(self.app_ref, "current_scale", 1.0) if self.app_ref else 1.0
        return int(28 * scale)

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
        if not self._redraw_pending:
            self._redraw_pending = True
            self.canvas.after_idle(self._deferred_redraw)

    def _deferred_redraw(self):
        self._redraw_pending = False
        self.redraw()

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

    def redraw(self):
        if self.df is None: return
        self.canvas.configure(bg=self.C.get("surface", "#1E1E1E"))
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

        try:
            y1_frac, y2_frac = self.canvas.yview()
        except Exception:
            y1_frac, y2_frac = 0.0, 1.0

        vis_top = y1_frac * total_h
        vis_bot = y2_frac * total_h

        first_vis = max(0, int((vis_top - self.HDR_H) / self.ROW_H))
        last_vis  = min(nrows - 1, int((vis_bot - self.HDR_H) / self.ROW_H) + 1)

        x = 0
        for ci, col, cw in vis_cols:
            is_col_sel = self._col_fully_selected(ci)
            bg = self.C.get("sel_bg", "#3D1F0A") if is_col_sel else self.C.get("header_bg", "#2D2D30")
            fg = self.C.get("sel_fg", "#E0E0E0") if is_col_sel else self.C.get("header_fg", "#E0E0E0")
            self.canvas.create_rectangle(x, 0, x + cw, self.HDR_H,
                                         fill=bg, outline=self.C.get("grid_line", "#3E3E42"), width=1)
            self.canvas.create_text(x + pad, self.HDR_H // 2,
                                    text=col, anchor="w",
                                    font=self.FONTS.get("HEAD", ("Segoe UI", 9, "bold")), fill=fg)
            x += cw

        for ri in range(first_vis, last_vis + 1):
            if ri >= nrows: break
            try:
                row = self.df.iloc[ri]
            except Exception:
                continue
            y   = self.HDR_H + ri * self.ROW_H
            row_bg = self.C.get("row_even", "#1E1E1E") if ri % 2 == 0 else self.C.get("row_odd", "#1A1A1A")
            x = 0
            for ci, col, cw in vis_cols:
                is_sel  = (ri, ci) in self.selected_cells
                col_sel = self._col_fully_selected(ci)

                if is_sel:              bg, fg = self.C.get("sel_bg", "#3D1F0A"),    self.C.get("sel_fg", "#E0E0E0")
                elif col_sel:           bg, fg = self.C.get("col_sel", "#1A1A1A"),   self.C.get("text", "#E0E0E0")
                elif ri == self.hovered_row: bg, fg = self.C.get("accent_pale", "#3D1F0A"), self.C.get("text", "#E0E0E0")
                else:                   bg, fg = row_bg,         self.C.get("text", "#E0E0E0")

                self.canvas.create_rectangle(x, y, x + cw, y + self.ROW_H,
                                             fill=bg, outline=self.C.get("grid_line", "#3E3E42"), width=1)
                try:
                    cell_text = self._sanitize_cell(row[col])
                except Exception:
                    cell_text = "ERROR"
                self.canvas.create_text(x + pad, y + self.ROW_H // 2,
                                        text=cell_text, anchor="w",
                                        font=self.FONTS.get("CELL", ("Segoe UI", 9)), fill=fg)
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
        self._col_menu = tk.Menu(self, tearoff=0, bg=self.C.get("surface", "#1E1E1E"),
                                 fg=self.C.get("text", "#E0E0E0"),
                                 font=self.FONTS.get("LABEL", ("Segoe UI", 9)), relief="flat")
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
