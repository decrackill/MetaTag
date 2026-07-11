"""
Image Sync — Renombrador de fotos desde Excel
Misna paleta de temas que MetaTag. Explorador de archivos nativo del SO.
"""

import csv
import json
import logging
import os
import platform
import re
import shutil
import subprocess
import threading
import traceback
from collections import OrderedDict
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

import pandas as pd
from PIL import Image, ImageTk

logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")
log = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════
#  TEMAS — Copia exacta de metatag_v8.py
# ══════════════════════════════════════════════════════════════════
THEMES = {
    "Arqueológico (Oscuro Refinado)": {
        "bg": "#121212", "surface": "#1E1E1E", "card": "#1A1A1A", "panel": "#2D2D30",
        "border": "#3E3E42", "border_light": "#252526",
        "accent": "#A67C52", "accent_hover": "#D4A574", "accent_light": "#7A4F2D",
        "accent_pale": "#3D1F0A", "header_bg": "#2D2D30",
        "header_fg": "#E0E0E0", "row_even": "#1E1E1E", "row_odd": "#1A1A1A",
        "sel_bg": "#3D1F0A", "sel_fg": "#E0E0E0", "col_sel": "#1A1A1A",
        "text": "#E0E0E0", "text2": "#AAAAAA", "text3": "#707070",
        "ok": "#4EC9B0", "err": "#F44747", "warn": "#CB4B16",
        "grid_line": "#3E3E42", "btn_ghost_bg": "#1A1A1A",
    },
    "Noche Total": {
        "bg": "#0A0A0A", "surface": "#111111", "card": "#0A0A0A", "panel": "#141414",
        "border": "#2A2A2A", "border_light": "#1E1E1E",
        "accent": "#BB86FC", "accent_hover": "#D0A8FF", "accent_light": "#6200EA",
        "accent_pale": "#1A0A2E", "header_bg": "#1A1A1A",
        "header_fg": "#E8E8E8", "row_even": "#111111", "row_odd": "#161616",
        "sel_bg": "#3700B3", "sel_fg": "#FFFFFF", "col_sel": "#1A0A2E",
        "text": "#E8E8E8", "text2": "#AAAAAA", "text3": "#666666",
        "ok": "#03DAC6", "err": "#CF6679", "warn": "#FF9800",
        "grid_line": "#1E1E1E", "btn_ghost_bg": "#1E1E1E",
    },
    "Carbón": {
        "bg": "#1E1E1E", "surface": "#252526", "card": "#1E1E1E", "panel": "#252526",
        "border": "#3E3E42", "border_light": "#2D2D30",
        "accent": "#569CD6", "accent_hover": "#79B8FF", "accent_light": "#264F78",
        "accent_pale": "#1E3A5F", "header_bg": "#007ACC",
        "header_fg": "#FFFFFF", "row_even": "#252526", "row_odd": "#2D2D30",
        "sel_bg": "#264F78", "sel_fg": "#FFFFFF", "col_sel": "#1E3A5F",
        "text": "#D4D4D4", "text2": "#9CDCFE", "text3": "#6A9955",
        "ok": "#4EC9B0", "err": "#F44747", "warn": "#D97706",
        "grid_line": "#3E3E42", "btn_ghost_bg": "#2D2D30",
    },
}

CURRENT_THEME = "Arqueológico (Oscuro Refinado)"
C = dict(THEMES[CURRENT_THEME])

FONTS = {
    "TITLE": ("Georgia", 13, "bold"),
    "H2": ("Georgia", 10, "bold"),
    "LABEL": ("Segoe UI", 9),
    "LABEL_B": ("Segoe UI", 9, "bold"),
    "BODY": ("Segoe UI", 10),
    "TINY": ("Segoe UI", 8),
    "MONO": ("Consolas", 9),
    "BIG_NUM": ("Georgia", 16, "bold"),
}

# ══════════════════════════════════════════════════════════════════
#  PERSISTENCIA
# ══════════════════════════════════════════════════════════════════
_STATE_FILE = Path(__file__).parent / ".image_sync_state.json"

def _load_state() -> dict:
    try:
        if _STATE_FILE.exists():
            return json.loads(_STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def _save_state(patch: dict) -> None:
    state = _load_state()
    state.update(patch)
    try:
        _STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

# ══════════════════════════════════════════════════════════════════
#  UTILIDADES
# ══════════════════════════════════════════════════════════════════
VALID_IMG_EXT = frozenset(
    {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tiff", ".heic", ".avif"}
)
SORT_OPTIONS = {
    "Orden numérico": "natural",
    "Nombre (A → Z)": "name_asc",
    "Nombre (Z → A)": "name_desc",
    "Fecha modificación ↑": "mtime_asc",
    "Fecha modificación ↓": "mtime_desc",
}

def _natural_key(p: Path) -> list:
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", p.stem)]

# ══════════════════════════════════════════════════════════════════
#  MINIATURAS
# ══════════════════════════════════════════════════════════════════
_THUMB_CACHE: OrderedDict[str, ImageTk.PhotoImage] = OrderedDict()
_THUMB_MAX = 100

def _get_thumb(path: Path, size=(50, 50)) -> ImageTk.PhotoImage | None:
    key = f"{path}|{size}"
    if key in _THUMB_CACHE:
        _THUMB_CACHE.move_to_end(key)
        return _THUMB_CACHE[key]
    try:
        img = Image.open(path)
        img.thumbnail(size, Image.BILINEAR)
        photo = ImageTk.PhotoImage(img)
        _THUMB_CACHE[key] = photo
        if len(_THUMB_CACHE) > _THUMB_MAX:
            _THUMB_CACHE.popitem(last=False)
        return photo
    except Exception:
        return None

# ══════════════════════════════════════════════════════════════════
#  MODELO
# ══════════════════════════════════════════════════════════════════
class RenameModel:
    def __init__(self):
        self.folder_path: Path | None = None
        self.excel_path: Path | None = None
        self.column_name: str | None = None
        self.sheet_name: str | None = None
        self.sort_mode: str = "natural"
        self._photos: list[Path] = []
        self._names: list[str] = []
        self.skipped_rows: list[int] = []
        self._undo_stack: list[tuple[list[tuple[Path, Path]], Path | None, bool]] = []

    @property
    def photos(self): return self._photos
    @property
    def names(self): return self._names
    @property
    def has_undo(self): return bool(self._undo_stack)

    def load_photos(self) -> int:
        if not self.folder_path:
            raise ValueError("No se ha seleccionado ninguna carpeta.")
        raw = [p for p in self.folder_path.iterdir()
               if p.is_file() and p.suffix.lower() in VALID_IMG_EXT]
        sorters = {
            "natural": lambda: raw.sort(key=_natural_key),
            "name_asc": lambda: raw.sort(key=lambda p: p.name.lower()),
            "name_desc": lambda: raw.sort(key=lambda p: p.name.lower(), reverse=True),
            "mtime_asc": lambda: raw.sort(key=lambda p: p.stat().st_mtime),
            "mtime_desc": lambda: raw.sort(key=lambda p: p.stat().st_mtime, reverse=True),
        }
        sorters.get(self.sort_mode, lambda: raw.sort(key=_natural_key))()
        self._photos = raw
        return len(raw)

    def load_sheets(self) -> list[str]:
        if not self.excel_path:
            raise ValueError("No se ha seleccionado ningún Excel.")
        return pd.ExcelFile(self.excel_path).sheet_names

    def load_columns(self) -> list[str]:
        if not self.excel_path:
            raise ValueError("No se ha seleccionado ningún Excel.")
        sheet = self.sheet_name or 0
        return list(pd.read_excel(self.excel_path, sheet_name=sheet, nrows=0).columns)

    def load_names(self) -> int:
        if not (self.excel_path and self.column_name):
            raise ValueError("Excel o columna no configurados.")
        sheet = self.sheet_name or 0
        df = pd.read_excel(self.excel_path, sheet_name=sheet, keep_default_na=False)
        names: list[str] = []
        self.skipped_rows = []
        for i, raw in enumerate(df[self.column_name]):
            text = "" if raw is None else str(raw).strip()
            if text == "" or text.lower() in ("nan", "nat", "none"):
                self.skipped_rows.append(i + 2)
                continue
            names.append(text)
        self._names = names
        return len(self._names)

    def build_preview(self, keep_ext: bool = True) -> list[tuple[str, str, Path]]:
        result = []
        for photo, name in zip(self._photos, self._names):
            if keep_ext:
                new_name = name + photo.suffix
            else:
                if "." in name:
                    new_name = name
                else:
                    new_name = name + photo.suffix
            result.append((photo.name, new_name, photo))
        return result

    def rename_all(self, on_progress, on_done, cancel_ev=None, copy_mode=False):
        total = min(len(self._photos), len(self._names))
        success = 0
        errors: list[str] = []
        batch: list[tuple[Path, Path]] = []
        dest_folder: Path | None = None

        if copy_mode and self._photos:
            dest_folder = self._photos[0].parent / "Renombradas"
            try:
                dest_folder.mkdir(exist_ok=True)
            except OSError as exc:
                on_done(0, [f"No se pudo crear carpeta de copias: {exc}"])
                return

        for i, (photo, name) in enumerate(zip(self._photos, self._names)):
            if cancel_ev and cancel_ev.is_set():
                errors.append("Cancelado por el usuario.")
                break
            new_name = name + photo.suffix
            dest = (dest_folder / new_name) if copy_mode and dest_folder else (photo.parent / new_name)
            try:
                if copy_mode:
                    shutil.copy2(photo, dest)
                else:
                    photo.rename(dest)
                batch.append((dest, photo))
                success += 1
            except OSError as exc:
                errors.append(f"{photo.name} → {new_name}  ({exc})")
            on_progress(i + 1, total, new_name)

        if batch:
            self._undo_stack.append((batch, dest_folder, copy_mode))
        on_done(success, errors)

    def undo_last(self, on_progress, on_done):
        if not self._undo_stack:
            on_done(0, ["No hay nada que deshacer."])
            return
        batch, dest_folder, copy_mode = self._undo_stack.pop()
        total = len(batch)
        success = 0
        errors: list[str] = []
        for i, (current, original) in enumerate(batch):
            try:
                if current.exists():
                    if copy_mode:
                        current.unlink()
                    else:
                        current.rename(original)
                    success += 1
                else:
                    errors.append(f"Archivo no encontrado: {current.name}")
            except OSError as exc:
                errors.append(f"{current.name} → {original.name}  ({exc})")
            on_progress(i + 1, total, original.name)
        if copy_mode and dest_folder and dest_folder.exists():
            try:
                if not any(dest_folder.iterdir()):
                    dest_folder.rmdir()
            except OSError:
                pass
        on_done(success, errors)

    def export_log(self, pairs, dest: Path):
        lines = [
            "=" * 62,
            "  LOG DE RENOMBRAMIENTO — Image Sync",
            f"  Fecha   : {datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}",
            f"  Carpeta : {self.folder_path}",
            "=" * 62, "",
        ]
        for orig, new, _ in pairs:
            lines.append(f"  {orig}  →  {new}")
        lines += ["", f"  Total: {len(pairs)} archivo(s)", ""]
        dest.write_text("\n".join(lines), encoding="utf-8")

    def export_preview_csv(self, pairs, dest: Path):
        with dest.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["original", "nuevo_nombre"])
            for orig, new, _ in pairs:
                w.writerow([orig, new])


# ══════════════════════════════════════════════════════════════════
#  APLICACIÓN
# ══════════════════════════════════════════════════════════════════
class ImageSyncApp(tk.Tk):
    def __init__(self, folder_arg: str | None = None, theme_arg: str | None = None,
                 excel_arg: str | None = None, parent_window=None):
        super().__init__()
        global C, CURRENT_THEME
        if theme_arg and theme_arg in THEMES:
            CURRENT_THEME = theme_arg
            C = dict(THEMES[CURRENT_THEME])

        self.configure(bg=C["bg"])
        self.minsize(900, 680)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(1100, int(sw * 0.70)), min(850, int(sh * 0.85))
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

        self._parent_window = parent_window
        self._model = RenameModel()
        self._cancel_ev: threading.Event | None = None
        self._last_pairs: list[tuple[str, str, Path]] = []
        self._thumb_refs: list[ImageTk.PhotoImage] = []
        self._thumb_win: tk.Toplevel | None = None
        self._thumb_lbl: tk.Label | None = None
        self._loading_columns = False
        self._simulated = False

        self._opt_keep_ext = tk.BooleanVar(value=True)
        self._opt_sort_alpha = tk.BooleanVar(value=False)
        self._opt_backup_log = tk.BooleanVar(value=True)
        self._opt_open_folder = tk.BooleanVar(value=False)

        self._apply_theme()
        self._build_ui()
        self._restore_state()
        self._sheet_var.trace_add("write", lambda *_: self._on_sheet_change())
        self._col_var.trace_add("write", lambda *_: self._on_col_change())
        self.bind("<Unmap>", self._on_minimize)
        self.bind("<Map>", self._on_restore)

        def _on_close():
            if self._parent_window:
                try:
                    self._parent_window.deiconify()
                except Exception:
                    pass
            self.destroy()
        self.protocol("WM_DELETE_WINDOW", _on_close)

        if folder_arg:
            self._entry_folder.delete(0, "end")
            self._entry_folder.insert(0, folder_arg)
            self._on_load_photos()
        if excel_arg:
            self._entry_excel.delete(0, "end")
            self._entry_excel.insert(0, excel_arg)
            self._on_load_excel()

    def _apply_theme(self):
        self.configure(bg=C["bg"])
        self.title(f"Image Sync — {CURRENT_THEME}")

    # ════════════════════════════════════════════════════════════
    #  BUILD UI
    # ════════════════════════════════════════════════════════════
    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", background=C["bg"], foreground=C["text"])
        style.configure("TFrame", background=C["bg"])
        style.configure("TLabel", background=C["bg"], foreground=C["text"], font=FONTS["LABEL"])
        style.configure("TLabelframe", background=C["bg"], foreground=C["accent"],
                         font=FONTS["LABEL_B"], bordercolor=C["border"])
        style.configure("TLabelframe.Label", background=C["bg"], foreground=C["accent"],
                         font=FONTS["LABEL_B"])
        style.configure("TButton", background=C["panel"], foreground=C["text"],
                         font=FONTS["LABEL_B"], borderwidth=1, relief="flat", padding=(10, 5))
        style.map("TButton",
                  background=[("active", C["accent_light"]), ("disabled", C["card"])],
                  foreground=[("active", C["accent_hover"]), ("disabled", C["text3"])])
        style.configure("Accent.TButton", background=C["accent"], foreground=C["bg"],
                         font=FONTS["LABEL_B"], borderwidth=0, relief="flat", padding=(12, 5))
        style.map("Accent.TButton",
                  background=[("active", C["accent_hover"]), ("disabled", C["card"])],
                  foreground=[("active", C["bg"]), ("disabled", C["text3"])])
        style.configure("Header.TFrame", background=C["header_bg"])
        style.configure("Header.TLabel", background=C["header_bg"], foreground=C["header_fg"],
                         font=FONTS["TITLE"])
        style.configure("Header2.TLabel", background=C["header_bg"], foreground=C["text2"],
                         font=FONTS["LABEL"])
        style.configure("Treeview", background=C["surface"], foreground=C["text"],
                         fieldbackground=C["surface"], font=FONTS["TINY"], borderwidth=0,
                         rowheight=26)
        style.configure("Treeview.Heading", background=C["header_bg"], foreground=C["accent"],
                         font=FONTS["LABEL_B"], borderwidth=0, relief="flat")
        style.map("Treeview",
                  background=[("selected", C["sel_bg"])],
                  foreground=[("selected", C["sel_fg"])])
        style.configure("TProgressbar", background=C["accent"], troughcolor=C["panel"],
                         borderwidth=0, thickness=8)
        style.configure("TCheckbutton", background=C["bg"], foreground=C["text2"],
                         font=FONTS["TINY"])
        style.map("TCheckbutton",
                  background=[("active", C["bg"])],
                  foreground=[("active", C["text"])])
        style.configure("TOptionMenu", background=C["panel"], foreground=C["text"],
                         font=FONTS["LABEL"])

        self._build_step_bar()
        self._build_header()
        self._build_summary_panel(self)
        self._build_folder_section(self)
        self._build_excel_section(self)
        self._build_options_panel(self)
        self._build_preview_section(self)
        self._build_action_section(self)
        self._build_log_section(self)
        self._build_footer()

        self.tree.tag_configure("ok", foreground=C["ok"])
        self.tree.tag_configure("warn", foreground=C["warn"])
        self.tree.tag_configure("err", foreground=C["err"])
        self.tree.tag_configure("even", background=C["row_even"])
        self.tree.tag_configure("odd", background=C["row_odd"])

    def _build_step_bar(self):
        pass

    def _draw_steps(self, steps):
        pass

    def _set_step(self, n):
        if not hasattr(self, "_step_pill_frames"):
            return
        self._current_step = n
        for i, (pill, num_lbl, txt_lbl) in enumerate(self._step_pill_frames):
            step_num = i + 1
            if step_num < n:
                pill.configure(highlightbackground=C["accent_light"])
                num_lbl.configure(bg=C["accent_light"], fg=C["bg"])
                txt_lbl.configure(bg=C["card"], fg=C["accent"])
            elif step_num == n:
                pill.configure(highlightbackground=C["accent"])
                num_lbl.configure(bg=C["accent"], fg=C["bg"])
                txt_lbl.configure(bg=C["card"], fg=C["accent_hover"])
            else:
                pill.configure(highlightbackground=C["border"])
                num_lbl.configure(bg=C["border"], fg=C["text3"])
                txt_lbl.configure(bg=C["card"], fg=C["text3"])

    def _build_header(self):
        hdr = tk.Frame(self, bg=C["header_bg"])
        hdr.pack(fill="x")

        top_row = tk.Frame(hdr, bg=C["header_bg"])
        top_row.pack(fill="x")

        tk.Label(top_row, text="  ✏  Image Sync",
                 bg=C["header_bg"], fg=C["header_fg"],
                 font=FONTS["TITLE"]).pack(side="left", padx=8, pady=(10, 2))

        tk.Label(top_row, text="Renombrador de fotos desde Excel",
                 bg=C["header_bg"], fg=C["text2"],
                 font=FONTS["LABEL"]).pack(side="left", padx=4, pady=(10, 2))

        tk.Label(top_row, text=f"  {CURRENT_THEME}  ",
                 bg=C["header_bg"], fg=C["text2"],
                 font=FONTS["TINY"]).pack(side="right", padx=8, pady=(10, 2))

        tk.Frame(hdr, bg=C["border_light"], height=1).pack(fill="x")

        steps_row = tk.Frame(hdr, bg=C["header_bg"])
        steps_row.pack(fill="x", padx=12, pady=(4, 8))

        STEPS = [
            ("1", "Emparejar"),
            ("2", "Validar"),
            ("3", "Vista previa"),
            ("4", "Renombrar"),
            ("5", "Resultado"),
        ]

        self._step_pill_frames = []

        for i, (num, label) in enumerate(STEPS):
            pill = tk.Frame(steps_row, bg=C["card"],
                            highlightthickness=1,
                            highlightbackground=C["border"])
            pill.pack(side="left")

            num_lbl = tk.Label(pill, text=f" {num} ",
                               bg=C["border"], fg=C["text3"],
                               font=FONTS["LABEL_B"],
                               padx=4, pady=2)
            num_lbl.pack(side="left")

            txt_lbl = tk.Label(pill, text=f" {label} ",
                               bg=C["card"], fg=C["text3"],
                               font=FONTS["TINY"],
                               padx=4, pady=2)
            txt_lbl.pack(side="left")

            self._step_pill_frames.append((pill, num_lbl, txt_lbl))

            if i < len(STEPS) - 1:
                tk.Label(steps_row, text=" ─── ",
                         bg=C["header_bg"], fg=C["border"],
                         font=FONTS["TINY"]).pack(side="left")

        self._step_items = []

    def _build_summary_panel(self, parent):
        frame = tk.Frame(parent, bg=C["surface"], highlightthickness=1,
                          highlightbackground=C["border"])
        frame.pack(fill="x", padx=14, pady=(8, 4))

        metrics = [
            ("Fotos", "0", "n_fotos"),
            ("Nombres", "0", "n_nombres"),
            ("Match", "0", "n_match"),
            ("Conflictos", "0", "n_conflicts"),
        ]
        self._summary_labels = {}
        for i, (label, val, key) in enumerate(metrics):
            col = tk.Frame(frame, bg=C["surface"])
            col.pack(side="left", expand=True, fill="both", padx=12, pady=6)
            tk.Label(col, text=label, bg=C["surface"], fg=C["text3"],
                      font=FONTS["TINY"]).pack(anchor="w")
            lbl = tk.Label(col, text=val, bg=C["surface"], fg=C["text"],
                            font=FONTS["BIG_NUM"])
            lbl.pack(anchor="w")
            self._summary_labels[key] = lbl

        self._summary_status = tk.Label(frame, text="❌ Faltan datos",
                                          bg=C["surface"], fg=C["err"],
                                          font=FONTS["LABEL_B"])
        self._summary_status.pack(side="right", padx=16)

    def _update_summary(self):
        n_fotos = len(self._model.photos)
        n_nombres = len(self._model.names)
        correspondencias = min(n_fotos, n_nombres)
        all_new = [name + photo.suffix for photo, name in zip(self._model.photos, self._model.names)]
        from collections import Counter
        dup_count = sum(1 for v in Counter(all_new).values() if v > 1)
        empty_count = sum(1 for name in self._model.names if not name.strip())
        conflictos = dup_count + empty_count

        order_file = self._model.folder_path / ".imgsync_order.json" if self._model.folder_path else None
        has_order = order_file and order_file.exists()

        self._summary_labels["n_fotos"].configure(text=str(n_fotos))
        self._summary_labels["n_nombres"].configure(text=str(n_nombres))
        self._summary_labels["n_match"].configure(text=str(correspondencias))
        self._summary_labels["n_conflicts"].configure(
            text=str(conflictos),
            fg=C["err"] if conflictos > 0 else C["text"])

        if n_fotos == 0 or n_nombres == 0:
            self._summary_status.configure(text="❌ Faltan datos", fg=C["err"])
        elif conflictos > 0:
            status = f"⚠ {conflictos} conflicto(s)"
            if has_order:
                status += "  [Orden MetaTag]"
            self._summary_status.configure(text=status, fg=C["warn"])
        else:
            status = "✓ Listo para renombrar"
            if has_order:
                status += "  [Orden MetaTag]"
            self._summary_status.configure(text=status, fg=C["ok"])

    def _build_folder_section(self, parent):
        frame = tk.Frame(parent, bg=C["surface"], highlightthickness=1,
                          highlightbackground=C["border"])
        frame.pack(fill="x", padx=14, pady=(0, 4))

        header = tk.Frame(frame, bg=C["surface"])
        header.pack(fill="x", padx=10, pady=(8, 0))
        tk.Label(header, text="1 · Carpeta de fotos", bg=C["surface"], fg=C["accent"],
                  font=FONTS["LABEL_B"]).pack(side="left")

        row = tk.Frame(frame, bg=C["surface"])
        row.pack(fill="x", padx=10, pady=(4, 4))
        self._entry_folder = tk.Entry(row, bg=C["panel"], fg=C["text"],
                                       insertbackground=C["text"], font=FONTS["BODY"],
                                       relief="flat", bd=2, highlightthickness=1,
                                       highlightbackground=C["border"],
                                       highlightcolor=C["accent"])
        self._entry_folder.pack(side="left", fill="x", expand=True, padx=(0, 8))

        def _mk_small_btn(parent, text, cmd):
            return tk.Button(parent, text=text, command=cmd,
                              bg=C["panel"], fg=C["text"], activebackground=C["accent_light"],
                              activeforeground=C["text"], font=FONTS["LABEL_B"],
                              relief="flat", bd=0, padx=8, pady=3, cursor="hand2",
                              highlightthickness=1, highlightbackground=C["border"],
                              highlightcolor=C["accent"])

        _mk_small_btn(row, "📂 Explorar", self._browse_folder).pack(side="left", padx=(0, 4))
        _mk_small_btn(row, "Cargar fotos", self._on_load_photos).pack(side="left")

        row2 = tk.Frame(frame, bg=C["surface"])
        row2.pack(fill="x", padx=10, pady=(0, 6))
        tk.Label(row2, text="Ordenar por:", bg=C["surface"], fg=C["text3"],
                  font=FONTS["TINY"]).pack(side="left")
        self._sort_var = tk.StringVar(value="Orden numérico")
        self._sort_menu = ttk.OptionMenu(row2, self._sort_var, "Orden numérico", *SORT_OPTIONS.keys())
        self._sort_menu.pack(side="left", padx=(6, 0))
        self._lbl_folder_status = tk.Label(row2, text="", bg=C["surface"], fg=C["text3"],
                                            font=FONTS["TINY"])
        self._lbl_folder_status.pack(side="right")

    def _build_excel_section(self, parent):
        frame = tk.Frame(parent, bg=C["surface"], highlightthickness=1,
                          highlightbackground=C["border"])
        frame.pack(fill="x", padx=14, pady=(0, 4))

        header = tk.Frame(frame, bg=C["surface"])
        header.pack(fill="x", padx=10, pady=(8, 0))
        tk.Label(header, text="2 · Archivo Excel", bg=C["surface"], fg=C["accent"],
                  font=FONTS["LABEL_B"]).pack(side="left")

        row = tk.Frame(frame, bg=C["surface"])
        row.pack(fill="x", padx=10, pady=(4, 4))
        self._entry_excel = tk.Entry(row, bg=C["panel"], fg=C["text"],
                                      insertbackground=C["text"], font=FONTS["BODY"],
                                      relief="flat", bd=2, highlightthickness=1,
                                      highlightbackground=C["border"],
                                      highlightcolor=C["accent"])
        self._entry_excel.pack(side="left", fill="x", expand=True, padx=(0, 8))

        def _mk_small_btn(parent, text, cmd):
            return tk.Button(parent, text=text, command=cmd,
                              bg=C["panel"], fg=C["text"], activebackground=C["accent_light"],
                              activeforeground=C["text"], font=FONTS["LABEL_B"],
                              relief="flat", bd=0, padx=8, pady=3, cursor="hand2",
                              highlightthickness=1, highlightbackground=C["border"],
                              highlightcolor=C["accent"])

        _mk_small_btn(row, "📂 Explorar", self._browse_excel).pack(side="left", padx=(0, 4))
        _mk_small_btn(row, "Cargar Excel", self._on_load_excel).pack(side="left")

        row2 = tk.Frame(frame, bg=C["surface"])
        row2.pack(fill="x", padx=10, pady=(0, 6))
        tk.Label(row2, text="Hoja:", bg=C["surface"], fg=C["text3"],
                  font=FONTS["TINY"]).pack(side="left")
        self._sheet_var = tk.StringVar()
        self._sheet_menu = ttk.OptionMenu(row2, self._sheet_var, "—")
        self._sheet_menu.pack(side="left", padx=(4, 12))
        self._sheet_menu.configure(state="disabled")
        tk.Label(row2, text="Columna:", bg=C["surface"], fg=C["text3"],
                  font=FONTS["TINY"]).pack(side="left")
        self._col_var = tk.StringVar()
        self._col_menu = ttk.OptionMenu(row2, self._col_var, "—")
        self._col_menu.pack(side="left", padx=(4, 0))
        self._col_menu.configure(state="disabled")
        self._lbl_excel_status = tk.Label(row2, text="", bg=C["surface"], fg=C["text3"],
                                           font=FONTS["TINY"])
        self._lbl_excel_status.pack(side="right")

    def _build_options_panel(self, parent):
        self._options_visible = tk.BooleanVar(value=False)
        container = tk.Frame(parent, bg=C["bg"])
        container.pack(fill="x", padx=14, pady=(0, 4))

        def _toggle():
            if self._options_visible.get():
                content.pack(fill="x", padx=14, pady=(0, 4))
                btn.configure(text="▾ Opciones")
            else:
                content.pack_forget()
                btn.configure(text="▸ Opciones")

        btn = tk.Button(container, text="▸ Opciones", bg=C["bg"], fg=C["text3"],
                         activebackground=C["bg"], activeforeground=C["text"],
                         font=FONTS["TINY"], relief="flat", bd=0, cursor="hand2",
                         command=lambda: [self._options_visible.set(not self._options_visible.get()), _toggle()])
        btn.pack(anchor="w")

        content = tk.Frame(container, bg=C["surface"], highlightthickness=1,
                            highlightbackground=C["border"])

        opts = [
            (self._opt_keep_ext, "Mantener extensión original"),
            (self._opt_sort_alpha, "Ordenar alfabéticamente antes de emparejar"),
            (self._opt_backup_log, "Crear log de seguridad al renombrar"),
            (self._opt_open_folder, "Abrir carpeta al finalizar"),
        ]
        for var, text in opts:
            tk.Checkbutton(content, text=text, variable=var,
                            bg=C["surface"], fg=C["text2"], activebackground=C["surface"],
                            activeforeground=C["text"], selectcolor=C["panel"],
                            font=FONTS["TINY"], bd=0, highlightthickness=0,
                            relief="flat").pack(anchor="w", padx=12, pady=2)

    def _build_preview_section(self, parent):
        frame = tk.Frame(parent, bg=C["surface"], highlightthickness=1,
                          highlightbackground=C["border"])
        frame.pack(fill="both", expand=True, padx=14, pady=(0, 4))

        header = tk.Frame(frame, bg=C["surface"])
        header.pack(fill="x", padx=10, pady=(8, 0))
        tk.Label(header, text="3 · Vista previa", bg=C["surface"], fg=C["accent"],
                  font=FONTS["LABEL_B"]).pack(side="left")

        search_frame = tk.Frame(frame, bg=C["surface"])
        search_frame.pack(fill="x", padx=10, pady=(6, 4))
        tk.Label(search_frame, text="🔍", bg=C["surface"], fg=C["text3"],
                  font=FONTS["BODY"]).pack(side="left")
        self._filter_var = tk.StringVar()
        self._filter_var.trace_add("write", lambda *_: self._apply_filter())
        search = tk.Entry(search_frame, textvariable=self._filter_var,
                          bg=C["panel"], fg=C["text"], insertbackground=C["text"],
                          font=FONTS["BODY"], relief="flat", bd=2, width=30,
                          highlightthickness=1, highlightbackground=C["border"],
                          highlightcolor=C["accent"])
        search.pack(side="left", padx=(4, 8))
        search.insert(0, "")
        search.configure(fg=C["text3"])
        def _on_focus_in(e):
            if search.get() == "":
                search.delete(0, "end")
                search.configure(fg=C["text"])
        def _on_focus_out(e):
            if search.get() == "":
                search.insert(0, "")
                search.configure(fg=C["text3"])
        search.bind("<FocusIn>", _on_focus_in)
        search.bind("<FocusOut>", _on_focus_out)

        self._lbl_count = tk.Label(search_frame, text="0 archivos", bg=C["surface"],
                                    fg=C["text3"], font=FONTS["TINY"])
        self._lbl_count.pack(side="right")

        tree_frame = tk.Frame(frame, bg=C["surface"])
        tree_frame.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        cols = ("num", "estado", "original", "nuevo")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=12)
        self.tree.heading("num", text="#", anchor="e")
        self.tree.heading("estado", text="Estado", anchor="center")
        self.tree.heading("original", text="Original", anchor="w")
        self.tree.heading("nuevo", text="Nuevo nombre", anchor="w")
        self.tree.column("num", width=40, anchor="e", stretch=False)
        self.tree.column("estado", width=60, anchor="center", stretch=False)
        self.tree.column("original", width=280, anchor="w")
        self.tree.column("nuevo", width=280, anchor="w")

        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.tree.bind("<Enter>", self._on_enter_tree)
        self.tree.bind("<Leave>", self._on_leave_tree)
        self.tree.bind("<Motion>", self._on_hover_thumb)
        self.tree.bind("<MouseWheel>", self._on_scroll_thumb)
        self.tree.bind("<Button-4>", self._on_scroll_thumb)
        self.tree.bind("<Button-5>", self._on_scroll_thumb)

    def _build_action_section(self, parent):
        frame = tk.Frame(parent, bg=C["bg"])
        frame.pack(fill="x", padx=14, pady=(0, 4))

        self._progress_canvas = tk.Canvas(frame, bg=C["panel"], height=10,
                                           highlightthickness=0, bd=0)
        self._progress_canvas.pack(fill="x", pady=(0, 4))
        self._progress_canvas.create_rectangle(0, 0, 0, 10, fill=C["accent"], outline="",
                                                tags="bar")

        self._lbl_status = tk.Label(frame, text="Listo para comenzar", bg=C["bg"],
                                     fg=C["text3"], font=FONTS["TINY"], anchor="w")
        self._lbl_status.pack(fill="x")

        btn_row = tk.Frame(frame, bg=C["bg"])
        btn_row.pack(fill="x", pady=(4, 0))

        self._copy_var = tk.BooleanVar()
        tk.Checkbutton(btn_row, text="Modo copiar", variable=self._copy_var,
                        bg=C["bg"], fg=C["text2"], activebackground=C["bg"],
                        activeforeground=C["text"], selectcolor=C["panel"],
                        font=FONTS["TINY"], bd=0, highlightthickness=0,
                        relief="flat").pack(side="left", padx=(0, 10))

        def _mk_btn(parent, text, cmd, accent=False, state="normal"):
            bg  = C["accent"]       if accent else C["panel"]
            fg  = C["bg"]           if accent else C["text"]
            abg = C["accent_hover"] if accent else C["accent_light"]
            btn = tk.Button(parent, text=text, command=cmd,
                             bg=bg, fg=fg, activebackground=abg, activeforeground=C["bg"],
                             disabledforeground=C["text3"], font=FONTS["LABEL_B"],
                             relief="flat", bd=0, padx=10, pady=4, cursor="hand2",
                             highlightthickness=1, highlightbackground=C["border"],
                             highlightcolor=C["accent"])
            if state == "disabled":
                btn.configure(state="disabled", bg=C["card"])
            return btn

        self._btn_log = _mk_btn(btn_row, "💾 Log", self._on_export_log, state="disabled")
        self._btn_log.pack(side="left", padx=(0, 4))
        self._btn_csv = _mk_btn(btn_row, "📤 CSV", self._on_export_csv, state="disabled")
        self._btn_csv.pack(side="left", padx=(0, 4))
        self._btn_cancel = _mk_btn(btn_row, "✖ Cancelar", self._on_cancel, state="disabled")
        self._btn_cancel.pack(side="left", padx=(0, 4))
        self._btn_undo = _mk_btn(btn_row, "↩ Deshacer", self._on_undo, state="disabled")
        self._btn_undo.pack(side="left", padx=(0, 4))
        self._btn_simulate = _mk_btn(btn_row, "👁 Simular", self._on_simulate)
        self._btn_simulate.pack(side="left", padx=(0, 4))
        self._btn_rename = _mk_btn(btn_row, "▶ Renombrar", self._on_rename,
                                    accent=True, state="disabled")
        self._btn_rename.pack(side="left")

    def _build_log_section(self, parent):
        frame = tk.Frame(parent, bg=C["surface"], highlightthickness=1,
                          highlightbackground=C["border"])
        frame.pack(fill="x", padx=14, pady=(0, 4))

        header = tk.Frame(frame, bg=C["surface"])
        header.pack(fill="x", padx=8, pady=(4, 0))
        tk.Label(header, text="Registro", bg=C["surface"], fg=C["text3"],
                  font=FONTS["TINY"]).pack(side="left")

        log_frame = tk.Frame(frame, bg=C["surface"])
        log_frame.pack(fill="x", padx=8, pady=(2, 6))

        self._log_widget = tk.Text(log_frame, bg=C["panel"], fg=C["text2"],
                                    font=FONTS["MONO"], height=4, state="disabled",
                                    relief="flat", bd=0, wrap="word",
                                    insertbackground=C["text2"],
                                    highlightthickness=1,
                                    highlightbackground=C["border"])
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self._log_widget.yview)
        self._log_widget.configure(yscrollcommand=log_scroll.set)
        self._log_widget.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

        self._log_widget.tag_configure("ok", foreground=C["ok"])
        self._log_widget.tag_configure("warn", foreground=C["warn"])
        self._log_widget.tag_configure("err", foreground=C["err"])
        self._log_widget.tag_configure("info", foreground=C["accent"])

    def _build_footer(self):
        ftr = tk.Frame(self, bg=C["header_bg"], height=24)
        ftr.pack(fill="x", side="bottom")
        ftr.pack_propagate(False)
        tk.Label(ftr, text="  Hover → miniatura  •  Ctrl+Z = Deshacer",
                  bg=C["header_bg"], fg=C["text3"], font=FONTS["TINY"]).pack(expand=True)
        self.bind_all("<Control-z>", lambda _: self._on_undo())

    # ════════════════════════════════════════════════════════════
    #  HELPERS UI
    # ════════════════════════════════════════════════════════════
    def _set_btn_state(self, btn: tk.Button, enabled: bool, accent: bool = False):
        if enabled:
            bg = C["accent"] if accent else C["panel"]
            fg = C["bg"]     if accent else C["text"]
            abg = C["accent_hover"] if accent else C["accent_light"]
            btn.configure(state="normal", bg=bg, fg=fg,
                          activebackground=abg, activeforeground=C["bg"])
        else:
            btn.configure(state="disabled", bg=C["card"], fg=C["text3"])

    def _log(self, msg: str, tag: str = ""):
        self._log_widget.configure(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self._log_widget.insert("end", f"[{ts}] {msg}\n", tag)
        self._log_widget.see("end")
        self._log_widget.configure(state="disabled")

    def _set_progress(self, frac, msg):
        w = self._progress_canvas.winfo_width()
        self._progress_canvas.delete("bar")
        self._progress_canvas.create_rectangle(0, 0, int(w * frac), 10,
                                                fill=C["accent"], outline="", tags="bar")
        self._lbl_status.configure(text=msg)

    # ════════════════════════════════════════════════════════════
    #  EXPLORADORES DE ARCHIVOS NATIVOS
    # ════════════════════════════════════════════════════════════
    def _browse_folder(self):
        path = None
        system = platform.system()
        if system == "Linux":
            try:
                result = subprocess.run(
                    ["zenity", "--file-selection", "--directory",
                     "--title=Selecciona la carpeta de fotos"],
                    capture_output=True, text=True, timeout=60)
                if result.returncode == 0:
                    path = result.stdout.strip()
            except FileNotFoundError:
                try:
                    result = subprocess.run(
                        ["kdialog", "--getexistingdirectory", str(Path.home()),
                         "--title", "Selecciona la carpeta de fotos"],
                        capture_output=True, text=True, timeout=60)
                    if result.returncode == 0:
                        path = result.stdout.strip()
                except FileNotFoundError:
                    path = filedialog.askdirectory(title="Selecciona la carpeta", mustexist=True)
        else:
            path = filedialog.askdirectory(title="Selecciona la carpeta", mustexist=True)
        if path:
            self._entry_folder.delete(0, "end")
            self._entry_folder.insert(0, path)

    def _browse_excel(self):
        path = None
        system = platform.system()
        if system == "Linux":
            try:
                result = subprocess.run(
                    ["zenity", "--file-selection",
                     "--title=Selecciona el archivo Excel",
                     "--file-filter=Excel (*.xlsx) | *.xlsx",
                     "--file-filter=Todos | *"],
                    capture_output=True, text=True, timeout=60)
                if result.returncode == 0:
                    path = result.stdout.strip()
            except FileNotFoundError:
                try:
                    result = subprocess.run(
                        ["kdialog", "--getopenfilename", str(Path.home()),
                         "*.xlsx | Excel (*.xlsx)",
                         "--title", "Selecciona el archivo Excel"],
                        capture_output=True, text=True, timeout=60)
                    if result.returncode == 0:
                        path = result.stdout.strip()
                except FileNotFoundError:
                    path = filedialog.askopenfilename(
                        title="Selecciona Excel",
                        filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
        else:
            path = filedialog.askopenfilename(
                title="Selecciona Excel",
                filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
        if path:
            self._entry_excel.delete(0, "end")
            self._entry_excel.insert(0, path)

    # ════════════════════════════════════════════════════════════
    #  HANDLERS
    # ════════════════════════════════════════════════════════════
    def _on_load_photos(self):
        self._simulated = False
        self._set_btn_state(self._btn_rename, False, accent=True)
        if hasattr(self, "_btn_simulate"):
            self._btn_simulate.configure(text="Simular")
        raw = self._entry_folder.get().strip()
        if not raw:
            messagebox.showwarning("Image Sync", "Selecciona primero una carpeta.")
            return
        path = Path(raw)
        if not path.is_dir():
            self._lbl_folder_status.configure(text="Ruta no existe", fg=C["err"])
            return
        self._model.folder_path = path
        self._model.sort_mode = SORT_OPTIONS.get(self._sort_var.get(), "natural")
        _save_state({"last_folder": str(path), "sort": self._model.sort_mode})
        try:
            n = self._model.load_photos()
        except Exception as exc:
            self._lbl_folder_status.configure(text=str(exc), fg=C["err"])
            return
        if n == 0:
            self._lbl_folder_status.configure(text="No se encontraron imágenes", fg=C["warn"])
            self._log(f"⚠ No se encontraron imágenes en {path.name}", "warn")
        else:
            self._lbl_folder_status.configure(
                text=f"{n} imagen{'es' if n != 1 else ''}  •  {self._sort_var.get()}", fg=C["ok"])
            self._log(f"📁 {n} imágenes cargadas desde {path.name}", "info")
        order_file = path / ".imgsync_order.json"
        if order_file.exists():
            try:
                order_data = json.loads(order_file.read_text(encoding="utf-8"))
                ordered_paths = [Path(p) for p in order_data.get("ordered", [])]
                valid_ordered = [
                    p for p in ordered_paths
                    if p.exists() and p.suffix.lower() in VALID_IMG_EXT
                ]
                if len(valid_ordered) == n:
                    self._model._photos = valid_ordered
                    self._lbl_folder_status.configure(
                        text=f"{n} imágenes  •  Orden de MetaTag aplicado",
                        foreground=C["ok"])
                    self._log(
                        f"[{datetime.now().strftime('%H:%M:%S')}] "
                        f"Orden de MetaTag cargado desde {order_file.name}\n", "info")
                elif len(valid_ordered) > 0:
                    ordered_set = set(valid_ordered)
                    missing = [p for p in self._model._photos if p not in ordered_set]
                    self._model._photos = valid_ordered + missing
                    self._lbl_folder_status.configure(
                        text=f"{n} imágenes  •  Orden parcial de MetaTag ({len(missing)} sin orden)",
                        foreground=C["warn"])
            except Exception as exc:
                log.warning("No se pudo leer orden de MetaTag: %s", exc)
        self._set_step(1)
        self._refresh_preview()

    def _on_load_excel(self):
        self._simulated = False
        self._set_btn_state(self._btn_rename, False, accent=True)
        if hasattr(self, "_btn_simulate"):
            self._btn_simulate.configure(text="Simular")
        raw = self._entry_excel.get().strip()
        if not raw:
            messagebox.showwarning("Image Sync", "Selecciona primero un archivo Excel.")
            return
        path = Path(raw)
        if not path.is_file() or path.suffix.lower() != ".xlsx":
            self._lbl_excel_status.configure(text="Debe ser .xlsx", fg=C["err"])
            return
        self._model.excel_path = path
        _save_state({"last_excel": str(path)})
        try:
            sheets = self._model.load_sheets()
        except Exception as exc:
            self._lbl_excel_status.configure(text=str(exc), fg=C["err"])
            return
        self._model.sheet_name = sheets[0]
        self._sheet_menu["menu"].delete(0, "end")
        for s in sheets:
            self._sheet_menu["menu"].add_command(label=s, command=lambda v=s: self._sheet_var.set(v))
        self._sheet_menu.configure(state="normal")
        self._sheet_var.set(sheets[0])
        self._load_columns()

    def _load_columns(self):
        try:
            cols = self._model.load_columns()
        except Exception as exc:
            self._lbl_excel_status.configure(text=str(exc), fg=C["err"])
            return
        self._loading_columns = True
        if len(cols) == 1:
            self._model.column_name = cols[0]
            self._col_menu.configure(state="disabled")
            self._lbl_excel_status.configure(text=f"Columna «{cols[0]}» auto-seleccionada.", fg=C["ok"])
            self._loading_columns = False
            self._load_names_and_preview()
        else:
            self._col_menu["menu"].delete(0, "end")
            for c in cols:
                self._col_menu["menu"].add_command(label=c, command=lambda v=c: self._col_var.set(v))
            self._col_menu.configure(state="normal")
            self._col_var.set(cols[0])
            self._model.column_name = cols[0]
            self._loading_columns = False
            self._lbl_excel_status.configure(text=f"{len(cols)} columnas — elige cuál.", fg=C["ok"])
            self._load_names_and_preview()

    def _on_sheet_change(self):
        sheet = self._sheet_var.get()
        if sheet and sheet != "—":
            self._model.sheet_name = sheet
            self._load_columns()

    def _on_col_change(self):
        if self._loading_columns:
            return
        col = self._col_var.get()
        if col and col != "—":
            self._model.column_name = col
            self._load_names_and_preview()

    def _load_names_and_preview(self):
        try:
            n = self._model.load_names()
        except Exception as exc:
            self._lbl_excel_status.configure(text=str(exc), fg=C["err"])
            return
        self._lbl_excel_status.configure(
            text=f"{n} nombre{'s' if n != 1 else ''} en «{self._model.column_name}».", fg=C["ok"])
        self._log(f"📊 Excel: {n} nombres en columna «{self._model.column_name}»", "info")
        if self._model.skipped_rows:
            rows = ", ".join(str(r) for r in self._model.skipped_rows[:10])
            self._log(f"⚠ Filas {rows} sin texto legible", "warn")
        self._set_step(2)
        self._refresh_preview()

    def _refresh_preview(self):
        self._last_pairs = self._model.build_preview(keep_ext=self._opt_keep_ext.get())
        self._populate_tree(self._last_pairs)
        has_data = bool(self._last_pairs)
        self._set_btn_state(self._btn_log, has_data)
        self._set_btn_state(self._btn_csv, has_data)
        self._update_summary()

    def _compute_status(self, idx, orig, new, all_new_names, photos):
        if not new or new == ".":
            return ("❌", "err")
        from collections import Counter
        dupes = [n for n, c in Counter(all_new_names).items() if c > 1]
        if new in dupes:
            return ("⚠ Dup", "warn")
        if idx < len(photos):
            photo = photos[idx]
            dest = photo.parent / new
            if dest.exists() and dest != photo:
                return ("🔁", "warn")
        return ("✅", "ok")

    def _populate_tree(self, pairs):
        self.tree.delete(*self.tree.get_children())
        all_new = [new for _, new, _ in pairs]
        photos = [p for _, _, p in pairs]
        for i, (orig, new, _) in enumerate(pairs):
            icon, tag = self._compute_status(i, orig, new, all_new, photos)
            parity = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", iid=str(i),
                              values=(i + 1, icon, orig, new),
                              tags=(tag, parity))

    def _apply_filter(self):
        q = self._filter_var.get().lower()
        visible = 0
        for i, (orig, new, _) in enumerate(self._last_pairs):
            match = not q or q in orig.lower() or q in new.lower()
            if match:
                try:
                    self.tree.reattach(str(i), "", "end")
                except Exception:
                    pass
                visible += 1
            else:
                try:
                    self.tree.detach(str(i))
                except Exception:
                    pass
        if q:
            self._lbl_count.configure(text=f"{visible} de {len(self._last_pairs)} resultados")
        else:
            self._lbl_count.configure(text=f"{len(self._last_pairs)} archivos")

    # ════════════════════════════════════════════════════════════
    #  MINIATURAS HOVER
    # ════════════════════════════════════════════════════════════
    def _on_enter_tree(self, _=None):
        if self._thumb_win:
            try:
                self._thumb_win.destroy()
            except Exception:
                pass
            self._thumb_win = None
            self._thumb_lbl = None
        self._thumb_win = tk.Toplevel(self)
        self._thumb_win.wm_overrideredirect(True)
        self._thumb_win.configure(bg=C["panel"])
        self._thumb_win.attributes("-topmost", True)
        self._thumb_lbl = tk.Label(self._thumb_win, bg=C["panel"])
        self._thumb_lbl.pack(padx=4, pady=4)
        self._thumb_win.withdraw()

    def _on_leave_tree(self, _=None):
        if self._thumb_win:
            try:
                self._thumb_win.destroy()
            except Exception:
                pass
            self._thumb_win = None
            self._thumb_lbl = None

    def _on_hover_thumb(self, event):
        if not self._thumb_win:
            return
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            self._thumb_win.withdraw()
            return
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or col not in ("#3", "#4"):
            self._thumb_win.withdraw()
            return
        col1_w = self.tree.column("#1", "width")
        col2_w = self.tree.column("#2", "width")
        col3_w = self.tree.column("#3", "width")
        col4_w = self.tree.column("#4", "width")
        max_x = col1_w + col2_w + col3_w + col4_w
        if event.x > max_x:
            self._thumb_win.withdraw()
            return
        try:
            cell_values = self.tree.item(item, "values")
            if not cell_values:
                self._thumb_win.withdraw()
                return
            col_idx = int(col.replace("#", "")) - 1
            cell_text = str(cell_values[col_idx]).strip() if col_idx < len(cell_values) else ""
            if not cell_text:
                self._thumb_win.withdraw()
                return
        except Exception:
            self._thumb_win.withdraw()
            return
        idx = int(item)
        if idx >= len(self._last_pairs):
            self._thumb_win.withdraw()
            return
        photo_path = self._last_pairs[idx][2]
        img = _get_thumb(photo_path, (220, 220))
        if img:
            self._thumb_lbl.configure(image=img)
            self._thumb_lbl._img = img
            x = event.x_root + 16
            y = event.y_root + 8
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            if x + 240 > sw:
                x = event.x_root - 246
            if y + 240 > sh:
                y = event.y_root - 246
            self._thumb_win.geometry(f"+{x}+{y}")
            self._thumb_win.deiconify()
        else:
            self._thumb_win.withdraw()

    def _on_scroll_thumb(self, event):
        self.after(80, lambda: self._on_hover_thumb(event))

    def _on_minimize(self, event):
        if event.widget is self and self._thumb_win:
            try:
                self._thumb_win.withdraw()
            except Exception:
                pass

    def _on_restore(self, event):
        pass

    # ════════════════════════════════════════════════════════════
    #  SIMULAR / RENOMBRAR / DESHACER
    # ════════════════════════════════════════════════════════════
    def _on_simulate(self):
        if not (self._model.photos and self._model.names):
            messagebox.showwarning("Image Sync", "Carga las fotos y el Excel primero.")
            return
        self._refresh_preview()
        self._update_summary()
        n = min(len(self._model.photos), len(self._model.names))
        self._log(f"[SIMULACIÓN] {n} archivos serán renombrados. Sin cambios en disco.", "info")
        self._lbl_status.configure(text=f"Simulación completada · {n} archivos", fg=C["accent"])
        self._simulated = True
        self._set_btn_state(self._btn_rename, True, accent=True)
        self._set_step(3)
        self._btn_simulate.configure(text="✓ Simulado")
        self._btn_rename.configure(text="▶ Renombrar")

    def _on_rename(self):
        if not self._simulated:
            messagebox.showwarning("Image Sync", "Primero ejecuta la simulación.")
            return
        if not (self._model.photos and self._model.names):
            return
        n_ph = len(self._model.photos)
        n_nm = len(self._model.names)
        will = min(n_ph, n_nm)
        all_new = [name + photo.suffix for photo, name in zip(self._model.photos, self._model.names)]
        from collections import Counter
        conflictos = sum(1 for v in Counter(all_new).values() if v > 1)
        if not self._confirm_rename_dialog(will, conflictos):
            return
        self._cancel_ev = threading.Event()
        self._btn_rename.configure(text="Renombrando…")
        self._set_btn_state(self._btn_rename, False)
        self._set_btn_state(self._btn_cancel, True)
        self._set_btn_state(self._btn_simulate, False)
        self._set_progress(0, "Iniciando…")
        self._set_step(4)
        threading.Thread(target=self._do_rename, daemon=True).start()

    def _confirm_rename_dialog(self, n, conflictos):
        dlg = tk.Toplevel(self)
        dlg.title("Confirmar renombramiento")
        dlg.configure(bg=C["bg"])
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        hdr = tk.Frame(dlg, bg=C["header_bg"], height=48)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="  ✏  Confirmar renombramiento", bg=C["header_bg"],
                  fg=C["header_fg"], font=FONTS["TITLE"]).pack(side="left", padx=12, pady=10)

        body = tk.Frame(dlg, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=24, pady=16)

        tk.Label(body, text=f"Se renombrarán {n} fotografías.",
                  bg=C["bg"], fg=C["text"], font=FONTS["H2"]).pack(anchor="w")
        tk.Label(body, text="Esta acción modificará permanentemente los nombres de los archivos.",
                  bg=C["bg"], fg=C["text3"], font=FONTS["LABEL"]).pack(anchor="w", pady=(4, 8))

        if conflictos > 0:
            tk.Label(body, text=f"⚠ {conflictos} conflicto(s) detectado(s) — revisa la tabla.",
                      bg=C["bg"], fg=C["warn"], font=FONTS["LABEL_B"]).pack(anchor="w", pady=(0, 8))

        btn_row = tk.Frame(body, bg=C["bg"])
        btn_row.pack(fill="x", pady=(12, 0))

        result = {"ok": False}

        def _cancel():
            result["ok"] = False
            dlg.destroy()

        def _confirm():
            result["ok"] = True
            dlg.destroy()

        tk.Button(btn_row, text="Cancelar", command=_cancel,
                   bg=C["panel"], fg=C["text"], activebackground=C["accent_light"],
                   activeforeground=C["text"], font=FONTS["LABEL_B"],
                   relief="flat", bd=0, padx=16, pady=6, cursor="hand2").pack(side="right", padx=(8, 0))
        tk.Button(btn_row, text="▶ Renombrar", command=_confirm,
                   bg=C["accent"], fg=C["bg"], activebackground=C["accent_hover"],
                   activeforeground=C["bg"], font=FONTS["LABEL_B"],
                   relief="flat", bd=0, padx=16, pady=6, cursor="hand2").pack(side="right")

        dlg.update_idletasks()
        w, h = 440, 220
        x = self.winfo_rootx() + (self.winfo_width() - w) // 2
        y = self.winfo_rooty() + (self.winfo_height() - h) // 2
        dlg.geometry(f"{w}x{h}+{x}+{y}")
        dlg.wait_window()
        return result["ok"]

    def _do_rename(self):
        def progress(cur, tot, name):
            self.after(0, lambda: self._set_progress(cur / tot, f"{cur}/{tot} — {name}"))
            if cur % 10 == 0:
                self.after(0, lambda n=name: self._log(f"✔ {n}", "ok"))
        def done(ok, errors):
            self.after(0, lambda: self._finish_rename(ok, errors))
        self._model.rename_all(progress, done, cancel_ev=self._cancel_ev,
                               copy_mode=self._copy_var.get())

    def _finish_rename(self, ok, errors):
        self._set_progress(1.0, f"Completado · {ok} renombradas.")
        self._btn_rename.configure(text="▶ Renombrar")
        self._set_btn_state(self._btn_rename, False)
        self._set_btn_state(self._btn_cancel, False)
        self._set_btn_state(self._btn_simulate, True)
        self._set_btn_state(self._btn_undo, self._model.has_undo)
        self._simulated = False
        self._btn_simulate.configure(text="Simular")
        self._btn_rename.configure(text="▶ Renombrar")
        self._set_step(5)

        if errors:
            for e in errors:
                self._log(f"✖ {e}", "err")
            self._log(f"⚠ {ok} OK · {len(errors)} con error", "warn")
        else:
            self._log(f"✓ {ok} foto{'s' if ok != 1 else ''} renombradas", "ok")

        if self._opt_backup_log.get() and self._model.folder_path:
            try:
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                log_path = self._model.folder_path / f"image_sync_log_{ts}.txt"
                self._model.export_log(self._last_pairs, log_path)
                self._log(f"💾 Log guardado: {log_path.name}", "info")
            except Exception:
                pass

        if self._opt_open_folder.get() and self._model.folder_path and not errors:
            try:
                system = platform.system()
                if system == "Linux":
                    subprocess.Popen(["xdg-open", str(self._model.folder_path)])
                elif system == "Windows":
                    os.startfile(str(self._model.folder_path))
            except Exception:
                pass

        if self._model.folder_path:
            try:
                self._model.load_photos()
            except Exception:
                pass
        self._refresh_preview()

    def _on_undo(self):
        if not self._model.has_undo:
            messagebox.showinfo("Image Sync", "No hay nada que deshacer.")
            return
        if not messagebox.askyesno("Deshacer", "¿Revertir el último lote?"):
            return
        self._set_btn_state(self._btn_undo, False)
        self._set_progress(0, "Deshaciendo…")
        threading.Thread(target=self._do_undo, daemon=True).start()

    def _do_undo(self):
        def progress(cur, tot, name):
            self.after(0, lambda: self._set_progress(cur / tot, f"Revirtiendo {cur}/{tot} — {name}"))
        def done(ok, errors):
            self.after(0, lambda: self._finish_undo(ok, errors))
        self._model.undo_last(progress, done)

    def _finish_undo(self, ok, errors):
        self._set_progress(0, f"Deshacer · {ok} revertidas.")
        self._set_btn_state(self._btn_undo, self._model.has_undo)
        self._log(f"↩ Deshecho: {ok} archivos revertidos", "info")
        if self._model.folder_path:
            try:
                self._model.load_photos()
            except Exception:
                pass
            self._refresh_preview()
        if errors:
            self._log(f"⚠ {len(errors)} error(es) al deshacer", "err")

    def _on_cancel(self):
        if self._cancel_ev:
            self._cancel_ev.set()
            self._set_btn_state(self._btn_cancel, False)
            self._lbl_status.configure(text="Cancelando…", fg=C["warn"])

    def _on_export_log(self):
        if not self._last_pairs:
            messagebox.showinfo("Image Sync", "No hay log para exportar.")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = filedialog.asksaveasfilename(
            title="Guardar log", defaultextension=".txt",
            initialfile=f"log_{ts}.txt", filetypes=[("Texto", "*.txt")])
        if dest:
            try:
                self._model.export_log(self._last_pairs, Path(dest))
                self._log(f"💾 Log guardado: {Path(dest).name}", "ok")
            except Exception as exc:
                self._log(f"✖ Error: {exc}", "err")

    def _on_export_csv(self):
        if not self._last_pairs:
            messagebox.showinfo("Image Sync", "No hay vista previa para exportar.")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = filedialog.asksaveasfilename(
            title="Guardar CSV", defaultextension=".csv",
            initialfile=f"preview_{ts}.csv", filetypes=[("CSV", "*.csv")])
        if dest:
            try:
                self._model.export_preview_csv(self._last_pairs, Path(dest))
                self._log(f"📤 CSV guardado: {Path(dest).name}", "ok")
            except Exception as exc:
                self._log(f"✖ Error: {exc}", "err")

    def _restore_state(self):
        state = _load_state()
        if "last_folder" in state:
            p = Path(state["last_folder"])
            if p.is_dir():
                self._entry_folder.insert(0, str(p))
        if "last_excel" in state:
            p = Path(state["last_excel"])
            if p.is_file():
                self._entry_excel.insert(0, str(p))
        if "sort" in state:
            for k, v in SORT_OPTIONS.items():
                if v == state["sort"]:
                    self._sort_var.set(k)
                    break


# ══════════════════════════════════════════════════════════════════
#  ENTRYPOINT
# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import argparse
    _parser = argparse.ArgumentParser()
    _parser.add_argument("--folder", default=None)
    _parser.add_argument("--excel", default=None)
    _parser.add_argument("--theme", default=None)
    _args, _ = _parser.parse_known_args()

    try:
        app = ImageSyncApp(folder_arg=_args.folder, theme_arg=_args.theme,
                           excel_arg=_args.excel)
        app.mainloop()
    except Exception:
        tb_str = traceback.format_exc()
        log.critical("Fallo crítico:\n%s", tb_str)
        try:
            crash_log = Path.home() / "image_sync_error.log"
            crash_log.write_text(tb_str, encoding="utf-8")
        except Exception:
            crash_log = None
        try:
            _root = tk.Tk()
            _root.withdraw()
            messagebox.showerror(
                "Error al iniciar — Image Sync",
                f"{tb_str[-700:]}\n\n" + (f"Log: {crash_log}" if crash_log else ""))
            _root.destroy()
        except Exception:
            print(tb_str)
