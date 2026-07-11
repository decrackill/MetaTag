"""
Image Sync — Renombrador de fotos desde Excel
Misma paleta de temas que MetaTag. Explorador de archivos nativo del SO.
"""

import csv
import json
import logging
import os
import re
import shutil
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

    def build_preview(self) -> list[tuple[str, str, Path]]:
        return [(photo.name, name + photo.suffix, photo)
                for photo, name in zip(self._photos, self._names)]

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
                 excel_arg: str | None = None):
        super().__init__()
        global C, CURRENT_THEME
        if theme_arg and theme_arg in THEMES:
            CURRENT_THEME = theme_arg
            C = dict(THEMES[CURRENT_THEME])

        self.configure(bg=C["bg"])
        self.minsize(820, 620)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(1050, int(sw * 0.68)), min(780, int(sh * 0.82))
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

        self._model = RenameModel()
        self._cancel_ev: threading.Event | None = None
        self._last_pairs: list[tuple[str, str, Path]] = []
        self._thumb_refs: list[ImageTk.PhotoImage] = []
        self._thumb_win: tk.Toplevel | None = None
        self._thumb_lbl: tk.Label | None = None

        self._apply_theme()
        self._build_ui()
        self._restore_state()
        self._loading_columns = False
        self._sheet_var.trace_add("write", lambda *_: self._on_sheet_change())
        self._col_var.trace_add("write", lambda *_: self._on_col_change())
        self.bind("<Unmap>", self._on_minimize)
        self.bind("<Map>", self._on_restore)
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
                  background=[("active", C["accent_light"]), ("disabled", C["card"]),
                              ("pressed", C["accent"])],
                  foreground=[("active", C["accent_hover"]), ("disabled", C["text3"]),
                              ("pressed", C["bg"])],
                  relief=[("pressed", "sunken")])
        style.configure("Accent.TButton", background=C["accent"], foreground=C["bg"],
                         font=FONTS["LABEL_B"], borderwidth=0, relief="flat", padding=(12, 5))
        style.map("Accent.TButton",
                  background=[("active", C["accent_hover"]), ("disabled", C["card"]),
                              ("pressed", C["accent_light"])],
                  foreground=[("active", C["bg"]), ("disabled", C["text3"]),
                              ("pressed", C["bg"])],
                  relief=[("pressed", "sunken")])
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

        # ── Header ──
        hdr = ttk.Frame(self, style="Header.TFrame")
        hdr.pack(fill="x")
        ttk.Label(hdr, text="  ✏  Image Sync", style="Header.TLabel").pack(side="left", padx=6, pady=10)
        ttk.Label(hdr, text="Renombrador de fotos desde Excel", style="Header2.TLabel").pack(side="left", padx=4)
        self._lbl_theme = ttk.Label(hdr, text=f"  🎨 {CURRENT_THEME}", style="Header2.TLabel")
        self._lbl_theme.pack(side="right", padx=12)

        # ── Body ──
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=14, pady=10)

        self._build_folder_section(body)
        self._build_excel_section(body)
        self._build_preview_section(body)
        self._build_action_section(body)

        # ── Footer ──
        ftr = ttk.Frame(self)
        ftr.pack(fill="x", side="bottom")
        ttk.Label(ftr, text="  Hover sobre fila → miniatura  •  Ctrl+Z = Deshacer",
                  foreground=C["text3"]).pack(pady=4)
        self.bind_all("<Control-z>", lambda _: self._on_undo())

    def _build_folder_section(self, parent):
        frame = ttk.LabelFrame(parent, text="  1 · Carpeta de fotos  ", padding=10)
        frame.pack(fill="x", pady=(0, 8))

        row = ttk.Frame(frame)
        row.pack(fill="x")
        self._entry_folder = tk.Entry(row, bg=C["surface"], fg=C["text"],
                                       insertbackground=C["text"], font=FONTS["BODY"],
                                       relief="flat", bd=2, highlightthickness=1,
                                       highlightbackground=C["border"],
                                       highlightcolor=C["accent"])
        self._entry_folder.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row, text="📂 Explorar", command=self._browse_folder).pack(side="left", padx=(0, 4))
        ttk.Button(row, text="Cargar fotos", style="Accent.TButton",
                   command=self._on_load_photos).pack(side="left")

        row2 = ttk.Frame(frame)
        row2.pack(fill="x", pady=(8, 0))
        ttk.Label(row2, text="Ordenar por:").pack(side="left")
        self._sort_var = tk.StringVar(value="Orden numérico")
        sort_menu = ttk.OptionMenu(row2, self._sort_var, "Orden numérico", *SORT_OPTIONS.keys())
        sort_menu.pack(side="left", padx=(6, 0))

        self._lbl_folder_status = ttk.Label(row2, text="", foreground=C["text2"])
        self._lbl_folder_status.pack(side="right")

    def _build_excel_section(self, parent):
        frame = ttk.LabelFrame(parent, text="  2 · Archivo Excel  ", padding=10)
        frame.pack(fill="x", pady=(0, 8))

        row = ttk.Frame(frame)
        row.pack(fill="x")
        self._entry_excel = tk.Entry(row, bg=C["surface"], fg=C["text"],
                                      insertbackground=C["text"], font=FONTS["BODY"],
                                      relief="flat", bd=2, highlightthickness=1,
                                      highlightbackground=C["border"],
                                      highlightcolor=C["accent"])
        self._entry_excel.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row, text="📂 Explorar", command=self._browse_excel).pack(side="left", padx=(0, 4))
        ttk.Button(row, text="Cargar Excel", style="Accent.TButton",
                   command=self._on_load_excel).pack(side="left")

        row2 = ttk.Frame(frame)
        row2.pack(fill="x", pady=(8, 0))
        ttk.Label(row2, text="Hoja:").pack(side="left")
        self._sheet_var = tk.StringVar()
        self._sheet_menu = ttk.OptionMenu(row2, self._sheet_var, "—")
        self._sheet_menu.pack(side="left", padx=(6, 12))
        self._sheet_menu.configure(state="disabled")

        ttk.Label(row2, text="Columna:").pack(side="left")
        self._col_var = tk.StringVar()
        self._col_menu = ttk.OptionMenu(row2, self._col_var, "—")
        self._col_menu.pack(side="left", padx=(6, 0))
        self._col_menu.configure(state="disabled")

        self._lbl_excel_status = ttk.Label(row2, text="", foreground=C["text2"])
        self._lbl_excel_status.pack(side="right")

    def _build_preview_section(self, parent):
        frame = ttk.LabelFrame(parent, text="  3 · Vista previa  ", padding=10)
        frame.pack(fill="both", expand=True, pady=(0, 8))

        toolbar = ttk.Frame(frame)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Label(toolbar, text="Buscar:").pack(side="left")
        self._filter_var = tk.StringVar()
        self._filter_var.trace_add("write", lambda *_: self._apply_filter())
        search = tk.Entry(toolbar, textvariable=self._filter_var, bg=C["surface"],
                          fg=C["text"], insertbackground=C["text"], font=FONTS["BODY"],
                          relief="flat", bd=2, width=28, highlightthickness=1,
                          highlightbackground=C["border"], highlightcolor=C["accent"])
        search.pack(side="left", padx=(6, 0))

        self._lbl_count = ttk.Label(toolbar, text="", foreground=C["text2"])
        self._lbl_count.pack(side="right")

        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True)

        cols = ("num", "original", "nuevo")
        self._tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=14)
        self._tree.heading("num", text="#", anchor="e")
        self._tree.heading("original", text="Original", anchor="w")
        self._tree.heading("nuevo", text="Nuevo nombre", anchor="w")
        self._tree.column("num", width=44, anchor="e")
        self._tree.column("original", width=320, anchor="w")
        self._tree.column("nuevo", width=320, anchor="w")

        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=scroll.set)
        self._tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self._tree.bind("<Enter>", self._on_enter_tree)
        self._tree.bind("<Leave>", self._on_leave_tree)
        self._tree.bind("<Motion>", self._on_hover_thumb)
        self._tree.bind("<MouseWheel>", self._on_scroll_thumb)
        self._tree.bind("<Button-4>", self._on_scroll_thumb)
        self._tree.bind("<Button-5>", self._on_scroll_thumb)

    def _build_action_section(self, parent):
        frame = ttk.Frame(parent)
        frame.pack(fill="x")

        self._progress = ttk.Progressbar(frame, mode="determinate")
        self._progress.pack(fill="x", pady=(0, 8))

        bottom = ttk.Frame(frame)
        bottom.pack(fill="x")

        self._lbl_status = ttk.Label(bottom, text="Listo para comenzar", foreground=C["text2"])
        self._lbl_status.pack(side="left")

        btn_row = ttk.Frame(bottom)
        btn_row.pack(side="right")

        self._copy_var = tk.BooleanVar()
        ttk.Checkbutton(btn_row, text="Modo copiar", variable=self._copy_var).pack(
            side="left", padx=(0, 10))

        self._btn_log = ttk.Button(btn_row, text="💾 Log", command=self._on_export_log,
                                    state="disabled")
        self._btn_log.pack(side="left", padx=(0, 4))
        self._btn_csv = ttk.Button(btn_row, text="📤 CSV", command=self._on_export_csv,
                                    state="disabled")
        self._btn_csv.pack(side="left", padx=(0, 4))
        self._btn_cancel = ttk.Button(btn_row, text="✖ Cancelar", command=self._on_cancel,
                                       state="disabled")
        self._btn_cancel.pack(side="left", padx=(0, 4))
        self._btn_undo = ttk.Button(btn_row, text="↩ Deshacer", command=self._on_undo,
                                     state="disabled")
        self._btn_undo.pack(side="left", padx=(0, 4))
        self._btn_rename = ttk.Button(btn_row, text="▶  Renombrar todo", style="Accent.TButton",
                                       command=self._on_rename, state="disabled")
        self._btn_rename.pack(side="left")

    # ── exploradores de archivos nativos ────────────────────────────────
    def _browse_folder(self):
        import subprocess, platform
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
                    path = filedialog.askdirectory(
                        title="Selecciona la carpeta de fotos", mustexist=True)
        else:
            path = filedialog.askdirectory(
                title="Selecciona la carpeta de fotos", mustexist=True)
        if path:
            self._entry_folder.delete(0, "end")
            self._entry_folder.insert(0, path)

    def _browse_excel(self):
        import subprocess, platform
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
                        title="Selecciona el archivo Excel",
                        filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
        else:
            path = filedialog.askopenfilename(
                title="Selecciona el archivo Excel",
                filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
        if path:
            self._entry_excel.delete(0, "end")
            self._entry_excel.insert(0, path)

    # ── handlers ───────────────────────────────────────────────────────
    def _on_load_photos(self):
        raw = self._entry_folder.get().strip()
        if not raw:
            messagebox.showwarning("Image Sync", "Selecciona primero una carpeta.")
            return
        path = Path(raw)
        if not path.is_dir():
            self._lbl_folder_status.configure(text="Ruta no existe", foreground=C["err"])
            return
        self._model.folder_path = path
        self._model.sort_mode = SORT_OPTIONS.get(self._sort_var.get(), "natural")
        _save_state({"last_folder": str(path), "sort": self._model.sort_mode})
        try:
            n = self._model.load_photos()
        except Exception as exc:
            self._lbl_folder_status.configure(text=str(exc), foreground=C["err"])
            return
        if n == 0:
            self._lbl_folder_status.configure(text="No se encontraron imágenes", foreground=C["warn"])
        else:
            self._lbl_folder_status.configure(
                text=f"{n} imagen{'es' if n != 1 else ''}  •  {self._sort_var.get()}",
                foreground=C["ok"])
        self._refresh_preview()

    def _on_load_excel(self):
        raw = self._entry_excel.get().strip()
        if not raw:
            messagebox.showwarning("Image Sync", "Selecciona primero un archivo Excel.")
            return
        path = Path(raw)
        if not path.is_file() or path.suffix.lower() != ".xlsx":
            self._lbl_excel_status.configure(text="Debe ser .xlsx", foreground=C["err"])
            return
        self._model.excel_path = path
        _save_state({"last_excel": str(path)})
        try:
            sheets = self._model.load_sheets()
        except Exception as exc:
            self._lbl_excel_status.configure(text=str(exc), foreground=C["err"])
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
            self._lbl_excel_status.configure(text=str(exc), foreground=C["err"])
            return
        self._loading_columns = True
        if len(cols) == 1:
            self._model.column_name = cols[0]
            self._col_menu.configure(state="disabled")
            self._lbl_excel_status.configure(
                text=f"Columna «{cols[0]}» auto-seleccionada.", foreground=C["ok"])
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
            self._lbl_excel_status.configure(
                text=f"{len(cols)} columnas — elige cuál.", foreground=C["ok"])
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
            self._lbl_excel_status.configure(text=str(exc), foreground=C["err"])
            return
        self._lbl_excel_status.configure(
            text=f"{n} nombre{'s' if n != 1 else ''} en «{self._model.column_name}».",
            foreground=C["ok"])
        if self._model.skipped_rows:
            rows = ", ".join(str(r) for r in self._model.skipped_rows[:10])
            messagebox.showwarning("Image Sync",
                f"Fila(s) {rows} sin texto legible.\nPuede ser celda vacía o fórmula sin valor.")
        self._refresh_preview()

    def _refresh_preview(self):
        self._last_pairs = self._model.build_preview()
        self._populate_tree(self._last_pairs)
        has_data = bool(self._last_pairs)
        self._btn_rename.configure(state="normal" if has_data else "disabled")
        self._btn_log.configure(state="normal" if has_data else "disabled")
        self._btn_csv.configure(state="normal" if has_data else "disabled")
        self._lbl_count.configure(text=f"{len(self._last_pairs)} archivos" if has_data else "")

    def _populate_tree(self, pairs):
        self._tree.delete(*self._tree.get_children())
        self._thumb_refs.clear()
        for i, (orig, new, _) in enumerate(pairs):
            self._tree.insert("", "end", iid=str(i), values=(i + 1, orig, new))

    def _apply_filter(self):
        q = self._filter_var.get().lower()
        visible = 0
        for i, (orig, new, _) in enumerate(self._last_pairs):
            match = not q or q in orig.lower() or q in new.lower()
            if match:
                try:
                    self._tree.reattach(str(i), "", "end")
                except Exception:
                    pass
                visible += 1
            else:
                try:
                    self._tree.detach(str(i))
                except Exception:
                    pass
        if q:
            self._lbl_count.configure(text=f"{visible} de {len(self._last_pairs)}")
        else:
            self._lbl_count.configure(text=f"{len(self._last_pairs)} archivos")

    # ── miniaturas hover ────────────────────────────────────────────────
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
        region = self._tree.identify_region(event.x, event.y)
        if region != "cell":
            self._thumb_win.withdraw()
            return
        item = self._tree.identify_row(event.y)
        col = self._tree.identify_column(event.x)
        if not item or col not in ("#2", "#3"):
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
        if event.widget is self:
            if self._thumb_win:
                try:
                    self._thumb_win.withdraw()
                except Exception:
                    pass

    def _on_restore(self, event):
        pass

    # ── renombrar ──────────────────────────────────────────────────────
    def _on_rename(self):
        if not (self._model.photos and self._model.names):
            messagebox.showwarning("Image Sync", "Carga las fotos y el Excel primero.")
            return
        n_ph = len(self._model.photos)
        n_nm = len(self._model.names)
        will = min(n_ph, n_nm)
        msg = f"Se renombrarán {will} foto{'s' if will != 1 else ''}."
        if n_nm < n_ph:
            msg += f"\n⚠ {n_ph - n_nm} foto(s) sin nombre — solo primeras {will}."
        if not messagebox.askyesno("Confirmar renombramiento", msg):
            return
        self._cancel_ev = threading.Event()
        self._btn_rename.configure(state="disabled", text="Renombrando…")
        self._btn_cancel.configure(state="normal")
        self._progress["value"] = 0
        threading.Thread(target=self._do_rename, daemon=True).start()

    def _do_rename(self):
        def progress(cur, tot, name):
            self.after(0, lambda: self._set_progress(cur / tot, f"{cur}/{tot} — {name}"))
        def done(ok, errors):
            self.after(0, lambda: self._finish_rename(ok, errors))
        self._model.rename_all(progress, done, cancel_ev=self._cancel_ev,
                               copy_mode=self._copy_var.get())

    def _finish_rename(self, ok, errors):
        self._progress["value"] = 100
        self._lbl_status.configure(text=f"Completado · {ok} renombradas.", foreground=C["ok"])
        self._btn_rename.configure(state="normal", text="▶  Renombrar todo")
        self._btn_cancel.configure(state="disabled")
        self._btn_undo.configure(state="normal" if self._model.has_undo else "disabled")
        if errors:
            messagebox.showwarning("Image Sync", f"⚠ {ok} OK · {len(errors)} con error.")
        else:
            messagebox.showinfo("Image Sync", f"✓ {ok} foto{'s' if ok != 1 else ''} renombradas.")
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
        self._btn_undo.configure(state="disabled")
        self._progress["value"] = 0
        threading.Thread(target=self._do_undo, daemon=True).start()

    def _do_undo(self):
        def progress(cur, tot, name):
            self.after(0, lambda: self._set_progress(cur / tot, f"Revirtiendo {cur}/{tot} — {name}"))
        def done(ok, errors):
            self.after(0, lambda: self._finish_undo(ok, errors))
        self._model.undo_last(progress, done)

    def _finish_undo(self, ok, errors):
        self._progress["value"] = 0
        self._lbl_status.configure(text=f"Deshacer · {ok} revertidas.", foreground=C["ok"])
        self._btn_undo.configure(state="normal" if self._model.has_undo else "disabled")
        if self._model.folder_path:
            try:
                self._model.load_photos()
            except Exception:
                pass
            self._refresh_preview()
        if errors:
            messagebox.showwarning("Image Sync", f"↩ {ok} OK · {len(errors)} con error.")
        else:
            messagebox.showinfo("Image Sync", f"↩ {ok} revertida{'s' if ok != 1 else ''}.")

    def _on_cancel(self):
        if self._cancel_ev:
            self._cancel_ev.set()
            self._btn_cancel.configure(state="disabled")
            self._lbl_status.configure(text="Cancelando…", foreground=C["warn"])

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
                messagebox.showinfo("Image Sync", f"Log guardado: {Path(dest).name}")
            except Exception as exc:
                messagebox.showerror("Image Sync", f"Error: {exc}")

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
                messagebox.showinfo("Image Sync", f"CSV guardado: {Path(dest).name}")
            except Exception as exc:
                messagebox.showerror("Image Sync", f"Error: {exc}")

    def _set_progress(self, frac, msg):
        self._progress["value"] = frac * 100
        self._lbl_status.configure(text=msg)

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
