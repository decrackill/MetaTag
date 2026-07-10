"""
MetaTag Visor v9.0 — Edición Definitiva (Código Completo y Expandido)
Paneles ajustables, Persistencia inteligente, Diseño Premium y Zoom Fluido.
Dependencias: pip install pillow piexif reportlab
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys
import os
import io
import json
import math
import datetime
import threading
import subprocess
import logging
import gc
from pathlib import Path

logging.basicConfig(
    filename='metatag_debug.log',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ══════════════════════════════════════════════════════════════════
#  VERIFICACIÓN DE DEPENDENCIAS
# ══════════════════════════════════════════════════════════════════
try:
    from PIL import Image, ImageTk, ImageOps, ImageDraw
    import piexif
    import piexif.helper
    PIL_OK = True
except ImportError:
    PIL_OK = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Image as RLImage,
        Table, TableStyle, HRFlowable
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False


# ══════════════════════════════════════════════════════════════════
#  TEMAS PREMIUM
# ══════════════════════════════════════════════════════════════════
THEMES = {
    "Oro Arqueológico": {
        "bg":           "#0A0A0B",
        "surface":      "#111114",
        "panel":        "#16161A",
        "panel2":       "#1C1C21",
        "border":       "#2E2E38",
        "accent":       "#D4A574",
        "accent_hover": "#F0C497",
        "text":         "#EDE8DF",
        "text2":        "#A1988C",
        "text3":        "#5E5750",
        "header_bg":    "#0F0F12",
        "header_fg":    "#D4A574",
        "json_section": "#D4A574",
        "exif_section": "#7EB8C9",
        "gps_section":  "#A8D4A0",
        "file_section": "#8A8A9A",
        "sel_bg":       "#3D2A10",
        "sel_fg":       "#EDE8DF",
        "btn_ghost":    "#16161A",
        "ok":           "#7EC894",
        "err":          "#E05050",
        "row_alt":      "#131318",
        "zoom_badge":   "#D4A574",
    },
    "Noche Total": {
        "bg":           "#000000",
        "surface":      "#0A0A0A",
        "panel":        "#111111",
        "panel2":       "#181818",
        "border":       "#2A2A2A",
        "accent":       "#BB86FC",
        "accent_hover": "#D0A8FF",
        "text":         "#FFFFFF",
        "text2":        "#AAAAAA",
        "text3":        "#555555",
        "header_bg":    "#050505",
        "header_fg":    "#BB86FC",
        "json_section": "#03DAC6",
        "exif_section": "#BB86FC",
        "gps_section":  "#88BBFF",
        "file_section": "#AAAAAA",
        "sel_bg":       "#3700B3",
        "sel_fg":       "#FFFFFF",
        "btn_ghost":    "#111111",
        "ok":           "#03DAC6",
        "err":          "#CF6679",
        "row_alt":      "#080808",
        "zoom_badge":   "#BB86FC",
    },
    "Carbón": {
        "bg":           "#1E1E1E",
        "surface":      "#252526",
        "panel":        "#2D2D30",
        "panel2":       "#333337",
        "border":       "#3E3E42",
        "accent":       "#569CD6",
        "accent_hover": "#79B8FF",
        "text":         "#D4D4D4",
        "text2":        "#9CDCFE",
        "text3":        "#666666",
        "header_bg":    "#007ACC",
        "header_fg":    "#FFFFFF",
        "json_section": "#4EC9B0",
        "exif_section": "#569CD6",
        "gps_section":  "#CE9178",
        "file_section": "#AAAAAA",
        "sel_bg":       "#264F78",
        "sel_fg":       "#FFFFFF",
        "btn_ghost":    "#2D2D30",
        "ok":           "#4EC9B0",
        "err":          "#F44747",
        "row_alt":      "#222222",
        "zoom_badge":   "#569CD6",
    },
}

# ══════════════════════════════════════════════════════════════════
#  CONFIGURACIONES Y CONSTANTES GLOBALES
# ══════════════════════════════════════════════════════════════════
if getattr(sys, "frozen", False):
    _BASE_DIR = Path(sys.executable).parent
else:
    _BASE_DIR = Path(__file__).parent

CONFIG_FILE = _BASE_DIR / "visor_config.json"

# Cargar tema inicial desde la configuración
try:
    if CONFIG_FILE.exists():
        _saved_config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        CURRENT_THEME = _saved_config.get("theme", "Oro Arqueológico")
        if CURRENT_THEME not in THEMES:
            CURRENT_THEME = "Oro Arqueológico"
    else:
        CURRENT_THEME = "Oro Arqueológico"
except Exception:
    CURRENT_THEME = "Oro Arqueológico"

C = dict(THEMES[CURRENT_THEME])
IMG_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".webp"}


# ══════════════════════════════════════════════════════════════════
#  DIÁLOGOS NATIVOS DEL SISTEMA OPERATIVO
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


def _native_save_file(title="Guardar archivo", default_name="", filetypes=None):
    if sys.platform == "linux":
        cmd = ["zenity", "--file-selection", "--save", "--confirm-overwrite",
               "--title", title]
        if default_name:
            cmd += ["--filename", default_name]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            if r.returncode == 0 and r.stdout.strip():
                return r.stdout.strip().split("\n")[0]
            if r.returncode == 1:
                return None
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
    return filedialog.asksaveasfilename(
        title=title, defaultextension=".pdf",
        initialfile=default_name, filetypes=filetypes or [("PDF", "*.pdf")])


def _format_file_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"

# Tipografía del programa
F_TITLE = ("Segoe UI", 14, "bold")
F_H2    = ("Segoe UI", 11, "bold")
F_BODY  = ("Segoe UI", 9)
F_BOLD  = ("Segoe UI", 9, "bold")
F_TINY  = ("Segoe UI", 8)
F_MICRO = ("Segoe UI", 7)

# Constantes del Zoom
ZOOM_MIN = 0.05
ZOOM_MAX = 50.0
ZOOM_STEP = 1.18
ZOOM_FIT = -1.0


# ══════════════════════════════════════════════════════════════════
#  CLASE PRINCIPAL DEL VISOR
# ══════════════════════════════════════════════════════════════════
class VisorApp(tk.Tk):

    def __init__(self, initial_image=None):
        super().__init__()

        # Configuración de la ventana principal
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = min(int(screen_width * 0.90), 1380)
        window_height = min(int(screen_height * 0.90), 860)
        
        pos_x = (screen_width - window_width) // 2
        pos_y = (screen_height - window_height) // 2
        
        self.geometry(f"{window_width}x{window_height}+{pos_x}+{pos_y}")
        self.minsize(960, 640)
        self.configure(bg=C["bg"])
        self.title("MetaTag Visor v9.0 — Inspector Arqueológico")

        # Variables de estado de la aplicación
        self.all_metadata = []
        self.folder_images = []
        self._all_folder_images = []
        self._current_folder = ""
        self.current_path = None
        self._pil_image = None

        # Variables de estado del Zoom y Paneo
        self._zoom_level = ZOOM_FIT
        self._pan_offset = [0, 0]
        self._pan_start  = [0, 0]
        self._panning    = False
        self._zoom_img_tk = None   
        self._zoom_img_tk_draft = None   
        self._zoom_debounce_job = None   
        self._render_gen = 0      

        # Variables del Comparador
        self._comp_window = None
        self._comp_thumbs = []
        self._comp_folder_a = ""
        self._comp_folder_b = ""

        # Construcción de la Interfaz
        self._build_styles()
        self._build_ui()

        # Iniciar sesión y persistencia
        self._init_session(initial_image)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ─────────────────────────────────────────────────────────────
    #  SISTEMA DE PERSISTENCIA
    # ─────────────────────────────────────────────────────────────
    def _init_session(self, initial_image):
        """Inicializa la sesión recordando la última carpeta o imagen abierta."""
        try:
            if CONFIG_FILE.exists():
                cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            else:
                cfg = {}
        except Exception:
            cfg = {}

        last_dir = cfg.get("last_folder", "")
        last_img = cfg.get("last_image", "")
        self._comp_folder_a = cfg.get("comp_folder_a", "")
        self._comp_folder_b = cfg.get("comp_folder_b", "")

        # Si el programa se abrió enviándole una imagen por parámetro (Abrir con...)
        if initial_image and initial_image != "STANDALONE" and os.path.exists(initial_image):
            self._hide_explorer_pane()
            self.load_image(initial_image)
            return

        # 1. Restaurar la carpeta si existe y tiene imágenes
        folder_loaded = False
        if last_dir and os.path.isdir(last_dir):
            self._open_folder_path(last_dir, auto_load=False)
            if self.folder_images:
                self._show_explorer_pane()
                folder_loaded = True

        # 2. Restaurar la imagen exacta
        if last_img and os.path.exists(last_img):
            if self.folder_images and last_img in self.folder_images:
                self.load_image(last_img)
                idx = self.folder_images.index(last_img)
                children = self.file_tree.get_children()
                if idx < len(children):
                    self.file_tree.selection_set(children[idx])
                    self.file_tree.see(children[idx])
            else:
                self._hide_explorer_pane()
                self.load_image(last_img)
        elif self.folder_images:
            self.load_image(self.folder_images[0])
            children = self.file_tree.get_children()
            if children:
                self.file_tree.selection_set(children[0])
        else:
            self._show_welcome()
            self._hide_explorer_pane()

    def _save_config(self):
        """Guarda el estado actual antes de cerrar."""
        try:
            last_folder_path = ""
            if self.current_path and self.folder_images:
                last_folder_path = str(Path(self.current_path).parent)

            cfg = {
                "last_image": self.current_path or "",
                "last_folder": last_folder_path,
                "theme": CURRENT_THEME,
                "comp_folder_a": getattr(self, "_comp_folder_a", ""),
                "comp_folder_b": getattr(self, "_comp_folder_b", ""),
            }
            CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as e:
            print(f"Error guardando configuración: {e}")

    def _on_close(self):
        self._save_config()
        self.destroy()

    # ─────────────────────────────────────────────────────────────
    #  ESTILOS TTK
    # ─────────────────────────────────────────────────────────────
    def _build_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        
        # Estilo del Treeview (Tabla)
        style.configure(
            "Treeview", 
            background=C["surface"], 
            foreground=C["text"], 
            fieldbackground=C["surface"], 
            rowheight=32, 
            font=F_BODY, 
            borderwidth=0
        )
        style.configure(
            "Treeview.Heading", 
            background=C["panel"], 
            foreground=C["accent"], 
            font=F_BOLD, 
            borderwidth=0, 
            padding=(10, 8)
        )
        style.map(
            "Treeview", 
            background=[("selected", C["sel_bg"])], 
            foreground=[("selected", C["sel_fg"])]
        )
        
        # Estilo de las barras de desplazamiento (Scrollbars)
        style.configure(
            "TScrollbar", 
            background=C["border"], 
            troughcolor=C["surface"], 
            arrowcolor=C["text"], 
            borderwidth=0, 
            relief="flat", 
            width=12, 
            arrowsize=10
        )
        style.map(
            "TScrollbar", 
            background=[("active", C["accent"]), ("!active", C["panel2"])]
        )

    # ─────────────────────────────────────────────────────────────
    #  CONSTRUCCIÓN DE LA INTERFAZ GRÁFICA (UI)
    # ─────────────────────────────────────────────────────────────
    def _build_ui(self):
        self._build_topbar()
        
        # Línea separadora bajo el topbar
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        # ── SPLITTER PRINCIPAL (Izquierda / Derecha) ──
        self.main_paned = tk.PanedWindow(
            self, 
            orient="horizontal", 
            bg=C["bg"], 
            sashwidth=8, 
            bd=0, 
            cursor="arrow",
            sashcursor="sb_h_double_arrow"
        )
        self.main_paned.pack(fill="both", expand=True, padx=10, pady=10)

        # ── SPLITTER IZQUIERDO (Imagen Arriba / Explorador Abajo) ──
        self.left_paned = tk.PanedWindow(
            self.main_paned, 
            orient="vertical", 
            bg=C["bg"], 
            sashwidth=8, 
            bd=0, 
            cursor="arrow",
            sashcursor="sb_v_double_arrow"
        )
        self.main_paned.add(self.left_paned, minsize=320, width=400)

        # Sección Superior Izquierda (Visor de Imagen)
        self.left_top = tk.Frame(self.left_paned, bg=C["panel"])
        self.left_paned.add(self.left_top, minsize=320)
        self._build_left_top(self.left_top)

        # Sección Inferior Izquierda (Explorador de Archivos)
        self.left_bottom = tk.Frame(self.left_paned, bg=C["panel"])
        self.left_paned.add(self.left_bottom, minsize=180)
        self._build_left_bottom(self.left_bottom)

        # ── PANEL DERECHO (Tabla de Metadatos) ──
        self.right_frame = tk.Frame(self.main_paned, bg=C["bg"])
        self.main_paned.add(self.right_frame, minsize=480)
        self._build_right(self.right_frame)

        # Pantalla de Bienvenida (Superpuesta al inicio)
        self._welcome_frame = tk.Frame(
            self.right_frame, 
            bg=C["surface"], 
            highlightthickness=1, 
            highlightbackground=C["border"]
        )

    def _build_topbar(self):
        bar = tk.Frame(self, bg=C["header_bg"], height=60)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        # Logo y Título
        logo_frame = tk.Frame(bar, bg=C["header_bg"])
        logo_frame.pack(side="left", padx=20, pady=10)
        tk.Label(logo_frame, text="⬡", font=("Segoe UI", 18), bg=C["header_bg"], fg=C["accent"]).pack(side="left")
        tk.Label(logo_frame, text="  Visor MetaTag", font=F_TITLE, bg=C["header_bg"], fg=C["header_fg"]).pack(side="left")
        tk.Label(logo_frame, text="  v9.0", font=F_TINY, bg=C["header_bg"], fg=C["text3"]).pack(side="left", pady=(8, 0))

        # Contenedor de Botones Derechos
        right_btns_frame = tk.Frame(bar, bg=C["header_bg"])
        right_btns_frame.pack(side="right", padx=15)

        # Selector de Temas
        self._theme_var = tk.StringVar(value=CURRENT_THEME)
        theme_pill = tk.Frame(right_btns_frame, bg=C["btn_ghost"], highlightthickness=1, highlightbackground=C["border"])
        theme_pill.pack(side="left", padx=(0, 20))
        
        for t in THEMES:
            is_cur = (t == CURRENT_THEME)
            bg = C["accent"] if is_cur else C["btn_ghost"]
            fg = C["bg"] if is_cur else C["text2"]
            b = tk.Button(
                theme_pill, text=f" {t} ", bg=bg, fg=fg, font=F_TINY, 
                relief="flat", bd=0, padx=12, pady=5, cursor="hand2", 
                command=lambda n=t: self._select_theme(n)
            )
            b.pack(side="left")

        # Botones de Acción
        actions = [
            ("📄 Archivo",      self._browse,          False),
            ("📁 Carpeta",      self._browse_folder,   False),
            ("⚖ Comparar",     self._open_comparator, False),
            ("⬇ PDF",          self._export_pdf,      True),
        ]
        
        for text, command, is_primary in actions:
            self._create_friendly_btn(right_btns_frame, text, command, is_primary).pack(side="left", padx=4)

    def _create_friendly_btn(self, parent, text, command, is_primary=False):
        """Crea un botón con marco dinámico (efecto hover moderno)"""
        bg_normal = C["accent"] if is_primary else C["panel"]
        fg_normal = C["bg"] if is_primary else C["text"]
        bg_hover = C["accent_hover"] if is_primary else C["panel2"]
        border_hover = C["accent"] if is_primary else C["text3"]
        
        frame = tk.Frame(parent, bg=C["border"], padx=1, pady=1)
        btn = tk.Button(
            frame, text=text, command=command, bg=bg_normal, fg=fg_normal, 
            font=F_BOLD, relief="flat", bd=0, padx=16, pady=7, cursor="hand2", 
            activebackground=bg_hover, activeforeground=fg_normal
        )
        btn.pack(expand=True, fill="both")
        
        # Efectos Hover
        def on_enter(e):
            btn.configure(bg=bg_hover)
            frame.configure(bg=border_hover)
            
        def on_leave(e):
            btn.configure(bg=bg_normal)
            frame.configure(bg=C["border"])
            
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        
        return frame

    def _build_left_top(self, parent):
        """Construye el visor de imágenes y controles de zoom."""
        # Contenedor del Canvas
        img_outer = tk.Frame(parent, bg=C["surface"], highlightthickness=1, highlightbackground=C["border"])
        img_outer.pack(fill="both", expand=True, padx=12, pady=(12, 6))

        self.img_canvas = tk.Canvas(img_outer, bg=C["bg"], highlightthickness=0)
        self.img_canvas.pack(fill="both", expand=True, padx=2, pady=2)

        # Botones de Zoom Rápidos
        zoom_bar = tk.Frame(parent, bg=C["panel"])
        zoom_bar.pack(fill="x", padx=12, pady=(0, 6))
        
        zoom_actions = [
            ("➕ Zoom In", lambda: self._zoom_by(ZOOM_STEP)), 
            ("➖ Zoom Out", lambda: self._zoom_by(1 / ZOOM_STEP)), 
            ("⛶ Ajustar", self._zoom_fit)
        ]
        
        for text, command in zoom_actions:
            b = tk.Button(
                zoom_bar, text=text, font=F_TINY, bg=C["panel2"], fg=C["text2"], 
                relief="flat", bd=0, padx=12, pady=5, cursor="hand2", command=command, 
                highlightthickness=1, highlightbackground=C["border"]
            )
            b.pack(side="left", padx=(0, 6))
            b.bind("<Enter>", lambda e, _b=b: _b.configure(bg=C["border"]))
            b.bind("<Leave>", lambda e, _b=b: _b.configure(bg=C["panel2"]))

        # Tarjeta de Información Básica
        self.info_card = tk.Frame(parent, bg=C["panel2"], highlightthickness=1, highlightbackground=C["border"])
        self.info_card.pack(fill="x", padx=12, pady=(0, 12))
        
        tk.Label(
            self.info_card, text="ARCHIVO ACTUAL", bg=C["panel2"], 
            fg=C["text3"], font=("Segoe UI", 7, "bold")
        ).pack(anchor="w", padx=12, pady=(8, 2))
        
        self.lbl_info = tk.Label(
            self.info_card, text="Sin selección.", bg=C["panel2"], 
            fg=C["text2"], font=F_BODY, justify="left", wraplength=340
        )
        self.lbl_info.pack(anchor="w", padx=12, pady=(0, 10))

        # Conectar los eventos del ratón al canvas
        self.after(100, self._bind_zoom_events)

    def _build_left_bottom(self, parent):
        """Construye el explorador de archivos nativo con Treeview."""
        # Barra de navegación superior
        nav_frame = tk.Frame(parent, bg=C["panel"])
        nav_frame.pack(fill="x", padx=8, pady=(6, 4))

        tk.Label(nav_frame, text="EXPLORADOR", bg=C["panel"],
                 fg=C["accent"], font=("Segoe UI", 9, "bold")).pack(side="left")

        self._lbl_count = tk.Label(nav_frame, text="", bg=C["panel"], fg=C["text3"], font=F_MICRO)
        self._lbl_count.pack(side="right", padx=(0, 4))

        # Barra de ruta actual
        self._path_var = tk.StringVar(value="")
        path_frame = tk.Frame(parent, bg=C["panel2"], highlightthickness=1, highlightbackground=C["border"])
        path_frame.pack(fill="x", padx=8, pady=(0, 4))

        self._path_label = tk.Label(
            path_frame, textvariable=self._path_var, bg=C["panel2"],
            fg=C["text2"], font=F_MICRO, anchor="w", wraplength=280)
        self._path_label.pack(side="left", fill="x", expand=True, ipady=4, padx=6)

        # Barra de búsqueda
        search_frame = tk.Frame(parent, bg=C["panel2"], highlightthickness=1, highlightbackground=C["border"])
        search_frame.pack(fill="x", padx=8, pady=(0, 4))

        tk.Label(search_frame, text=" 🔍 ", bg=C["panel2"], fg=C["text3"], font=F_TINY).pack(side="left")
        self._folder_search = tk.StringVar()
        self._folder_search.trace_add("write", self._filter_folder)
        tk.Entry(
            search_frame, textvariable=self._folder_search, bg=C["panel2"],
            fg=C["text"], insertbackground=C["accent"], relief="flat", bd=0, font=F_BODY
        ).pack(side="left", fill="x", expand=True, ipady=5)

        # Treeview de archivos
        tree_frame = tk.Frame(parent, bg=C["surface"], highlightthickness=1, highlightbackground=C["border"])
        tree_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        cols = ("size", "date")
        self.file_tree = ttk.Treeview(
            tree_frame, columns=cols, show="tree headings", selectmode="browse")

        self.file_tree.heading("#0", text="Nombre", anchor="w")
        self.file_tree.heading("size", text="Tamaño", anchor="e")
        self.file_tree.heading("date", text="Fecha", anchor="w")

        self.file_tree.column("#0", width=180, minwidth=100, anchor="w")
        self.file_tree.column("size", width=65, minwidth=50, anchor="e", stretch=False)
        self.file_tree.column("date", width=75, minwidth=60, anchor="w", stretch=False)

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side="right", fill="y")
        self.file_tree.pack(fill="both", expand=True)

        self.file_tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.file_tree.bind("<Double-Button-1>", self._on_tree_double_click)

        # Mapa de ruta por item ID
        self._tree_item_paths = {}

    def _build_right(self, parent):
        """Construye la tabla principal de metadatos."""
        # Buscador superior
        search_frame = tk.Frame(parent, bg=C["bg"])
        search_frame.pack(fill="x", pady=(0, 12))
        
        tk.Label(search_frame, text="🔍 Búsqueda rápida:", bg=C["bg"], fg=C["text2"], font=F_BODY).pack(side="left", padx=(0, 10))
        
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", self._filter_tree)
        
        entry_wrap = tk.Frame(search_frame, bg=C["surface"], highlightthickness=1, highlightbackground=C["border"])
        entry_wrap.pack(side="left", fill="x", expand=True)
        
        tk.Entry(
            entry_wrap, textvariable=self._search_var, bg=C["surface"], 
            fg=C["text"], insertbackground=C["accent"], relief="flat", bd=0, font=F_BODY
        ).pack(fill="x", ipady=6, padx=8)

        # Tabla Treeview
        tree_frame = tk.Frame(parent, bg=C["border"], highlightthickness=1, highlightbackground=C["border"])
        tree_frame.pack(fill="both", expand=True)
        
        self.tree = ttk.Treeview(tree_frame, columns=("Campo", "Valor"), show="headings", selectmode="browse")
        self.tree.heading("Campo", text="PROPIEDAD / ETIQUETA", anchor="w")
        self.tree.heading("Valor", text="INFORMACIÓN EXTRACTADA", anchor="w")
        
        self.tree.column("Campo", width=250, anchor="w", minwidth=140)
        self.tree.column("Valor", width=500, anchor="w", minwidth=200)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        # Configuración de Etiquetas (Tags) de colores en la tabla
        self.tree.tag_configure("odd",  background=C["surface"])
        self.tree.tag_configure("even", background=C["row_alt"])
        self.tree.tag_configure("hdr_json", foreground=C["json_section"], background=C["panel2"], font=F_BOLD)
        self.tree.tag_configure("hdr_exif", foreground=C["exif_section"], background=C["panel2"], font=F_BOLD)
        self.tree.tag_configure("hdr_gps",  foreground=C["gps_section"],  background=C["panel2"], font=F_BOLD)
        self.tree.tag_configure("hdr_file", foreground=C["file_section"], background=C["panel2"], font=F_BOLD)

    # ─────────────────────────────────────────────────────────────
    #  LÓGICA DE INTERFAZ GENERAL Y MANEJO DE PANELES
    # ─────────────────────────────────────────────────────────────
    def _hide_explorer_pane(self):
        """Oculta el panel inferior izquierdo si solo se abre una imagen."""
        if str(self.left_bottom) in self.left_paned.panes():
            self.left_paned.remove(self.left_bottom)

    def _show_explorer_pane(self):
        """Muestra el panel inferior izquierdo."""
        if str(self.left_bottom) not in self.left_paned.panes():
            self.left_paned.add(self.left_bottom, minsize=150)

    def _show_welcome(self):
        """Muestra la pantalla de inicio sobre la tabla vacía."""
        f = self._welcome_frame
        for w in f.winfo_children(): 
            w.destroy()
            
        f.place(relx=0.5, rely=0.5, anchor="center", width=560, height=320)

        tk.Label(f, text="⬡ METATAG PRO", font=("Segoe UI", 28, "bold"), bg=C["surface"], fg=C["accent"]).pack(pady=(50, 5))
        tk.Label(f, text="Auditoría y Clasificación Cerámica", font=("Segoe UI", 11), bg=C["surface"], fg=C["text2"]).pack(pady=(0, 40))

        row = tk.Frame(f, bg=C["surface"])
        row.pack()
        
        self._create_friendly_btn(row, "📄 Archivo Único", self._browse, False).pack(side="left", padx=10)
        self._create_friendly_btn(row, "📁 Carpeta Completa", self._browse_folder, True).pack(side="left", padx=10)

    def _hide_welcome(self):
        """Oculta la pantalla de bienvenida."""
        if self._welcome_frame.winfo_ismapped():
            self._welcome_frame.place_forget()

    def _select_theme(self, name: str):
        """Cambia el tema dinámicamente y reordena los paneles."""
        global CURRENT_THEME, C
        CURRENT_THEME = name
        C = dict(THEMES[CURRENT_THEME])
        self._save_config()
        
        # Guardar estado de la sesión actual
        saved_path = self.current_path
        saved_folder = self.folder_images[:]
        
        # Limpiar ventana completa
        for w in self.winfo_children(): 
            w.destroy()
        
        # Reconstruir UI con nuevos colores
        self.configure(bg=C["bg"])
        self._build_styles()
        self._build_ui()
        
        # Restaurar estado
        if saved_folder:
            self.folder_images = saved_folder
            self._all_folder_images = saved_folder[:]
            self._show_explorer_pane()
            self._refresh_file_tree()
            
        if saved_path and os.path.exists(saved_path):
            self.load_image(saved_path)
            if saved_path in self.folder_images:
                children = self.file_tree.get_children()
                idx = self.folder_images.index(saved_path)
                if idx < len(children):
                    self.file_tree.selection_set(children[idx])
                    self.file_tree.see(children[idx])
        elif not self.folder_images:
            self._show_welcome()

    # ─────────────────────────────────────────────────────────────
    #  NAVEGACIÓN DE ARCHIVOS
    # ─────────────────────────────────────────────────────────────
    def _browse(self):
        """Abre un solo archivo mediante diálogo nativo del SO."""
        file_path = _native_file_open(
            title="Seleccionar imagen",
            filetypes=[("Imágenes", "*.jpg *.jpeg *.png *.tif *.tiff *.webp"),
                       ("Todos", "*.*")])
        if file_path and os.path.exists(file_path):
            self._hide_explorer_pane()
            self.folder_images = []
            self.load_image(file_path)
        else:
            if not self.folder_images and not self.current_path:
                self._show_welcome()

    def _browse_folder(self):
        """Abre una carpeta mediante diálogo nativo del SO."""
        folder_path = _native_folder_open(title="Seleccionar carpeta de estudio")
        if folder_path and os.path.isdir(folder_path):
            self._open_folder_path(folder_path, auto_load=True)
        elif not self.folder_images:
            self._show_welcome()

    def _open_folder_path(self, folder: str, auto_load: bool = True):
        """Procesa una ruta de directorio y extrae las imágenes válidas."""
        path_obj = Path(folder)
        images = []
        try:
            for item in path_obj.iterdir():
                if item.is_file() and item.suffix.lower() in IMG_EXTS:
                    images.append(str(item))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer la carpeta:\n{e}")
            return

        images = sorted(images, key=lambda p: Path(p).name.lower())

        if not images:
            if auto_load:
                messagebox.showinfo("Carpeta vacía", "No hay imágenes compatibles en esta carpeta.")
            return

        self.folder_images = images
        self._all_folder_images = images[:]
        self._current_folder = folder

        self._show_explorer_pane()
        self._path_var.set(folder)
        self._refresh_file_tree()

        if auto_load:
            self.file_tree.selection_set(self.file_tree.get_children()[0])
            self.load_image(self.folder_images[0])

    def _filter_folder(self, *_):
        """Filtra la lista de archivos mediante la barra de búsqueda."""
        search_term = self._folder_search.get().lower()
        all_items = getattr(self, "_all_folder_images", self.folder_images)

        if search_term:
            self.folder_images = [p for p in all_items if search_term in Path(p).name.lower()]
        else:
            self.folder_images = all_items[:]

        self._refresh_file_tree()

    def _refresh_file_tree(self):
        """Actualiza el Treeview del explorador con nombre, tamaño y fecha."""
        self.file_tree.delete(*self.file_tree.get_children())
        self._tree_item_paths.clear()

        for i, filepath in enumerate(self.folder_images):
            p = Path(filepath)
            try:
                stat = p.stat()
                size_str = _format_file_size(stat.st_size)
                date_str = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%d/%m/%y")
            except Exception:
                size_str = "—"
                date_str = "—"

            bg = C["row_alt"] if i % 2 != 0 else C["surface"]
            tag = f"row_{i}"
            item_id = self.file_tree.insert(
                "", "end", text=f"  {p.name}", values=(size_str, date_str),
                tags=(tag,))
            self.file_tree.tag_configure(tag, background=bg)
            self._tree_item_paths[item_id] = filepath

        total_size = sum(Path(f).stat().st_size for f in self.folder_images
                         if Path(f).exists())
        self._lbl_count.configure(
            text=f"{len(self.folder_images)} archivos · {_format_file_size(total_size)}")

    def _on_tree_select(self, event):
        """Evento al seleccionar un archivo en el Treeview."""
        sel = self.file_tree.selection()
        if not sel:
            return
        path = self._tree_item_paths.get(sel[0])
        if path and os.path.exists(path):
            self.load_image(path)

    def _on_tree_double_click(self, event):
        """Doble clic: si es carpeta, navegar a ella."""
        sel = self.file_tree.selection()
        if not sel:
            return
        path = self._tree_item_paths.get(sel[0])
        if path and os.path.isdir(path):
            self._open_folder_path(path, auto_load=True)


    # ─────────────────────────────────────────────────────────────
    #  CARGA DE IMAGEN PRINCIPAL (100% EN SEGUNDO PLANO)
    # ─────────────────────────────────────────────────────────────
    def load_image(self, path: str):
        """
        Inicia la carga en segundo plano con debounce de 80ms.
        Si el usuario hace clic 10 veces en 500ms, solo se procesa
        la última — las anteriores son canceladas antes de lanzar
        cualquier hilo, manteniendo el main loop libre.
        """
        if not PIL_OK: 
            messagebox.showerror("Error", "Faltan dependencias. Ejecuta: pip install pillow piexif")
            return

        # Cancelar carga anterior pendiente (debounce)
        if hasattr(self, "_load_debounce_job") and self._load_debounce_job:
            self.after_cancel(self._load_debounce_job)
            self._load_debounce_job = None

        # Respuesta visual INSTANTÁNEA — sin esperar al disco
        self._hide_welcome()
        self.current_path = path
        self.title(f"⬡ MetaTag Visor v9 — {Path(path).name}")

        # Limpiar canvas inmediatamente para feedback visual
        self.img_canvas.delete("img")
        self.img_canvas.delete("zoom_badge")
        self._pil_image   = None
        self._zoom_img_tk = None
        self._zoom_level  = ZOOM_FIT
        self._pan_offset  = [0, 0]

        # Incrementar generación ahora — cualquier hilo anterior
        # compara contra este valor y se auto-cancela
        self._render_gen += 1
        gen = self._render_gen

        # Lanzar el hilo real después de 80ms de silencio
        def _launch():
            self._load_debounce_job = None
            threading.Thread(
                target=self._process_image_thread,
                args=(path, gen),
                daemon=True
            ).start()

        self._load_debounce_job = self.after(80, _launch)

    def _process_image_thread(self, path, gen):
        """Hilo en segundo plano: lee del disco, descomprime y extrae metadatos."""
        try:
            raw_img = Image.open(path)
            raw_img.load()                    # fuerza descompresión AQUÍ, en el hilo
            img = ImageOps.exif_transpose(raw_img)
            img = img.copy()                  # desacopla del file handle
            
            if gen != self._render_gen:
                return

            self.after(0, lambda: self._apply_image_to_ui(img, path, gen))

            info_dict = raw_img.info

            # Extraer metadatos en lista LOCAL al hilo
            local_meta = []

            local_meta.append(("header", "json", "🏺  DATOS ARQUEOLÓGICOS (JSON)"))
            local_meta.extend(self._extract_json_local(raw_img, info_dict))

            local_meta.append(("header", "exif", "📸  CONFIGURACIÓN CÁMARA (EXIF)"))
            local_meta.extend(self._extract_exif_local(info_dict))

            local_meta.append(("header", "gps", "🌍  UBICACIÓN (GPS)"))
            local_meta.extend(self._extract_gps_local(info_dict))

            local_meta.append(("header", "file", "📄  INFO ARCHIVO SO"))
            local_meta.extend([
                ("file_val", "Nombre del Archivo",        Path(path).name),
                ("file_val", "Ubicación del Directorio",  str(Path(path).parent)),
                ("file_val", "Peso Total",                f"{os.path.getsize(path)/1024:.1f} KB"),
            ])

            if gen == self._render_gen:
                self.after(0, lambda m=local_meta: self._apply_extracted_metadata(m))

        except Exception as e:
            logging.error(f"Error en hilo de imagen/metadatos: {e}", exc_info=True)
            if gen == self._render_gen:
                self.after(0, lambda: self.lbl_info.configure(text="Error de lectura de archivo."))

    def _apply_image_to_ui(self, img, path, gen):
        """Pinta la imagen en pantalla inmediatamente después de ser leída."""
        if gen != self._render_gen:
            return

        self._pil_image = img
        self._redraw_zoom()
        
        kb = os.path.getsize(path) / 1024
        info_text = (
            f"📄 {Path(path).name}\n"
            f"📁 {Path(path).parent.name}\n\n"
            f"⚖  {kb:.1f} KB   |   📐 {img.size[0]}×{img.size[1]} px\n"
            f"🎨 Formato: {img.format or '—'}  |  Modo: {img.mode}"
        )
        self.lbl_info.configure(text=info_text)

    def _apply_extracted_metadata(self, metadata_list):
        """Actualiza la tabla de la UI con los datos extraídos."""
        self.all_metadata = metadata_list
        self._filter_tree()

    # ─────────────────────────────────────────────────────────────
    #  MOTOR DE EXTRACCIÓN DE METADATOS
    # ─────────────────────────────────────────────────────────────
    def _extract_all_metadata(self, path):
        """Coordina la extracción de todos los bloques de datos y los clasifica."""
        self.all_metadata = []
        try:
            with Image.open(path) as img:
                info_dict = img.info
                
                # 1. Metadatos Arqueológicos (MetaTag JSON)
                self.all_metadata.append(("header", "json", "🏺  DATOS ARQUEOLÓGICOS (JSON)"))
                self._extract_json(img, info_dict)
                
                # 2. Configuración de Cámara (EXIF Técnico)
                self.all_metadata.append(("header", "exif", "📸  CONFIGURACIÓN CÁMARA (EXIF)"))
                self._extract_exif(info_dict)
                
                # 3. Ubicación (GPS)
                self.all_metadata.append(("header", "gps", "🌍  UBICACIÓN (GPS)"))
                self._extract_gps(info_dict)
                
                # 4. Información nativa del archivo del SO
                self.all_metadata.append(("header", "file", "📄  INFO ARCHIVO SO"))
                self.all_metadata.extend([
                    ("file_val", "Nombre del Archivo", Path(path).name),
                    ("file_val", "Ubicación del Directorio", str(Path(path).parent)),
                    ("file_val", "Peso Total", f"{os.path.getsize(path)/1024:.1f} KB")
                ])
                
        except Exception as e:
            print(f"Error general en metadatos: {e}")

    def _extract_json(self, img, info):
        """Extrae el JSON incrustado en el UserComment (jpg) o en Comment (png)."""
        data = None
        ext = Path(self.current_path).suffix.lower()
        
        try:
            if ext in (".jpg", ".jpeg") and "exif" in info:
                exif_data = piexif.load(info["exif"])
                user_comment = exif_data.get("Exif", {}).get(piexif.ExifIFD.UserComment)
                if user_comment:
                    decoded_str = piexif.helper.UserComment.load(user_comment)
                    data = json.loads(decoded_str)
                    
            elif ext == ".png":
                # Extracción desde metadatos text de PNG
                comment = getattr(img, "text", {}).get("Comment", "")
                if comment:
                    data = json.loads(comment)
        except Exception:
            pass
            
        if data and isinstance(data, dict):
            for key, value in data.items():
                self.all_metadata.append(("json_val", key, str(value)))
        else:
            self.all_metadata.append(("json_val", "Aviso del Sistema", "No se detectó un archivo JSON de MetaTag válido en esta imagen."))

    def _extract_exif(self, info):
        """Extrae de forma robusta las etiquetas EXIF básicas de cámara."""
        if "exif" not in info:
            self.all_metadata.append(("exif_val", "Aviso", "La imagen no contiene datos EXIF técnicos."))
            return
            
        try:
            exif_dict = piexif.load(info["exif"])
            found_data = False
            
            # Etiquetas que no aportan a la tabla (saltarlas)
            skip_tags = {piexif.ExifIFD.UserComment, piexif.ExifIFD.MakerNote}
            skip_0th = {piexif.ImageIFD.ImageDescription, 40092, 40094}
            
            ifd_names = {"0th": "Image", "Exif": "Exif", "1st": "Image"}

            for ifd in ("0th", "Exif", "1st"):
                if ifd not in exif_dict:
                    continue
                    
                for tag_id, value in exif_dict[ifd].items():
                    if ifd == "Exif" and tag_id in skip_tags: continue
                    if ifd == "0th" and tag_id in skip_0th: continue
                    
                    try:
                        tag_name = piexif.TAGS[ifd_names[ifd]][tag_id]["name"]
                    except KeyError:
                        tag_name = f"Tag_ID_{tag_id}"
                        
                    cleaned_value = self._clean_exif_value(value, tag_name)
                    if cleaned_value:
                        self.all_metadata.append(("exif_val", tag_name, cleaned_value))
                        found_data = True
                        
            if not found_data:
                self.all_metadata.append(("exif_val", "Aviso", "No hay datos legibles de configuración de cámara."))
                
        except Exception as e:
            self.all_metadata.append(("exif_val", "Error Crítico", f"El diccionario EXIF está corrupto: {e}"))

    def _extract_gps(self, info):
        """Extrae y formatea las coordenadas GPS a Grados, Minutos y Segundos."""
        if "exif" not in info:
            self.all_metadata.append(("gps_val", "Aviso", "La imagen no contiene encabezado GPS."))
            return
            
        try:
            exif_dict = piexif.load(info["exif"])
            gps_data = exif_dict.get("GPS", {})
            
            if not gps_data:
                self.all_metadata.append(("gps_val", "Aviso", "El bloque GPS está vacío."))
                return
                
            labels = {
                piexif.GPSIFD.GPSLatitudeRef:  "Referencia Latitud",
                piexif.GPSIFD.GPSLatitude:     "Latitud Exacta",
                piexif.GPSIFD.GPSLongitudeRef: "Referencia Longitud",
                piexif.GPSIFD.GPSLongitude:    "Longitud Exacta",
                piexif.GPSIFD.GPSAltitude:     "Altitud sobre el mar (m)",
            }
            
            coord_tags = {piexif.GPSIFD.GPSLatitude, piexif.GPSIFD.GPSLongitude}
            found_gps = False
            
            for tag_id, tag_label in labels.items():
                if tag_id not in gps_data: continue
                
                raw_value = gps_data[tag_id]
                
                # Conversión matemática de raciones EXIF a Grados Minutos Segundos
                if tag_id in coord_tags and isinstance(raw_value, tuple) and len(raw_value) == 3:
                    try:
                        degrees = raw_value[0][0] / raw_value[0][1]
                        minutes = raw_value[1][0] / raw_value[1][1]
                        seconds = raw_value[2][0] / raw_value[2][1]
                        formatted_val = f"{degrees:.0f}° {minutes:.0f}' {seconds:.2f}\""
                    except Exception:
                        formatted_val = str(raw_value)
                else:
                    formatted_val = self._clean_exif_value(raw_value, "GPS")
                    
                self.all_metadata.append(("gps_val", tag_label, formatted_val))
                found_gps = True
                
            if not found_gps:
                self.all_metadata.append(("gps_val", "Aviso", "El bloque GPS no tiene coordenadas legibles."))
                
        except Exception:
            self.all_metadata.append(("gps_val", "Error GPS", "Error procesando el subdirectorio de ubicación."))

    # ── Versiones thread-safe: retornan lista, no mutan self ─────
    def _extract_json_local(self, img, info) -> list:
        rows = []
        data = None
        ext  = Path(self.current_path).suffix.lower()
        try:
            if ext in (".jpg", ".jpeg") and "exif" in info:
                ed = piexif.load(info["exif"])
                uc = ed.get("Exif", {}).get(piexif.ExifIFD.UserComment)
                if uc:
                    data = json.loads(piexif.helper.UserComment.load(uc))
            elif ext == ".png":
                c = getattr(img, "text", {}).get("Comment", "")
                if c: data = json.loads(c)
        except Exception:
            pass
        if data and isinstance(data, dict):
            rows.extend(("json_val", k, str(v)) for k, v in data.items())
        else:
            rows.append(("json_val", "Aviso del Sistema",
                         "No se detectó JSON de MetaTag en esta imagen."))
        return rows

    def _extract_exif_local(self, info) -> list:
        rows = []
        if "exif" not in info:
            rows.append(("exif_val", "Aviso", "Sin datos EXIF técnicos."))
            return rows
        try:
            exif_dict  = piexif.load(info["exif"])
            found      = False
            skip_exif  = {piexif.ExifIFD.UserComment, piexif.ExifIFD.MakerNote}
            skip_0th   = {piexif.ImageIFD.ImageDescription, 40092, 40094}
            ifd_names  = {"0th": "Image", "Exif": "Exif", "1st": "Image"}
            for ifd in ("0th", "Exif", "1st"):
                if ifd not in exif_dict: continue
                for tid, val in exif_dict[ifd].items():
                    if ifd == "Exif" and tid in skip_exif: continue
                    if ifd == "0th"  and tid in skip_0th:  continue
                    try:    name = piexif.TAGS[ifd_names[ifd]][tid]["name"]
                    except: name = f"Tag_ID_{tid}"
                    cleaned = self._clean_exif_value(val, name)
                    if cleaned:
                        rows.append(("exif_val", name, cleaned))
                        found = True
            if not found:
                rows.append(("exif_val", "Aviso", "No hay datos legibles de cámara."))
        except Exception as e:
            rows.append(("exif_val", "Error", f"EXIF corrupto: {e}"))
        return rows

    def _extract_gps_local(self, info) -> list:
        rows = []
        if "exif" not in info:
            rows.append(("gps_val", "Aviso", "Sin encabezado GPS."))
            return rows
        try:
            gps = piexif.load(info["exif"]).get("GPS", {})
            if not gps:
                rows.append(("gps_val", "Aviso", "Bloque GPS vacío."))
                return rows
            labels = {
                piexif.GPSIFD.GPSLatitudeRef:  "Ref. Latitud",
                piexif.GPSIFD.GPSLatitude:     "Latitud Exacta",
                piexif.GPSIFD.GPSLongitudeRef: "Ref. Longitud",
                piexif.GPSIFD.GPSLongitude:    "Longitud Exacta",
                piexif.GPSIFD.GPSAltitude:     "Altitud (m)",
            }
            coord = {piexif.GPSIFD.GPSLatitude, piexif.GPSIFD.GPSLongitude}
            found = False
            for tid, label in labels.items():
                if tid not in gps: continue
                v = gps[tid]
                if tid in coord and isinstance(v, tuple) and len(v) == 3:
                    try:
                        d,m,s = v[0][0]/v[0][1], v[1][0]/v[1][1], v[2][0]/v[2][1]
                        fv = f"{d:.0f}° {m:.0f}' {s:.2f}\""
                    except: fv = str(v)
                else:
                    fv = self._clean_exif_value(v, "GPS")
                rows.append(("gps_val", label, fv))
                found = True
            if not found:
                rows.append(("gps_val", "Aviso", "Sin coordenadas legibles."))
        except Exception:
            rows.append(("gps_val", "Error", "Error procesando GPS."))
        return rows

    def _clean_exif_value(self, value, tag_name: str) -> str:
        """Limpia la basura binaria y decodifica las raciones matemáticas de EXIF."""
        if tag_name in {"MakerNote", "ComponentsConfiguration", "DeviceSettingDescription", "DNGVersion"}:
            return "<Datos binarios opacos del fabricante>"
            
        if isinstance(value, bytes):
            if len(value) > 30:
                return f"<Blob binario de {len(value)} bytes>"
            cleaned_bytes = value.replace(b"\x00", b"").strip()
            try:
                return cleaned_bytes.decode("utf-8", errors="ignore")
            except Exception:
                return "<Cadena de bytes no decodificable>"
                
        # Manejo de raciones (Numerador, Denominador) comunes en Velocidad de Obturador o Apertura
        if isinstance(value, tuple) and len(value) == 2 and isinstance(value[0], int) and isinstance(value[1], int):
            num, den = value[0], value[1]
            if den == 0:   return "0"
            if den == 1:   return str(num)
            if num == 1:   return f"1/{den}"
            try:           return f"{num/den:.2f}"
            except Exception: return f"{num}/{den}"
            
        if isinstance(value, tuple):
            return ", ".join(str(v) for v in value)
            
        return str(value)

    # ─────────────────────────────────────────────────────────────
    #  ACTUALIZACIÓN DE TABLA Y FILTROS
    # ─────────────────────────────────────────────────────────────
    def _filter_tree(self, *_args):
        """Vuelca la lista `all_metadata` en el TreeView, aplicando la búsqueda activa."""
        self.tree.delete(*self.tree.get_children())
        query = self._search_var.get().lower()
        is_odd_row = True
        
        # Mapeo de colores y estilos por sección
        color_map = {
            "json_val": C["json_section"], 
            "exif_val": C["exif_section"], 
            "gps_val": C["gps_section"], 
            "file_val": C["file_section"]
        }
        header_map = {
            "json": "hdr_json", 
            "exif": "hdr_exif", 
            "gps": "hdr_gps", 
            "file": "hdr_file"
        }

        for type_tag, key, val in self.all_metadata:
            if type_tag == "header":
                # Si no hay texto de búsqueda, mostramos la cabecera
                if not query: 
                    self.tree.insert("", "end", values=(val, ""), tags=(header_map.get(key),))
            else:
                # Si hay texto de búsqueda, verificamos coincidencia
                if query in key.lower() or query in val.lower():
                    bg_tag = "odd" if is_odd_row else "even"
                    
                    item_id = self.tree.insert("", "end", values=(key, val), tags=(bg_tag,))
                    
                    # Generamos un tag de color único para esta fila
                    fg_color = color_map.get(type_tag, C["text"])
                    self.tree.tag_configure(f"fg_{item_id}", foreground=fg_color)
                    self.tree.item(item_id, tags=(bg_tag, f"fg_{item_id}"))
                    
                    is_odd_row = not is_odd_row

    # ─────────────────────────────────────────────────────────────
    #  SISTEMA DE ZOOM Y PANEO (RECORTE VIEWPORT OPTIMIZADO)
    # ─────────────────────────────────────────────────────────────
    def _bind_zoom_events(self):
        """Enlaza el ratón al sistema de visualización."""
        canvas = self.img_canvas
        
        # Rueda de ratón (OS independiente)
        canvas.bind("<MouseWheel>", self._on_wheel_windows)
        canvas.bind("<Button-4>", self._on_wheel_up_linux)
        canvas.bind("<Button-5>", self._on_wheel_down_linux)
        
        # Arrastre (Paneo)
        canvas.bind("<ButtonPress-1>", self._pan_start_event)
        canvas.bind("<B1-Motion>", self._pan_drag_event)
        canvas.bind("<ButtonRelease-1>", self._pan_end_event)
        
        # Doble clic para ajustar
        canvas.bind("<Double-Button-1>", lambda e: self._zoom_fit())
        
        # Cambio de cursor al pasar sobre la imagen
        canvas.bind("<Enter>", lambda e: canvas.configure(cursor="fleur") if self._pil_image else None)
        canvas.bind("<Leave>", lambda e: canvas.configure(cursor=""))

        # EVENTO NUEVO: Redibujado adaptativo al redimensionar paneles
        canvas.bind("<Configure>", self._on_canvas_resize)

    def _on_canvas_resize(self, event):
        """Dispara un redibujado de la imagen cuando el panel cambia de tamaño, evitando deformaciones."""
        if self._pil_image:
            if getattr(self, "_resize_job", None):
                self.after_cancel(self._resize_job)
            self._resize_job = self.after(150, lambda: self._redraw_zoom(final_pass=True))

    def _on_wheel_windows(self, event):
        factor = ZOOM_STEP if event.delta > 0 else 1 / ZOOM_STEP
        self._zoom_at(event.x, event.y, factor)

    def _on_wheel_up_linux(self, event):
        self._zoom_at(event.x, event.y, ZOOM_STEP)

    def _on_wheel_down_linux(self, event):
        self._zoom_at(event.x, event.y, 1 / ZOOM_STEP)

    def _pan_start_event(self, event):
        if self._pil_image:
            self._panning = True
            self._pan_start = [event.x, event.y]
            self.img_canvas.configure(cursor="fleur")

    def _pan_drag_event(self, event):
        if not self._panning: return
        
        # Calcular delta
        delta_x = event.x - self._pan_start[0]
        delta_y = event.y - self._pan_start[1]
        
        # Sumar al offset general
        self._pan_offset[0] += delta_x
        self._pan_offset[1] += delta_y
        
        # Actualizar inicio
        self._pan_start = [event.x, event.y]
        
        # Redibujar la vista
        self._redraw_zoom()

    def _pan_end_event(self, event):
        self._panning = False

    def _zoom_at(self, cursor_x: int, cursor_y: int, factor: float):
        """Aplica el zoom asegurándose que el píxel bajo el ratón se quede quieto."""
        if not self._pil_image: return
        
        old_zoom = self._current_zoom_value()
        new_zoom = max(ZOOM_MIN, min(ZOOM_MAX, old_zoom * factor))
        
        if abs(new_zoom - old_zoom) < 1e-4: return

        ratio = new_zoom / old_zoom
        canvas_w = self.img_canvas.winfo_width()
        canvas_h = self.img_canvas.winfo_height()
        
        # Fórmula de traslación afín para mantener el centro
        self._pan_offset[0] = ratio * self._pan_offset[0] + (1 - ratio) * (cursor_x - canvas_w / 2)
        self._pan_offset[1] = ratio * self._pan_offset[1] + (1 - ratio) * (cursor_y - canvas_h / 2)
        
        self._zoom_level = new_zoom
        self._redraw_zoom()
        self._show_zoom_badge(new_zoom)

    def _zoom_by(self, factor: float):
        """Botón de lupa (zoom centrado en el medio de la pantalla)."""
        center_x = self.img_canvas.winfo_width() // 2
        center_y = self.img_canvas.winfo_height() // 2
        self._zoom_at(center_x, center_y, factor)

    def _zoom_fit(self):
        """Restablece la imagen para que quepa exacta en el marco actual."""
        self._zoom_level = ZOOM_FIT
        self._pan_offset = [0, 0]
        self._redraw_zoom()
        self._show_zoom_badge(self._current_zoom_value())

    def _current_zoom_value(self) -> float:
        """Devuelve el valor numérico absoluto del zoom."""
        if self._zoom_level == ZOOM_FIT and self._pil_image:
            cw = max(self.img_canvas.winfo_width(), 60)
            ch = max(self.img_canvas.winfo_height(), 60)
            return min((cw - 4) / self._pil_image.width, (ch - 4) / self._pil_image.height)
        return self._zoom_level if self._zoom_level != ZOOM_FIT else 1.0

    def _redraw_zoom(self, final_pass: bool = False):
        """
        El núcleo gráfico del Visor.
        Calcula qué parte de la imagen choca con la ventana (Viewport) y SOLO 
        recorta, reescala y dibuja ese pedazo. Previene congelamientos.
        """
        if not self._pil_image: return

        # Cancelamos hilos obsoletos incrementando el generador
        self._render_gen += 1
        current_gen = self._render_gen
        
        z = self._current_zoom_value()
        
        canvas_w = self.img_canvas.winfo_width()
        canvas_h = self.img_canvas.winfo_height()
        
        # Precaución si la UI aún no ha dibujado
        if canvas_w < 10 or canvas_h < 10: 
            canvas_w, canvas_h = 800, 600

        # Punto central desplazado por el paneo
        center_x = canvas_w // 2 + int(self._pan_offset[0])
        center_y = canvas_h // 2 + int(self._pan_offset[1])
        
        img_w, img_h = self._pil_image.size
        
        # Dimensiones que tendría la imagen completa
        scaled_full_w = img_w * z
        scaled_full_h = img_h * z

        # 1. Bounding box de la imagen completa en el canvas virtual
        img_cx0 = center_x - scaled_full_w / 2
        img_cy0 = center_y - scaled_full_h / 2
        
        # 2. Intersección entre la Imagen Virtual y el Canvas Físico (Viewport)
        viewport_x0 = max(0, img_cx0)
        viewport_y0 = max(0, img_cy0)
        viewport_x1 = min(canvas_w, img_cx0 + scaled_full_w)
        viewport_y1 = min(canvas_h, img_cy0 + scaled_full_h)

        # Si la imagen está fuera de pantalla, borrar
        if viewport_x1 <= viewport_x0 or viewport_y1 <= viewport_y0:
            self.img_canvas.delete("img")
            return

        # 3. Mapear ese pedazo visible (pantalla) hacia los píxeles de la imagen original
        orig_x0 = int((viewport_x0 - img_cx0) / z)
        orig_y0 = int((viewport_y0 - img_cy0) / z)
        orig_x1 = int((viewport_x1 - img_cx0) / z)
        orig_y1 = int((viewport_y1 - img_cy0) / z)

        # Limitar bordes
        orig_x0, orig_y0 = max(0, min(img_w, orig_x0)), max(0, min(img_h, orig_y0))
        orig_x1, orig_y1 = max(0, min(img_w, orig_x1)), max(0, min(img_h, orig_y1))

        if orig_x1 <= orig_x0 or orig_y1 <= orig_y0: return
        
        # Tamaño final del recorte escalado
        crop_render_w = int((orig_x1 - orig_x0) * z)
        crop_render_h = int((orig_y1 - orig_y0) * z)
        
        if crop_render_w < 1 or crop_render_h < 1: return

        # Punto donde anclaremos el recorte en el canvas
        place_x = viewport_x0 + (viewport_x1 - viewport_x0) / 2
        place_y = viewport_y0 + (viewport_y1 - viewport_y0) / 2

        if final_pass:
            # ── PASO FINAL: Render de Alta Calidad (LANCZOS) en hilo secundario ──
            def _do_lanczos_thread():
                if current_gen != self._render_gen: return # Si el usuario movió, abortar
                try:
                    crop = self._pil_image.crop((orig_x0, orig_y0, orig_x1, orig_y1))
                    scaled = crop.resize((crop_render_w, crop_render_h), Image.LANCZOS)
                    
                    if current_gen != self._render_gen: return # Comprobar de nuevo tras proceso pesado
                    
                    tk_img = ImageTk.PhotoImage(scaled)
                    
                    # Llamar al hilo principal de UI de forma segura
                    self.after(0, lambda: self._apply_zoom_image_safely(tk_img, place_x, place_y, current_gen))
                except Exception:
                    pass
                    
            threading.Thread(target=_do_lanczos_thread, daemon=True).start()
            
        else:
            # ── PASO BORRADOR: Rápido (NEAREST) mientras arrastramos ──
            try:
                crop = self._pil_image.crop((orig_x0, orig_y0, orig_x1, orig_y1))
                draft = crop.resize((crop_render_w, crop_render_h), Image.NEAREST)
                self._zoom_img_tk_draft = ImageTk.PhotoImage(draft)
                
                self.img_canvas.delete("img")
                self.img_canvas.create_image(
                    place_x, place_y, anchor="center", 
                    image=self._zoom_img_tk_draft, tags="img"
                )
            except Exception:
                pass

            # Programamos el paso final
            if self._zoom_debounce_job:
                self.after_cancel(self._zoom_debounce_job)
            self._zoom_debounce_job = self.after(100, lambda: self._redraw_zoom(final_pass=True))

    def _apply_zoom_image_safely(self, tk_img, final_x: float, final_y: float, request_gen: int):
        """Asigna la imagen procesada por el hilo a la variable de Tkinter."""
        if request_gen != self._render_gen:
            return
        self._zoom_img_tk = tk_img
        self.img_canvas.delete("img")
        self.img_canvas.create_image(final_x, final_y, anchor="center", image=self._zoom_img_tk, tags="img")

    def _show_zoom_badge(self, zoom_val: float):
        """Dibuja un porcentaje temporal en la esquina de la imagen."""
        self.img_canvas.delete("zoom_badge")
        
        canvas_w = self.img_canvas.winfo_width()
        self.img_canvas.create_text(
            canvas_w - 12, 12, 
            text=f"{zoom_val*100:.0f}%", anchor="ne", 
            fill=C["zoom_badge"], font=F_BOLD, tags="zoom_badge"
        )
        
        if getattr(self, "_zoom_hide_job", None):
            self.after_cancel(self._zoom_hide_job)
            
        self._zoom_hide_job = self.after(1500, lambda: self.img_canvas.delete("zoom_badge"))


    # ─────────────────────────────────────────────────────────────
    #  HERRAMIENTA DE COMPARACIÓN DE IMÁGENES
    # ─────────────────────────────────────────────────────────────
    def _open_comparator(self):
        """Diálogo para seleccionar dos carpetas, luego abre vista dedicada."""
        if self._comp_window and self._comp_window.winfo_exists():
            self._comp_window.lift()
            return

        win = tk.Toplevel(self)
        self._comp_window = win
        win.title("⚖ Seleccionar carpetas para comparar")
        win.configure(bg=C["bg"])
        win.resizable(False, False)
        win.attributes("-topmost", True)

        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        ww = 600
        wh = 300
        win.geometry(f"{ww}x{wh}+{(sw - ww) // 2}+{(sh - wh) // 2}")

        hdr = tk.Frame(win, bg=C["header_bg"], height=50)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="⚖  Selecciona las dos carpetas a comparar", font=F_TITLE,
                 bg=C["header_bg"], fg=C["header_fg"]).pack(side="left", padx=20, pady=10)

        body = tk.Frame(win, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=20, pady=10)

        folder_a_var = tk.StringVar(value=getattr(self, "_comp_folder_a", ""))
        folder_b_var = tk.StringVar(value=getattr(self, "_comp_folder_b", ""))
        count_a_var = tk.StringVar()
        count_b_var = tk.StringVar()

        def count_images(folder):
            if not folder or not os.path.isdir(folder):
                return 0
            return sum(1 for f in Path(folder).iterdir()
                       if f.is_file() and f.suffix.lower() in IMG_EXTS)

        def update_count(var, count_var):
            folder = var.get()
            n = count_images(folder)
            count_var.set(f"{n} imágenes" if n else "Sin imágenes" if folder else "")

        def pick_folder(var, count_var):
            p = _native_folder_open(title="Seleccionar carpeta")
            if p:
                var.set(p)
                update_count(var, count_var)

        for label, var, count_var, col in [
                ("Carpeta A (Izq):", folder_a_var, count_a_var, C["json_section"]),
                ("Carpeta B (Der):", folder_b_var, count_b_var, C["exif_section"])]:
            row = tk.Frame(body, bg=C["bg"])
            row.pack(fill="x", pady=4)
            tk.Label(row, text=label, bg=C["bg"], fg=col, font=F_BOLD, width=16, anchor="w").pack(side="left")
            tk.Entry(row, textvariable=var, bg=C["surface"], fg=C["text"],
                     relief="solid", bd=1, font=F_TINY, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 4))
            tk.Button(row, text="Examinar", bg=C["panel2"], fg=C["text2"], font=F_TINY,
                      relief="flat", bd=0, padx=8, cursor="hand2",
                      command=lambda v=var, c=count_var: pick_folder(v, c)).pack(side="right")
            lbl = tk.Label(row, textvariable=count_var, bg=C["bg"], fg=C["text3"], font=F_TINY)
            lbl.pack(side="right", padx=(0, 6))

        # Cargar conteos iniciales
        update_count(folder_a_var, count_a_var)
        update_count(folder_b_var, count_b_var)

        def start_compare():
            a, b = folder_a_var.get(), folder_b_var.get()
            if not a or not os.path.isdir(a):
                return messagebox.showwarning("Falta", "Selecciona la Carpeta A.")
            if not b or not os.path.isdir(b):
                return messagebox.showwarning("Falta", "Selecciona la Carpeta B.")
            na, nb = count_images(a), count_images(b)
            if na == 0:
                return messagebox.showwarning("Vacía", f"No hay imágenes en:\n{a}")
            if nb == 0:
                return messagebox.showwarning("Vacía", f"No hay imágenes en:\n{b}")
            self._comp_folder_a = a
            self._comp_folder_b = b
            self._save_config()
            win.destroy()
            self._comp_window = None
            self._open_image_comparison(a, b)

        tk.Button(body, text="⚖  Abrir Comparador de Imágenes", bg=C["accent"], fg="#FFF5E8",
                  font=F_BOLD, relief="flat", bd=0, padx=20, pady=10, cursor="hand2",
                  activebackground=C["accent_hover"],
                  command=start_compare).pack(pady=(15, 0))

    def _open_image_comparison(self, folder_a: str, folder_b: str):
        """Visor doble de comparación: imagen + metadatos en cada lado."""
        imgs_a = sorted([str(f) for f in Path(folder_a).iterdir()
                         if f.is_file() and f.suffix.lower() in IMG_EXTS],
                        key=lambda p: p.lower())
        imgs_b = sorted([str(f) for f in Path(folder_b).iterdir()
                         if f.is_file() and f.suffix.lower() in IMG_EXTS],
                        key=lambda p: p.lower())

        if not imgs_a:
            return messagebox.showwarning("Vacía", f"No hay imágenes en:\n{folder_a}")
        if not imgs_b:
            return messagebox.showwarning("Vacía", f"No hay imágenes en:\n{folder_b}")

        self.withdraw()

        win = tk.Toplevel(self)
        self._comp_window = win
        self._comp_thumbs = []
        win.title(f"⚖ {Path(folder_a).name}  vs  {Path(folder_b).name}")
        win.configure(bg=C["bg"])
        win.protocol("WM_DELETE_WINDOW", lambda: self._close_comparison(win))

        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f"{sw}x{sh}+0+0")
        try:
            win.attributes("-fullscreen", True)
        except Exception:
            pass

        # ── Estado ──
        state = {"idx_a": 0, "idx_b": 0, "thumbs": []}

        # ── Header ──
        hdr = tk.Frame(win, bg=C["header_bg"], height=50)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text=f"⚖  {Path(folder_a).name}  vs  {Path(folder_b).name}",
                 font=F_TITLE, bg=C["header_bg"], fg=C["header_fg"]).pack(side="left", padx=20, pady=10)

        nav_frame = tk.Frame(hdr, bg=C["header_bg"])
        nav_frame.pack(side="left", padx=20)

        lbl_counter = tk.Label(nav_frame, text="", font=F_BOLD, bg=C["header_bg"], fg=C["text"])
        lbl_counter.pack(side="left", padx=10)

        def update_counter():
            total = max(len(imgs_a), len(imgs_b))
            lbl_counter.configure(text=f"{state['idx_a']+1} / {total}")

        tk.Button(hdr, text="  ⬇ PDF  ", bg=C["accent"], fg="#FFF5E8",
                  font=F_BOLD, relief="flat", bd=0, padx=16, pady=6, cursor="hand2",
                  activebackground=C["accent_hover"],
                  command=lambda: self._export_comparison_pdf(
                      imgs_a, imgs_b, Path(folder_a).name, Path(folder_b).name)
                  ).pack(side="right", padx=(0, 8), pady=8)

        tk.Button(hdr, text="  Cerrar  ", bg=C["err"], fg="#FFFFFF",
                  font=F_BOLD, relief="flat", bd=0, padx=16, pady=6, cursor="hand2",
                  command=lambda: self._close_comparison(win)).pack(side="right", padx=20, pady=8)

        # ── Body: dos paneles lado a lado ──
        body = tk.Frame(win, bg=C["border"])
        body.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        panels = []
        for side, folder_name, imgs, idx_key, accent_col in [
            ("left", folder_a, imgs_a, "idx_a", C["json_section"]),
            ("right", folder_b, imgs_b, "idx_b", C["exif_section"])]:

            panel = tk.Frame(body, bg=C["bg"])
            panel.pack(side=side, fill="both", expand=True, padx=4)

            # Título de carpeta
            tk.Label(panel, text=f"  {Path(folder_name).name}  ({len(imgs)} imgs)",
                     bg=C["bg"], fg=accent_col, font=F_BOLD, anchor="w").pack(fill="x", padx=8, pady=(6, 4))

            # ── Zona FIJA: imagen + info de archivo (no scrollea) ──
            img_zone = tk.Frame(panel, bg=C["surface"],
                                highlightthickness=1,
                                highlightbackground=C["border"])
            img_zone.pack(fill="x", padx=4, pady=(0, 4))

            img_lbl = tk.Label(img_zone, bg=C["surface"])
            img_lbl.pack(pady=(8, 4), padx=8)

            name_lbl = tk.Label(img_zone, text="", bg=C["surface"], fg=C["text2"],
                                font=F_TINY, anchor="w", wraplength=int(sw * 0.42))
            name_lbl.pack(fill="x", padx=12)

            info_lbl = tk.Label(img_zone, text="", bg=C["surface"], fg=C["text3"],
                                font=F_MICRO, anchor="w", wraplength=int(sw * 0.42))
            info_lbl.pack(fill="x", padx=12, pady=(0, 6))

            # ── Zona SCROLLABLE: solo la tabla de metadatos ──
            meta_outer = tk.Frame(panel, bg=C["border"])
            meta_outer.pack(fill="both", expand=True, padx=4, pady=(0, 4))

            tree_vsb = ttk.Scrollbar(meta_outer, orient="vertical")
            tree = ttk.Treeview(meta_outer, columns=("Campo", "Valor"),
                                show="headings", selectmode="browse",
                                yscrollcommand=tree_vsb.set)
            tree.heading("Campo", text="Campo", anchor="w")
            tree.heading("Valor", text="Información", anchor="w")
            col_w = max(120, int(sw * 0.12))
            tree.column("Campo", width=col_w, anchor="w", minwidth=80)
            tree.column("Valor", width=col_w * 2, anchor="w", minwidth=100)
            tree_vsb.configure(command=tree.yview)
            tree_vsb.pack(side="right", fill="y")
            tree.pack(fill="both", expand=True)

            tree.tag_configure("hdr", foreground=accent_col, font=F_BOLD, background=C["panel2"])
            tree.tag_configure("odd",  background=C["surface"])
            tree.tag_configure("even", background=C["row_alt"])

            # canvas y inner son dummies para mantener compatibilidad con show_current
            canvas = meta_outer
            inner  = meta_outer

            panels.append({
                "canvas": canvas, "inner": inner, "img_lbl": img_lbl,
                "name_lbl": name_lbl, "info_lbl": info_lbl, "tree": tree,
                "folder_name": folder_name, "imgs": imgs, "idx_key": idx_key,
                "accent": accent_col, "img_zone": img_zone,
            })

        # ── Barra de navegación inferior ──
        nav_bar = tk.Frame(win, bg=C["panel"], height=50)
        nav_bar.pack(fill="x")
        nav_bar.pack_propagate(False)

        def go_prev():
            if state["idx_a"] > 0:
                state["idx_a"] -= 1
                state["idx_b"] = min(state["idx_b"], len(imgs_b) - 1)
                if state["idx_b"] > 0: state["idx_b"] -= 1
                show_current()

        def go_next():
            max_idx = max(len(imgs_a), len(imgs_b)) - 1
            if state["idx_a"] < max_idx:
                state["idx_a"] += 1
                state["idx_b"] = min(state["idx_b"] + 1, len(imgs_b) - 1)
                show_current()

        tk.Button(nav_bar, text="  ←  Anterior  ", bg=C["panel2"], fg=C["text"],
                  font=F_BOLD, relief="flat", bd=0, padx=16, pady=8, cursor="hand2",
                  command=go_prev).pack(side="left", padx=20, pady=6)

        tk.Button(nav_bar, text="  Siguiente  →  ", bg=C["panel2"], fg=C["text"],
                  font=F_BOLD, relief="flat", bd=0, padx=16, pady=8, cursor="hand2",
                  command=go_next).pack(side="left", padx=4, pady=6)

        # ── Función para mostrar la imagen actual ──
        def show_current():
            update_counter()
            for p in panels:
                imgs = p["imgs"]
                idx_key = p["idx_key"]
                idx = state[idx_key]
                if idx >= len(imgs):
                    p["img_lbl"].configure(image="", text="[Sin imagen]")
                    p["name_lbl"].configure(text="")
                    p["info_lbl"].configure(text="")
                    p["tree"].delete(*p["tree"].get_children())
                    continue

                path = imgs[idx]
                p["name_lbl"].configure(text=Path(path).name)

                try:
                    size_kb = os.path.getsize(path) / 1024
                    img = Image.open(path)
                    img = ImageOps.exif_transpose(img)
                    w, h = img.size
                    fmt = img.format or "—"
                    p["info_lbl"].configure(
                        text=f"{size_kb:.1f} KB  |  {w}×{h} px  |  {fmt}")

                    # Thumbnail para la tabla
                    thumb = img.copy()
                    thumb_w = max(320, int(sw * 0.42) - 32)
                    thumb_h = max(200, int(sh * 0.38) - 16)
                    thumb.thumbnail((thumb_w, thumb_h), Image.LANCZOS)
                    tk_img = ImageTk.PhotoImage(thumb)
                    p["img_lbl"].configure(image=tk_img, text="")
                    p["img_lbl"]._img_ref = tk_img
                except Exception as e:
                    p["img_lbl"].configure(image="", text=f"Error: {e}")
                    p["info_lbl"].configure(text="")

                # Llenar tabla de metadatos
                p["tree"].delete(*p["tree"].get_children())
                try:
                    old_path = self.current_path
                    old_meta = self.all_metadata[:]
                    self.current_path = path
                    self._extract_all_metadata(path)
                    meta = self.all_metadata[:]
                    self.current_path = old_path
                    self.all_metadata = old_meta

                    is_odd = True
                    for tag_type, key, val in meta:
                        if tag_type == "header":
                            p["tree"].insert("", "end", values=(val, ""), tags=("hdr",))
                        else:
                            bg_tag = "odd" if is_odd else "even"
                            p["tree"].insert("", "end", values=(key, val), tags=(bg_tag,))
                            is_odd = not is_odd
                except Exception:
                    pass

                # Al cambiar imagen, la tabla de metadatos vuelve al principio
                if p["tree"].get_children():
                    p["tree"].see(p["tree"].get_children()[0])

        show_current()

        win.bind("<Left>", lambda e: go_prev())
        win.bind("<Right>", lambda e: go_next())
        win.bind("<Escape>", lambda e: self._close_comparison(win))

    def _show_full_image(self, path: str):
        """Abre imagen a tamaño completo en una ventana emergente."""
        win = tk.Toplevel(self)
        win.title(Path(path).name)
        win.configure(bg="#000000")
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f"{sw}x{sh}+0+0")
        try:
            win.attributes("-fullscreen", True)
        except Exception:
            pass

        canvas = tk.Canvas(win, bg="#000000", highlightthickness=0)
        canvas.pack(fill="both", expand=True)

        try:
            img = Image.open(path)
            img = ImageOps.exif_transpose(img)
            img.thumbnail((sw - 20, sh - 60), Image.LANCZOS)
            tk_img = ImageTk.PhotoImage(img)
            canvas.create_image(sw // 2, sh // 2, anchor="center", image=tk_img)
            canvas._img_ref = tk_img
        except Exception as e:
            canvas.create_text(sw // 2, sh // 2, text=f"Error: {e}",
                               fill="#F44747", font=F_BODY, justify="center")

        tk.Label(win, text=f"  {Path(path).name}  —  Doble clic o Escape para cerrar  ",
                 bg="#111111", fg="#AAAAAA", font=F_TINY).pack(fill="x")

        win.bind("<Double-Button-1>", lambda e: win.destroy())
        win.bind("<Escape>", lambda e: win.destroy())

    def _close_comparison(self, comp_win):
        """Cierra la ventana de comparación y restaura el Visor."""
        comp_win.destroy()
        self._comp_window = None
        self._comp_thumbs = []
        self.deiconify()
        self.lift()
        self.focus_force()

    def _export_comparison_pdf(self, imgs_a: list, imgs_b: list, name_a: str, name_b: str):
        """Exporta un PDF con análisis comparativo de metadatos."""
        if not REPORTLAB_OK:
            return messagebox.showerror("Falta", "Instala reportlab: pip install reportlab")

        output_path = _native_save_file(
            title="Guardar PDF de Comparación",
            default_name=f"Comparacion_{name_a}_vs_{name_b}.pdf")
        if not output_path:
            return

        total = min(len(imgs_a), len(imgs_b))
        if total == 0:
            return messagebox.showwarning("Sin datos", "No hay pares de imágenes para comparar.")

        # ── Ventana de progreso ──
        prog = tk.Toplevel(self)
        prog.title("Exportando PDF...")
        prog.configure(bg=C["bg"])
        prog.resizable(False, False)
        prog.attributes("-topmost", True)
        pw, ph = 420, 160
        prog.geometry(f"{pw}x{ph}+{(sw - pw)//2 if (sw:=self.winfo_screenwidth()) else 400}+{(sh - ph)//2 if (sh:=self.winfo_screenheight()) else 300}")
        prog.grab_set()

        tk.Label(prog, text="Generando PDF de comparación...", bg=C["bg"],
                 fg=C["accent"], font=F_BOLD).pack(pady=(20, 8))
        lbl_status = tk.Label(prog, text="Preparando...", bg=C["bg"], fg=C["text2"], font=F_BODY)
        lbl_status.pack()

        bar_outer = tk.Frame(prog, bg=C["border"], height=12)
        bar_outer.pack(fill="x", padx=30, pady=12)
        bar_fill = tk.Frame(bar_outer, bg=C["accent"], height=12)
        bar_fill.place(x=0, y=0, relheight=1, width=0)

        cancelled = threading.Event()

        def on_cancel():
            cancelled.set()
            prog.destroy()

        tk.Button(prog, text="Cancelar", bg=C["panel2"], fg=C["text"], font=F_BODY,
                  relief="flat", bd=0, cursor="hand2", command=on_cancel).pack()

        def update_prog(frac, text):
            try:
                bar_outer.update_idletasks()
                w = bar_outer.winfo_width() or 360
                bar_fill.place(width=int(w * frac))
                lbl_status.configure(text=text)
                prog.update_idletasks()
            except Exception:
                pass

        def worker():
            try:
                def extract_meta(path):
                    try:
                        with Image.open(path) as img:
                            info = img.info
                            ext = Path(path).suffix.lower()
                            data = None
                            if ext in (".jpg", ".jpeg") and "exif" in info:
                                exif_d = piexif.load(info["exif"])
                                uc = exif_d.get("Exif", {}).get(piexif.ExifIFD.UserComment)
                                if uc:
                                    data = json.loads(piexif.helper.UserComment.load(uc))
                            elif ext == ".png":
                                c = getattr(img, "text", {}).get("Comment", "")
                                if c:
                                    data = json.loads(c)
                            if data and isinstance(data, dict):
                                return {k: str(v) for k, v in data.items()}
                    except Exception:
                        pass
                    return {}

                doc = SimpleDocTemplate(
                    output_path, pagesize=A4,
                    topMargin=1.5*cm, bottomMargin=1.5*cm,
                    leftMargin=1.5*cm, rightMargin=1.5*cm)

                styles = getSampleStyleSheet()
                CLR_GOLD = colors.HexColor("#D4A574")
                CLR_BROWN = colors.HexColor("#5C3518")
                CLR_PAPER = colors.HexColor("#FAF7F2")
                CLR_STRIPE = colors.HexColor("#F3EDE5")
                CLR_LINE = colors.HexColor("#D4C8B4")

                s_title = ParagraphStyle("T", parent=styles["Title"], fontName="Helvetica-Bold",
                                         fontSize=16, textColor=CLR_BROWN, alignment=TA_CENTER, spaceAfter=4)
                s_sub = ParagraphStyle("S", parent=styles["Normal"], fontName="Helvetica-Oblique",
                                       fontSize=9, textColor=CLR_GOLD, alignment=TA_CENTER, spaceAfter=10)
                s_section = ParagraphStyle("Sec", parent=styles["Normal"], fontName="Helvetica-Bold",
                                           fontSize=10, textColor=CLR_BROWN, spaceBefore=10, spaceAfter=4)
                s_key = ParagraphStyle("K", parent=styles["Normal"], fontName="Helvetica-Bold",
                                       fontSize=7.5, textColor=colors.HexColor("#2A1505"))
                s_val = ParagraphStyle("V", parent=styles["Normal"], fontName="Helvetica",
                                       fontSize=7.5, textColor=CLR_BROWN)
                s_name = ParagraphStyle("N", parent=styles["Normal"], fontName="Helvetica-Bold",
                                        fontSize=8, textColor=CLR_BROWN, alignment=TA_CENTER)

                flowables = []

                # Portada
                flowables.append(Paragraph("ANÁLISIS COMPARATIVO DE METADATOS", s_title))
                flowables.append(Paragraph(
                    f"{name_a}  vs  {name_b}  —  {datetime.datetime.now():%d/%m/%Y %H:%M}", s_sub))
                flowables.append(HRFlowable(width="100%", thickness=2, color=CLR_GOLD, spaceAfter=10))

                # Resumen general
                total_same = total_diff = total_new_a = total_new_b = 0

                all_pair_data = []
                for i in range(total):
                    if cancelled.is_set():
                        self.after(0, prog.destroy)
                        return
                    self.after(0, update_prog, (i + 0.5) / total, f"Analizando imagen {i+1}/{total}...")

                    meta_a = extract_meta(imgs_a[i])
                    meta_b = extract_meta(imgs_b[i])
                    all_pair_data.append((imgs_a[i], imgs_b[i], meta_a, meta_b))

                    keys_a, keys_b = set(meta_a), set(meta_b)
                    total_same += len(keys_a & keys_b)
                    total_diff += sum(1 for k in keys_a & keys_b if meta_a[k] != meta_b[k])
                    total_new_a += len(keys_a - keys_b)
                    total_new_b += len(keys_b - keys_a)

                self.after(0, update_prog, 0.8, "Generando documento PDF...")

                # Resumen ejecutivo
                flowables.append(Paragraph("RESUMEN EJECUTIVO", s_section))
                summary_data = [
                    [Paragraph("<b>Métrica</b>", s_key), Paragraph("<b>Cantidad</b>", s_key)],
                    [Paragraph("Pares comparados", s_key), Paragraph(str(total), s_val)],
                    [Paragraph("Campos idénticos", s_key), Paragraph(f'<font color="#7EC894">{total_same}</font>', s_val)],
                    [Paragraph("Campos diferentes", s_key), Paragraph(f'<font color="#E05050">{total_diff}</font>', s_val)],
                    [Paragraph(f"Sólo en {name_a}", s_key), Paragraph(f'<font color="#7EB8C9">{total_new_a}</font>', s_val)],
                    [Paragraph(f"Sólo en {name_b}", s_key), Paragraph(f'<font color="#BB86FC">{total_new_b}</font>', s_val)],
                ]
                summary_table = Table(summary_data, colWidths=[10*cm, 6*cm])
                summary_table.setStyle(TableStyle([
                    ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, CLR_STRIPE]),
                    ("GRID", (0, 0), (-1, -1), 0.5, CLR_LINE),
                    ("BOX", (0, 0), (-1, -1), 1, CLR_GOLD),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ]))
                flowables.append(summary_table)
                flowables.append(Spacer(1, 0.5*cm))

                # Detalle por par
                for i, (path_a, path_b, meta_a, meta_b) in enumerate(all_pair_data):
                    if cancelled.is_set():
                        self.after(0, prog.destroy)
                        return
                    self.after(0, update_prog, 0.8 + 0.2 * (i / total), f"Escribiendo par {i+1}/{total}...")

                    flowables.append(HRFlowable(width="100%", thickness=1, color=CLR_LINE, spaceBefore=8))
                    flowables.append(Paragraph(
                        f"Par {i+1}/{total}:  {Path(path_a).name}  vs  {Path(path_b).name}", s_section))

                    # Imágenes lado a lado
                    try:
                        imgs_row = []
                        for p in [(path_a, "A"), (path_b, "B")]:
                            img = ImageOps.exif_transpose(Image.open(p[0]))
                            if img.mode not in ("RGB", "L"):
                                img = img.convert("RGB")
                            img.thumbnail((8*cm, 5*cm), Image.LANCZOS)
                            buf = io.BytesIO()
                            img.save(buf, format="JPEG", quality=85)
                            buf.seek(0)
                            rl_img = RLImage(buf)
                            rl_img.drawWidth = min(8*cm, rl_img.drawWidth)
                            rl_img.drawHeight = min(5*cm, rl_img.drawHeight)
                            imgs_row.append([
                                Paragraph(f"<b>{p[1]}: {Path(p[0]).name}</b>", s_name),
                                rl_img
                            ])
                        img_table = Table(
                            [[imgs_row[0][0], imgs_row[1][0]],
                             [imgs_row[0][1], imgs_row[1][1]]],
                            colWidths=[8.5*cm, 8.5*cm])
                        img_table.setStyle(TableStyle([
                            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("BOX", (0, 0), (-1, -1), 1, CLR_GOLD),
                            ("BACKGROUND", (0, 0), (-1, -1), CLR_PAPER),
                            ("TOPPADDING", (0, 0), (-1, -1), 6),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                        ]))
                        flowables.append(img_table)
                        flowables.append(Spacer(1, 0.3*cm))
                    except Exception:
                        pass

                    # Análisis del par
                    keys_a, keys_b = set(meta_a), set(meta_b)
                    identical = keys_a & keys_b
                    identical_same = {k for k in identical if meta_a[k] == meta_b[k]}
                    identical_diff = identical - identical_same
                    only_a = keys_a - keys_b
                    only_b = keys_b - keys_a

                    if identical_diff or only_a or only_b:
                        analysis_items = []
                        if identical_diff:
                            analysis_items.append(
                                f'<font color="#E05050"><b>Diferentes ({len(identical_diff)}):</b></font> ' +
                                ", ".join(sorted(identical_diff)))
                        if only_a:
                            analysis_items.append(
                                f'<font color="#7EB8C9"><b>Sólo en A ({len(only_a)}):</b></font> ' +
                                ", ".join(sorted(only_a)))
                        if only_b:
                            analysis_items.append(
                                f'<font color="#BB86FC"><b>Sólo en B ({len(only_b)}):</b></font> ' +
                                ", ".join(sorted(only_b)))
                        for item in analysis_items:
                            flowables.append(Paragraph(f"• {item}", s_val))
                        flowables.append(Spacer(1, 0.2*cm))

                    # Tabla de metadatos
                    all_keys = sorted(keys_a | keys_b)
                    if all_keys:
                        tbl_data = [[Paragraph("<b>Campo</b>", s_key),
                                     Paragraph(f"<b>{name_a}</b>", s_key),
                                     Paragraph(f"<b>{name_b}</b>", s_key)]]
                        for k in all_keys:
                            va = meta_a.get(k, "—")
                            vb = meta_b.get(k, "—")
                            ca = "#7EC894" if va == vb else ("#E05050" if va != "—" and vb != "—" else "#7EB8C9")
                            cb = "#7EC894" if va == vb else ("#E05050" if va != "—" and vb != "—" else "#BB86FC")
                            tbl_data.append([
                                Paragraph(k, s_key),
                                Paragraph(f'<font color="{ca}">{va}</font>', s_val),
                                Paragraph(f'<font color="{cb}">{vb}</font>', s_val),
                            ])
                        meta_table = Table(tbl_data, colWidths=[5*cm, 6*cm, 6*cm])
                        meta_table.setStyle(TableStyle([
                            ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, CLR_STRIPE]),
                            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E5DED5")),
                            ("BOX", (0, 0), (-1, -1), 1, CLR_LINE),
                            ("TOPPADDING", (0, 0), (-1, -1), 3),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                            ("LEFTPADDING", (0, 0), (-1, -1), 4),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ]))
                        flowables.append(meta_table)

                # Pie
                flowables.append(Spacer(1, 1*cm))
                flowables.append(HRFlowable(width="100%", thickness=1, color=CLR_LINE))
                flowables.append(Paragraph(
                    "<b>MetaTag Visor v9.0</b>  •  Análisis generado automáticamente", s_sub))

                self.after(0, update_prog, 0.95, "Guardando archivo...")
                doc.build(flowables)

                def finish():
                    try:
                        prog.destroy()
                    except Exception:
                        pass
                    messagebox.showinfo("PDF Exportado", f"Comparación guardada en:\n\n{output_path}")

                self.after(0, finish)

            except Exception as e:
                def show_err():
                    try:
                        prog.destroy()
                    except Exception:
                        pass
                    messagebox.showerror("Error PDF", str(e))
                self.after(0, show_err)

        threading.Thread(target=worker, daemon=True).start()


    # ─────────────────────────────────────────────────────────────
    #  MOTOR DE EXPORTACIÓN A PDF (REPORTLAB)
    # ─────────────────────────────────────────────────────────────
    def _export_pdf(self):
        """Inicia el proceso para compilar una ficha técnica completa."""
        if not self.current_path:
            messagebox.showwarning("Aviso", "Se requiere una imagen cargada para generar el reporte.")
            return
            
        if not REPORTLAB_OK:
            messagebox.showerror(
                "Módulo Faltante", 
                "El subsistema PDF requiere la librería ReportLab.\nInstálala en la terminal con:\n\npip install reportlab"
            )
            return
            
        default_name = f"Ficha_Arqueologica_{Path(self.current_path).stem}.pdf"
        output_path = _native_save_file(
            title="Guardar Ficha Arqueológica",
            default_name=default_name)
        
        if not output_path:
            return
            
        self.config(cursor="watch")
        try:
            self._generate_pdf_document(output_path)
            messagebox.showinfo("Proceso Exitoso", f"La ficha ha sido compilada correctamente en:\n\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error en Compilación PDF", f"Falló la creación del documento:\n{e}")
        finally:
            self.config(cursor="")

    def _prepare_image_for_pdf(self, image_path):
        """Redimensiona la imagen para que quepa en un A4 sin pesar tanto."""
        try:
            img = Image.open(image_path)
            img = ImageOps.exif_transpose(img)
            
            max_width_cm = 17.0
            max_width_px = int(max_width_cm * 71.43)
            
            if img.width > max_width_px:
                ratio = max_width_px / img.width
                new_size = (max_width_px, int(img.height * ratio))
                img = img.resize(new_size, Image.LANCZOS)
                
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
                
            img_buffer = io.BytesIO()
            img.save(img_buffer, format="JPEG", quality=85)
            img_buffer.seek(0)
            
            pdf_image = RLImage(img_buffer)
            pdf_image.drawWidth = min(13.5*cm, pdf_image.drawWidth)
            pdf_image.drawHeight = min(10*cm, pdf_image.drawHeight)
            return pdf_image
            
        except Exception as e:
            logging.error(f"Error preparando imagen para PDF: {e}", exc_info=True)
            return None

    def _generate_pdf_document(self, output_path: str):
        """Construye el PDF página por página con estilos y tablas complejas."""
        document = SimpleDocTemplate(
            output_path, 
            pagesize=A4,
            topMargin=1.5*cm, 
            bottomMargin=1.5*cm,
            leftMargin=2*cm,  
            rightMargin=2*cm
        )
        
        stylesheet = getSampleStyleSheet()
        
        # Paleta de colores para el papel
        CLR_GOLD   = colors.HexColor("#D4A574")
        CLR_BROWN  = colors.HexColor("#5C3518")
        CLR_PAPER  = colors.HexColor("#FAF7F2")
        CLR_STRIPE = colors.HexColor("#F3EDE5")
        CLR_LINE   = colors.HexColor("#D4C8B4")

        # Configuración de estilos tipográficos PDF
        style_title = ParagraphStyle(
            "MainTitle", parent=stylesheet["Title"],
            fontName="Helvetica-Bold", fontSize=18,
            textColor=CLR_BROWN, spaceAfter=6, alignment=TA_CENTER
        )
        style_subtitle = ParagraphStyle(
            "Sub", parent=stylesheet["Normal"],
            fontName="Helvetica-Oblique", fontSize=10,
            textColor=CLR_GOLD, spaceAfter=15, alignment=TA_CENTER
        )
        style_section = ParagraphStyle(
            "SecHeader", parent=stylesheet["Normal"],
            fontName="Helvetica-Bold", fontSize=11,
            textColor=CLR_BROWN, spaceBefore=15, spaceAfter=6
        )
        style_key = ParagraphStyle(
            "TKey", parent=stylesheet["Normal"],
            fontName="Helvetica-Bold", fontSize=8.5,
            textColor=colors.HexColor("#2A1505")
        )
        style_val = ParagraphStyle(
            "TVal", parent=stylesheet["Normal"],
            fontName="Helvetica", fontSize=8.5,
            textColor=CLR_BROWN
        )

        flowables = []
        
        # 1. Cabecera del Documento
        flowables.append(Paragraph("FICHA DE REGISTRO ARQUEOLÓGICO", style_title))
        flowables.append(Paragraph(f"MetaTag Visor v9.0  —  Fecha de Impresión: {datetime.datetime.now():%d/%m/%Y %H:%M}", style_subtitle))
        flowables.append(HRFlowable(width="100%", thickness=2, color=CLR_GOLD, spaceAfter=15))

        # 2. Inserción Segura de Imagen Fotográfica
        if PIL_OK and self.current_path:
            pdf_image = self._prepare_image_for_pdf(self.current_path)
            if pdf_image:
                img_table = Table([[pdf_image]], colWidths=[16*cm])
                img_table.setStyle(TableStyle([
                    ("ALIGN",         (0,0), (-1,-1), "CENTER"),
                    ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
                    ("BOX",           (0,0), (-1,-1), 1.5, CLR_GOLD),
                    ("BACKGROUND",    (0,0), (-1,-1), CLR_PAPER),
                    ("TOPPADDING",    (0,0), (-1,-1), 10),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 10),
                ]))
                flowables.append(img_table)
                flowables.append(Spacer(1, 0.5*cm))

        # 3. Metadatos del Archivo de Identificación
        flowables.append(Paragraph(f"<b>Identificador de Archivo:</b> {Path(self.current_path).name}", style_subtitle))
        flowables.append(HRFlowable(width="100%", thickness=1, color=CLR_LINE, spaceAfter=10))

        # 4. Creación Iterativa de Secciones de Tabla
        SECTION_DEFS = {
            "json": ("🏺 Extracto Arqueológico (MetaTag)", "#D4A574"),
            "exif": ("📸 Parámetros de Exposición de Cámara", "#7EB8C9"),
            "gps":  ("🌍 Coordenadas Geográficas y Altitud", "#A8D4A0"),
            "file": ("📄 Diagnóstico Físico del Archivo", "#8A8A9A"),
        }
        
        current_rows = []
        current_title = None
        current_border_col = "#D4A574"

        def flush_table_section():
            """Función interna para vaciar el buffer de filas y pintar la tabla de una sección."""
            nonlocal current_rows
            if not current_rows: 
                return
                
            flowables.append(Paragraph(current_title, style_section))
            
            # Formatear filas en Párrafos para soportar saltos de línea automáticos
            formatted_data = []
            for k, v in current_rows:
                formatted_data.append([Paragraph(k, style_key), Paragraph(str(v), style_val)])
                
            data_table = Table(formatted_data, colWidths=[5.5*cm, 10.5*cm])
            
            # Aplicar la piel de la tabla
            data_table.setStyle(TableStyle([
                ("BACKGROUND",    (0,0),  (0,-1),   CLR_PAPER), # Columna Izq fondo suave
                ("ROWBACKGROUNDS",(0,0),  (-1,-1),  [colors.white, CLR_STRIPE]), # Filas cebra
                ("GRID",          (0,0),  (-1,-1),  0.5, colors.HexColor("#E5DED5")), # Cuadrícula
                ("BOX",           (0,0),  (-1,-1),  1, colors.HexColor(current_border_col[:7])), # Borde según sección
                ("TOPPADDING",    (0,0),  (-1,-1),  5),
                ("BOTTOMPADDING", (0,0),  (-1,-1),  5),
                ("LEFTPADDING",   (0,0),  (-1,-1),  8),
                ("RIGHTPADDING",  (0,0),  (-1,-1),  8),
                ("VALIGN",        (0,0),  (-1,-1),  "TOP"),
            ]))
            
            flowables.append(data_table)
            flowables.append(Spacer(1, 0.3*cm))
            current_rows = []

        # Recorrer todos los metadatos y dividirlos por sus headers
        for tag_type, key, value in self.all_metadata:
            if tag_type == "header":
                flush_table_section()
                sect_info = SECTION_DEFS.get(key, (value, "#D4A574"))
                current_title = sect_info[0]
                current_border_col = sect_info[1]
                
            # Excluir de la tabla impresa los avisos de campos vacíos
            elif key not in {"(Sin datos JSON)", "(Sin EXIF)", "(Sin GPS)", "(Error EXIF)"}:
                current_rows.append((key, value))
                
        flush_table_section() # Volcar la última sección que quedó en memoria

        # 5. Pie de Página
        flowables.append(Spacer(1, 1*cm))
        flowables.append(HRFlowable(width="100%", thickness=1, color=CLR_LINE, spaceBefore=10))
        flowables.append(Paragraph("<b>Programa de Conservación Cerámica</b>  •  Universidad del Magdalena", style_subtitle))
        
        document.build(flowables)


# ══════════════════════════════════════════════════════════════════
#  ARRANQUE PRINCIPAL DE LA APLICACIÓN
# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    passed_image = sys.argv[1] if len(sys.argv) > 1 else None
    app = VisorApp(initial_image=passed_image)
    app.mainloop()
