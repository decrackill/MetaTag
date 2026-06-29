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
from pathlib import Path

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
        self._pinned_path = None
        self._pinned_meta = []
        self._comp_window = None

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

        # Si el programa se abrió enviándole una imagen por parámetro (Abrir con...)
        if initial_image and initial_image != "STANDALONE" and os.path.exists(initial_image):
            self._hide_explorer_pane()
            self.load_image(initial_image)
            return

        # 1. Restaurar la carpeta si existe
        if last_dir and os.path.isdir(last_dir):
            self._open_folder_path(last_dir, auto_load=False)
            self._show_explorer_pane()  # Asegurar que el explorador se muestre
        
        # 2. Restaurar la imagen exacta
        if last_img and os.path.exists(last_img):
            if self.folder_images and last_img in self.folder_images:
                self.load_image(last_img)
                idx = self.folder_images.index(last_img)
                self.file_listbox.selection_clear(0, "end")
                self.file_listbox.selection_set(idx)
                self.file_listbox.see(idx)
            else:
                # La imagen existe pero no pertenece a la carpeta cargada
                self._hide_explorer_pane()
                self.load_image(last_img)
        elif self.folder_images:
            # Si hay carpeta pero no imagen guardada, cargar la primera
            self.load_image(self.folder_images[0])
            self.file_listbox.selection_set(0)
        else:
            # Si no hay nada, mostrar pantalla de bienvenida
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
            ("📌 Fijar",        self._pin_current,     False),
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
        """Construye el explorador de la carpeta."""
        # Encabezado del explorador
        exp_hdr = tk.Frame(parent, bg=C["panel"])
        exp_hdr.pack(fill="x", padx=12, pady=(4, 6))
        
        tk.Label(
            exp_hdr, text="EXPLORADOR", bg=C["panel"], 
            fg=C["accent"], font=("Segoe UI", 9, "bold")
        ).pack(side="left")
        
        self._lbl_count = tk.Label(exp_hdr, text="", bg=C["panel"], fg=C["text3"], font=F_MICRO)
        self._lbl_count.pack(side="right")

        # Barra de búsqueda de archivos
        search_frame = tk.Frame(parent, bg=C["panel2"], highlightthickness=1, highlightbackground=C["border"])
        search_frame.pack(fill="x", padx=12, pady=(0, 8))
        
        tk.Label(search_frame, text=" 🔍 ", bg=C["panel2"], fg=C["text3"], font=F_TINY).pack(side="left")
        
        self._folder_search = tk.StringVar()
        self._folder_search.trace_add("write", self._filter_folder)
        
        tk.Entry(
            search_frame, textvariable=self._folder_search, bg=C["panel2"], 
            fg=C["text"], insertbackground=C["accent"], relief="flat", bd=0, font=F_BODY
        ).pack(side="left", fill="x", expand=True, ipady=6)

        # Lista de archivos (Listbox)
        listbox_frame = tk.Frame(parent, bg=C["surface"], highlightthickness=1, highlightbackground=C["border"])
        listbox_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        
        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical")
        
        self.file_listbox = tk.Listbox(
            listbox_frame, bg=C["surface"], fg=C["text"], 
            selectbackground=C["sel_bg"], selectforeground=C["sel_fg"], 
            relief="flat", bd=0, font=F_BODY, activestyle="none", 
            yscrollcommand=scrollbar.set
        )
        
        scrollbar.configure(command=self.file_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.file_listbox.pack(fill="both", expand=True, padx=2, pady=2)
        
        self.file_listbox.bind("<<ListboxSelect>>", self._on_list_select)

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
            self._refresh_listbox()
            
        if saved_path and os.path.exists(saved_path):
            self.load_image(saved_path)
            if saved_path in self.folder_images:
                idx = self.folder_images.index(saved_path)
                self.file_listbox.selection_clear(0, "end")
                self.file_listbox.selection_set(idx)
                self.file_listbox.see(idx)
        elif not self.folder_images:
            self._show_welcome()

    # ─────────────────────────────────────────────────────────────
    #  NAVEGACIÓN DE ARCHIVOS
    # ─────────────────────────────────────────────────────────────
    def _browse(self):
        """Abre un solo archivo mediante diálogo."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Imágenes", "*.jpg *.jpeg *.png *.tif *.tiff *.webp")]
        )
        if file_path:
            self._hide_explorer_pane()
            self.folder_images = []
            self.load_image(file_path)

    def _browse_folder(self):
        """Abre una carpeta mediante diálogo."""
        folder_path = filedialog.askdirectory(title="Seleccionar carpeta de estudio")
        if folder_path:
            self._open_folder_path(folder_path, auto_load=True)

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
        
        self._show_explorer_pane()
        self._refresh_listbox()
        
        if auto_load:
            self.file_listbox.selection_clear(0, "end")
            self.file_listbox.selection_set(0)
            self.load_image(self.folder_images[0])

    def _filter_folder(self, *_):
        """Filtra la lista de archivos mediante la barra de búsqueda inferior."""
        search_term = self._folder_search.get().lower()
        all_items = getattr(self, "_all_folder_images", self.folder_images)
        
        if search_term:
            self.folder_images = [p for p in all_items if search_term in Path(p).name.lower()]
        else:
            self.folder_images = all_items[:]
            
        self._refresh_listbox()

    def _refresh_listbox(self):
        """Actualiza la visualización gráfica del Listbox de archivos mediante carga en bloque."""
        self.file_listbox.delete(0, "end")
        
        # Inserción en bloque para optimizar el rendimiento de la GUI
        items_to_insert = [f"  {Path(filepath).name}" for filepath in self.folder_images]
        if items_to_insert:
            self.file_listbox.insert("end", *items_to_insert)
            
        # Aplicar colores alternos a las filas
        for i in range(len(items_to_insert)):
            bg_color = C["row_alt"] if i % 2 != 0 else C["surface"]
            self.file_listbox.itemconfigure(i, background=bg_color)
            
        self._lbl_count.configure(text=f"{len(self.folder_images)} archivos")

    def _on_list_select(self, event):
        """Evento lanzado al hacer clic en un archivo del explorador."""
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            self.load_image(self.folder_images[index])


    # ─────────────────────────────────────────────────────────────
    #  CARGA DE IMAGEN PRINCIPAL
    # ─────────────────────────────────────────────────────────────
    def load_image(self, path: str):
        """Carga la imagen, actualiza la información y dispara la extracción de metadatos."""
        if not PIL_OK: 
            messagebox.showerror("Error", "Faltan dependencias. Ejecuta: pip install pillow piexif")
            return
            
        self._hide_welcome()
        self.current_path = path
        
        # Resetear variables de zoom
        self._zoom_level = ZOOM_FIT
        self._pan_offset = [0, 0]

        try:
            # Abrir y corregir orientación EXIF automáticamente
            raw_img = Image.open(path)
            img = ImageOps.exif_transpose(raw_img)
            self._pil_image = img.copy()
            
            # Dar tiempo al canvas para calcular dimensiones antes del primer render
            self.after(30, self._redraw_zoom)
            
            # Actualizar tarjeta de información
            kb = os.path.getsize(path) / 1024
            info_text = (
                f"📄 {Path(path).name}\n"
                f"📁 {Path(path).parent.name}\n\n"
                f"⚖  {kb:.1f} KB   |   📐 {img.size[0]}×{img.size[1]} px\n"
                f"🎨 Formato: {img.format or '—'}  |  Modo: {img.mode}"
            )
            self.lbl_info.configure(text=info_text)
            
        except Exception as e:
            self.img_canvas.delete("all")
            self.img_canvas.create_text(
                200, 150, text=f"Error cargando la imagen:\n{e}", 
                fill=C["err"], font=F_BODY, justify="center"
            )
            self._pil_image = None
            self.lbl_info.configure(text="Error de lectura de archivo.")

        # Iniciar motor de extracción
        self._extract_all_metadata(path)
        self._filter_tree()
        self.title(f"⬡ MetaTag Visor v9 — {Path(path).name}")

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
    #  HERRAMIENTA DE COMPARACIÓN DE METADATOS
    # ─────────────────────────────────────────────────────────────
    def _pin_current(self):
        """Fija en memoria los datos actuales para compararlos con el siguiente archivo."""
        if not self.current_path: 
            messagebox.showwarning("Aviso", "No hay ninguna imagen cargada.")
            return
            
        self._pinned_path = self.current_path
        self._pinned_meta = self.all_metadata[:]
        
        messagebox.showinfo(
            "Muestra Fijada", 
            f"Se ha memorizado la pieza:\n{Path(self.current_path).name}\n\n"
            "Ahora puedes hacer clic en otra imagen en el explorador y seleccionar 'Comparar'."
        )

    def _open_comparator(self):
        """Abre la ventana dedicada a mostrar el Delta entre dos cerámicas."""
        if not self._pinned_path:
            messagebox.showwarning("Aviso", "Primero usa el botón 'Fijar' en una imagen de referencia.")
            return
            
        if not self.current_path or self.current_path == self._pinned_path:
            messagebox.showwarning("Aviso", "Abre una imagen diferente a la fijada para hacer la comparación.")
            return
            
        if self._comp_window and self._comp_window.winfo_exists():
            self._comp_window.lift()
            return

        COLORS = {
            "same": "#7EC894", 
            "diff": "#E05050",
            "only_a": "#7EB8C9", 
            "only_b": "#BB86FC"
        }

        # Setup ventana TopLevel
        win = tk.Toplevel(self)
        self._comp_window = win
        win.title("⚖ Comparador Analítico de Metadatos")
        win.configure(bg=C["bg"])
        
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        ww, wh = min(1100, int(sw*0.85)), min(720, int(sh*0.85))
        win.geometry(f"{ww}x{wh}+{(sw-ww)//2}+{(sh-wh)//2}")

        # Topbar Comparador
        hdr = tk.Frame(win, bg=C["header_bg"], height=55)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        
        tk.Label(
            hdr, text="⚖  Análisis Comparativo", font=F_TITLE, 
            bg=C["header_bg"], fg=C["header_fg"]
        ).pack(side="left", padx=20, pady=10)
        
        # Leyenda
        leg = tk.Frame(hdr, bg=C["header_bg"])
        leg.pack(side="right", padx=20)
        for txt, col in [("● Coincide", COLORS["same"]), ("● Discrepa", COLORS["diff"]), ("● Sólo en Base", COLORS["only_a"]), ("● Sólo en Nueva", COLORS["only_b"])]:
            tk.Label(leg, text=txt, bg=C["header_bg"], fg=col, font=F_BOLD).pack(side="left", padx=10)

        # Nombres de archivos a comparar
        names_frame = tk.Frame(win, bg=C["panel"])
        names_frame.pack(fill="x", padx=15, pady=(15, 5))
        tk.Label(names_frame, text=f"📌 A (Base): {Path(self._pinned_path).name}", bg=C["panel"], fg=C["accent"], font=F_BOLD, anchor="w").pack(side="left", fill="x", expand=True, padx=15, pady=10)
        tk.Label(names_frame, text=f"🔵 B (Nueva): {Path(self.current_path).name}", bg=C["panel"], fg=C["accent"], font=F_BOLD, anchor="w").pack(side="left", fill="x", expand=True, padx=15, pady=10)

        # Tabla del Comparador
        table_frame = tk.Frame(win, bg=C["border"], highlightthickness=1, highlightbackground=C["border"])
        table_frame.pack(fill="both", expand=True, padx=15, pady=(0, 5))
        
        cols = ("Propiedad", "Valor en A", "Valor en B")
        tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")
        
        for cid, width_val in zip(cols, [220, 360, 360]):
            tree.heading(cid, text=cid, anchor="w")
            tree.column(cid, width=width_val, anchor="w")
            
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True)

        # Lógica de comparación de diccionarios
        def convert_to_dict(meta_list):
            ignore_keys = {"(Sin datos JSON)", "(Sin EXIF)", "(Sin GPS)", "(Sin datos GPS)", "(Error EXIF)", ""}
            return {key: val for t_tag, key, val in meta_list if t_tag != "header" and key not in ignore_keys}

        dict_a = convert_to_dict(self._pinned_meta)
        dict_b = convert_to_dict(self.all_metadata)
        
        count_same = count_diff = count_only_a = count_only_b = 0

        # Popular tabla cruzando los datos
        for key in sorted(set(dict_a.keys()) | set(dict_b.keys())):
            val_a = dict_a.get(key, "—")
            val_b = dict_b.get(key, "—")
            
            if val_a == "—":            
                color = COLORS["only_b"]
                count_only_b += 1
            elif val_b == "—":          
                color = COLORS["only_a"]
                count_only_a += 1
            elif val_a == val_b:           
                color = COLORS["same"]
                count_same += 1
            else:                    
                color = COLORS["diff"]
                count_diff += 1
                
            tag_name = f"comp_color_{color[1:]}"
            tree.insert("", "end", values=(key, val_a, val_b), tags=(tag_name,))
            tree.tag_configure(tag_name, foreground=color)

        # Barra de Resumen Inferior
        summary_bar = tk.Frame(win, bg=C["panel"])
        summary_bar.pack(fill="x", padx=15, pady=(0, 15))
        
        summary_txt = (
            f"Total Analizados: {count_same + count_diff + count_only_a + count_only_b}   |   "
            f"Iguales: {count_same}   |   Diferentes: {count_diff}   |   "
            f"Faltantes en B: {count_only_a}   |   Nuevos en B: {count_only_b}"
        )
        tk.Label(summary_bar, text=summary_txt, bg=C["panel"], fg=C["text2"], font=F_BODY).pack(padx=15, pady=8)


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
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", 
            initialfile=default_name, 
            filetypes=[("Documento PDF", "*.pdf")]
        )
        
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
            try:
                img_to_print = ImageOps.exif_transpose(Image.open(self.current_path))
                if img_to_print.mode not in ("RGB", "L"):
                    img_to_print = img_to_print.convert("RGB")
                    
                # Redimensionar para no exceder ancho de página
                max_w_px = int(13.5 * cm * 37.8)
                max_h_px = int(10 * cm * 37.8)
                img_to_print.thumbnail((max_w_px, max_h_px), Image.LANCZOS)
                
                # Volcar a buffer RAM (evita problemas de locking en disco)
                buffer = io.BytesIO()
                img_to_print.save(buffer, format="JPEG", quality=92)
                buffer.seek(0)
                
                pdf_image = RLImage(buffer)
                pdf_image.drawWidth  = min(13.5*cm, pdf_image.drawWidth)
                pdf_image.drawHeight = min(10*cm,  pdf_image.drawHeight)
                
                # Enmarcar la foto en una tabla estética de una celda
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
                
            except Exception as e:
                print(f"Alerta: No se pudo inyectar imagen en PDF: {e}")

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
