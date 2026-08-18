"""
renombrar_fotos_gui.py — Image Sync (Renombrador de Fotos v4.1)

Aplicación de escritorio para renombrar fotos en lote usando nombres
de un archivo Excel (.xlsx) o CSV. Construida con CustomTkinter.

Funcionalidades principales:
- Carga de fotos desde carpeta con soporte de 9 formatos
- Lectura de Excel (.xlsx) y CSV/TSV
- Vista previa con miniaturas y detección de duplicados
- Modo copia (no modifica originales)
- Deshacer multinivel
- Atajos de teclado: Ctrl+O/E/Z, Escape, Ctrl+Enter

Autor: Deivis
"""
from __future__ import annotations

import csv
import json
import logging
import os
import platform
import re
import shutil
import string
import subprocess
import threading
import tkinter as tk
from tkinter import ttk
import traceback
import uuid
from collections import Counter, OrderedDict
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import customtkinter as ctk
import pandas as pd
from PIL import Image

# ── matching seguro (opcional) ──────────────────────────────────────────────
# El Renombrador puede funcionar de forma totalmente independiente (modo
# posicional). Si el proyecto MetaTag está presente, importa el motor de
# emparejamiento seguro (puro, sin Tkinter) para el modo "matching seguro".
# El script vive en src/ junto a metatag_matching.py: se inserta su propio
# directorio en sys.path para funcionar sin depender del cwd de invocación.
try:
    import sys
    _PROJECT_SRC = Path(__file__).resolve().parent
    if str(_PROJECT_SRC) not in sys.path:
        sys.path.insert(0, str(_PROJECT_SRC))
    from metatag_matching import ImageMatcher
except Exception:
    ImageMatcher = None

# ── tokens técnicos de tema (fuente de verdad: metatag_theme.py) ──
import metatag_theme as mt
from metatag_theme import (
    CustomTkinterThemeAdapter, compute_font_scale, scaled_size,
    DEFAULT_THEME, THEME_ORDER,
)
from metatag_responsive import PROFILE

# Estados posibles de cada fila de la vista previa / plan de renombrado.
#   ok          → se renombrará sin obstáculos.
#   ya_correcto → el archivo ya tiene exactamente ese nombre.
#   existe      → el destino es un archivo EXTERNO (no pertenece al lote) que
#                 ya existe: jamás se sobreescribe (bloquea).
#   conflicto   → colisión interna NO resoluble dentro del lote (bloquea).
#   duplicado   → dos filas compiten por el mismo nombre destino (se omite).
#   not_found   → matching: no se halló la fotografía de ese registro.
#   sin_foto    → posicional: hay más registros que fotografías.
#   ambiguo     → matching: varios archivos compiten por la misma clave.
#   error       → fallo del motor de emparejamiento (sin tocar disco).
PLAN_STATES = ("ok", "ya_correcto", "existe", "conflicto", "duplicado",
               "not_found", "sin_foto", "ambiguo", "error")
STATE_LABELS = {
    "ok": "✓ Correcto", "ya_correcto": "✓ Correcto", "existe": "⚠ Ya existe",
    "conflicto": "⚠ Conflicto",
    "duplicado": "⚠ Duplicado", "not_found": "✕ Sin foto",
    "sin_foto": "✕ Sin foto", "ambiguo": "⚠ Ambiguo", "error": "✕ Error",
}
# Los colores por estado viven en el tema técnico (C["state_bg"] / C["state_fg"]),
# derivados de los semánticos canónicos de MetaTag (ok / err / warn).

# ── logging ────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")
log = logging.getLogger(__name__)

# ── tema ───────────────────────────────────────────────────────────────────
# Los colores provienen de metatag_theme (los 3 temas canónicos de MetaTag).
# C es un alias global que se refresca al cambiar de tema (y la vista se
# reconstruye completa para que ningún widget conserve colores del tema viejo).
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

CURRENT_THEME: str = DEFAULT_THEME
_THEME_ADAPTER = CustomTkinterThemeAdapter()
C = _THEME_ADAPTER.palette(CURRENT_THEME)

def _refresh_button_constants() -> None:
    global BTN_SECONDARY, BTN_DANGER, BTN_PRIMARY
    BTN_SECONDARY = {"fg_color": C["surface2"], "hover_color": C["accent"]}
    BTN_DANGER    = {"fg_color": C["surface2"], "hover_color": C["red"]}
    BTN_PRIMARY   = {"fg_color": C["accent"],   "hover_color": C["accent2"]}

_refresh_button_constants()

# ── fuentes (se inicializan en MainView.__init__ después de crear la ventana Tk) ──
FONT_XS: ctk.CTkFont
FONT_XS_SM: ctk.CTkFont
FONT_SM: ctk.CTkFont
FONT_SM_BD: ctk.CTkFont
FONT_MD: ctk.CTkFont
FONT_MD_BD: ctk.CTkFont
FONT_LG_BD: ctk.CTkFont
FONT_HD: ctk.CTkFont
FONT_TITLE: ctk.CTkFont

def _init_fonts(screen_width: int) -> None:
    global FONT_XS, FONT_XS_SM, FONT_SM, FONT_SM_BD, FONT_MD, FONT_MD_BD, FONT_LG_BD, FONT_HD, FONT_TITLE
    _family = "Segoe UI" if platform.system() == "Windows" else (
        ".SF NS Text" if platform.system() == "Darwin" else "Noto Sans")
    # Misma lógica de escalado que MetaTag (misma referencia, mismo rango y la
    # fórmula max(floor, int(base*scale))) pero con las bases propias del
    # Renombrador (9–18). En pantalla de referencia (1920px) queda idéntico
    # al escalado actual.
    _s = compute_font_scale(screen_width)
    FONT_XS    = ctk.CTkFont(_family, scaled_size(9,  _s, 6))
    FONT_XS_SM = ctk.CTkFont(_family, scaled_size(10, _s, 6))
    FONT_SM    = ctk.CTkFont(_family, scaled_size(11, _s, 7))
    FONT_SM_BD = ctk.CTkFont(_family, scaled_size(11, _s, 7), "bold")
    FONT_MD    = ctk.CTkFont(_family, scaled_size(12, _s, 7))
    FONT_MD_BD = ctk.CTkFont(_family, scaled_size(12, _s, 7), "bold")
    FONT_LG_BD = ctk.CTkFont(_family, scaled_size(13, _s, 8), "bold")
    FONT_HD    = ctk.CTkFont(_family, scaled_size(14, _s, 8), "bold")
    FONT_TITLE = ctk.CTkFont(_family, scaled_size(18, _s, 10), "bold")

# ── persistencia ───────────────────────────────────────────────────────────
_STATE_FILE = Path(__file__).parent / ".renombrador_state.json"

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
    import tempfile
    try:
        fd, tmp = tempfile.mkstemp(dir=_STATE_FILE.parent, suffix=".tmp")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
        os.replace(tmp, _STATE_FILE)
    except Exception as exc:
        log.warning("No se pudo guardar estado: %s", exc)

def _safe_cancel_after(widget, job_id) -> None:
    if job_id:
        try:
            widget.after_cancel(job_id)
        except Exception:
            pass

def _make_option_menu(parent, variable, values, *, width=185, height=28,
                      font=None, fg=None, btn_fg=None,
                      command=None) -> ctk.CTkOptionMenu:
    return ctk.CTkOptionMenu(
        parent, variable=variable, values=values, width=width, height=height,
        font=font or FONT_SM, fg_color=fg or C["surface"],
        button_color=btn_fg or C["surface2"],
        button_hover_color=C["accent"], command=command)

def _show_toplevel(dlg) -> None:
    dlg.deiconify()
    dlg.update_idletasks()
    try:
        dlg.wait_visibility()
        dlg.grab_set()
    except Exception:
        pass

def _place_tip_near_pointer(tip, widget, offset: int = 18) -> None:
    """Coloca un Toplevel flotante junto al CURSOR (no a la derecha del
    widget), y lo ajusta para que nunca quede fuera de pantalla.

    Se mide el tamaño real del tooltip y se decide el lado:
    - abajo-derecha por defecto;
    - si desborda la pantalla por la derecha → a la izquierda del cursor;
    - si desborda por abajo → por arriba del cursor.
    """
    try:
        tip.update_idletasks()
        tw, th = tip.winfo_reqwidth(), tip.winfo_reqheight()
        px = widget.winfo_pointerx()
        py = widget.winfo_pointery()
        sw = tip.winfo_screenwidth()
        sh = tip.winfo_screenheight()
    except Exception:
        try:
            tip.geometry(f"+10+10")
        except Exception:
            pass
        return
    x = px + offset
    if x + tw > sw:
        x = px - offset - tw
    y = py + offset
    if y + th > sh:
        y = py - offset - th
    x = max(0, min(x, sw - tw))
    y = max(0, min(y, sh - th))
    tip.geometry(f"+{x}+{y}")

# ── caché de miniaturas (LRU, máx 150 entradas) ───────────────────────────
_THUMB_CACHE: OrderedDict[str, ctk.CTkImage] = OrderedDict()
_THUMB_MAX = 150
_THUMB_LOCK = threading.Lock()

def _get_thumb(path: Path, size: tuple[int, int] = (56, 56)) -> Optional[ctk.CTkImage]:
    """Devuelve miniatura cacheada o la genera. Nunca lanza excepción."""
    key = f"{path}|{size}"
    with _THUMB_LOCK:
        if key in _THUMB_CACHE:
            _THUMB_CACHE.move_to_end(key)
            return _THUMB_CACHE[key]
    try:
        img = Image.open(path)
        img.thumbnail(size, Image.BILINEAR)
        ctk_img = ctk.CTkImage(img, size=img.size)
        with _THUMB_LOCK:
            _THUMB_CACHE[key] = ctk_img
            if len(_THUMB_CACHE) > _THUMB_MAX:
                _THUMB_CACHE.popitem(last=False)
        return ctk_img
    except Exception as exc:
        log.debug("Thumbnail error for %s: %s", path, exc)
        return None

# ── utilidades ─────────────────────────────────────────────────────────────
_RE_NATURAL_SPLIT = re.compile(r"(\d+)")

def _natural_key(p: Path) -> list:
    """Ordenamiento natural: foto2 antes que foto10."""
    return [int(t) if t.isdigit() else t.lower()
            for t in _RE_NATURAL_SPLIT.split(p.stem)]

def _get_exif_date(p: Path) -> float:
    """Extrae fecha EXIF de una foto. Retorna 0 si no hay EXIF."""
    try:
        img = Image.open(p)
        exif = img.getexif()
        if exif:
            val = exif.get(306)  # DateTime, nivel raíz
            if not val:
                ifd = exif.get_ifd(0x8769)
                val = ifd.get(36867) or ifd.get(36868)  # DateTimeOriginal / Digitized
            if val:
                return datetime.strptime(str(val), "%Y:%m:%d %H:%M:%S").timestamp()
    except Exception:
        pass
    return 0.0

def _detect_drives() -> list[str]:
    """
    Detecta unidades/raíces disponibles según el SO. NUNCA lanza excepción.

    FIX #2 (real): la versión anterior usaba ``Path(...).exists()``, que en
    Windows puede lanzar una excepción no controlada (en vez de devolver
    False) cuando la letra corresponde a una unidad "no lista" — un lector
    de tarjetas SD vacío, un DVD sin disco, una unidad de red desconectada.
    Eso rompía el explorador de archivos COMPLETO con un solo dispositivo
    así presente, dando la falsa impresión de que "no detecta los discos".
    ``os.path.exists`` sí absorbe ese error correctamente, y aun así se
    envuelve todo en try/except por máxima seguridad.
    """
    try:
        if platform.system() == "Windows":
            drives: list[str] = []
            for letter in string.ascii_uppercase:
                drive = f"{letter}:\\"
                try:
                    if os.path.exists(drive):
                        drives.append(drive)
                except Exception:
                    continue   # unidad presente pero no lista/accesible: se omite, no se rompe nada
            return drives or ["C:\\"]

        # Linux / macOS
        candidates = ["/", str(Path.home())]
        for base in ("/mnt", "/media", "/Volumes"):
            p = Path(base)
            try:
                if p.is_dir():
                    candidates += [str(c) for c in p.iterdir() if c.is_dir()]
            except Exception:
                pass
        return list(dict.fromkeys(candidates))   # deduplicar manteniendo orden
    except Exception as exc:
        log.warning("No se pudieron detectar las unidades de disco: %s", exc)
        return ["C:\\"] if platform.system() == "Windows" else ["/"]

SORT_OPTIONS: dict[str, str] = {
    "Orden numérico":        "natural",
    "Nombre (A → Z)":        "name_asc",
    "Nombre (Z → A)":        "name_desc",
    "Fecha modificación ↑":  "mtime_asc",
    "Fecha modificación ↓":  "mtime_desc",
    "Fecha creación ↑":      "ctime_asc",
    "Fecha creación ↓":      "ctime_desc",
    "Fecha foto ↑":          "exif_asc",
    "Fecha foto ↓":          "exif_desc",
}
VALID_IMG_EXT: frozenset[str] = frozenset(
    {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tiff", ".heic", ".avif"}
)

# Caracteres prohibidos en nombres de archivo (Windows y otros FS).
_INVALID_FILENAME_CHARS: frozenset[str] = frozenset('\\/:*?"<>|')

def _invalid_name_chars(name: str) -> list[str]:
    """Devuelve la lista ordenada de caracteres inválidos presentes en `name`."""
    return sorted({c for c in name if c in _INVALID_FILENAME_CHARS})

# Pasos del flujo (indicador bajo la cabecera). Números U+2460…U+2465:
# glifos ampliamente disponibles (verificado con fc-list), a diferencia de
# los emojis de color U+1Fxxx que se rompen en Linux sin fuente de emoji.
_STEP_NAMES: tuple[str, ...] = (
    "Emparejar fotografías", "Validar correspondencias", "Vista previa",
    "Renombrar", "Resultado",
)
_STEP_DIGITS: tuple[str, ...] = ("①", "②", "③", "④", "⑤")

def _state_text_color(kind: str) -> str:
    return {"ok": C["green"], "warn": C["yellow"], "error": C["red"]}.get(
        kind, C["subtext"])


# ===========================================================================
#  MODEL
# ===========================================================================
class RenameModel:
    """Lógica de negocio completamente desacoplada de la UI (MVC)."""

    def __init__(self) -> None:
        self.folder_path:  Optional[Path] = None
        self.excel_path:   Optional[Path] = None
        self.column_name:  Optional[str]  = None
        self.sheet_name:   Optional[str]  = None
        self.sort_mode:    str             = "natural"
        # Posicional = primera foto ↔ primer nombre (compatibilidad standalone).
        # Matching seguro = cada nombre busca SU fotografía (recomendado).
        self.matching_mode: bool = False
        # Opciones de renombrado (defaults idénticos a los checkboxes de la UI).
        self.keep_extension: bool = True
        self._matcher      = None          # ImageMatcher (carga perezosa)
        self._photos:      list[Path]      = []
        self._names:       list[str]       = []
        self._plan:        list[dict]      = []   # último plan construido
        self.skipped_rows: list[int]       = []   # filas de Excel ignoradas (vacías/no leíbles) — FIX #6
        # Caché del DataFrame del Excel: se lee UNA vez por (archivo, hoja) y
        # se reutiliza al cambiar de columna, de orden o al reconstruir la UI
        # (evita releer el .xlsx varias veces — ver FIX de rendimiento).
        self._df:          Optional[pd.DataFrame] = None
        self._df_key:      Optional[tuple]        = None
        # historial multinivel para deshacer: lista de lotes
        self._undo_stack:  list[tuple[list[tuple[Path, Path]], Optional[Path], bool]] = []

    # ── matching seguro ────────────────────────────────────────────────────
    @property
    def matching_available(self) -> bool:
        return self._get_matcher() is not None

    def _get_matcher(self):
        if self._matcher is None and ImageMatcher is not None:
            try:
                self._matcher = ImageMatcher()
            except Exception:
                self._matcher = None
        return self._matcher

    @staticmethod
    def _same_file(a: Optional[Path], b: Optional[Path]) -> bool:
        """True si 'a' y 'b' son el MISMO archivo lógico (mismo inodo)."""
        if not a or not b:
            return False
        try:
            return os.path.exists(a) and os.path.exists(b) and os.path.samefile(a, b)
        except OSError:
            try:
                return a.resolve() == b.resolve()
            except OSError:
                return False

    # ── propiedades ────────────────────────────────────────────────────────
    @property
    def photos(self) -> list[Path]:
        return self._photos

    @property
    def names(self) -> list[str]:
        return self._names

    @property
    def has_undo(self) -> bool:
        return bool(self._undo_stack)

    # ── carga ──────────────────────────────────────────────────────────────
    def load_photos(self) -> int:
        """Escanea la carpeta y ordena las imágenes."""
        if not self.folder_path:
            raise ValueError("No se ha seleccionado ninguna carpeta.")
        # FIX: el matcher cachea el índice de la carpeta; si las fotos
        # cambiaron (p. ej. tras renombrar) la caché quedaría obsoleta.
        if self._matcher is not None:
            self._matcher._invalidate(str(self.folder_path))
        raw = [p for p in self.folder_path.iterdir()
               if p.is_file() and p.suffix.lower() in VALID_IMG_EXT]
        {
            "natural":   lambda: raw.sort(key=_natural_key),
            "name_asc":  lambda: raw.sort(key=lambda p: p.name.lower()),
            "name_desc": lambda: raw.sort(key=lambda p: p.name.lower(), reverse=True),
            "mtime_asc": lambda: raw.sort(key=lambda p: p.stat().st_mtime),
            "mtime_desc":lambda: raw.sort(key=lambda p: p.stat().st_mtime, reverse=True),
            "ctime_asc": lambda: raw.sort(key=lambda p: p.stat().st_ctime),
            "ctime_desc":lambda: raw.sort(key=lambda p: p.stat().st_ctime, reverse=True),
            "exif_asc":  lambda: raw.sort(key=_get_exif_date),
            "exif_desc": lambda: raw.sort(key=_get_exif_date, reverse=True),
        }.get(self.sort_mode, lambda: raw.sort(key=_natural_key))()
        self._photos = raw
        self._plan = []
        return len(raw)

    def load_sheets(self) -> list[str]:
        """Retorna las hojas del Excel o None para CSV."""
        if not self.excel_path:
            raise ValueError("No se ha seleccionado ningún archivo.")
        ext = self.excel_path.suffix.lower()
        if ext in (".csv", ".tsv", ".txt"):
            return []
        return pd.ExcelFile(self.excel_path).sheet_names

    def clear_excel_data(self) -> None:
        """Invalida caché y datos derivados al cambiar de archivo/hoja."""
        self._df = None
        self._df_key = None
        self._names = []
        self.skipped_rows = []
        self._plan = []

    def _read_df(self) -> pd.DataFrame:
        """
        Lee el DataFrame completo del Excel/CSV UNA vez por (archivo, hoja) y
        lo cachea. Las columnas, hojas y nombres se derivan de esta caché.
        """
        key = (str(self.excel_path), self.sheet_name or 0)
        if self._df is not None and self._df_key == key:
            return self._df
        if not self.excel_path:
            raise ValueError("No se ha seleccionado ningún archivo.")
        ext = self.excel_path.suffix.lower()
        if ext in (".csv", ".tsv", ".txt"):
            sep = "\t" if ext == ".tsv" else ","
            df = pd.read_csv(self.excel_path, sep=sep, keep_default_na=False,
                             encoding="utf-8")
        else:
            sheet = self.sheet_name or 0
            df = pd.read_excel(self.excel_path, sheet_name=sheet,
                               keep_default_na=False)
        self._df = df
        self._df_key = key
        return df

    def load_columns(self) -> list[str]:
        """Retorna las columnas de la hoja activa o del CSV."""
        if not self.excel_path:
            raise ValueError("No se ha seleccionado ningún archivo.")
        return list(self._read_df().columns)

    @staticmethod
    def _normalize_excel_value(raw) -> Optional[str]:
        """
        Convierte un valor de celda a un nombre de archivo seguro.

        Reglas:
        - None / vacío / "nan" / "nat" / "none" → None (se omite la fila).
        - Valores numéricos enteros (1, 1.0) → "1" (nunca "1.0").
        - Se conservan los ceros iniciales legítimos tal como vienen en el
          archivo ("001" permanece "001", nunca "1" ni "1.0").
        """
        if raw is None:
            return None
        if isinstance(raw, float):
            try:
                if raw != raw:      # NaN
                    return None
                text = str(int(raw)) if raw.is_integer() else str(raw)
            except (ValueError, OverflowError):
                text = str(raw)
        else:
            text = str(raw)
        text = text.strip()
        if text == "" or text.lower() in ("nan", "nat", "none"):
            return None
        # Flotante con formato entero legible: "1.0" → "1", "2.00" → "2".
        if re.fullmatch(r"-?\d+\.0+", text):
            text = text.split(".")[0]
        return text

    def load_names(self) -> int:
        """
        Lee los nombres de la columna seleccionada.

        FIX bug #6 (real): la versión anterior usaba ``dropna()`` / una
        exclusión silenciosa de NaN, lo que podía "perder" una fila válida
        sin avisar — por ejemplo: una celda con fórmula cuyo valor
        calculado no quedó guardado dentro del .xlsx (pandas la lee como
        NaN aunque en Excel se vea con texto), o un valor que por
        casualidad coincide con la lista de "nulos" por defecto de pandas
        (como "NA" o "N/A", que SÍ pueden aparecer en códigos de catálogo).

        Ahora:
          · se lee con keep_default_na=False, para no tratar como vacíos
            textos legítimos que coincidan con esa lista por casualidad,
          · los valores se normalizan (nunca "1.0" ni "nan" ni espacios),
            conservando ceros iniciales legítimos,
          · se guarda en self.skipped_rows el número REAL de fila de Excel
            de cualquier celda que sí haya quedado vacía, para poder
            avisarle al usuario exactamente cuál revisar.
        """
        if not (self.excel_path and self.column_name):
            raise ValueError("Archivo o columna no configurados.")
        df = self._read_df()
        if self.column_name not in df.columns:
            raise ValueError(
                f"La columna «{self.column_name}» no existe en la hoja "
                f"«{self.sheet_name or 0}».")

        names: list[str] = []
        self.skipped_rows = []
        for i, raw in enumerate(df[self.column_name]):
            text = self._normalize_excel_value(raw)
            if text is None:
                self.skipped_rows.append(i + 2)   # +1 índice 0→1, +1 por la fila de encabezado
                continue
            names.append(text)

        self._names = names
        return len(self._names)

    # ── acceso público ─────────────────────────────────────────────────────
    def set_name(self, index: int, name: str) -> None:
        if 0 <= index < len(self._names):
            self._names[index] = name

    # ── preview ────────────────────────────────────────────────────────────
    def build_preview(self) -> list[tuple[str, str, Optional[Path], bool, str]]:
        """
        Devuelve lista de (nombre_original, nombre_nuevo, path_foto, es_duplicado, estado).

        El estado representa el resultado REAL de la operación sobre esa fila
        ANTES de ejecutarla: ok / ya_correcto / conflicto / duplicado /
        not_found / ambiguo / error. Nunca debe aparecer una fila como
        segura en la preview y fallar después por una condición previsible.
        """
        plan = self._build_plan()
        pairs: list[tuple[str, str, Optional[Path], bool, str]] = []
        for item in plan:
            src = item["src"]
            orig = src.name if src else "—"
            pairs.append((orig, item["new"], src, item["is_dup"], item["state"]))
        return pairs

    def _error_plan(self, reason: str) -> list[dict]:
        """Plan de error: una fila por nombre, sin tocar disco nunca.

        Se usa cuando el modo matching seguro está activo pero NO se puede
        emparejar (motor no disponible o carpeta no seleccionada). Un plan
        de error NUNCA puede degradar a posicional: renombrar por posición
        con el matching activo renombraría fotos equivocadas en silencio.
        """
        plan: list[dict] = []
        for i, name in enumerate(self._names):
            plan.append({
                "src": None, "new": name, "state": "error", "reason": reason,
                "method": "error", "candidates": [], "is_dup": False,
                "name": name, "row": i + 2,
            })
        self._plan = plan
        return plan

    def _build_plan(self) -> list[dict]:
        """Construye el plan de renombrado (posicional o por matching seguro).

        Seguridad: si el matching seguro está activo pero el motor no está
        disponible o no hay carpeta, devuelve un plan de ERROR. Queda
        PROHIBIDO el fallback posicional silencioso (renombraría por
        posición fotos que el usuario esperaba emparejadas por nombre).
        """
        if self.matching_mode:
            if not self.matching_available:
                return self._error_plan(
                    "el motor de matching seguro no está disponible "
                    "(no se pudo importar metatag_matching)")
            if not self.folder_path or not self.folder_path.is_dir():
                return self._error_plan(
                    "no hay una carpeta de imágenes seleccionada")
            return self._build_plan_matching()
        return self._build_plan_positional()

    def _build_plan_positional(self) -> list[dict]:
        """
        Plan posicional (compatibilidad standalone): primera foto ↔ primer
        nombre. NO garantiza identidad de fotografías; se documenta en la UI.

        Los nombres del Excel que se quedan SIN fotografía aparecen como
        filas explícitas con estado "sin_foto" (NUNCA se descartan en
        silencio): la vista previa y el resumen las muestran para que el
        usuario sepa exactamente cuántos registros quedan sin imagen.
        """
        seen_dest: dict[str, int] = {}
        plan: list[dict] = []
        for i, name in enumerate(self._names):
            bad = _invalid_name_chars(name)
            if bad:
                bad_txt = " ".join(bad)
                plan.append({
                    "src": None, "new": name, "state": "error",
                    "reason": f"el nombre contiene caracteres inválidos ({bad_txt})",
                    "method": None, "candidates": [], "is_dup": False,
                    "name": name, "row": i + 2,
                })
                continue
            if i >= len(self._photos):
                plan.append({
                    "src": None, "new": name, "state": "sin_foto",
                    "reason": "no hay fotografía para este registro",
                    "method": None, "candidates": [], "is_dup": False,
                    "name": name, "row": i + 2,
                })
                continue
            photo = self._photos[i]
            new_full = name + photo.suffix if self.keep_extension else name
            is_dup = seen_dest.get(new_full, 0) > 0
            seen_dest[new_full] = seen_dest.get(new_full, 0) + 1
            dest = photo.parent / new_full
            if is_dup:
                state, reason = "duplicado", "el nombre destino ya está usado en el lote"
            elif self._same_file(photo, dest):
                state, reason = "ya_correcto", "el archivo ya tiene este nombre"
            elif dest.exists():
                state, reason = "conflicto", "el destino ya existe en la carpeta"
            else:
                state, reason = "ok", ""
            plan.append({
                "src": photo, "new": new_full, "state": state, "reason": reason,
                "method": None, "candidates": [], "is_dup": is_dup,
                "name": name, "row": i + 2,
            })
        self._plan = plan
        return self._enable_batch_swaps(plan)

    @staticmethod
    def _enable_batch_swaps(plan: list[dict]) -> list[dict]:
        """
        Clasifica las colisiones de destino del lote:

        - Destino es un archivo EXTERNO (ya existe y NO se renombra en este
          lote) → estado "existe": jamás se sobreescribe (bloquea).
        - Destino es un archivo del MISMO lote → se resuelve con el
          renombrado en dos fases (FASE 1: temporales) → "ok".

        Cada foto del lote aparece una sola vez (out-degree ≤ 1): el grafo
        src→dest es un conjunto de cadenas y ciclos. Una cadena termina en
        una fila "ok" (destino libre) o en un archivo externo (irresoluble);
        un ciclo puro (A→B, B→A; A→B→C→A) es siempre resoluble con las dos
        fases. Se recorre con memoización + detección de ciclo.
        """
        rows = [it for it in plan
                if it["src"] and it["state"] in ("ok", "conflicto")]
        by_src: dict[Path, dict] = {it["src"]: it for it in rows}
        memo: dict[Path, bool] = {}

        def movable(src: Path, path: frozenset) -> bool:
            if src in memo:
                return memo[src]
            it = by_src.get(src)
            if it is None:
                memo[src] = True      # destino libre (no es un archivo del lote)
                return True
            if it["state"] == "ok":
                memo[src] = True
                return True
            if src in path:
                return True           # ciclo puro → resoluble en dos fases
            dest = it["src"].parent / it["new"]
            if dest not in by_src:
                memo[src] = False     # archivo externo → no se toca
                return False
            val = movable(dest, path | {src})
            memo[src] = val
            return val

        for it in rows:
            if it["state"] != "conflicto":
                continue
            dest = it["src"].parent / it["new"]
            if dest not in by_src and dest.exists():
                it["state"] = "existe"
                it["reason"] = "el destino ya existe y NO se renombra en este lote"
            elif movable(it["src"], frozenset()):
                it["state"] = "ok"
                it["reason"] = "destino renombrado en el mismo lote (swap)"
        return plan

    def _build_plan_matching(self) -> list[dict]:
        """
        Plan por MATCHING SEGURO: cada nombre del Excel busca SU fotografía.

        Reglas de seguridad:
        - "not_found"  → no se renombra nada.
        - "ambiguo"    → no se elige candidato; se listan todos.
        - reuso (un archivo ya emparejado por otra fila) → "duplicado".
        - Prohibido el fallback posicional silencioso.
        """
        folder = self.folder_path
        matcher = self._get_matcher()
        if not folder or not folder.is_dir() or matcher is None:
            return self._error_plan(
                "no se puede emparejar: carpeta o motor de matching no "
                "disponibles")

        used: set[str] = set()
        seen_dest: dict[str, int] = {}
        plan: list[dict] = []

        for i, name in enumerate(self._names):
            row_no = i + 2   # +1 encabezado, +1 índice 0→1 (convención de load_names)
            bad = _invalid_name_chars(name)
            if bad:
                bad_txt = " ".join(bad)
                plan.append({
                    "src": None, "new": name, "state": "error",
                    "reason": f"el nombre contiene caracteres inválidos ({bad_txt})",
                    "method": "invalid", "candidates": [], "is_dup": False,
                    "name": name, "row": row_no,
                })
                continue
            try:
                path, status, candidates = matcher.find_image_ex(name, str(folder))
            except Exception:
                path, status, candidates = None, "not_found", []

            if status != "ok":
                plan.append({
                    "src": None, "new": name,
                    "state": "ambiguo" if status == "ambiguous" else "not_found",
                    "reason": ("varios archivos compiten por la misma clave "
                               "(no se elige ninguno)" if status == "ambiguous"
                               else "no se encontró ninguna fotografía para este nombre"),
                    "method": status, "candidates": candidates, "is_dup": False,
                    "name": name, "row": row_no,
                })
                continue

            src = Path(path)
            src_key = str(src)
            if src_key in used:
                plan.append({
                    "src": src, "new": name + src.suffix, "state": "duplicado",
                    "reason": "la fotografía ya está emparejada con otra fila (reuso)",
                    "method": "reuso", "candidates": [], "is_dup": True,
                    "name": name, "row": row_no,
                })
                continue
            used.add(src_key)

            new_full = name + src.suffix if self.keep_extension else name
            is_dup = seen_dest.get(new_full, 0) > 0
            seen_dest[new_full] = seen_dest.get(new_full, 0) + 1
            dest = src.parent / new_full
            if is_dup:
                state, reason = "duplicado", "el nombre destino ya está usado en el lote"
            elif self._same_file(src, dest):
                state, reason = "ya_correcto", "el archivo ya tiene este nombre"
            elif dest.exists():
                state, reason = "conflicto", "el destino ya existe en la carpeta"
            else:
                state, reason = "ok", ""
            plan.append({
                "src": src, "new": new_full, "state": state, "reason": reason,
                "method": "ok", "candidates": [], "is_dup": is_dup,
                "name": name, "row": row_no,
            })

        self._plan = plan
        return self._enable_batch_swaps(plan)

    @staticmethod
    def _skip_text(item: dict) -> str:
        """Mensaje legible para filas omitidas (nunca destructivo)."""
        src = item["src"]
        src_name = src.name if src else "—"
        state = item["state"]
        if state == "not_found":
            return f"Omitida: no se encontró fotografía para «{item['name']}»"
        if state == "sin_foto":
            return f"Omitida (sin fotografía): «{item['name']}» no tiene imagen"
        if state == "existe":
            return (f"Omitida (ya existe): {src_name} → {item['new']} "
                    f"({item.get('reason') or 'el destino ya existe y no se toca'})")
        if state == "ambiguo":
            cands = ", ".join(Path(c).name for c in item.get("candidates") or [])
            return (f"Omitida (ambigua): «{item['name']}» coincide con varios archivos "
                    f"({cands}); no se eligió ninguno")
        if state == "error":
            return f"Omitida (error): {item.get('reason') or 'matching no disponible'}"
        if state == "duplicado":
            return f"Omitida (duplicado): {src_name} → {item['new']} ({item.get('reason') or 'nombre ya usado'})"
        if state == "conflicto":
            return f"Omitida (conflicto): {src_name} → {item['new']} ({item.get('reason') or 'el destino ya existe'})"
        return f"Omitida: {src_name} → {item['new']}"

    # ── validación de bloqueo (única fuente para UI y on_rename) ──────────
    _BLOCKING_STATES = ("existe", "conflicto", "duplicado", "not_found",
                        "sin_foto", "ambiguo", "error")

    def rename_blocked(self, plan: Optional[list[dict]] = None) -> tuple[bool, str]:
        """
        Devuelve (bloqueado, motivo). La UI deshabilita «Renombrar todo» y
        on_rename vuelve a comprobar aquí (también cubre los atajos Ctrl+Enter).

        - Modo posicional: bloquea si el conteo no cuadra O hay cualquier
          fila en estado bloqueante (conflicto, duplicado, sin fotografía…).
        - Modo matching seguro: bloquea solo si hay conflictos reales
          (conflicto/duplicado/ambiguo/error); las filas "not_found" se
          omiten sin bloquear (cada nombre busca SU foto, se renombra lo
          que coincide con seguridad).

        ``plan`` opcional: permite evaluar un plan recién calculado (p. ej.
        desde un hilo de fondo) sin depender de ``self._plan``.
        """
        if not (self._photos and self._names):
            return True, "No hay fotografías y nombres cargados."
        if not self.matching_mode and len(self._photos) != len(self._names):
            return True, (f"El conteo no coincide: {len(self._photos)} "
                          f"fotografías vs {len(self._names)} registros.")
        plan = plan if plan is not None else (self._plan if self._plan
                                              else self._build_plan())
        if not plan:
            return True, "No hay correspondencias que renombrar."
        if self.matching_mode:
            bad = [p for p in plan if p["state"] in
                   ("conflicto", "duplicado", "ambiguo", "error")]
        else:
            bad = [p for p in plan if p["state"] in self._BLOCKING_STATES]
        if bad:
            estados = ", ".join(sorted({p["state"] for p in bad}))
            return True, (f"Existen {len(bad)} fila(s) con estado "
                          f"bloqueante ({estados}).")
        return False, ""

    # ── renombramiento ─────────────────────────────────────────────────────
    @staticmethod
    def _emit(on_log: Optional[Callable[[str], None]], line: str) -> None:
        """Emite una línea al panel de registro si el callback está presente."""
        if on_log:
            try:
                on_log(line)
            except Exception:
                pass

    def rename_all(
        self,
        on_progress: Callable[[int, int, str], None],
        on_done: Callable[[int, list[str]], None],
        cancel_ev: Optional[threading.Event] = None,
        copy_mode: bool = False,
        plan: Optional[list[dict]] = None,
        on_log: Optional[Callable[[str], None]] = None,
    ) -> None:
        """Renombra (o copia) archivos en un hilo secundario.

        Reglas de seguridad (obligatorias):
        - "ya_correcto": el destino es EL MISMO archivo → no se toca, sin error.
        - "conflicto": el destino existe y es OTRO archivo → NO se sobreescribe,
          NO se elimina, NO se modifica (en ningún modo, incluido copia).
        - "not_found" / "ambiguo" / "duplicado" / "error" / "sin_foto": se
          omiten, nunca se renombra.
        - Renombrado en DOS FASES para evitar colisiones (a→b mientras
          b→a): FASE 1 mueve/copia a nombres temporales únicos, FASE 2
          convierte temporales en definitivos. Si algo falla a mitad de
          camino, los temporales se revierten a su origen (jamás se dejan
          huérfanos) y se registra exactamente lo que se ejecutó.
        - Re-chequeo TOCTOU justo antes de cada toque a disco.
        """
        if plan is None:
            plan = self._build_plan()
        total   = len(plan) or 1
        success = 0
        errors: list[str] = []
        batch:  list[tuple[Path, Path]] = []
        seen_dest: dict[str, int] = {}

        dest_folder: Optional[Path] = None
        if copy_mode and self._photos:
            dest_folder = self._photos[0].parent / "Renombradas"
            try:
                dest_folder.mkdir(exist_ok=True)
            except OSError as exc:
                on_done(0, [f"No se pudo crear carpeta de copias: {exc}"])
                return

        # ── pre-selección: qué filas se ejecutan de verdad ─────────────────
        skip_states = self._BLOCKING_STATES
        # Orígenes que se renombran en ESTE lote: sus destinos quedan libres
        # en la FASE 1 (temporales), por lo que no son conflictos reales.
        work_srcs = {item["src"] for item in plan
                     if item["state"] == "ok" and item["src"]}
        work: list[tuple[int, Path, Path, dict]] = []
        for i, item in enumerate(plan):
            if cancel_ev and cancel_ev.is_set():
                errors.append("Cancelado por el usuario.")
                break
            src      = item["src"]
            new_full = item["new"]
            state    = item["state"]

            if state in skip_states:
                errors.append(self._skip_text(item))
                self._emit(on_log, f"— {src.name if src else '—'} → {new_full}  [{state}]")
                on_progress(i + 1, total, f"[{state}] {new_full}")
                continue

            if state == "ya_correcto":
                success += 1
                self._emit(on_log, f"✓ {src.name}  (ya correcto)")
                on_progress(i + 1, total, f"[ya correcto] {new_full}")
                continue

            if src is None:
                continue

            if new_full in seen_dest:
                errors.append(f"Saltado duplicado: {src.name} → {new_full} (nombre ya usado)")
                log.warning("DUP  %s → %s (skip)", src.name, new_full)
                self._emit(on_log, f"— {src.name} → {new_full}  [duplicado]")
                on_progress(i + 1, total, f"[duplicado] {new_full}")
                continue
            seen_dest[new_full] = 1

            dest = (dest_folder / new_full) if copy_mode and dest_folder \
                   else (src.parent / new_full)
            if dest.exists():
                if self._same_file(dest, src):
                    success += 1   # el destino es el propio origen → ya correcto
                    self._emit(on_log, f"✓ {src.name}  (ya correcto)")
                    on_progress(i + 1, total, f"[ya correcto] {new_full}")
                    continue
                if dest in work_srcs:
                    # El destino se renombra en este mismo lote: la FASE 1 lo
                    # libera antes de la FASE 2 (swap A→B, B→C).
                    log.info("SWAP  %s → %s (destino en lote)", src.name, new_full)
                    work.append((i, src, dest, item))
                    continue
                errors.append(f"Saltado: {src.name} → {new_full} (el destino ya existe, no se sobreescribe)")
                log.warning("CONFLICT  %s → %s (exists)", src.name, new_full)
                self._emit(on_log, f"— {src.name} → {new_full}  [conflicto]")
                on_progress(i + 1, total, f"[conflicto] {new_full}")
                continue
            work.append((i, src, dest, item))

        # ── FASE 1: mover/copiar a nombres temporales únicos ───────────────
        staged: list[tuple[int, Path, Path, Path, dict]] = []
        for j, (plan_i, src, dest, item) in enumerate(work):
            if cancel_ev and cancel_ev.is_set():
                errors.append("Cancelado por el usuario.")
                break
            new_full = item["new"]
            temp = dest.parent / f".metatag_tmp_{uuid.uuid4().hex[:10]}{dest.suffix}"
            try:
                if copy_mode:
                    shutil.copy2(src, temp)
                else:
                    src.rename(temp)
                staged.append((plan_i, temp, dest, src, item))
            except PermissionError:
                errors.append(f"{src.name} → {new_full}  (permiso denegado)")
                log.error("PERM  %s", src.name)
                self._emit(on_log, f"✕ No fue posible renombrar {src.name}")
            except OSError as exc:
                errors.append(f"{src.name} → {new_full}  ({exc})")
                log.error("ERR %s", exc)
                self._emit(on_log, f"✕ No fue posible renombrar {src.name}")
            on_progress(plan_i + 1, total, new_full)

        # ── FASE 2: temporales → definitivos ───────────────────────────────
        for plan_i, temp, dest, src, item in staged:
            new_full = item["new"]
            if cancel_ev and cancel_ev.is_set():
                self._rollback_temp(temp, src, copy_mode)
                errors.append("Cancelado por el usuario.")
                log.warning("ROLLBACK  %s (cancelado)", src.name)
                self._emit(on_log, f"— {src.name} revertido (cancelado)")
                on_progress(plan_i + 1, total, f"[cancelado] {new_full}")
                continue
            if dest.exists():
                # Otro archivo ocupó la ruta tras la FASE 1: no se sobreescribe.
                self._rollback_temp(temp, src, copy_mode)
                errors.append(f"Saltado: {src.name} → {new_full} (el destino ya existe, no se sobreescribe)")
                log.warning("CONFLICT  %s → %s (exists)", src.name, new_full)
                self._emit(on_log, f"✕ No fue posible renombrar {src.name} (el destino ya existe)")
                on_progress(plan_i + 1, total, f"[conflicto] {new_full}")
                continue
            try:
                temp.rename(dest)
                batch.append((dest, src))
                success += 1
                log.info("OK  %s → %s", src.name, new_full)
                self._emit(on_log, f"✓ {src.name} → {new_full}")
            except PermissionError:
                self._rollback_temp(temp, src, copy_mode)
                errors.append(f"{src.name} → {new_full}  (permiso denegado)")
                log.error("PERM  %s", src.name)
                self._emit(on_log, f"✕ No fue posible renombrar {src.name}")
            except OSError as exc:
                self._rollback_temp(temp, src, copy_mode)
                errors.append(f"{src.name} → {new_full}  ({exc})")
                log.error("ERR %s", exc)
                self._emit(on_log, f"✕ No fue posible renombrar {src.name}")
            on_progress(plan_i + 1, total, new_full)

        if batch:
            self._undo_stack.append((batch, dest_folder, copy_mode))
        on_done(success, errors)

    @staticmethod
    def _rollback_temp(temp: Path, src: Path, copy_mode: bool) -> None:
        """Revierte un temporal a su origen (nunca deja archivos huérfanos)."""
        try:
            if copy_mode:
                if temp.exists():
                    temp.unlink()
            else:
                if temp.exists():
                    temp.rename(src)
        except OSError as exc:
            log.error("Rollback fallido para %s: %s", temp, exc)

    # ── deshacer ───────────────────────────────────────────────────────────
    def undo_last(
        self,
        on_progress: Callable[[int, int, str], None],
        on_done: Callable[[int, list[str]], None],
        on_log: Optional[Callable[[str], None]] = None,
    ) -> None:
        """Revierte el último lote.

        Seguridad: antes de restaurar a la ruta original se comprueba que esa
        ruta no haya sido ocupada por OTRO archivo creado después del
        renombrado. Si existe otro archivo → NO se sobreescribe: se registra
        conflicto de undo y se preservan ambos.
        """
        if not self._undo_stack:
            on_done(0, ["No hay nada que deshacer."])
            return
        batch, dest_folder, copy_mode = self._undo_stack.pop()
        total   = len(batch)
        success = 0
        errors: list[str] = []
        for i, (current, original) in enumerate(batch):
            try:
                if not current.exists():
                    errors.append(f"Archivo no encontrado: {current.name}")
                    self._emit(on_log, f"— no se encontró {current.name}")
                elif copy_mode:
                    current.unlink()
                    success += 1
                    self._emit(on_log, f"↩ eliminada copia {current.name}")
                else:
                    if original.exists():
                        if self._same_file(original, current):
                            success += 1   # ya restaurado (mismo archivo)
                            self._emit(on_log, f"↩ {current.name} ya estaba restaurado")
                        else:
                            errors.append(
                                f"Conflicto al deshacer: no se sobreescribe "
                                f"{original.name} (hay otro archivo en esa ruta)")
                            self._emit(on_log, f"✕ conflicto al restaurar {original.name}")
                    else:
                        current.rename(original)
                        success += 1
                        self._emit(on_log, f"↩ {current.name} → {original.name}")
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

    # ── exportar log ───────────────────────────────────────────────────────
    def export_log(self, pairs: list[tuple[str, str, Path, bool, str]], dest: Path) -> None:
        with dest.open("w", encoding="utf-8") as f:
            f.write("=" * 62 + "\n")
            f.write("  LOG DE RENOMBRAMIENTO — Renombrador de Fotos v4\n")
            f.write(f"  Fecha   : {datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}\n")
            f.write(f"  Carpeta : {self.folder_path}\n")
            f.write("=" * 62 + "\n\n")
            for orig, new, _, _, state in pairs:
                label = f"  [{state}] " if state and state != "ok" else "  "
                f.write(f"{label}{orig}  →  {new}\n")
            f.write(f"\n  Total: {len(pairs)} archivo(s)\n")

    def export_preview_csv(self, pairs: list[tuple[str, str, Path, bool, str]], dest: Path) -> None:
        with dest.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["indice", "estado", "original", "nuevo_nombre",
                        "ruta_original", "ruta_destino", "resultado"])
            for i, (orig, new, path, is_dup, state) in enumerate(pairs, start=1):
                ruta_orig = str(path) if path is not None else ""
                ruta_dest = (str(path.parent / new) if path is not None else "")
                resultado = ("Sí" if is_dup else
                             ("✓" if state in ("ok", "ya_correcto") else state))
                w.writerow([i, state, orig, new, ruta_orig, ruta_dest, resultado])

    def write_backup(self, dest: Path, batch: list[tuple[Path, Path]],
                     copy_mode: bool) -> None:
        """Escribe el registro JSON de la última operación (original→nuevo).

        ``batch`` es la tupla (destino, origen) del _undo_stack: aquí se
        invierte a (original → nuevo) para que el registro sea legible.
        """
        entries = [
            {
                "original": src.name,
                "nuevo": cur.name,
                "ruta_original": str(src),
                "ruta_destino": str(cur),
                "modo": "copia" if copy_mode else "renombrado",
            }
            for cur, src in batch
        ]
        payload = {
            "creado": datetime.now().isoformat(timespec="seconds"),
            "carpeta": str(self.folder_path or ""),
            "modo": "copia" if copy_mode else "renombrado",
            "archivos": entries,
        }
        dest.write_text(json.dumps(payload, ensure_ascii=False, indent=2),
                        encoding="utf-8")


# ===========================================================================
#  SCROLL SUAVE
# ===========================================================================
class SmoothScroller:
    """
    Scroll con inercia para CTkScrollableFrame.
    
    - Activa el listener SOLO mientras el cursor está dentro (bind <Enter>/<Leave>)
      → evita interferir con otros widgets.
    - Usa yview_moveto() con interpolación → animación fluida, no brusca.
    - No recorre los hijos al construirse ni necesita rebind al agregar filas.
    """
    STEP    = 0.03
    EASE    = 0.15
    TICK_MS = 16

    def __init__(self, sf: ctk.CTkScrollableFrame) -> None:
        self._canvas   = sf._parent_canvas
        self._target:  Optional[float] = None
        self._running: bool = False
        self._widget   = sf

        sf.bind("<Enter>", self._arm,   add="+")
        sf.bind("<Leave>", self._disarm, add="+")

    def _arm(self, _=None) -> None:
        self._canvas.bind_all("<MouseWheel>", self._wheel_generic)
        self._canvas.bind_all("<Button-4>",   self._wheel_up)
        self._canvas.bind_all("<Button-5>",   self._wheel_down)

    def _disarm(self, _=None) -> None:
        for ev in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            try:
                self._canvas.unbind_all(ev)
            except Exception:
                pass

    def _wheel_generic(self, e) -> None:
        self._scroll(-1 if e.delta > 0 else 1)

    def _wheel_up(self, _)   -> None: self._scroll(-1)
    def _wheel_down(self, _) -> None: self._scroll(1)

    def _scroll(self, direction: int) -> None:
        top, bot = self._canvas.yview()
        if top <= 0.0 and bot >= 1.0:
            return
        if not self._running or self._target is None:
            self._target = top
        self._target = max(0.0, min(1.0, self._target + direction * self.STEP))
        if not self._running:
            self._running = True
            self._animate()

    def _animate(self) -> None:
        top, _ = self._canvas.yview()
        new_top = top + (self._target - top) * self.EASE
        if abs(self._target - new_top) < 0.0008:
            new_top = self._target
            self._running = False
        else:
            self._widget.after(self.TICK_MS, self._animate)
        self._canvas.yview_moveto(max(0.0, min(1.0, new_top)))


# ===========================================================================
#  TOOLTIP DE IMAGEN  — FIX BUG #1
# ===========================================================================
class ImageTooltip:
    """
    Miniatura flotante al hacer hover sobre una fila de la vista previa.

    Correcciones v4 respecto a v3:
    - Se destruye al salir del WIDGET PADRE (no del tooltip mismo, lo que
      causaba que quedara "pegado" porque el Leave del Toplevel disparaba
      antes de que el cursor volviera al padre).
    - Usa after() para mostrar (delay 300 ms) → evita parpadeos al pasar
      el cursor rápido entre filas.
    - Destrucción garantizada con protocol WM_DELETE_WINDOW y bind <Destroy>.
    """

    DELAY_MS = 300

    # Seguro global: garantiza que NUNCA pueda haber más de un tooltip
    # flotante vivo al mismo tiempo, sin importar qué secuencia rara de
    # eventos ocurra (scroll rápido, filtrado, cambio de tema, etc.).
    _ACTIVE: "list[ctk.CTkToplevel]" = []

    @classmethod
    def _close_active(cls) -> None:
        while cls._ACTIVE:
            tip = cls._ACTIVE.pop()
            try:
                tip.destroy()
            except Exception:
                pass

    def __init__(self, widget: ctk.CTkFrame, path: Path) -> None:
        self._widget = widget
        self._path   = path
        self._tip:   Optional[ctk.CTkToplevel] = None
        self._job:   Optional[str] = None

        widget.bind("<Enter>",   self._schedule_show, add="+")
        widget.bind("<Leave>",   self._cancel_and_hide, add="+")
        widget.bind("<Destroy>", self._cancel_and_hide, add="+")

    def _schedule_show(self, _=None) -> None:
        self._cancel_job()
        self._job = self._widget.after(self.DELAY_MS, self._show)

    def _cancel_and_hide(self, _=None) -> None:
        self._cancel_job()
        self._destroy_tip()

    def _cancel_job(self) -> None:
        _safe_cancel_after(self._widget, self._job)
        self._job = None

    def _show(self) -> None:
        self._job = None
        if self._tip or not self._path.exists():
            return
        ImageTooltip._close_active()   # por si quedó alguno de otra fila, cerrarlo primero
        ctk_img = _get_thumb(self._path, (180, 180))
        if not ctk_img:
            return
        try:
            tip = ctk.CTkToplevel(self._widget.winfo_toplevel())
            tip.wm_overrideredirect(True)
            tip.configure(fg_color=C["surface"])
            tip.attributes("-topmost", True)
            ctk.CTkLabel(tip, image=ctk_img, text="",
                         fg_color=C["surface"]).pack(padx=6, pady=6)
            _place_tip_near_pointer(tip, self._widget, offset=18)
            tip.protocol("WM_DELETE_WINDOW", self._destroy_tip)
            self._tip = tip
            ImageTooltip._ACTIVE.append(tip)
        except Exception as exc:
            log.debug("Tooltip error: %s", exc)

    def _destroy_tip(self) -> None:
        if self._tip:
            try:
                if self._tip in ImageTooltip._ACTIVE:
                    ImageTooltip._ACTIVE.remove(self._tip)
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


# ===========================================================================
#  COMPONENTES UI
# ===========================================================================
class ToolTip:
    """Tooltip simple para widgets."""

    def __init__(self, widget, text: str) -> None:
        self._widget = widget
        self._text = text
        self._tip: Optional[ctk.CTkToplevel] = None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def _show(self, _=None) -> None:
        if self._tip:
            return
        try:
            tip = ctk.CTkToplevel(self._widget)
            tip.wm_overrideredirect(True)
            tip.configure(fg_color=C["surface2"])
            ctk.CTkLabel(tip, text=self._text, font=FONT_SM,
                         text_color=C["text"], fg_color="transparent",
                         wraplength=250).pack(padx=8, pady=4)
            x = self._widget.winfo_rootx() + self._widget.winfo_width() // 2
            y = self._widget.winfo_rooty() + self._widget.winfo_height() + 5
            tip.geometry(f"+{x}+{y}")
            self._tip = tip
        except Exception:
            pass

    def _hide(self, _=None) -> None:
        if self._tip:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


class StatusBadge(ctk.CTkLabel):
    """Pastilla circular de estado (idle / ok / warn / error / loading)."""

    _S: dict[str, tuple[str, str, str]] = {
        "idle":    ("surface2", "subtext",  "—"),
        "ok":      ("green",    "bg",       "✓"),
        "warn":    ("yellow",   "bg",       "!"),
        "error":   ("red",      "bg",       "✗"),
        "loading": ("accent",   "bg",       "…"),
    }

    def __init__(self, master, **kw) -> None:
        super().__init__(master, text="—", width=28, height=28,
                         corner_radius=14,
                         font=FONT_SM_BD, **kw)
        self._state = "idle"
        self.set_state("idle")

    def set_state(self, state: str) -> None:
        self._state = state
        bg_k, fg_k, icon = self._S.get(state, self._S["idle"])
        self.configure(fg_color=C[bg_k], text_color=C[fg_k], text=icon)


class Toast(ctk.CTkFrame):
    """Notificación flotante en esquina inferior-derecha. Auto-destruye en 3.5 s."""

    _ACTIVE: "list[Toast]" = []

    def __init__(self, master, message: str, kind: str = "ok") -> None:
        _color = {"ok": C["green"], "error": C["red"], "warn": C["yellow"]}
        _icon  = {"ok": "✓", "error": "✗", "warn": "!"}
        border = _color.get(kind, C["accent"])
        super().__init__(master, fg_color=C["surface"], corner_radius=12,
                         border_width=1, border_color=border)
        ctk.CTkLabel(self, text=_icon.get(kind, "•"), width=26,
                     font=FONT_LG_BD,
                     text_color=border, fg_color="transparent"
                     ).pack(side="left", padx=(12, 4), pady=10)
        ctk.CTkLabel(self, text=message,
                     font=FONT_SM,
                     text_color=C["text"], fg_color="transparent",
                     wraplength=270).pack(side="left", padx=(0, 16), pady=10)

        y_offset = -16
        for t in Toast._ACTIVE:
            try:
                y_offset += t.winfo_reqheight() + 8
            except Exception:
                pass
        self.place(relx=1.0, rely=1.0, anchor="se", x=-16, y=y_offset)
        Toast._ACTIVE.append(self)
        self.after(3500, self._bye)

    def _bye(self) -> None:
        try:
            if self in Toast._ACTIVE:
                Toast._ACTIVE.remove(self)
            self.place_forget(); self.destroy()
        except Exception:
            pass


# ===========================================================================
#  EXPLORADOR DE ARCHIVOS  — FIX BUG #2 (multi-disco)
# ===========================================================================
class FileBrowser(ctk.CTkToplevel):
    """
    Explorador dark integrado que reemplaza el diálogo nativo feo de tkinter.

    Mejoras v4:
    - Selector de unidades/raíces (detecta todos los discos en Windows,
      puntos de montaje en Linux, volúmenes en macOS).
    - Carga de filas en chunks (70/batch, 20 ms entre chunks) → no congela
      aunque haya cientos de archivos.
    - Scroll fluido via SmoothScroller.
    - grab_set() llamado DESPUÉS de deiconify + wait_visibility → sin crash.
    """

    CHUNK = 70

    def __init__(self, master, mode: str = "folder",
                 filetypes: Optional[list[str]] = None,
                 title: str = "Seleccionar") -> None:
        super().__init__(master)
        self.transient(master)
        self.withdraw()
        self.title(title)
        self.configure(fg_color=C["bg"])
        self.resizable(True, True)

        self._mode      = mode
        self._exts      = frozenset(filetypes or [".xlsx"])
        self._result:   Optional[str] = None
        self._current   = Path.home()
        self._pending:  list[Path] = []
        self._chunk_job: Optional[str] = None
        self._selected: Optional[Path] = None

        self._drives = _detect_drives()

        self._build()
        self._resize_to_screen(master)
        _show_toplevel(self)
        self._navigate(self._current)

    def _resize_to_screen(self, master) -> None:
        """Ajusta tamaño y posición al padre / pantalla."""
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = max(560, min(720, int(sw * 0.50)))
        h  = max(380, min(540, int(sh * 0.55)))
        px = master.winfo_rootx() + (master.winfo_width()  - w) // 2
        py = master.winfo_rooty() + (master.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{px}+{py}")
        self.minsize(480, 340)

    # ── construcción UI ────────────────────────────────────────────────────
    def _build(self) -> None:
        # barra superior
        top = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=0, height=50)
        top.pack(fill="x"); top.pack_propagate(False)

        ctk.CTkButton(top, text="↑ Subir", width=80, height=30,
                      **BTN_SECONDARY,
                      font=FONT_SM,
                      command=self._go_up).pack(side="left", padx=(10, 4), pady=10)

        # selector de disco/raíz (FIX #2): solo se muestra si hay más de uno
        if len(self._drives) > 1:
            self._drive_var = ctk.StringVar(value=self._drives[0])
            _make_option_menu(
                top, self._drive_var, self._drives, width=130, height=30,
                font=FONT_XS_SM, fg=C["surface2"], btn_fg=C["surface3"],
                command=lambda v: self._navigate(Path(v))
            ).pack(side="left", padx=4, pady=10)
        else:
            # un solo disco disponible: se usa directamente, sin mostrar selector
            self._drive_var = ctk.StringVar(value=self._drives[0] if self._drives else "C:\\")

        self._path_var = ctk.StringVar()
        pe = ctk.CTkEntry(top, textvariable=self._path_var, height=30,
                          font=FONT_SM,
                          fg_color=C["surface2"], border_color=C["border"],
                          text_color=C["text"])
        pe.pack(side="left", fill="x", expand=True, padx=4, pady=10)
        pe.bind("<Return>", lambda _: self._navigate(Path(self._path_var.get())))

        ctk.CTkButton(top, text="Ir", width=40, height=30,
                      **BTN_SECONDARY,
                      font=FONT_SM,
                      command=lambda: self._navigate(Path(self._path_var.get()))
                      ).pack(side="left", padx=(0, 10), pady=10)

        # lista scrollable
        lf = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=0)
        lf.pack(fill="both", expand=True)
        self._sf = ctk.CTkScrollableFrame(lf, fg_color="transparent",
                                          scrollbar_button_color=C["surface2"])
        self._sf.pack(fill="both", expand=True)
        SmoothScroller(self._sf)

        # barra inferior
        bot = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=0, height=52)
        bot.pack(fill="x", side="bottom"); bot.pack_propagate(False)

        self._lbl_sel = ctk.CTkLabel(bot, text="Nada seleccionado",
                                     font=FONT_SM,
                                     text_color=C["subtext"], fg_color="transparent")
        self._lbl_sel.pack(side="left", padx=14)

        ctk.CTkButton(bot, text="Cancelar", width=88, height=32,
                      **BTN_DANGER,
                      font=FONT_SM,
                      command=self.destroy).pack(side="right", padx=(4, 10), pady=10)

        self._btn_ok = ctk.CTkButton(
            bot, text="Seleccionar", width=110, height=32,
            **BTN_PRIMARY,
            text_color=C["accent_text"], font=FONT_SM_BD,
            command=self._confirm, state="disabled")
        self._btn_ok.pack(side="right", padx=4, pady=10)

    # ── navegación ─────────────────────────────────────────────────────────
    def _navigate(self, path: Path) -> None:
        try:
            valid = path.is_dir()
        except OSError:
            self._info(f"No se puede acceder a «{path}».\n¿La unidad está conectada y lista?")
            return
        if not valid:
            return
        self._current = path
        self._path_var.set(str(path))
        self._selected = None
        self._lbl_sel.configure(text="Nada seleccionado")
        self._btn_ok.configure(state="disabled")
        self._cancel_chunk()
        for w in self._sf.winfo_children():
            w.destroy()

        if self._mode == "folder":
            self._add_row("▸  (esta carpeta)", path, is_dir=True, current=True)

        try:
            entries = sorted(path.iterdir(),
                             key=lambda p: (not p.is_dir(), p.name.lower()))
        except OSError as exc:
            self._info(f"No se pudo leer esta ubicación: {exc}"); return

        self._pending = [
            e for e in entries
            if not e.name.startswith(".")
            and (e.is_dir() or
                 (self._mode == "file" and e.suffix.lower() in self._exts))
        ]
        self._load_chunk()

    def _load_chunk(self) -> None:
        chunk, self._pending = self._pending[:self.CHUNK], self._pending[self.CHUNK:]
        for e in chunk:
            label = f"▸  {e.name}" if e.is_dir() else f"   {e.name}"
            self._add_row(label, e, is_dir=e.is_dir())
        if self._pending:
            self._chunk_job = self.after(20, self._load_chunk)

    def _cancel_chunk(self) -> None:
        _safe_cancel_after(self, self._chunk_job)
        self._chunk_job = None
        self._pending = []

    def _add_row(self, label: str, path: Path,
                 is_dir: bool, current: bool = False) -> None:
        idx = len(self._sf.winfo_children())
        bg  = C["surface2"] if current else (C["surface"] if idx % 2 == 0 else C["bg"])

        row = ctk.CTkFrame(self._sf, fg_color=bg, corner_radius=3, height=34)
        row.pack(fill="x", pady=1, padx=3); row.pack_propagate(False)

        lbl = ctk.CTkLabel(row, text=label, anchor="w",
                           font=FONT_SM,
                           text_color=C["accent"] if current else C["text"],
                           fg_color="transparent")
        lbl.pack(side="left", fill="x", expand=True, padx=10)

        def _click(p=path, d=is_dir, cur=current):
            if d and not cur:
                self._navigate(p)
            else:
                self._select(p)

        for w in (row, lbl):
            w.bind("<Button-1>", lambda _, f=_click: f())
            w.bind("<Enter>",    lambda _, r=row: r.configure(fg_color=C["surface3"]))
            w.bind("<Leave>",    lambda _, r=row, b=bg: r.configure(fg_color=b))

    def _info(self, msg: str) -> None:
        ctk.CTkLabel(self._sf, text=msg, fg_color="transparent",
                     font=FONT_SM,
                     text_color=C["subtext"]).pack(pady=16)

    def _select(self, path: Path) -> None:
        self._selected = path
        self._lbl_sel.configure(text=f"✓  {path.name or str(path)}")
        self._btn_ok.configure(state="normal")

    def _go_up(self) -> None:
        self._navigate(self._current.parent)

    def _confirm(self) -> None:
        if self._selected:
            self._result = str(self._selected)
        self.grab_release()
        self.destroy()

    def get_result(self) -> Optional[str]:
        self.wait_window()
        return self._result


# ===========================================================================
#  PATH SELECTOR
# ===========================================================================
class PathSelector(ctk.CTkFrame):
    """Entrada de ruta + botón Explorar → abre FileBrowser."""

    def __init__(self, master, placeholder: str,
                 mode: str = "folder",
                 filetypes: Optional[list[str]] = None,
                 on_change: Optional[Callable[[str], None]] = None,
                 **kw) -> None:
        super().__init__(master, fg_color="transparent", **kw)
        self._mode      = mode
        self._exts      = filetypes or [".xlsx"]
        self._on_change = on_change
        self.columnconfigure(0, weight=1)

        self._entry = ctk.CTkEntry(
            self, placeholder_text=placeholder, height=38,
            font=FONT_MD,
            fg_color=C["surface"], border_color=C["border"],
            text_color=C["text"])
        self._entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._entry.bind("<Return>", lambda _: self._notify())

        ctk.CTkButton(self, text="Explorar", width=118, height=38,
                      font=FONT_MD,
                      **BTN_SECONDARY,
                      command=self._browse).grid(row=0, column=1)

    def _browse(self) -> None:
        title = ("Selecciona la carpeta de fotos" if self._mode == "folder"
                 else "Selecciona el archivo Excel")
        br = FileBrowser(self.winfo_toplevel(),
                         mode=self._mode, filetypes=self._exts, title=title)
        r = br.get_result()
        if r:
            self.set(r)
            self._notify()

    def _notify(self) -> None:
        if self._on_change:
            self._on_change(self.get())

    def get(self) -> str:
        return self._entry.get().strip()

    def set(self, v: str) -> None:
        self._entry.delete(0, "end")
        self._entry.insert(0, v)


# ===========================================================================
#  VISTA PREVIA — FIX BUG #4 (lag al scrollear)
# ===========================================================================
class PreviewTable(ctk.CTkFrame):
    """
    Tabla de vista previa VIRTUALIZADA.

    Reemplaza el render widget-per-row (chunks asíncronos) por un viewport
    de altura FIJA con tk.Canvas + tk.Scrollbar vertical propios:

        _all_pairs  → modelo (fuente de verdad, N dicts, intacto)
        _filtered   → índices de _all_pairs que pasan el filtro
        _first      → primera fila lógica materializada
        pool (_rows)→ O(viewport) slots reutilizables sobre el canvas

    El número de widgets vivos NO depende de N. El scroll desplaza el
    canvas y recicla slots (solo ``configure``), sin destruir/crear filas.
    La edición escribe en el modelo; el slot es solo un espejo. Rueda del
    ratón sobre la tabla desplaza SOLO la tabla (devuelve ``"break"`` para
    frenar el ``bind_all`` del SmoothScroller exterior).

    API pública conservada: render, filter, _apply_filter, update_dup_states,
    set_edit_mode, _cancel_jobs, _configure_grid, _build_header, _rows.
    """

    ROW_H = PROFILE.table_row_h
    BUFFER = 5
    _MIN_H = 240
    _MAX_H = 900

    # Configuración de columnas maestra (sincronizada header ↔ filas)
    _COL_NUM_W   = 36    # Columna # (índice)
    _COL_STATE_W = 96    # Columna Estado (ancho fijo)
    _COL_ORIG_W  = None  # Columna Original (dinámica, weight=1)
    _COL_NEW_W   = None  # Columna Nuevo nombre (dinámica, weight=1)

    def __init__(self, master, on_name_change: Optional[Callable[[int, str, Path], None]] = None,
                 on_filter: Optional[Callable[[int], None]] = None, **kw) -> None:
        super().__init__(master, fg_color=C["surface"], corner_radius=10, **kw)
        self._all_pairs: list[dict[str, object]] = []
        self._filtered: list[int] = []
        self._first: int = 0
        self._rows: list[dict[str, object]] = []      # pool de slots
        self._selected: set[int] = set()              # índices en _all_pairs
        self._edit_mode: bool = False
        self._on_name_change = on_name_change
        self._on_filter = on_filter
        self._current_query: str = ""
        self._filter_job: Optional[str] = None
        self._syncing: bool = False
        self._hdr: Optional[ctk.CTkFrame] = None
        self._cv: Optional[tk.Canvas] = None
        self._sb: Optional[tk.Scrollbar] = None
        self._tip: Optional[ctk.CTkToplevel] = None
        self._tip_job: Optional[str] = None
        self._hover_path: Optional[Path] = None

        # La altura la controla el padre (CTkScrollableFrame o similar).
        # En modo CTkScrollableFrame, calculamos altura desde la pantalla.
        sh = master.winfo_screenheight()
        h = max(self._MIN_H, min(self._MAX_H, int(sh * 0.45)))
        self._adaptive_height: int = h
        self.pack_propagate(False)
        self.configure(height=h)

        self._lbl_empty = ctk.CTkLabel(
            self, fg_color="transparent",
            text="Aquí aparecerá la vista previa.\n"
                 "Pasa el cursor sobre una fila para ver la miniatura.",
            font=FONT_MD, text_color=C["subtext"], justify="center")
        self._lbl_empty.pack(pady=28)

        self.bind("<Configure>", self._on_frame_configure)

        # Capturar TODOS los eventos de scroll DENTRO de PreviewTable.
        # Sin esto, los bind_all de CTkScrollableFrame y SmoothScroller
        # capturan scroll sobre widgets sin binding propio (_lbl_empty, _hdr,
        # etc.) causando doble scroll.  "break" detiene la cadena antes de
        # que llegue a bind_all.
        # NOTA: usamos tk.Frame.bind() en lugar de self.bind() porque
        # CTkFrame.bind() redirige a su _canvas interno, no al frame.
        for _ev in ("<Button-4>", "<Button-5>", "<MouseWheel>"):
            tk.Frame.bind(self, _ev, self._on_wheel, add="+")

    # ── API pública ────────────────────────────────────────────────────────
    def render(self, pairs: list[tuple[str, str, Optional[Path], bool, str]],
               empty_message: str = "Sin datos para mostrar.") -> None:
        """Reemplaza el contenido de la tabla. Síncrono y acotado a O(viewport)."""
        self._cancel_jobs()
        self._close_tip()
        self._current_query = ""
        self._all_pairs = [
            {
                "orig": orig, "new": new, "photo_path": photo_path,
                "photo_index": i, "is_dup": is_dup, "state": state,
            }
            for i, (orig, new, photo_path, is_dup, state) in enumerate(pairs)
        ]
        self._filtered = list(range(len(self._all_pairs)))
        self._selected = {s for s in self._selected if s < len(self._all_pairs)}
        self._first = 0
        self._notify_filter()

        if not self._all_pairs:
            self._hide_table()
            self._lbl_empty.configure(text=empty_message)
            self._lbl_empty.pack(pady=24)
            return

        # Altura adaptativa: crece con el contenido (hasta _MAX_H) en vez de
        # mostrar siempre un viewport pequeño con mucho espacio vacío abajo.
        h2 = min(self._MAX_H, max(self._MIN_H, len(self._all_pairs) * self.ROW_H + 34))
        if h2 != self._adaptive_height:
            self._adaptive_height = h2
            self.configure(height=h2)

        self._show_table()
        self._cv.delete("all")
        self._build_pool()
        self._cv.yview_moveto(0)
        self._sync_scrollregion()
        self._sync_viewport()

        # Reconstruir el pool si el canvas aún no estaba mapeado cuando se
        # construyó (winfo_height()=1 → pool fantasma): se reintenta tras el
        # primer layout y de nuevo después de 80 ms por si <Configure> tardó.
        def _check_pool() -> None:
            real_size = self._pool_size()
            if real_size != len(self._rows):
                self._cv.delete("all")
                self._build_pool()
                self._sync_scrollregion()
            self._sync_viewport()
        self.after_idle(_check_pool)
        self.after(80, _check_pool)

    def filter(self, query: str) -> None:
        """Aplica filtro de texto sobre los pares ya cargados (debounced)."""
        _safe_cancel_after(self, self._filter_job)
        self._filter_job = self.after(100, lambda: self._apply_filter(query))

    def _apply_filter(self, query: str) -> None:
        self._filter_job = None
        q = query.lower().strip()
        self._current_query = q
        self._filtered = [
            i for i, row in enumerate(self._all_pairs)
            if not q or q in row["orig"].lower() or q in row.get("new", "").lower()
        ]
        # La selección se conserva; el viewport se reclampa a la nueva altura.
        self._sync_scrollregion()
        self._sync_viewport()
        self._notify_filter()

    def _notify_filter(self) -> None:
        if self._on_filter:
            try:
                self._on_filter(len(self._filtered))
            except Exception:
                pass

    def update_dup_states(self, pairs: list[tuple[str, str, Optional[Path], bool, str]]) -> None:
        """Refresca estados/duplicados en el MODELO. No destruye widgets."""
        n = len(self._all_pairs)
        changed = False
        for i, (_, _, _, is_dup, state) in enumerate(pairs):
            if i >= n:
                break
            row = self._all_pairs[i]
            if row["is_dup"] != is_dup or row["state"] != state:
                row["is_dup"] = is_dup
                row["state"] = state
                changed = True
        if changed:
            for slot in self._rows:
                slot["index"] = -1
            self._sync_viewport()

    def set_edit_mode(self, enabled: bool) -> None:
        """Alterna edición inline SIN reconstruir la tabla (solo slots vivos)."""
        if self._edit_mode == enabled:
            return
        self._edit_mode = enabled
        if not self._all_pairs:
            return
        for slot in self._rows:
            slot["index"] = -1
        self._sync_viewport()

    def _set_row_selected(self, pair_index: int, selected: bool) -> None:
        """Selección por ÍNDICE LÓGICO (no por widget). La UI actual no la
        dispara; queda disponible y sobrevive scroll/filtro/tema/recycling."""
        if selected:
            self._selected.add(pair_index)
        else:
            self._selected.discard(pair_index)
        for slot in self._rows:
            if slot["pair_index"] == pair_index:
                slot["index"] = -1
        self._sync_viewport()

    def _scroll_by(self, dy: int) -> None:
        """Desplaza el viewport dy píxeles y recicla los slots afectados."""
        n = len(self._filtered)
        if n == 0:
            return
        total = n * self.ROW_H
        view_h = self._cv.winfo_height()
        max_px = max(0, total - view_h)
        current_px = self._pixel_offset()
        new_px = max(0, min(max_px, current_px + dy))
        if new_px == current_px:
            return
        self._cv.yview_moveto(new_px / max_px if max_px else 0.0)
        # Tras un scroll la fila bajo el cursor ya no es la misma: cerrar el
        # tooltip evita que quede mostrando la foto de la fila ANTERIOR.
        self._close_tip()
        self._hover_path = None
        self._sync_viewport()

    def _pixel_offset(self) -> int:
        return int(self._cv.canvasy(0))

    # ── internos ──────────────────────────────────────────────────────────
    def _cancel_jobs(self) -> None:
        for attr in ("_filter_job", "_tip_job"):
            _safe_cancel_after(self, getattr(self, attr, None))
            setattr(self, attr, None)

    def _configure_grid(self, frame: ctk.CTkFrame) -> None:
        frame.columnconfigure(0, minsize=self._COL_NUM_W)
        frame.columnconfigure(1, minsize=self._COL_STATE_W)
        frame.columnconfigure(2, weight=1)
        frame.columnconfigure(3, weight=1)

    def _build_header(self) -> None:
        hdr = ctk.CTkFrame(self, fg_color=C["surface2"], corner_radius=0, height=30)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        self._configure_grid(hdr)
        cols = [
            (0, "#",           "w",  4),
            (1, "Estado",      "w",  8),
            (2, "Original",    "w",  8),
            (3, "Nuevo nombre","w",  8),
        ]
        for col, txt, sticky, padx in cols:
            ctk.CTkLabel(hdr, text=txt, width=1,
                         font=FONT_SM_BD, text_color=C["accent"],
                         fg_color="transparent").grid(row=0, column=col,
                                                      sticky=sticky, padx=padx,
                                                      pady=4)
        self._hdr = hdr

    def _show_table(self) -> None:
        if self._lbl_empty.winfo_manager():
            self._lbl_empty.pack_forget()
        if self._hdr is None:
            self._build_header()
        if self._cv is None:
            canvas = tk.Canvas(self, bg=C["surface"], highlightthickness=0,
                               bd=0, yscrollincrement=self.ROW_H)
            sb = tk.Scrollbar(self, orient="vertical", command=self._on_sb,
                              troughcolor=C["surface2"], highlightthickness=0)
            canvas.configure(yscrollcommand=sb.set)
            canvas.bind("<MouseWheel>", self._on_wheel, add="+")
            canvas.bind("<Button-4>", self._on_wheel, add="+")
            canvas.bind("<Button-5>", self._on_wheel, add="+")
            canvas.bind("<Configure>", self._on_canvas_resize, add="+")
            canvas.bind("<Enter>", self._close_tip, add="+")
            canvas.bind("<Leave>", self._close_tip, add="+")
            canvas.bind("<Motion>", self._on_motion, add="+")
            self._cv = canvas
            self._sb = sb
        # Repack en orden canónico (header → scrollbar → canvas).
        for w in (self._hdr, self._sb, self._cv):
            try:
                w.pack_forget()
            except Exception:
                pass
        self._hdr.pack(fill="x")
        self._sb.pack(side="right", fill="y")
        self._cv.pack(side="left", fill="both", expand=True)

    def _hide_table(self) -> None:
        for w in (self._cv, self._sb, self._hdr):
            if w is not None:
                try:
                    w.pack_forget()
                except Exception:
                    pass

    def _build_pool(self) -> None:
        for slot in self._rows:
            try:
                slot["frame"].destroy()
            except Exception:
                pass
        self._rows = []
        for k in range(self._pool_size()):
            slot = self._make_slot(k)
            item = self._cv.create_window(
                0, k * self.ROW_H, window=slot["frame"], anchor="nw",
                height=self.ROW_H)
            slot["item"] = item
            self._rows.append(slot)

    def _pool_size(self) -> int:
        if self._cv is None:
            return 0
        visible = max(1, int(self._view_height() // self.ROW_H) + 2)  # +2 filas parciales
        return min(len(self._filtered), visible + 2 * self.BUFFER)

    def _make_slot(self, k: int) -> dict[str, object]:
        row = ctk.CTkFrame(self._cv, fg_color=C["surface"], corner_radius=0,
                           height=self.ROW_H)
        row.pack_propagate(False)
        self._configure_grid(row)

        num = ctk.CTkLabel(row, text="", width=1, font=FONT_SM,
                           text_color=C["overlay"], fg_color="transparent")
        num.grid(row=0, column=0, sticky="w", padx=4)

        state_lbl = ctk.CTkLabel(row, text="", width=1, anchor="w", font=FONT_SM,
                                 text_color=C["overlay"], fg_color="transparent")
        state_lbl.grid(row=0, column=1, sticky="w", padx=8)

        orig = ctk.CTkLabel(row, text="", width=1, anchor="w", font=FONT_SM,
                            text_color=C["subtext"], fg_color="transparent")
        orig.grid(row=0, column=2, sticky="w", padx=8)

        var = ctk.StringVar()
        new_lbl = ctk.CTkLabel(row, text="", width=1, anchor="w", font=FONT_SM,
                               text_color=C["text"], fg_color="transparent")
        new_lbl.grid(row=0, column=3, sticky="w", padx=8)
        new_entry = ctk.CTkEntry(row, textvariable=var, width=1, height=28,
                                 font=FONT_SM, fg_color=C["surface"],
                                 border_color=C["border"], text_color=C["text"])
        new_entry.grid(row=0, column=3, sticky="ew", padx=8)
        new_entry.grid_remove()

        slot = {
            "frame": row, "num_widget": num, "orig_widget": orig,
            "new_widget": new_lbl,
            "new_entry": new_entry, "new_var": var, "state_widget": state_lbl,
            "index": -1, "pair_index": -1, "item": None,
        }
        # La rueda sobre una fila también debe desplazar SOLO la tabla
        # (las filas son hijas del canvas y capturan el evento).
        row.bind("<MouseWheel>", self._on_wheel, add="+")
        row.bind("<Button-4>", self._on_wheel, add="+")
        row.bind("<Button-5>", self._on_wheel, add="+")
        row.bind("<Enter>", self._close_tip, add="+")
        row.bind("<Leave>", self._close_tip, add="+")
        row.bind("<Motion>", self._on_slot_motion(slot), add="+")
        var.trace_add("write", self._make_trace(slot))
        return slot

    def _make_trace(self, slot: dict[str, object]):
        def _on_change(*_):
            if self._syncing or slot["pair_index"] < 0:
                return
            value = slot["new_var"].get()
            safe = Path(value).name if value else value
            if safe != value:
                slot["new_var"].set(safe)
            self._write_model(slot, safe)
        return _on_change

    def _write_model(self, slot: dict[str, object], safe: str) -> None:
        """Persiste un nombre sanitizado en el modelo y notifica al controlador."""
        row_data = self._all_pairs[slot["pair_index"]]
        row_data["new"] = safe
        if self._on_name_change:
            self._on_name_change(slot["pair_index"], safe, row_data["photo_path"])

    def _commit_slot(self, slot: dict[str, object]) -> None:
        """Persiste el valor pendiente del slot en su fila del modelo."""
        pi = slot["pair_index"]
        if pi < 0:
            return
        value = slot["new_var"].get()
        row = self._all_pairs[pi]
        if value != row["new"]:
            safe = Path(value).name if value else value
            row["new"] = safe
            if self._on_name_change:
                self._on_name_change(pi, safe, row["photo_path"])

    def _sync_viewport(self) -> None:
        if self._cv is None:
            return
        n = len(self._filtered)
        view_h = self._view_height()
        if n == 0:
            for slot in self._rows:
                self._hide_slot(slot)
            return
        total = n * self.ROW_H
        max_px = max(0, total - view_h)
        if self._pixel_offset() > max_px:
            self._cv.yview_moveto(max_px / total if total else 0.0)
        first = max(0, int(self._pixel_offset() // self.ROW_H) - self.BUFFER)
        # El límite del pool es visible + 2 (+2 cubre filas parciales de
        # pantalla); `last` usa el MISMO tamaño para no dejar slots muertos.
        last = min(n, first + self._pool_size())
        self._first = first
        for k, slot in enumerate(self._rows):
            logical = first + k
            if logical < last:
                if slot["index"] != logical:
                    self._rebind_slot(slot, logical)
                else:
                    try:
                        self._cv.itemconfigure(slot["item"], state="normal")
                    except Exception:
                        pass
            else:
                self._hide_slot(slot)

    def _rebind_slot(self, slot: dict[str, object], logical: int) -> None:
        self._commit_slot(slot)
        i = self._filtered[logical]
        row = self._all_pairs[i]
        state = row.get("state", "ok")
        is_dup = row.get("is_dup", False)

        bg = C["state_bg"].get(state) if state != "ok" else None
        if bg is None:
            bg = C["dup_bg"] if is_dup else (C["surface"] if i % 2 == 0 else C["bg"])
        if i in self._selected:
            slot["frame"].configure(fg_color=bg, border_width=1, border_color=C["accent"])
        else:
            slot["frame"].configure(fg_color=bg, border_width=0)
        slot["num_widget"].configure(text=f"{i+1:>3}.")
        slot["orig_widget"].configure(text=row["orig"])
        new_color = self._new_color(state, is_dup)
        slot["new_widget"].configure(text=row["new"], text_color=new_color)
        slot["new_entry"].configure(border_color=C["red"] if is_dup else C["border"],
                                    text_color=new_color)
        self._syncing = True
        try:
            slot["new_var"].set(row["new"])
        finally:
            self._syncing = False
        slot["state_widget"].configure(text=STATE_LABELS.get(state, ""),
                                       text_color=C["state_fg"].get(state, C["overlay"]))
        slot["photo_path"] = row["photo_path"]
        slot["index"] = logical
        slot["pair_index"] = i
        if self._edit_mode:
            slot["new_entry"].grid()
            slot["new_widget"].grid_remove()
        else:
            slot["new_widget"].grid()
            slot["new_entry"].grid_remove()
        # FIX: reposicionar la fila física en su coordenada lógica. El pool se
        # construye en k*ROW_H (solo k=0 está en su sitio); sin este coords()
        # las filas quedan todas encima de la primera y el scroll "mueve" los
        # widgets pero la tabla parece no desplazarse.
        try:
            self._cv.coords(slot["item"], 0, logical * self.ROW_H)
        except Exception:
            pass
        try:
            self._cv.itemconfigure(slot["item"], state="normal")
        except Exception:
            pass

    def _hide_slot(self, slot: dict[str, object]) -> None:
        self._commit_slot(slot)
        slot["index"] = -1
        slot["pair_index"] = -1
        try:
            self._cv.itemconfigure(slot["item"], state="hidden")
        except Exception:
            pass

    def _view_height(self) -> int:
        """Altura del viewport del canvas, estable incluso antes del primer
        layout: si el canvas aún no está mapeado usa la altura FIJA de la
        tabla (así el pool de slots no se construye con un tamaño fantasma)."""
        view_h = self._cv.winfo_height()
        if view_h < 2:
            view_h = max(self._MIN_H, self._adaptive_height - 34)
        return view_h

    def _sync_scrollregion(self) -> None:
        if self._cv is None:
            return
        n = len(self._filtered)
        total = n * self.ROW_H
        w = self._cv.winfo_width()
        if w < 2:
            w = 1
        self._cv.configure(scrollregion=(0, 0, w, total))
        view_h = self._view_height()
        max_px = max(0, total - view_h)
        if self._pixel_offset() > max_px:
            self._cv.yview_moveto(max_px / total if total else 0.0)

    @staticmethod
    def _new_color(state: str, is_dup: bool) -> str:
        if state != "ok":
            return C["state_fg"].get(state, C["text"])
        return C["red"] if is_dup else C["text"]

    # ── scroll / viewport ─────────────────────────────────────────────────
    def _on_sb(self, *args) -> None:
        self._cv.yview(*args)
        self._close_tip()
        self._hover_path = None
        self._sync_viewport()

    def _on_wheel(self, event) -> Optional[str]:
        step = self.ROW_H * 3
        if event.num == 4:
            self._scroll_by(-step)
        elif event.num == 5:
            self._scroll_by(step)
        elif event.delta:
            self._scroll_by(-step if event.delta > 0 else step)
        # Scroll chaining: solo devolver "break" si la tabla tiene contenido
        # scrollable.  Si está vacía o ya no puede scrollear en la dirección
        # pedida, devolver None para que el evento llegue al CTkScrollableFrame
        # exterior (el usuario puede seguir bajando/subiendo la página).
        if not self._all_pairs:
            return None
        top, bot = self._cv.yview()
        scrolling_up = (event.num == 4) or (event.delta and event.delta > 0)
        scrolling_down = (event.num == 5) or (event.delta and event.delta < 0)
        at_top = top <= 0.0
        at_bottom = bot >= 1.0
        if (scrolling_up and at_top) or (scrolling_down and at_bottom):
            return None
        return "break"

    def _on_frame_configure(self, event) -> None:
        """Adaptar canvas y pool cuando el padre redimensiona este frame."""
        if event.widget is not self:
            return
        new_h = event.height
        if new_h < 2:
            return
        self._adaptive_height = new_h
        if self._cv is not None:
            old_pool = len(self._rows)
            new_pool = self._pool_size()
            if new_pool != old_pool:
                self._cv.delete("all")
                self._build_pool()
            self._sync_scrollregion()
            self._sync_viewport()

    def _on_canvas_resize(self, event) -> None:
        w = event.width if event.width > 1 else 1
        for slot in self._rows:
            try:
                self._cv.itemconfigure(slot["item"], width=w)
            except Exception:
                pass
        # Si el canvas recién se mapeó, el pool pudo haberse construido con un
        # tamaño fantasma (winfo_height=1): reconstruirlo con el tamaño real
        # para que el viewport siempre tenga slots de sobra (visible + buffer).
        if self._rows and len(self._rows) != self._pool_size():
            self._cv.delete("all")
            self._build_pool()
        self._sync_scrollregion()
        self._sync_viewport()

    # ── tooltip por fila lógica (datos leídos del modelo, sin refs a slots) ─
    def _on_motion(self, event) -> None:
        n = len(self._filtered)
        if n == 0:
            self._close_tip()
            self._hover_path = None
            return
        # canvasy ya devuelve la coordenada ABSOLUTA del canvas (incluye el
        # scroll). Dividir por ROW_H da el índice lógico directo: NO se debe
        # sumar self._first (eso duplicaría el offset tras hacer scroll).
        logical = int(self._cv.canvasy(event.y) // self.ROW_H)
        if logical < 0 or logical >= n:
            self._close_tip()
            self._hover_path = None
            return
        path = self._all_pairs[self._filtered[logical]]["photo_path"]
        self._set_hover(path)

    def _on_slot_motion(self, slot: dict[str, object]):
        def _handler(_e=None):
            # pair_index apunta al elemento REAL del modelo (_all_pairs), es
            # inmutable para esa fila y no depende del orden de _filtered ni
            # del scroll. Leer slot["index"] aquí podía mostrar la imagen de
            # la fila "anterior" del slot durante scroll rápido (reciclo).
            pi = slot["pair_index"]
            if pi < 0 or pi >= len(self._all_pairs):
                return
            path = self._all_pairs[pi]["photo_path"]
            self._set_hover(path)
        return _handler

    def _set_hover(self, path) -> None:
        if path is None:
            self._close_tip()
            self._hover_path = None
            return
        if path == self._hover_path:
            return
        self._hover_path = path
        self._close_tip()
        if not path.exists():
            return
        _safe_cancel_after(self, self._tip_job)
        self._tip_job = self.after(ImageTooltip.DELAY_MS, lambda: self._show_tip(path))

    def _show_tip(self, path: Path) -> None:
        self._tip_job = None
        if path != self._hover_path or self._tip:
            return
        try:
            ImageTooltip._close_active()
            ctk_img = _get_thumb(path, (180, 180))
            if not ctk_img:
                return
            tip = ctk.CTkToplevel(self.winfo_toplevel())
            tip.wm_overrideredirect(True)
            tip.configure(fg_color=C["surface"])
            tip.attributes("-topmost", True)
            ctk.CTkLabel(tip, image=ctk_img, text="", fg_color=C["surface"]).pack(padx=6, pady=6)
            _place_tip_near_pointer(tip, self, offset=18)
            tip.protocol("WM_DELETE_WINDOW", self._close_tip)
            self._tip = tip
            ImageTooltip._ACTIVE.append(tip)
        except Exception as exc:
            log.debug("Tooltip error: %s", exc)

    def _close_tip(self, _=None) -> None:
        _safe_cancel_after(self, self._tip_job)
        self._tip_job = None
        if self._tip is not None:
            try:
                if self._tip in ImageTooltip._ACTIVE:
                    ImageTooltip._ACTIVE.remove(self._tip)
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None
#  DIÁLOGO DE CONFIRMACIÓN INTEGRADO
# ===========================================================================
class ConfirmDialog(ctk.CTkToplevel):
    """Reemplaza messagebox.askyesno con un diálogo dark estilizado."""

    def __init__(self, master, title: str, message: str,
                 ok_text: str = "Confirmar") -> None:
        super().__init__(master)
        self.transient(master)
        self.withdraw()
        self.title(title)
        self.resizable(False, False)
        self.configure(fg_color=C["bg"])
        self.result = False
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        ctk.CTkLabel(self, text=message,
                     font=FONT_MD,
                     text_color=C["text"], wraplength=380,
                     justify="center").pack(padx=28, pady=(28, 16))

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(pady=(0, 22))
        ctk.CTkButton(btn_row, text="Cancelar", width=100, height=32,
                      **BTN_DANGER,
                      font=FONT_MD,
                      command=self._cancel).pack(side="left", padx=8)
        ctk.CTkButton(btn_row, text=ok_text, width=110, height=32,
                      **BTN_PRIMARY,
                      text_color=C["accent_text"],
                      font=FONT_MD_BD,
                      command=self._ok).pack(side="left", padx=8)

        self.update_idletasks()
        w, h = 440, 180
        px = master.winfo_rootx() + (master.winfo_width()  - w) // 2
        py = master.winfo_rooty() + (master.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{px}+{py}")
        _show_toplevel(self)

    def _ok(self) -> None:
        self.result = True
        self.grab_release()
        self.destroy()

    def _cancel(self) -> None:
        self.grab_release()
        self.destroy()

    @classmethod
    def ask(cls, master, title: str, message: str, ok_text: str = "Confirmar") -> bool:
        dlg = cls(master, title, message, ok_text)
        dlg.wait_window()
        return dlg.result


# ===========================================================================
#  VISTA PRINCIPAL  — FIX #3 (resolución adaptativa)
# ===========================================================================
class MainView(ctk.CTk):
    """Vista principal. Delega lógica al Controller (MVC)."""

    def __init__(self, controller: "AppController") -> None:
        super().__init__()
        _init_fonts(self.winfo_screenwidth())
        self._ctrl = controller
        self.title("MetaTag v8.9 — Image Sync")
        self.configure(fg_color=C["bg"])

        # FIX #3: tamaño adaptativo a la pantalla
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = max(780, min(1100, int(sw * 0.68)))
        h  = max(600, min(900,  int(sh * 0.80)))
        px = (sw - w) // 2
        py = (sh - h) // 2
        self.geometry(f"{w}x{h}+{px}+{py}")
        self.minsize(PROFILE.min_w, PROFILE.min_h)

        self.report_callback_exception = self._tk_error_handler
        self._build()
        self._bind_shortcuts()

    # ── construcción ───────────────────────────────────────────────────────
    def _build(self) -> None:
        # header
        hdr = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=0, height=64)
        hdr.pack(fill="x"); hdr.pack_propagate(False)

        title_col = ctk.CTkFrame(hdr, fg_color="transparent")
        title_col.pack(side="left", padx=22, pady=(8, 4))
        ctk.CTkLabel(title_col, text="Image Sync",
                     font=FONT_TITLE,
                     text_color=C["text"], anchor="w").pack(anchor="w")
        ctk.CTkLabel(title_col,
                     text="Sincroniza los nombres de las fotografías con los "
                          "registros del Excel.",
                     font=FONT_XS_SM,
                     text_color=C["subtext"], anchor="w").pack(anchor="w")

        self._btn_undo = ctk.CTkButton(
            hdr, text="↩  Deshacer", width=120, height=32,
            font=FONT_SM,
            fg_color=C["surface2"], hover_color=C["yellow"],
            text_color=C["text"],
            command=self._ctrl.on_undo, state="disabled")
        self._btn_undo.pack(side="right", padx=(0, 14))

        self._theme_var = ctk.StringVar(value=CURRENT_THEME)
        _make_option_menu(
            hdr, self._theme_var, list(THEME_ORDER),
            width=210, height=30, fg=C["surface2"], btn_fg=C["surface3"],
            command=self._ctrl.on_theme_change
        ).pack(side="right", padx=(0, 8))

        # indicador de pasos (①…⑤)
        self._steps = ctk.CTkFrame(self, fg_color=C["bg"], height=30)
        self._steps.pack(fill="x"); self._steps.pack_propagate(False)
        self._step_labels: list[ctk.CTkLabel] = []
        self._steps_done = 0
        for i, sname in enumerate(_STEP_NAMES, start=1):
            lbl = ctk.CTkLabel(
                self._steps, text=f"{_STEP_DIGITS[i - 1]} {sname}",
                font=FONT_XS_SM, text_color=C["overlay"],
                fg_color="transparent")
            lbl.pack(side="left", pady=6)
            if i < len(_STEP_NAMES):
                ctk.CTkLabel(self._steps, text="→", font=FONT_XS_SM,
                             text_color=C["border"],
                             fg_color="transparent").pack(side="left",
                                                          padx=(6, 6), pady=6)
            else:
                lbl.pack_configure(padx=(0, 4))
            self._step_labels.append(lbl)
        self.update_steps(0)

        # scroll principal
        self._scroll = ctk.CTkScrollableFrame(
            self, fg_color="transparent",
            scrollbar_button_color=C["surface2"])
        self._scroll.pack(fill="both", expand=True, padx=20, pady=12)
        self._scroll.columnconfigure(0, weight=1)
        SmoothScroller(self._scroll)

        c = self._scroll

        # § 1 carpeta
        self._sec(c, "1 · Carpeta de fotos", 0)
        self._folder_sel = PathSelector(
            c, placeholder="/home/usuario/Fotos  o  D:\\Fotos",
            mode="folder", on_change=self._ctrl.on_folder_path_changed)
        self._folder_sel.grid(row=1, column=0, sticky="ew", pady=(0, 5))

        sort_row = ctk.CTkFrame(c, fg_color="transparent")
        sort_row.grid(row=2, column=0, sticky="ew", pady=(0, 3))
        ctk.CTkLabel(sort_row, text="Ordenar por:",
                     font=FONT_SM,
                     text_color=C["subtext"]).pack(side="left")
        self._sort_var = ctk.StringVar(value="Orden numérico")
        _make_option_menu(
            sort_row, self._sort_var, list(SORT_OPTIONS.keys()),
            width=185, command=self._ctrl.on_sort_change
        ).pack(side="left", padx=8)

        i1 = ctk.CTkFrame(c, fg_color="transparent")
        i1.grid(row=3, column=0, sticky="ew", pady=(0, 12))
        self._badge_folder = StatusBadge(i1); self._badge_folder.pack(side="left")
        self._lbl_folder = ctk.CTkLabel(i1, text="Selecciona una carpeta",
                                        font=FONT_SM,
                                        text_color=C["subtext"]); self._lbl_folder.pack(side="left", padx=7)
        self._btn_load_folder = ctk.CTkButton(
            i1, text="Cargar fotos", width=108, height=26,
            font=FONT_SM,
            **BTN_SECONDARY,
            command=self._ctrl.on_load_photos)
        self._btn_load_folder.pack(side="right")

        self._div(c, 4)

        # § 2 excel
        self._sec(c, "2 · Archivo Excel", 5)
        self._excel_sel = PathSelector(
            c, placeholder="/home/usuario/nombres.xlsx  o  D:\\nombres.xlsx",
            mode="file", filetypes=[".xlsx", ".csv", ".tsv", ".txt"],
            on_change=self._ctrl.on_excel_path_changed)
        self._excel_sel.grid(row=6, column=0, sticky="ew", pady=(0, 5))

        i2 = ctk.CTkFrame(c, fg_color="transparent")
        i2.grid(row=7, column=0, sticky="ew", pady=(0, 6))
        self._badge_excel = StatusBadge(i2); self._badge_excel.pack(side="left")
        self._lbl_excel = ctk.CTkLabel(i2, text="Selecciona el archivo Excel",
                                       font=FONT_SM,
                                       text_color=C["subtext"]); self._lbl_excel.pack(side="left", padx=7)
        self._btn_load_excel = ctk.CTkButton(
            i2, text="Cargar Excel", width=108, height=26,
            font=FONT_SM,
            **BTN_SECONDARY,
            command=self._ctrl.on_load_excel)
        self._btn_load_excel.pack(side="right")

        # selector hoja (oculto)
        self._sheet_frame = ctk.CTkFrame(c, fg_color="transparent")
        self._sheet_frame.grid(row=8, column=0, sticky="ew", pady=(0, 3))
        ctk.CTkLabel(self._sheet_frame, text="Hoja:",
                     font=FONT_SM,
                     text_color=C["subtext"]).pack(side="left")
        self._sheet_var = ctk.StringVar()
        self._sheet_menu = _make_option_menu(
            self._sheet_frame, self._sheet_var, ["—"],
            width=175, command=self._ctrl.on_sheet_selected)
        self._sheet_menu.pack(side="left", padx=7)
        self._sheet_frame.grid_remove()

        # selector columna (oculto)
        self._col_frame = ctk.CTkFrame(c, fg_color="transparent")
        self._col_frame.grid(row=9, column=0, sticky="ew", pady=(0, 12))
        ctk.CTkLabel(self._col_frame, text="Columna con los nombres:",
                     font=FONT_SM,
                     text_color=C["subtext"]).pack(side="left")
        self._col_var = ctk.StringVar()
        self._col_menu = _make_option_menu(
            self._col_frame, self._col_var, ["—"],
            width=195, command=self._ctrl.on_column_selected)
        self._col_menu.pack(side="left", padx=7)
        self._col_frame.grid_remove()

        self._build_summary(c, 10)

        self._div(c, 11)

        # § 3 vista previa
        self._sec(c, "3 · Vista previa", 12)

        filter_row = ctk.CTkFrame(c, fg_color="transparent")
        filter_row.grid(row=13, column=0, sticky="ew", pady=(0, 5))
        ctk.CTkLabel(filter_row, text="Buscar:",
                     font=FONT_SM,
                     text_color=C["subtext"]).pack(side="left")
        self._filter_var = ctk.StringVar()
        self._filter_var.trace_add("write", lambda *_: self._ctrl.on_filter_change(self._filter_var.get()))
        ctk.CTkEntry(filter_row, textvariable=self._filter_var,
                     width=220, height=28,
                     font=FONT_SM,
                     fg_color=C["surface"], border_color=C["border"],
                     text_color=C["text"]).pack(side="left", padx=8)

        self._edit_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(filter_row, text="Modo editar",
                        variable=self._edit_var,
                        font=FONT_SM,
                        text_color=C["subtext"],
                        command=lambda: self._ctrl.on_edit_mode_change(self._edit_var.get())
                        ).pack(side="left", padx=(8, 0))

        self._lbl_filter_count = ctk.CTkLabel(
            filter_row, text="0 resultados", font=FONT_SM,
            text_color=C["subtext"])
        self._lbl_filter_count.pack(side="right", padx=6)

        self._preview = PreviewTable(
            c, on_name_change=self._ctrl.on_preview_name_changed,
            on_filter=self._set_filter_count)
        self._preview.grid(row=14, column=0, sticky="ew", pady=(2, 12))

        self._div(c, 15)

        # § 4 opciones
        self._sec(c, "4 · Opciones", 16)
        opt = ctk.CTkFrame(c, fg_color=C["surface"], corner_radius=10)
        opt.grid(row=17, column=0, sticky="ew", pady=(2, 10))
        pad = ctk.CTkFrame(opt, fg_color="transparent")
        pad.pack(fill="x", padx=12, pady=10)

        box_row = ctk.CTkFrame(pad, fg_color="transparent")
        box_row.pack(fill="x", anchor="w")
        self._keep_ext_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            box_row, text="Mantener extensión original",
            variable=self._keep_ext_var, font=FONT_SM,
            text_color=C["text"],
            command=lambda: self._ctrl.on_keep_ext_change(self._keep_ext_var.get())
        ).pack(side="left")
        self._backup_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            box_row, text="Crear registro/backup de la operación",
            variable=self._backup_var, font=FONT_SM,
            text_color=C["text"],
            command=lambda: self._ctrl.on_backup_change(self._backup_var.get())
        ).pack(side="left", padx=(18, 0))
        self._open_folder_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            box_row, text="Abrir carpeta al finalizar",
            variable=self._open_folder_var, font=FONT_SM,
            text_color=C["text"],
            command=lambda: self._ctrl.on_open_folder_change(self._open_folder_var.get())
        ).pack(side="left", padx=(18, 0))

        self._lbl_sort_echo = ctk.CTkLabel(
            pad, text="", font=FONT_SM, text_color=C["subtext"], anchor="w")
        self._lbl_sort_echo.pack(fill="x", anchor="w", pady=(8, 0))
        self.set_option_sort(self._sort_var.get())

        self._div(c, 18)

        # § 5 renombrar
        self._sec(c, "5 · Renombrar", 19)
        self._progress = ctk.CTkProgressBar(c, height=9,
                                             fg_color=C["surface"],
                                             progress_color=C["accent"])
        self._progress.grid(row=20, column=0, sticky="ew", pady=(3, 3))
        self._progress.set(0)

        act = ctk.CTkFrame(c, fg_color="transparent")
        act.grid(row=21, column=0, sticky="ew", pady=(0, 12))

        self._lbl_prog = ctk.CTkLabel(act, text="Listo para comenzar",
                                      font=FONT_SM,
                                      text_color=C["subtext"])
        self._lbl_prog.pack(side="left")

        # copia, cancelar, log, preview csv, renombrar — de derecha a izquierda
        self._btn_rename = ctk.CTkButton(
            act, text="▶  Renombrar todo", width=165, height=36,
            font=FONT_MD_BD,
            **BTN_PRIMARY,
            text_color=C["accent_text"],
            command=self._ctrl.on_rename, state="disabled")
        self._btn_rename.pack(side="right")

        self._btn_sim = ctk.CTkButton(
            act, text="Simular", width=90, height=36,
            font=FONT_SM,
            fg_color=C["surface2"], hover_color=C["accent2"],
            command=self._ctrl.on_simulate, state="disabled")
        self._btn_sim.pack(side="right", padx=(0, 6))

        self._btn_cancel = ctk.CTkButton(
            act, text="✖ Cancelar", width=100, height=36,
            font=FONT_SM,
            **BTN_DANGER,
            command=self._ctrl.on_cancel, state="disabled")
        self._btn_cancel.pack(side="right", padx=(0, 6))

        self._btn_log = ctk.CTkButton(
            act, text="Guardar log", width=100, height=36,
            font=FONT_SM,
            fg_color=C["surface2"], hover_color=C["accent2"],
            command=self._ctrl.on_export_log, state="disabled")
        self._btn_log.pack(side="right", padx=(0, 6))

        self._btn_csv = ctk.CTkButton(
            act, text="CSV", width=64, height=36,
            font=FONT_SM,
            fg_color=C["surface2"], hover_color=C["accent2"],
            command=self._ctrl.on_export_csv, state="disabled")
        self._btn_csv.pack(side="right", padx=(0, 6))

        self._copy_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(act, text="Modo copiar",
                        variable=self._copy_var,
                        font=FONT_SM,
                        text_color=C["subtext"]).pack(side="right", padx=(0, 10))

        # Matching seguro: cada nombre busca SU fotografía (recomendado).
        # Disponible solo si el motor puro de MetaTag está presente.
        self._match_var = ctk.BooleanVar(value=False)
        self._btn_match = ctk.CTkCheckBox(
            act, text="Matching seguro", variable=self._match_var,
            font=FONT_SM, text_color=C["subtext"],
            command=self._ctrl.on_matching_toggle)
        self._btn_match.pack(side="right", padx=(0, 10))
        if not self._ctrl.matching_available():
            self._btn_match.configure(state="disabled")

        self._div(c, 22)

        # § 6 registro
        self._sec(c, "6 · Registro", 23)
        log_frame = ctk.CTkFrame(c, fg_color=C["surface"], corner_radius=10)
        log_frame.grid(row=24, column=0, sticky="ew", pady=(2, 4))
        self._log_box = ctk.CTkTextbox(log_frame, height=110,
                                       font=FONT_XS_SM,
                                       fg_color=C["surface"],
                                       text_color=C["subtext"],
                                       border_spacing=6)
        self._log_box.pack(fill="both", expand=True, padx=8, pady=(8, 4))
        self._log_box.configure(state="disabled")
        log_bar = ctk.CTkFrame(log_frame, fg_color="transparent")
        log_bar.pack(fill="x", padx=8, pady=(0, 6))
        ctk.CTkLabel(log_bar, text="Historial de la última operación",
                     font=FONT_XS_SM,
                     text_color=C["subtext"]).pack(side="left")
        ctk.CTkButton(log_bar, text="Limpiar", width=70, height=26,
                      font=FONT_SM,
                      **BTN_SECONDARY,
                      command=self.clear_log).pack(side="right")

        # footer
        ftr = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=0, height=26)
        ftr.pack(fill="x", side="bottom"); ftr.pack_propagate(False)
        ctk.CTkLabel(
            ftr,
            text="Los archivos se renombran en su misma ubicación  •  No se eliminan  •  Hover → miniatura  •  Ctrl+Z = Deshacer",
            font=FONT_XS,
            text_color=C["subtext"]).pack(expand=True)

    def _sec(self, p, title: str, row: int) -> None:
        ctk.CTkLabel(p, text=title,
                     font=FONT_HD,
                     text_color=C["accent"]).grid(
            row=row, column=0, sticky="w", pady=(10, 3))

    def _div(self, p, row: int) -> None:
        ctk.CTkFrame(p, height=1, fg_color=C["border"]).grid(
            row=row, column=0, sticky="ew", pady=6)

    def _build_summary(self, p, row: int) -> None:
        """Panel de resumen en vivo: Fotografías / Registros / Correspondencias /
        Conflictos / Estado. Cada celda tiene dos líneas: caption pequeña arriba
        y valor en negrita abajo. Lo rellena `update_summary` desde el modelo.

        Definiciones formales de los contadores (fuente única: el plan):
          Fotografías    = archivos de imagen válidos en la carpeta.
          Registros      = nombres válidos extraídos del Excel/CSV.
          Correspondencias = registros con foto asociada (src != None).
          Conflictos     = correspondencias en estado bloqueante.
        """
        f = ctk.CTkFrame(p, fg_color="transparent")
        f.grid(row=row, column=0, sticky="ew", pady=(0, 6))
        for col in range(5):
            f.columnconfigure(col, weight=1, uniform="sum")

        def _stat(col: int, caption: str) -> tuple[StatusBadge, ctk.CTkLabel]:
            cell = ctk.CTkFrame(f, fg_color="transparent")
            cell.grid(row=0, column=col, sticky="ew", padx=4)
            ctk.CTkLabel(cell, text=caption, font=FONT_XS_SM,
                         text_color=C["subtext"], anchor="w").pack(fill="x", anchor="w")
            line = ctk.CTkFrame(cell, fg_color="transparent")
            line.pack(fill="x", anchor="w")
            b = StatusBadge(line); b.pack(side="left")
            lbl = ctk.CTkLabel(line, text="—", font=FONT_MD_BD,
                               text_color=C["text"])
            lbl.pack(side="left", padx=5)
            return b, lbl

        self._sum_fotos, self._sum_fotos_lbl = _stat(0, "Fotografías")
        self._sum_regs,  self._sum_regs_lbl  = _stat(1, "Registros")
        self._sum_corr,  self._sum_corr_lbl  = _stat(2, "Correspondencias")
        self._sum_conf,  self._sum_conf_lbl  = _stat(3, "Conflictos")
        self._sum_state, self._sum_state_lbl = _stat(4, "Estado")
        self.update_summary(None, None, None, None, "Sin datos", "idle")

    def update_summary(self, fotos, registros, correspondencias, conflictos,
                       estado: str, kind: str) -> None:
        """kind ∈ {idle, ok, warn, error} controla el color del Estado."""
        def _set(badge, lbl, val) -> None:
            badge.set_state("ok" if val else "idle")
            lbl.configure(text=("—" if val is None else str(val)))
        _set(self._sum_fotos, self._sum_fotos_lbl, fotos)
        _set(self._sum_regs,  self._sum_regs_lbl,  registros)
        _set(self._sum_corr,  self._sum_corr_lbl,  correspondencias)
        _set(self._sum_conf,  self._sum_conf_lbl,  conflictos)
        self._sum_state.set_state(kind if kind in self._sum_state._S else "idle")
        self._sum_state_lbl.configure(text=estado,
                                      text_color=_state_text_color(kind))

    def _bind_shortcuts(self) -> None:
        self.bind_all("<Control-z>", lambda _: self._ctrl.on_undo())
        self.bind_all("<Control-o>", lambda _: self._ctrl.on_load_photos())
        self.bind_all("<Control-e>", lambda _: self._ctrl.on_load_excel())
        self.bind_all("<Escape>", lambda _: self._ctrl.on_cancel())
        self.bind_all("<Control-Return>", lambda _: self._ctrl.on_rename())
        self.after(100, self._set_initial_focus)

    def _set_initial_focus(self) -> None:
        try:
            self._folder_sel._entry.focus_set()
        except Exception:
            pass

    # ── cambio de tema: reconstrucción completa sin perder estado ─────────
    def rebuild_theme(self) -> None:
        """Destruye la UI y la reconstruye con el nuevo C, conservando el
        estado visible del usuario (rutas, opciones, badges, botones, etc.)."""
        state = self._capture_ui_state()
        try:
            self._preview._cancel_jobs()
        except Exception:
            pass
        for w in self.winfo_children():
            w.destroy()
        self._build()
        self._bind_shortcuts()
        self._restore_ui_state(state)

    def _capture_ui_state(self) -> dict:
        return {
            "folder": self._folder_sel.get(),
            "excel": self._excel_sel.get(),
            "sort": self._sort_var.get(),
            "edit": self._edit_var.get(),
            "copy": self._copy_var.get(),
            "match": self._match_var.get(),
            "filter": self._filter_var.get(),
            "sheet_visible": self._sheet_frame.winfo_manager() == "grid",
            "sheet_values": list(self._sheet_menu.cget("values") or []),
            "sheet_value": self._sheet_var.get(),
            "col_visible": self._col_frame.winfo_manager() == "grid",
            "col_values": list(self._col_menu.cget("values") or []),
            "col_value": self._col_var.get(),
            "folder_badge": self._badge_folder._state,
            "folder_msg": self._lbl_folder.cget("text"),
            "excel_badge": self._badge_excel._state,
            "excel_msg": self._lbl_excel.cget("text"),
            "progress": self._progress.get(),
            "progress_msg": self._lbl_prog.cget("text"),
            "undo": (self._btn_undo.cget("state"), self._btn_undo.cget("text")),
            "rename": (self._btn_rename.cget("state"), self._btn_rename.cget("text")),
            "cancel": self._btn_cancel.cget("state"),
            "log": self._btn_log.cget("state"),
            "csv": self._btn_csv.cget("state"),
            "sim": self._btn_sim.cget("state"),
            "sum_fotos": self._sum_fotos_lbl.cget("text"),
            "sum_regs": self._sum_regs_lbl.cget("text"),
            "sum_corr": self._sum_corr_lbl.cget("text"),
            "sum_conf": self._sum_conf_lbl.cget("text"),
            "sum_state": self._sum_state_lbl.cget("text"),
            "sum_state_kind": self._sum_state._state,
            "keep_ext": self._keep_ext_var.get(),
            "backup": self._backup_var.get(),
            "open_folder": self._open_folder_var.get(),
            "steps": self._steps_done,
            "log_text": self.get_log_text(),
        }

    def _restore_ui_state(self, state: dict) -> None:
        self._folder_sel.set(state["folder"])
        self._excel_sel.set(state["excel"])
        self._sort_var.set(state["sort"])
        self._edit_var.set(state["edit"])
        self._copy_var.set(state["copy"])
        self._match_var.set(state["match"])
        if state["sheet_visible"] and state["sheet_values"]:
            self._sheet_menu.configure(values=state["sheet_values"])
            if state["sheet_value"] in state["sheet_values"]:
                self._sheet_var.set(state["sheet_value"])
            else:
                self._sheet_var.set(state["sheet_values"][0])
            self._sheet_frame.grid()
        if state["col_visible"] and state["col_values"]:
            self._col_menu.configure(values=state["col_values"])
            if state["col_value"] in state["col_values"]:
                self._col_var.set(state["col_value"])
            else:
                self._col_var.set(state["col_values"][0])
            self._col_frame.grid()
        self._badge_folder.set_state(state["folder_badge"])
        self._lbl_folder.configure(text=state["folder_msg"])
        self._badge_excel.set_state(state["excel_badge"])
        self._lbl_excel.configure(text=state["excel_msg"])
        self._progress.set(state["progress"])
        self._lbl_prog.configure(text=state["progress_msg"])
        self._btn_undo.configure(state=state["undo"][0], text=state["undo"][1])
        self._btn_rename.configure(state=state["rename"][0], text=state["rename"][1])
        self._btn_cancel.configure(state=state["cancel"])
        self._btn_log.configure(state=state["log"])
        self._btn_csv.configure(state=state["csv"])
        self._btn_sim.configure(state=state.get("sim", "disabled"))
        self._sum_fotos_lbl.configure(text=state.get("sum_fotos", "—"))
        self._sum_regs_lbl.configure(text=state.get("sum_regs", "—"))
        self._sum_corr_lbl.configure(text=state.get("sum_corr", "—"))
        self._sum_conf_lbl.configure(text=state.get("sum_conf", "—"))
        self._sum_state_lbl.configure(text=state.get("sum_state", "—"))
        self._sum_state.set_state(state.get("sum_state_kind", "idle"))
        self._filter_var.set(state["filter"])
        self._keep_ext_var.set(state.get("keep_ext", True))
        self._backup_var.set(state.get("backup", True))
        self._open_folder_var.set(state.get("open_folder", True))
        self.update_steps(state.get("steps", 0))
        self.set_option_sort(self._sort_var.get())
        if state.get("log_text"):
            self.append_log_lines(state["log_text"].splitlines())

    # ── API para el Controller ─────────────────────────────────────────────
    def get_folder_path(self) -> str:   return self._folder_sel.get()
    def get_excel_path(self) -> str:    return self._excel_sel.get()
    def get_column(self) -> str:        return self._col_var.get()
    def get_sort_mode(self) -> str:     return SORT_OPTIONS.get(self._sort_var.get(), "natural")
    def get_copy_mode(self) -> bool:    return self._copy_var.get()
    def get_matching_mode(self) -> bool: return self._match_var.get()
    def get_edit_mode(self) -> bool:    return self._edit_var.get()
    def get_keep_ext_option(self) -> bool:  return self._keep_ext_var.get()
    def get_backup_option(self) -> bool:    return self._backup_var.get()
    def get_open_folder_option(self) -> bool: return self._open_folder_var.get()

    def set_btn_folder_load(self, enabled: bool) -> None:
        self._set_btn(self._btn_load_folder, enabled)

    def set_btn_excel_load(self, enabled: bool) -> None:
        self._set_btn(self._btn_load_excel, enabled)

    def set_option_sort(self, text: str) -> None:
        try:
            self._lbl_sort_echo.configure(
                text=f"Orden de emparejamiento: {text}")
        except Exception:
            pass

    def update_steps(self, done: int) -> None:
        """Colorea el indicador ①…⑤: completado (verde) / activo (acento) /
        pendiente (apagado)."""
        self._steps_done = max(0, min(len(self._step_labels), done))
        for i, lbl in enumerate(self._step_labels, start=1):
            if i <= self._steps_done:
                color = C["green"]
            elif i == self._steps_done + 1:
                color = C["accent"]
            else:
                color = C["overlay"]
            try:
                lbl.configure(text_color=color)
            except Exception:
                pass

    def get_steps_done(self) -> int:
        return self._steps_done

    def _set_filter_count(self, n: int) -> None:
        try:
            self._lbl_filter_count.configure(
                text=f"{n} resultado{'s' if n != 1 else ''}")
        except Exception:
            pass

    def append_log_line(self, line: str) -> None:
        self.append_log_lines([line])

    def append_log_lines(self, lines: list[str]) -> None:
        try:
            self._log_box.configure(state="normal")
            for ln in lines:
                self._log_box.insert("end", ln + "\n")
            self._log_box.configure(state="disabled")
            self._log_box.see("end")
        except Exception:
            pass

    def clear_log(self) -> None:
        try:
            self._log_box.configure(state="normal")
            self._log_box.delete("1.0", "end")
            self._log_box.configure(state="disabled")
        except Exception:
            pass

    def get_log_text(self) -> str:
        try:
            return self._log_box.get("1.0", "end")
        except Exception:
            return ""

    def set_folder_status(self, state: str, msg: str) -> None:
        self._set_section_status(self._badge_folder, self._lbl_folder, state, msg)

    def set_excel_status(self, state: str, msg: str) -> None:
        self._set_section_status(self._badge_excel, self._lbl_excel, state, msg)

    @staticmethod
    def _set_section_status(badge: StatusBadge, label: ctk.CTkLabel,
                            state: str, msg: str) -> None:
        badge.set_state(state)
        label.configure(text=msg)

    def _show_selector(self, menu, var, frame, values: list[str]) -> None:
        menu.configure(values=values)
        var.set(values[0])
        frame.grid()

    def show_sheets(self, sheets: list[str]) -> None:
        self._show_selector(self._sheet_menu, self._sheet_var, self._sheet_frame, sheets)

    def hide_sheets(self) -> None:
        self._sheet_frame.grid_remove()

    def show_columns(self, cols: list[str]) -> None:
        self._show_selector(self._col_menu, self._col_var, self._col_frame, cols)

    def hide_columns(self) -> None:
        self._col_frame.grid_remove()

    def render_preview(self, pairs: list[tuple[str, str, Optional[Path], bool, str]],
                       empty_message: str = "Sin datos para mostrar.") -> None:
        self._preview.render(pairs, empty_message=empty_message)

    def filter_preview(self, query: str) -> None:
        self._preview.filter(query)

    def set_edit_mode(self, enabled: bool) -> None:
        self._preview.set_edit_mode(enabled)

    def set_progress(self, value: float, msg: str) -> None:
        self._progress.set(value); self._lbl_prog.configure(text=msg)

    def _set_btn(self, btn, enabled: bool) -> None:
        btn.configure(state="normal" if enabled else "disabled")

    def set_btn_rename(self, enabled: bool, label: Optional[str] = None) -> None:
        self._set_btn(self._btn_rename, enabled)
        if label: self._btn_rename.configure(text=label)

    def set_btn_undo(self, enabled: bool, count: int = 0) -> None:
        self._set_btn(self._btn_undo, enabled)
        label = f"↩ Deshacer ({count})" if count > 0 else "↩ Deshacer"
        self._btn_undo.configure(text=label)

    def set_btn_cancel(self, enabled: bool) -> None:
        self._set_btn(self._btn_cancel, enabled)

    def set_btn_log(self, enabled: bool) -> None:
        self._set_btn(self._btn_log, enabled)
        self._set_btn(self._btn_csv, enabled)

    def set_btn_csv(self, enabled: bool) -> None:
        self._set_btn(self._btn_csv, enabled)

    def set_btn_sim(self, enabled: bool) -> None:
        self._set_btn(self._btn_sim, enabled)

    def toast(self, msg: str, kind: str = "ok") -> None:
        Toast(self, msg, kind)

    def confirm(self, title: str, msg: str, ok_text: str = "Confirmar") -> bool:
        return ConfirmDialog.ask(self, title, msg, ok_text)

    def ask_save_folder(self, default_name: str) -> Optional[str]:
        br = FileBrowser(self, mode="folder", title="Elige dónde guardar")
        folder = br.get_result()
        return str(Path(folder) / default_name) if folder else None

    # ── FIX #5: manejador global de errores de Tk ─────────────────────────
    def _tk_error_handler(self, exc_type, exc_val, exc_tb) -> None:
        tb_str = "".join(traceback.format_exception(exc_type, exc_val, exc_tb))
        log.error("Excepción no manejada en UI:\n%s", tb_str)
        self.after(0, self._show_error_dialog,
                   f"{exc_type.__name__}: {exc_val}", tb_str)

    def _show_error_dialog(self, title: str, details: str) -> None:
        dlg = ctk.CTkToplevel(self)
        dlg.title("⚠ Error en la aplicación")
        dlg.configure(fg_color=C["bg"])
        dlg.resizable(True, True)

        ctk.CTkLabel(dlg, text="Ha ocurrido un error inesperado.\nEl programa puede continuar funcionando.",
                     font=FONT_MD_BD,
                     text_color=C["red"], fg_color="transparent",
                     justify="center").pack(padx=20, pady=(18, 6))

        ctk.CTkLabel(dlg, text=title,
                     font=FONT_SM,
                     text_color=C["text"], fg_color="transparent",
                     wraplength=480, justify="center").pack(padx=20, pady=(0, 8))

        tb_box = ctk.CTkTextbox(dlg, width=520, height=180,
                                font=ctk.CTkFont("Courier New", scaled_size(10, _s, 8)),
                                fg_color=C["surface"], text_color=C["subtext"])
        tb_box.pack(padx=18, pady=(0, 10), fill="both", expand=True)
        tb_box.insert("end", details[:3000])
        tb_box.configure(state="disabled")

        btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_row.pack(pady=(0, 16))

        def _copy():
            self.clipboard_clear(); self.clipboard_append(details)
            self.toast("Detalles copiados al portapapeles.")

        ctk.CTkButton(btn_row, text="Copiar detalles", width=140,
                      command=_copy).pack(side="left", padx=8)
        ctk.CTkButton(btn_row, text="Cerrar", width=90,
                      command=dlg.destroy).pack(side="left", padx=8)

        dlg.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = 560, 380
        dlg.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        _show_toplevel(dlg)


# ===========================================================================
#  CONTROLLER
# ===========================================================================
class AppController:
    """Orquesta Model ↔ View. Toda la lógica de interacción vive aquí."""

    def __init__(self) -> None:
        self._model       = RenameModel()
        self._view        = MainView(controller=self)
        self._last_pairs: list[tuple[str, str, Optional[Path], bool, str]] = []
        self._cancel_ev:  Optional[threading.Event] = None
        self._dup_recalc_job: Optional[str] = None
        # Generación de la última recarga: descarta resultados obsoletos de
        # hilos de fondo que terminaron DESPUÉS de un recálculo más nuevo.
        self._sync_gen: int = 0
        # Guardas contra cargas concurrentes (fotos/Excel) y buffer del panel
        # de registro (líneas llegadas desde hilos de fondo).
        self._loading_photos: bool = False
        self._loading_excel: bool = False
        self._pending_sort: bool = False
        self._log_pending: list[str] = []
        # Pin del indicador de pasos tras finalizar: evita que el recálculo
        # de fondo posterior degrade el paso "Resultado" durante ~4 s.
        self._step_pinned: bool = False
        self._restore_state()

    def run(self) -> None:
        self._view.mainloop()

    # ── restaurar estado ──────────────────────────────────────────────────
    def _restore_state(self) -> None:
        global C
        state = _load_state()
        if "theme" in state:
            try:
                self._apply_theme(state["theme"])
            except Exception:
                pass
        if "sort" in state:
            for k, v in SORT_OPTIONS.items():
                if v == state["sort"]:
                    self._view._sort_var.set(k); break
        if "last_folder" in state:
            p = Path(state["last_folder"])
            if p.is_dir():
                self._view._folder_sel.set(str(p))
        if "last_excel" in state:
            p = Path(state["last_excel"])
            if p.is_file():
                self._view._excel_sel.set(str(p))

    # ── handlers ──────────────────────────────────────────────────────────
    def _guard(self, condition, msg: str, kind: str = "error") -> bool:
        if condition:
            self._view.toast(msg, kind)
            return True
        return False

    def on_load_photos(self) -> None:
        if self._loading_photos:
            self._view.toast("Las fotos ya se están cargando…", "warn")
            return
        raw = self._view.get_folder_path()
        if self._guard(not raw, "Selecciona primero una carpeta."): return
        path = Path(raw)
        if not path.is_dir():
            self._view.set_folder_status("error", "La ruta no existe."); return

        self._model.folder_path = path
        self._model.sort_mode   = self._view.get_sort_mode()
        _save_state({"last_folder": str(path), "sort": self._model.sort_mode})
        self._loading_photos = True
        self._view.set_btn_folder_load(False)
        self._view.set_folder_status("loading", "Cargando fotos...")
        threading.Thread(target=self._load_photos_bg, daemon=True).start()

    def _load_photos_bg(self) -> None:
        def on_done(n, err):
            self._loading_photos = False
            self._view.set_btn_folder_load(True)
            if err:
                self._view.set_folder_status("error", err)
                return
            if n == 0:
                self._view.set_folder_status("warn", "No se encontraron imágenes.")
            else:
                sort_label = self._view._sort_var.get()
                self._view.set_folder_status("ok", f"{n} imagen{'es' if n!=1 else ''}  ·  {sort_label}")
            if self._pending_sort:
                self._pending_sort = False
                self._apply_sort()
                return
            self._refresh_preview()
        def _work():
            try:
                n = self._model.load_photos()
                self._view.after(0, on_done, n, None)
            except Exception as exc:
                self._view.after(0, on_done, 0, str(exc))
        self._run_safely(_work)

    def on_load_excel(self) -> None:
        if self._loading_excel:
            self._view.toast("El Excel ya se está cargando…", "warn")
            return
        raw = self._view.get_excel_path()
        if self._guard(not raw, "Selecciona primero un archivo Excel."): return
        path = Path(raw)
        valid_exts = {".xlsx", ".csv", ".tsv", ".txt"}
        if not path.is_file() or path.suffix.lower() not in valid_exts:
            self._view.set_excel_status("error", f"Formato no soportado: {path.suffix}"); return

        self._model.excel_path = path
        _save_state({"last_excel": str(path)})
        self._loading_excel = True
        self._view.set_btn_excel_load(False)
        try:
            sheets = self._call_model(self._model.load_sheets,
                                      lambda s, m: self._view.set_excel_status(s, f"Error leyendo Excel: {m}"))
            if sheets is None:
                return

            if len(sheets) > 1:
                self._view.set_excel_status("ok", f"{len(sheets)} hojas — elige una.")
                self._view.show_sheets(sheets)
                self._model.sheet_name = sheets[0]
            else:
                self._model.sheet_name = sheets[0] if sheets else None
                self._view.hide_sheets()
            self._load_columns()
        finally:
            self._loading_excel = False
            self._view.set_btn_excel_load(True)

    def on_sheet_selected(self, sheet: str) -> None:
        self._model.sheet_name = sheet
        self._model.clear_excel_data()
        self._load_columns()

    def on_column_selected(self, column: str) -> None:
        self._model.column_name = column
        self._load_names_and_preview()

    def on_filter_change(self, query: str) -> None:
        self._view.filter_preview(query)

    # ── handlers reactivos (ruta / orden / matching) ──────────────────────
    def on_sort_change(self, _value: str) -> None:
        """Reordena las fotos en vivo al cambiar el criterio de orden."""
        if not self._model.folder_path:
            return
        if self._loading_photos:
            self._pending_sort = True
            self._view.toast("El nuevo orden se aplicará al terminar la carga…", "warn")
            return
        self._apply_sort()

    def _apply_sort(self) -> None:
        """Aplica el orden actual del menú a las fotos y refresca la UI."""
        self._model.sort_mode = self._view.get_sort_mode()
        _save_state({"sort": self._model.sort_mode})
        self._view.set_option_sort(self._view._sort_var.get())
        try:
            n = self._model.load_photos()
        except Exception as exc:
            self._view.set_folder_status("error", str(exc))
            return
        self._view.set_folder_status(
            "ok", f"{n} imagen{'es' if n != 1 else ''}  ·  {self._view._sort_var.get()}")
        self._update_sync_state(notify=False)

    def on_folder_path_changed(self, raw: str) -> None:
        """El usuario escribió otra carpeta (Enter) o la eligió por el diálogo."""
        path = Path(raw)
        if not raw or not path.is_dir():
            if raw:
                self._view.set_folder_status("error", "La ruta no existe.")
                self._update_sync_state(notify=False)
            return
        if self._model.folder_path == path:
            return
        if self._loading_photos:
            self._view.toast("Espera: las fotos se están cargando…", "warn")
            return
        self._model.folder_path = path
        self._model.sort_mode = self._view.get_sort_mode()
        self._model._plan = []
        _save_state({"last_folder": str(path), "sort": self._model.sort_mode})
        self._loading_photos = True
        self._view.set_btn_folder_load(False)
        self._view.set_folder_status("loading", "Cargando fotos...")
        threading.Thread(target=self._load_photos_bg, daemon=True).start()

    def on_excel_path_changed(self, raw: str) -> None:
        """El usuario escribió otro Excel (Enter) o lo eligió por el diálogo."""
        path = Path(raw)
        valid_exts = {".xlsx", ".csv", ".tsv", ".txt"}
        if not raw or not path.is_file() or path.suffix.lower() not in valid_exts:
            if raw:
                self._view.set_excel_status("error", f"Archivo no válido: {path.suffix or 'sin extensión'}")
                self._update_sync_state(notify=False)
            return
        if self._model.excel_path == path:
            return
        if self._loading_excel:
            self._view.toast("Espera: el Excel se está cargando…", "warn")
            return
        self._model.excel_path = path
        self._model.clear_excel_data()
        _save_state({"last_excel": str(path)})
        self._loading_excel = True
        self._view.set_btn_excel_load(False)
        try:
            sheets = self._call_model(self._model.load_sheets,
                                      lambda s, m: self._view.set_excel_status(s, f"Error leyendo Excel: {m}"))
            if sheets is None:
                return
            if len(sheets) > 1:
                self._view.set_excel_status("ok", f"{len(sheets)} hojas — elige una.")
                self._view.show_sheets(sheets)
                self._model.sheet_name = sheets[0]
            else:
                self._model.sheet_name = sheets[0] if sheets else None
                self._view.hide_sheets()
            self._load_columns()
        finally:
            self._loading_excel = False
            self._view.set_btn_excel_load(True)

    def on_simulate(self) -> None:
        """Refresca el plan y el resumen SIN tocar ningún archivo."""
        self._update_sync_state(notify=True)
        self._view.toast("Simulación — no se modificó ningún archivo.", "ok")

    def on_rename(self) -> None:
        if self._guard(not (self._model.photos and self._model.names),
                       "Carga las fotos y el Excel primero."): return

        # Con la vista previa en segundo plano puede haber un recálculo en
        # vuelo: no bloquear la UI reconstruyendo el plan aquí (atajo Ctrl+Enter).
        if self._view._sum_state_lbl.cget("text").startswith("Calculando"):
            self._view.toast("Espera: se está calculando la vista previa…", "warn")
            return

        blocked, reason = self._model.rename_blocked()
        if blocked:
            self._view.toast(f"Renombrado bloqueado: {reason}", "error")
            self._update_sync_state(notify=False)
            return

        plan = self._model._plan
        if plan:
            will = sum(1 for it in plan if it["state"] == "ok")
        else:
            will = len(self._model.photos)
        copy = self._view.get_copy_mode()
        verb = "copiarán" if copy else "renombrarán"
        msg  = (f"Se {verb} {will} fotografía{'s' if will != 1 else ''}.\n"
                f"Esta operación modificará los nombres de los archivos.\n"
                f"¿Desea continuar?")

        if not self._view.confirm("¿Confirmar renombramiento?", msg,
                                  ok_text="Renombrar"):
            return

        self._cancel_ev = threading.Event()
        self._view.update_steps(4)
        self._view.append_log_line(f"Iniciando renombrado de {will} archivo(s)…")
        self._view.set_btn_rename(False, "Renombrando…")
        self._view.set_btn_cancel(True)
        self._view.set_progress(0, "Iniciando…")
        threading.Thread(target=lambda: self._run_safely(self._do_rename),
                         daemon=True).start()

    def on_undo(self) -> None:
        if self._guard(not self._model.has_undo, "No hay nada que deshacer.", "warn"): return
        if not self._view.confirm("¿Deshacer?", "¿Revertir el último lote de renombramientos?"):
            return
        self._view.set_btn_undo(False)
        self._view.set_progress(0, "Deshaciendo…")
        threading.Thread(target=lambda: self._run_safely(self._do_undo),
                         daemon=True).start()

    def on_cancel(self) -> None:
        if self._cancel_ev:
            self._cancel_ev.set()
            self._view.toast("Cancelando…", "warn")
            self._view.set_btn_cancel(False)

    def on_export_log(self) -> None:
        self._export(self._model.export_log, "log", "log", "txt")

    def on_export_csv(self) -> None:
        self._export(self._model.export_preview_csv, "vista previa", "preview", "csv")

    def _export(self, export_fn, label: str, prefix: str, ext: str) -> None:
        if not self._last_pairs:
            self._view.toast(f"No hay {label} para exportar.", "warn"); return
        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = self._view.ask_save_folder(f"{prefix}_{ts}.{ext}")
        if not dest: return
        try:
            export_fn(self._last_pairs, Path(dest))
            self._view.toast(f"{label.title()} guardado: {Path(dest).name}", "ok")
        except Exception as exc:
            self._view.toast(f"Error: {exc}", "error")

    def _apply_theme(self, theme: str) -> None:
        """Aplica un tema canónico de MetaTag y reconstruye la vista.

        No guarda estado ni muestra toast. Si el nombre no es un tema
        canónico (p. ej. un valor antiguo de un estado previo) se normaliza
        al tema por defecto en vez de fallar.
        """
        global C, CURRENT_THEME
        if theme not in THEME_ORDER:
            theme = DEFAULT_THEME
        CURRENT_THEME = theme
        C = _THEME_ADAPTER.palette(theme)
        _refresh_button_constants()
        if self._dup_recalc_job:
            try:
                self._view.after_cancel(self._dup_recalc_job)
            except Exception:
                pass
            self._dup_recalc_job = None
        self._view.rebuild_theme()
        self._resync_after_rebuild()

    def _resync_after_rebuild(self) -> None:
        """Re-renderiza la vista previa y el filtro tras reconstruir la UI."""
        if self._last_pairs:
            self._view.render_preview(self._last_pairs)
            self._view.set_edit_mode(self._view.get_edit_mode())
            query = self._view._filter_var.get()
            if query:
                self._view.filter_preview(query)

    def on_theme_change(self, theme: str) -> None:
        self._apply_theme(theme)
        _save_state({"theme": CURRENT_THEME})
        self._view.toast(f"Tema cambiado a {CURRENT_THEME}.", "ok")

    # ── helpers ───────────────────────────────────────────────────────────
    _EXCEL_ERRORS = {
        "InvalidFileError": "El archivo no es un .xlsx válido.",
        "BadZipFile": "El archivo parece estar corrupto o no es un Excel.",
        "FileNotFoundError": "El archivo no fue encontrado.",
        "PermissionError": "No hay permiso para leer el archivo.",
    }

    def _call_model(self, fn, set_status, *args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except Exception as exc:
            friendly = self._EXCEL_ERRORS.get(type(exc).__name__, str(exc))
            set_status("error", friendly)
            return None

    def _load_columns(self) -> None:
        cols = self._call_model(self._model.load_columns, self._view.set_excel_status)
        if cols is None: return

        if len(cols) == 1:
            self._model.column_name = cols[0]
            self._view.hide_columns()
            self._view.set_excel_status("ok", f"Excel cargado · columna «{cols[0]}» auto-seleccionada.")
            self._load_names_and_preview()
        else:
            # Auto-selección de la primera columna: `show_columns` fija el
            # menú con var.set() (que NO dispara el command de CTkOptionMenu),
            # por eso la selección se aplica aquí en el controlador.
            self._model.column_name = cols[0]
            self._view.show_columns(cols)
            self._view.set_excel_status(
                "ok", f"{len(cols)} columnas — usando «{cols[0]}» (puedes cambiarla).")
            self._load_names_and_preview()

    def _load_names_and_preview(self) -> None:
        n = self._call_model(self._model.load_names, self._view.set_excel_status)
        if n is None: return
        self._view.set_excel_status(
            "ok", f"Excel · {n} nombre{'s' if n!=1 else ''} en «{self._model.column_name}».")

        # FIX #6: si alguna fila del Excel quedó vacía/no leíble, decir EXACTAMENTE cuál
        if self._model.skipped_rows:
            rows = self._model.skipped_rows
            rows_txt = ", ".join(str(r) for r in rows[:10])
            extra = f" y {len(rows) - 10} más" if len(rows) > 10 else ""
            self._view.toast(
                f"⚠ Fila(s) de Excel {rows_txt}{extra} sin texto legible en esa columna.\n"
                f"Revísala(s): puede ser una celda vacía o una fórmula cuyo valor "
                f"no quedó guardado al exportar el archivo.", "warn")

        self._update_sync_state(notify=False)

    def _update_sync_state(self, notify: bool = False) -> None:
        """ÚNICA ruta central de recálculo.

        Decide el estado completo de la interfaz a partir del modelo:
        - vista previa (o mensaje de qué falta exactamente),
        - panel de resumen (Fotografías / Registros / Correspondencias /
          Conflictos / Estado),
        - estado de los botones (renombrar / simular / csv / log).

        Todos los handlers (columna, orden, ruta, matching, cargas…) terminan
        aquí, por lo que cualquier cambio de datos se refleja de inmediato.

        El cálculo del plan (build_preview, en especial el MATCHING seguro
        que compara similitud fotos × nombres) corre en un HILO DE FONDO
        para no congelar la interfaz; `_update_sync_state_finish` aplica el
        resultado. Un contador de generación descarta resultados obsoletos.
        """
        m = self._model

        # ── estado incompleto: aún no hay correspondencias que mostrar ─────
        if not (m.photos and m.names and m.column_name):
            # Invalida cualquier cálculo de fondo aún en vuelo.
            self._sync_gen += 1
            n_ph, n_nm = len(m.photos), len(m.names)
            if n_ph and not n_nm:
                empty_msg = ("Carga el archivo Excel y elige la columna con "
                             "los nombres en la sección 2.")
            elif n_nm and not n_ph:
                empty_msg = "Carga la carpeta de fotografías (sección 1)."
            else:
                empty_msg = ("Carga las fotografías y el archivo Excel para "
                             "ver la vista previa.")
            self._view.render_preview([], empty_message=empty_msg)
            self._view.update_summary(
                n_ph or None, n_nm or None, None, None,
                "Sin correspondencias", "idle")
            self._update_steps_if_allowed(0)
            self._view.set_btn_rename(False, "▶  Renombrar todo")
            self._view.set_btn_sim(False)
            self._view.set_btn_log(False)
            self._view.set_btn_csv(False)
            self._last_pairs = []
            return

        # Nueva generación: los resultados de hilos anteriores quedan viejos.
        self._sync_gen += 1
        gen = self._sync_gen
        self._view.render_preview([], empty_message="Calculando vista previa…")
        self._view.update_summary(len(m.photos), len(m.names),
                                  None, None, "Calculando…", "idle")
        self._update_steps_if_allowed(1)
        self._view.set_btn_rename(False, "Calculando…")
        self._view.set_btn_sim(False)
        self._view.set_btn_log(False)
        self._view.set_btn_csv(False)

        def _bg() -> None:
            try:
                plan = m._build_plan()
            except Exception as exc:
                # Sin cálculo nuevo: restaurar el último plan conocido para no
                # dejar la UI atascada en "Calculando…".
                plan_old = m._plan if m._plan is not None else []
                pairs_old = list(self._last_pairs) if self._last_pairs else []
                self._view.after(0, lambda: self._view.toast(str(exc), "error"))
                self._view.after(0, lambda: self._update_sync_state_finish(
                    plan_old, pairs_old, gen, notify))
                return
            pairs: list[tuple[str, str, Optional[Path], bool, str]] = []
            for item in plan:
                src = item["src"]
                pairs.append((src.name if src else "—", item["new"], src,
                              item["is_dup"], item["state"]))
            self._view.after(0, lambda: self._update_sync_state_finish(
                plan, pairs, gen, notify))

        threading.Thread(target=_bg, daemon=True).start()

    def _update_sync_state_finish(
        self,
        plan: list[dict],
        pairs: list[tuple[str, str, Optional[Path], bool, str]],
        gen: int,
        notify: bool,
    ) -> None:
        """Aplica (en el hilo de la UI) el resultado de un cálculo de fondo.
        Si ya se lanzó un recálculo más nuevo, se descarta este resultado."""
        m = self._model
        if gen != self._sync_gen:
            return  # resultado obsoleto
        m._plan = plan
        self._last_pairs = pairs
        self._view.render_preview(pairs)

        # ── Modelo único de contadores ──────────────────────────────────
        # TODOS los contadores derivan de UNA sola pasada sobre el plan.
        # Definiciones formales:
        #
        # Fotografías   = archivos de imagen válidos en la carpeta.
        # Registros     = nombres válidos extraídos del Excel/CSV.
        # Correspondencias = registros con una foto asociada (src != None),
        #   sin importar si se renombrarán o no. Equivale a:
        #   ok + ya_correcto + existe + conflicto + duplicado.
        # Conflictos     = correspondencias en estado bloqueante:
        #   existe + conflicto + duplicado + ambiguo + error.
        #
        # Invariante: Registros == Correspondencias + (not_found + sin_foto
        #   + ambiguo + error).  Los estados son mutuamente excluyentes.
        _BLOCKING = ("existe", "conflicto", "duplicado", "ambiguo", "error")
        state_counter = Counter(item["state"] for item in plan)
        n_corr = sum(1 for item in plan if item["src"] is not None)
        n_conf = sum(state_counter[s] for s in _BLOCKING)
        faltan = len(m.names) - n_corr
        blocked, reason = m.rename_blocked(plan)
        if n_conf:
            kind = "warn"
            estado = (f"⚠ {n_conf} conflicto(s) sin resolver"
                      if not blocked else reason)
        elif blocked:
            if faltan > 0:
                kind = "warn"
                estado = (f"⚠ Correspondencia incompleta: faltan "
                          f"{faltan} fotografías")
            else:
                kind = "error"
                estado = reason
        elif faltan > 0:
            kind = "warn"
            estado = f"⚠ {faltan} registros sin fotografía (se omitirán)"
        else:
            kind = "ok"
            estado = "✓ Listo para renombrar"
        self._view.update_summary(len(m.photos), len(m.names),
                                  n_corr, n_conf, estado, kind)

        can_act = not blocked
        done = 1 if not pairs else (3 if can_act else 2)
        self._update_steps_if_allowed(done)
        self._view.set_btn_rename(can_act, "▶  Renombrar todo")
        self._view.set_btn_sim(True)
        self._view.set_btn_log(True)
        self._view.set_btn_csv(True)

        dup_count = sum(1 for _, _, _, is_dup, _ in pairs if is_dup)
        if notify and dup_count:
            self._view.toast(
                f"⚠ {dup_count} nombre{'s' if dup_count != 1 else ''} duplicado"
                f"{'s' if dup_count != 1 else ''} — se omitirán al renombrar.",
                "warn")
        if notify and blocked:
            self._view.toast(f"Renombrado bloqueado: {reason}", "warn")

    def _refresh_preview(self) -> None:
        """Compatibilidad: delegación a la ruta central (sin notificar)."""
        self._update_sync_state(notify=False)

    # ── threads ───────────────────────────────────────────────────────────
    def _run_safely(self, fn: Callable[[], None]) -> None:
        """
        Ejecuta ``fn`` (el trabajo de un hilo de fondo). Si lanza una
        excepción no controlada, en vez de morir en silencio y dejar la
        app trabada (botones deshabilitados para siempre, barra de
        progreso congelada), la reporta con el mismo diálogo de error de
        la UI y reactiva los controles. FIX #5.
        """
        try:
            fn()
        except Exception as exc:
            tb_str = traceback.format_exc()
            log.error("Excepción no controlada en hilo de fondo:\n%s", tb_str)
            self._view.after(0, self._view._show_error_dialog,
                             f"{type(exc).__name__}: {exc}", tb_str)
            self._view.after(0, self._unstick_ui)

    def _unstick_ui(self) -> None:
        """Reactiva los controles si un hilo de fondo murió a mitad de camino."""
        self._loading_photos = False
        self._loading_excel = False
        self._pending_sort = False
        self._view.set_btn_folder_load(True)
        self._view.set_btn_excel_load(True)
        self._view.set_btn_rename(True, "▶  Renombrar todo")
        self._view.set_btn_cancel(False)
        self._view.set_btn_undo(self._model.has_undo)
        self._view.set_progress(0, "Ocurrió un error — revisa el detalle.")

    # ── helpers compartidos para operaciones async ──────────────────────────
    def _reload_photos_and_preview(self) -> None:
        if self._model.folder_path:
            try:
                self._model.load_photos()
            except Exception:
                pass
        self._refresh_preview()

    def _run_async(self, model_fn, progress_prefix: str, finish_fn, **model_kwargs) -> None:
        def progress(cur, tot, name):
            frac = cur / tot if tot else 0
            label = f"{progress_prefix}{cur}/{tot}  —  {name}" if progress_prefix else f"{cur}/{tot}  —  {name}"
            self._view.after(0, self._view.set_progress, frac, label)
        def done(ok, errors):
            self._view.after(0, finish_fn, ok, errors)
        model_fn(progress, done, **model_kwargs)

    def _finish_operation(self, progress_val: float, progress_text: str,
                          btn_setup, toast_fn) -> None:
        self._view.set_progress(progress_val, progress_text)
        btn_setup()
        self._reload_photos_and_preview()
        toast_fn()

    # ── rename / undo ──────────────────────────────────────────────────────
    def _do_rename(self) -> None:
        plan = getattr(self._model, "_plan", None) or None
        self._run_async(self._model.rename_all, "Renombrando fotografías  ",
                        self._finish_rename,
                        cancel_ev=self._cancel_ev, copy_mode=self._view.get_copy_mode(),
                        plan=plan, on_log=self._log_bg_line)

    def _do_undo(self) -> None:
        self._run_async(self._model.undo_last, "Revirtiendo ", self._finish_undo,
                        on_log=self._log_bg_line)

    # ── panel de registro (buffer hilo de fondo → UI) ──────────────────────
    def _log_bg_line(self, line: str) -> None:
        """Acepta líneas desde hilos de fondo: se acumulan y se vuelcan a la UI
        en lotes (evita inundar Tk con miles de after())."""
        self._log_pending.append(line)
        if len(self._log_pending) >= 20:
            lines, self._log_pending = self._log_pending, []
            self._view.after(0, self._view.append_log_lines, lines)

    def _flush_log(self) -> None:
        if self._log_pending:
            lines, self._log_pending = self._log_pending, []
            try:
                self._view.append_log_lines(lines)
            except Exception:
                pass

    # ── indicador de pasos ─────────────────────────────────────────────────
    def _update_steps_if_allowed(self, done: int) -> None:
        if not self._step_pinned:
            self._view.update_steps(done)

    def _unpin_steps(self) -> None:
        self._step_pinned = False

    # ── backup JSON de la última operación ─────────────────────────────────
    def _write_backup(self) -> None:
        try:
            if not self._view.get_backup_option():
                return
            stack = getattr(self._model, "_undo_stack", None)
            if not stack:
                return
            batch, _dest_folder, copy_mode = stack[-1]
            if not batch:
                return
            folder = self._model.folder_path
            if not folder or not folder.is_dir():
                return
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest = folder / f".metatag_backup_{ts}.json"
            self._model.write_backup(dest, batch, copy_mode)
            self._view.append_log_line(f"Backup: {dest.name}")
        except Exception as exc:
            log.warning("No se pudo escribir el backup: %s", exc)

    def on_edit_mode_change(self, enabled: bool) -> None:
        self._view.set_edit_mode(enabled)

    def on_keep_ext_change(self, enabled: bool) -> None:
        """Refleja «Mantener extensión original» en el modelo y recalcula."""
        self._model.keep_extension = enabled
        self._update_sync_state(notify=False)

    def on_backup_change(self, enabled: bool) -> None:
        """Nada que calcular: solo se consulta al finalizar una operación."""
        if not enabled:
            self._view.toast("Backup desactivado: no se creará el registro JSON.", "warn")

    def on_open_folder_change(self, enabled: bool) -> None:
        if not enabled:
            self._view.toast("Al finalizar ya no se abrirá la carpeta.", "warn")

    def matching_available(self) -> bool:
        return self._model.matching_available

    def on_matching_toggle(self) -> None:
        """Activa/desactiva el emparejamiento seguro (nombre → SU foto)."""
        enabled = self._view.get_matching_mode()
        self._model.matching_mode = enabled
        if enabled:
            self._view.toast(
                "Matching seguro: ON — cada nombre del Excel buscará SU "
                "fotografía. Si no hay coincidencia clara, la fila se omitirá.",
                "ok")
        else:
            self._view.toast(
                "Modo posicional: la 1ª foto ↔ 1er nombre (puede no coincidir "
                "la persona).", "warn")
        self._update_sync_state(notify=True)

    def on_preview_name_changed(self, index: int, new_name: str, photo_path: Path) -> None:
        if 0 <= index < len(self._model._names):
            self._model._names[index] = new_name
        if hasattr(self, '_dup_recalc_job') and self._dup_recalc_job:
            self._view.after_cancel(self._dup_recalc_job)
        self._dup_recalc_job = self._view.after(150, self._recalc_dups)

    def _recalc_dups(self) -> None:
        self._dup_recalc_job = None
        pairs = self._model.build_preview()
        self._last_pairs = pairs
        self._view._preview.update_dup_states(pairs)

    def _open_folder(self, path: Path) -> None:
        try:
            if platform.system() == "Windows":
                os.startfile(str(path))
            elif platform.system() == "Darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except Exception:
            self._view.toast("No se pudo abrir la carpeta.", "warn")

    def _finish_rename(self, ok: int, errors: list[str]) -> None:
        folder = self._model.folder_path
        total = ok + len(errors)
        if errors:
            text = f"{ok}/{total} · {len(errors)} con error"
        else:
            text = f"{ok}/{total} · ✓ Renombrado completado"
        self._write_backup()
        def btn_setup():
            self._view.set_btn_rename(True, "▶  Renombrar todo")
            self._view.set_btn_cancel(False)
            self._view.set_btn_undo(True, len(self._model._undo_stack))
        def toast_fn():
            if errors:
                self._view.toast(f"⚠ {ok} OK · {len(errors)} con error.", "warn")
            elif folder and self._view.get_open_folder_option():
                self._view.toast(f"✓ {ok} foto{'s' if ok!=1 else ''} renombradas.", "ok")
                self._open_folder(folder)
            else:
                self._view.toast(f"✓ {ok} foto{'s' if ok!=1 else ''} renombradas.", "ok")
        self._step_pinned = True
        self._view.update_steps(5)
        self._view.after(4000, self._unpin_steps)
        self._finish_operation(1.0, text, btn_setup, toast_fn)
        self._flush_log()
        self._view.append_log_line(f"{ok} correctos · {len(errors)} errores")

    def _finish_undo(self, ok: int, errors: list[str]) -> None:
        def btn_setup():
            count = len(self._model._undo_stack) if self._model.has_undo else 0
            self._view.set_btn_undo(self._model.has_undo, count)
        def toast_fn():
            if errors:
                self._view.toast(f"↩ {ok} OK · {len(errors)} con error.", "warn")
            else:
                self._view.toast(f"↩ {ok} revertida{'s' if ok!=1 else ''}.", "ok")
        self._step_pinned = True
        self._view.update_steps(3)
        self._view.after(4000, self._unpin_steps)
        self._finish_operation(0, f"Deshacer completo · {ok} revertidas.", btn_setup, toast_fn)
        self._flush_log()


# ===========================================================================
#  ENTRYPOINT
# ===========================================================================
if __name__ == "__main__":
    # Neutralizar XIM↔iBus ANTES de crear el primer Tk() (FASE 3B.1): el
    # Renombrador se beneficia al lanzarse desde MetaTag por herencia de
    # entorno, y también al ejecutarse de forma independiente.
    try:
        from metatag_xim import neutralize_xim_for_tk
        neutralize_xim_for_tk()
    except Exception:
        pass
    try:
        app = AppController()
        app.run()
    except Exception:
        tb_str = traceback.format_exc()
        log.critical("Fallo crítico al iniciar la aplicación:\n%s", tb_str)

        crash_log: Optional[Path] = None
        try:
            crash_log = Path.home() / "renombrador_fotos_error.log"
            crash_log.write_text(tb_str, encoding="utf-8")
        except Exception:
            crash_log = None

        try:
            import tkinter as _tk
            from tkinter import messagebox as _mb
            _root = _tk.Tk()
            _root.withdraw()
            _mb.showerror(
                "Error al iniciar — Renombrador de Fotos",
                "La aplicación no pudo iniciarse.\n\n"
                f"{tb_str[-700:]}\n\n"
                + (f"Detalle completo guardado en:\n{crash_log}" if crash_log else "")
            )
            _root.destroy()
        except Exception:
            print(tb_str)