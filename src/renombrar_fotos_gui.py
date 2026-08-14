"""
renombrar_fotos_gui.py — Renombrador de Fotos v4.0

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
import traceback
from collections import OrderedDict
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

# Estados posibles de cada fila de la vista previa / plan de renombrado.
PLAN_STATES = ("ok", "ya_correcto", "conflicto", "duplicado",
               "not_found", "ambiguo", "error")
STATE_LABELS = {
    "ok": "", "ya_correcto": "Ya correcto", "conflicto": "Conflicto",
    "duplicado": "Duplicado", "not_found": "No encontrada",
    "ambiguo": "Ambiguo", "error": "Error",
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
        self._matcher      = None          # ImageMatcher (carga perezosa)
        self._photos:      list[Path]      = []
        self._names:       list[str]       = []
        self._plan:        list[dict]      = []   # último plan construido
        self.skipped_rows: list[int]       = []   # filas de Excel ignoradas (vacías/no leíbles) — FIX #6
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
        return len(raw)

    def load_sheets(self) -> list[str]:
        """Retorna las hojas del Excel o None para CSV."""
        if not self.excel_path:
            raise ValueError("No se ha seleccionado ningún archivo.")
        ext = self.excel_path.suffix.lower()
        if ext in (".csv", ".tsv", ".txt"):
            return []
        return pd.ExcelFile(self.excel_path).sheet_names

    def load_columns(self) -> list[str]:
        """Retorna las columnas de la hoja activa o del CSV."""
        if not self.excel_path:
            raise ValueError("No se ha seleccionado ningún archivo.")
        ext = self.excel_path.suffix.lower()
        if ext in (".csv", ".tsv", ".txt"):
            sep = "\t" if ext == ".tsv" else ","
            df = pd.read_csv(self.excel_path, sep=sep, nrows=0, encoding="utf-8")
            return list(df.columns)
        sheet = self.sheet_name or 0
        return list(pd.read_excel(self.excel_path, sheet_name=sheet, nrows=0).columns)

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
          · se guarda en self.skipped_rows el número REAL de fila de Excel
            de cualquier celda que sí haya quedado vacía, para poder
            avisarle al usuario exactamente cuál revisar (en vez de un
            misterioso "40 de 41").
        """
        if not (self.excel_path and self.column_name):
            raise ValueError("Archivo o columna no configurados.")
        ext = self.excel_path.suffix.lower()
        if ext in (".csv", ".tsv", ".txt"):
            sep = "\t" if ext == ".tsv" else ","
            df = pd.read_csv(self.excel_path, sep=sep, keep_default_na=False, encoding="utf-8")
        else:
            sheet = self.sheet_name or 0
            df = pd.read_excel(self.excel_path, sheet_name=sheet, keep_default_na=False)

        names: list[str] = []
        self.skipped_rows = []
        for i, raw in enumerate(df[self.column_name]):
            text = "" if raw is None else str(raw).strip()
            if text == "" or text.lower() in ("nan", "nat", "none"):
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
        """
        seen_dest: dict[str, int] = {}
        plan: list[dict] = []
        for photo, name in zip(self._photos, self._names):
            new_full = name + photo.suffix
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
                "name": name, "row": None,
            })
        self._plan = plan
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

            new_full = name + src.suffix
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
        return plan

    @staticmethod
    def _skip_text(item: dict) -> str:
        """Mensaje legible para filas omitidas (nunca destructivo)."""
        src = item["src"]
        src_name = src.name if src else "—"
        state = item["state"]
        if state == "not_found":
            return f"Omitida: no se encontró fotografía para «{item['name']}»"
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

    # ── renombramiento ─────────────────────────────────────────────────────
    def rename_all(
        self,
        on_progress: Callable[[int, int, str], None],
        on_done: Callable[[int, list[str]], None],
        cancel_ev: Optional[threading.Event] = None,
        copy_mode: bool = False,
        plan: Optional[list[dict]] = None,
    ) -> None:
        """Renombra (o copia) archivos en un hilo secundario.

        Reglas de seguridad (obligatorias):
        - "ya_correcto": el destino es EL MISMO archivo → no se toca, sin error.
        - "conflicto": el destino existe y es OTRO archivo → NO se sobreescribe,
          NO se elimina, NO se modifica (en ningún modo, incluido copia).
        - "not_found" / "ambiguo" / "duplicado" / "error": se omiten, nunca se renombra.
        - Re-chequeo TOCTOU justo antes de tocar disco.
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

        for i, item in enumerate(plan):
            if cancel_ev and cancel_ev.is_set():
                errors.append("Cancelado por el usuario.")
                break
            src      = item["src"]
            new_full = item["new"]
            state    = item["state"]

            if state in ("not_found", "ambiguo", "duplicado", "conflicto",
                         "error"):
                errors.append(self._skip_text(item))
                on_progress(i + 1, total, f"[{state}] {new_full}")
                continue

            if state == "ya_correcto":
                success += 1
                on_progress(i + 1, total, f"[ya correcto] {new_full}")
                continue

            if src is None:
                continue

            new_name = new_full
            if new_name in seen_dest:
                seen_dest[new_name] += 1
                errors.append(f"Saltado duplicado: {src.name} → {new_name} (nombre ya usado)")
                log.warning("DUP  %s → %s (skip)", src.name, new_name)
                on_progress(i + 1, total, f"[duplicado] {new_name}")
                continue
            seen_dest[new_name] = 1

            dest = (dest_folder / new_name) if copy_mode and dest_folder \
                   else (src.parent / new_name)
            if dest.exists():
                if self._same_file(dest, src):
                    success += 1   # el destino es el propio origen → ya correcto
                    on_progress(i + 1, total, f"[ya correcto] {new_name}")
                    continue
                errors.append(f"Saltado: {src.name} → {new_name} (el destino ya existe, no se sobreescribe)")
                log.warning("CONFLICT  %s → %s (exists)", src.name, new_name)
                on_progress(i + 1, total, f"[conflicto] {new_name}")
                continue
            try:
                if copy_mode:
                    shutil.copy2(src, dest)
                else:
                    src.rename(dest)
                batch.append((dest, src))
                success += 1
                log.info("OK  %s → %s", src.name, new_name)
            except PermissionError:
                errors.append(f"{src.name} → {new_name}  (permiso denegado)")
                log.error("PERM  %s", src.name)
            except OSError as exc:
                errors.append(f"{src.name} → {new_name}  ({exc})")
                log.error("ERR %s", exc)
            on_progress(i + 1, total, new_name)

        if batch:
            self._undo_stack.append((batch, dest_folder, copy_mode))
        on_done(success, errors)

    # ── deshacer ───────────────────────────────────────────────────────────
    def undo_last(
        self,
        on_progress: Callable[[int, int, str], None],
        on_done: Callable[[int, list[str]], None],
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
                elif copy_mode:
                    current.unlink()
                    success += 1
                else:
                    if original.exists():
                        if self._same_file(original, current):
                            success += 1   # ya restaurado (mismo archivo)
                        else:
                            errors.append(
                                f"Conflicto al deshacer: no se sobreescribe "
                                f"{original.name} (hay otro archivo en esa ruta)")
                    else:
                        current.rename(original)
                        success += 1
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
            w.writerow(["original", "nuevo_nombre", "duplicado", "estado"])
            for orig, new, _, is_dup, state in pairs:
                w.writerow([orig, new, "Sí" if is_dup else "", state])


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
            x = self._widget.winfo_rootx() + self._widget.winfo_width() + 10
            y = self._widget.winfo_rooty()
            tip.geometry(f"+{x}+{y}")
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
            self._add_row("📂  (esta carpeta)", path, is_dir=True, current=True)

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
            label = f"📁  {e.name}" if e.is_dir() else f"📄  {e.name}"
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

        ctk.CTkButton(self, text="📂  Explorar", width=118, height=38,
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
    Tabla de vista previa con renderizado diferido por lotes.

    Bug #4 corregido: antes se creaban TODOS los widgets de golpe en el
    hilo principal → la UI se congelaba con listas grandes.
    Solución: se renderizan en chunks de 30 filas con after(0) entre cada
    batch, permitiendo que Tk procese eventos intermedios. Resultado: la
    tabla aparece progresivamente sin congelar nada.
    Miniaturas: se cargan en un thread de fondo y se insertan via after().
    """

    CHUNK = 30

    def __init__(self, master, on_name_change: Optional[Callable[[int, str, Path], None]] = None, **kw) -> None:
        super().__init__(master, fg_color=C["surface"], corner_radius=10, **kw)
        self._all_pairs: list[dict[str, object]] = []
        self._rows: list[dict[str, object]] = []
        self._pending: list[tuple[int, dict[str, object]]] = []
        self._chunk_job: Optional[str] = None
        self._thumb_q:   list[tuple[ctk.CTkLabel, Path]] = []
        self._thumb_job: Optional[str] = None
        self._thumb_run: int = 0
        self._thumb_cancel: Optional[threading.Event] = None
        self._edit_mode: bool = False
        self._on_name_change = on_name_change
        self._current_query: str = ""
        self._filter_job: Optional[str] = None

        self._lbl_empty = ctk.CTkLabel(
            self, fg_color="transparent",
            text="Aquí aparecerá la vista previa.\n"
                 "Pasa el cursor sobre una fila para ver la miniatura.",
            font=FONT_MD, text_color=C["subtext"],
            justify="center")
        self._lbl_empty.pack(pady=28)

    # ── API pública ────────────────────────────────────────────────────────
    def render(self, pairs: list[tuple[str, str, Optional[Path], bool, str]]) -> None:
        """Reemplaza el contenido de la tabla de forma no-bloqueante."""
        self._cancel_jobs()
        self._current_query = ""
        self._all_pairs = [
            {
                "orig": orig,
                "new": new,
                "photo_path": photo_path,
                "photo_index": i,
                "is_dup": is_dup,
                "state": state,
            }
            for i, (orig, new, photo_path, is_dup, state) in enumerate(pairs)
        ]
        self._rows = []
        for w in self.winfo_children():
            w.destroy()

        if not self._all_pairs:
            self._lbl_empty = ctk.CTkLabel(
                self, fg_color="transparent",
                text="Sin datos para mostrar.",
                font=FONT_MD,
                text_color=C["subtext"])
            self._lbl_empty.pack(pady=24)
            return

        self._build_header()
        self._thumb_run += 1
        self._pending = list(enumerate(self._all_pairs))
        self._thumb_q = []
        self._load_chunk()

    def filter(self, query: str) -> None:
        """Aplica filtro de texto sobre los pares ya cargados."""
        _safe_cancel_after(self, self._filter_job)
        self._filter_job = self.after(100, lambda: self._apply_filter(query))

    def _apply_filter(self, query: str) -> None:
        self._filter_job = None
        q = query.lower().strip()
        self._current_query = q
        for row in self._rows:
            row["visible"] = not q or q in row["orig_text"] or q in row["new_text"]
        self._refresh_visible_rows()

    def update_dup_states(self, pairs: list[tuple[str, str, Optional[Path], bool, str]]) -> None:
        """Recolorea solo filas cuyo estado o duplicado cambió. No destruye widgets."""
        for row_data in self._rows:
            idx = row_data["photo_index"]
            if idx >= len(pairs):
                continue
            _, _, _, is_dup, state = pairs[idx]
            changed = False
            if row_data.get("is_dup") != is_dup:
                row_data["is_dup"] = is_dup
                changed = True
            if row_data.get("state") != state:
                row_data["state"] = state
                changed = True
            if not changed:
                continue

            bg = C["state_bg"].get(state) if state != "ok" else None
            if bg is None:
                bg = C["dup_bg"] if is_dup else (C["surface"] if idx % 2 == 0 else C["bg"])
            row_data["frame"].configure(fg_color=bg)

            arrow = row_data.get("arrow_widget")
            if arrow is not None:
                arrow.configure(text_color=self._arrow_color(state, is_dup))

            new_widget = row_data.get("new_widget")
            if new_widget is not None:
                color = self._new_color(state, is_dup)
                if isinstance(new_widget, ctk.CTkEntry):
                    new_widget.configure(border_color=color if state != "ok" else C["border"],
                                         text_color=color)
                else:
                    new_widget.configure(text_color=color)

            state_lbl = row_data.get("state_widget")
            if state_lbl is not None:
                state_lbl.configure(text=STATE_LABELS.get(state, ""),
                                    text_color=C["state_fg"].get(state, C["overlay"]))

    # ── internos ──────────────────────────────────────────────────────────
    def _cancel_jobs(self) -> None:
        for attr in ("_chunk_job", "_thumb_job", "_filter_job"):
            _safe_cancel_after(self, getattr(self, attr, None))
            setattr(self, attr, None)
        self._pending = []
        self._thumb_q = []

    # Configuración de columnas maestra (sincronizada header ↔ filas)
    _COL_NUM_W   = 36    # Columna # (índice)
    _COL_ORIG_W  = None  # Columna Original (dinámica, weight=1)
    _COL_ARROW_W = 24    # Columna → (flecha)
    _COL_NEW_W   = None  # Columna Nuevo nombre (dinámica, weight=1)
    _COL_THUMB_W = 52    # Columna miniatura

    def _configure_grid(self, frame: ctk.CTkFrame) -> None:
        """Aplica la misma configuración de grilla a header y filas."""
        frame.columnconfigure(0, minsize=self._COL_NUM_W)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, minsize=self._COL_ARROW_W)
        frame.columnconfigure(3, weight=1)
        frame.columnconfigure(4, minsize=self._COL_THUMB_W)

    def _build_header(self) -> None:
        hdr = ctk.CTkFrame(self, fg_color=C["surface2"], corner_radius=0, height=30)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        self._configure_grid(hdr)
        cols = [
            (0, "#",           "w",  4),
            (1, "Original",    "w",  8),
            (2, "→",           "w",  4),
            (3, "Nuevo nombre","w",  8),
            (4, "Estado",      "w",  8),
        ]
        for col, txt, sticky, padx in cols:
            lbl = ctk.CTkLabel(hdr, text=txt, width=1,
                               font=FONT_SM_BD, text_color=C["accent"],
                               fg_color="transparent")
            lbl.grid(row=0, column=col, sticky=sticky, padx=padx, pady=4)

    @staticmethod
    def _arrow_color(state: str, is_dup: bool) -> str:
        if state != "ok":
            return C["state_fg"].get(state, C["red"])
        return C["red"] if is_dup else C["accent2"]

    @staticmethod
    def _new_color(state: str, is_dup: bool) -> str:
        if state != "ok":
            return C["state_fg"].get(state, C["text"])
        return C["red"] if is_dup else C["text"]

    def _load_chunk(self) -> None:
        chunk, self._pending = self._pending[:self.CHUNK], self._pending[self.CHUNK:]
        for _, row_data in chunk:
            self._add_row(row_data)
        if self._pending:
            self._chunk_job = self.after(5, self._load_chunk)
        else:
            # iniciar carga de miniaturas en segundo plano
            self._chunk_job = None
            self._refresh_visible_rows()
            self._schedule_thumbs()

    def _add_row(self, row_data: dict[str, object]) -> None:
        i = row_data["photo_index"]
        orig = row_data["orig"]
        new = row_data["new"]
        photo_path = row_data["photo_path"]
        is_dup = row_data.get("is_dup", False)
        state = row_data.get("state", "ok")

        bg = C["state_bg"].get(state) if state != "ok" else None
        if bg is None:
            bg = C["dup_bg"] if is_dup else (C["surface"] if i % 2 == 0 else C["bg"])
        row = ctk.CTkFrame(self, fg_color=bg, corner_radius=0, height=32)
        row.pack_propagate(False)
        self._configure_grid(row)

        # Columna #
        ctk.CTkLabel(row, text=f"{i+1:>3}.", width=1,
                     font=FONT_SM, text_color=C["overlay"],
                     fg_color="transparent").grid(row=0, column=0,
                                                  sticky="w", padx=4)

        # Columna Original
        ctk.CTkLabel(row, text=orig, width=1, anchor="w",
                     font=FONT_SM, text_color=C["subtext"],
                     fg_color="transparent").grid(row=0, column=1,
                                                  sticky="w", padx=8)

        # Columna →
        arrow_color = self._arrow_color(state, is_dup)
        arrow_lbl = ctk.CTkLabel(row, text="→", width=1,
                     font=FONT_SM, text_color=arrow_color,
                     fg_color="transparent")
        arrow_lbl.grid(row=0, column=2, sticky="w", padx=4)

        # Columna Nuevo nombre
        new_color = self._new_color(state, is_dup)
        if self._edit_mode:
            new_var = ctk.StringVar(value=new)
            new_widget = ctk.CTkEntry(row, textvariable=new_var,
                                      width=1, height=28, font=FONT_SM,
                                      fg_color=C["surface"],
                                      border_color=C["red"] if is_dup else C["border"],
                                      text_color=new_color)
            new_widget.grid(row=0, column=3, sticky="ew", padx=8)
            def _on_change(*_):
                value = new_var.get()
                safe = Path(value).name if value else value
                if safe != value:
                    new_var.set(safe)
                    return
                row_data["new"] = safe
                row_data["new_text"] = safe.lower()
                if self._on_name_change:
                    self._on_name_change(i, safe, photo_path)
            new_var.trace_add("write", _on_change)
        else:
            new_widget = ctk.CTkLabel(row, text=new, width=1, anchor="w",
                                      font=FONT_SM, text_color=new_color,
                                      fg_color="transparent")
            new_widget.grid(row=0, column=3, sticky="w", padx=8)

        # Columna Estado
        state_lbl = ctk.CTkLabel(row, text=STATE_LABELS.get(state, ""), width=1,
                                 anchor="w", font=FONT_SM,
                                 text_color=C["state_fg"].get(state, C["overlay"]),
                                 fg_color="transparent")
        state_lbl.grid(row=0, column=4, sticky="w", padx=8)

        # Miniaturas: se muestran en el tooltip hover (ya no hay columna dedicada).

        visible = not self._current_query or self._current_query in orig.lower() or self._current_query in row_data.get("new", "").lower()
        row_data["visible"] = visible
        self._rows.append({
            "frame": row,
            "orig_text": orig.lower(),
            "new_text": row_data.get("new", "").lower(),
            "photo_index": i,
            "visible": visible,
            "new_widget": new_widget,
            "arrow_widget": arrow_lbl,
            "state_widget": state_lbl,
            "is_dup": is_dup,
            "state": state,
        })
        if visible:
            row.pack(fill="x")
        if photo_path is not None:
            row._tooltip_path = photo_path
            row.bind("<Enter>", self._lazy_tooltip, add="+")



    def _lazy_tooltip(self, event) -> None:
        row = event.widget
        if not hasattr(row, "_tooltip_path") or hasattr(row, "_tooltip"):
            return
        row._tooltip = ImageTooltip(row, row._tooltip_path)
        row._tooltip._schedule_show()

    # ── miniaturas en segundo plano ────────────────────────────────────────
    def _schedule_thumbs(self) -> None:
        """Lanza un thread que pre-genera las miniaturas y las inserta via after()."""
        if not self._thumb_q:
            return
        if self._thumb_cancel:
            self._thumb_cancel.set()
        self._thumb_cancel = threading.Event()
        cancel = self._thumb_cancel
        run_id = self._thumb_run
        queue = list(self._thumb_q)
        self._thumb_q = []

        def _worker():
            for lbl, path in queue:
                if cancel.is_set():
                    return
                img = _get_thumb(path, (52, 52))
                if img:
                    def _update(l=lbl, im=img, run=run_id):
                        try:
                            if run != self._thumb_run:
                                return
                            if l.winfo_exists():
                                l.configure(image=im)
                                l._thumb_image = im
                        except Exception:
                            pass
                    try:
                        self.after(0, _update)
                    except Exception:
                        pass

        threading.Thread(target=_worker, daemon=True).start()

    def _refresh_visible_rows(self) -> None:
        for row in self._rows:
            row["frame"].pack_forget()
        for row in self._rows:
            if row["visible"]:
                row["frame"].pack(fill="x")

    def set_edit_mode(self, enabled: bool) -> None:
        if self._edit_mode == enabled:
            return
        self._edit_mode = enabled
        if not self._all_pairs:
            return
        query = self._current_query
        self.render([
            (row["orig"], row["new"], row["photo_path"], row.get("is_dup", False),
             row.get("state", "ok"))
            for row in self._all_pairs])
        if query:
            self.filter(query)


# ===========================================================================
#  DIÁLOGO DE CONFIRMACIÓN INTEGRADO
# ===========================================================================
class ConfirmDialog(ctk.CTkToplevel):
    """Reemplaza messagebox.askyesno con un diálogo dark estilizado."""

    def __init__(self, master, title: str, message: str) -> None:
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
        ctk.CTkButton(btn_row, text="Confirmar", width=110, height=32,
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
    def ask(cls, master, title: str, message: str) -> bool:
        dlg = cls(master, title, message)
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
        self.title("MetaTag v8.9 — Renombrador de Fotos")
        self.configure(fg_color=C["bg"])

        # FIX #3: tamaño adaptativo a la pantalla
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = max(780, min(1100, int(sw * 0.68)))
        h  = max(600, min(900,  int(sh * 0.80)))
        px = (sw - w) // 2
        py = (sh - h) // 2
        self.geometry(f"{w}x{h}+{px}+{py}")
        self.minsize(700, 540)

        self.report_callback_exception = self._tk_error_handler
        self._build()
        self._bind_shortcuts()

    # ── construcción ───────────────────────────────────────────────────────
    def _build(self) -> None:
        # header
        hdr = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=0, height=56)
        hdr.pack(fill="x"); hdr.pack_propagate(False)

        ctk.CTkLabel(hdr, text="🖼  Renombrador de Fotos",
                     font=FONT_TITLE,
                     text_color=C["text"]).pack(side="left", padx=22, pady=12)
        ctk.CTkLabel(hdr, text="desde Excel · integrado en MetaTag v8.9",
                     font=FONT_MD,
                     text_color=C["accent"]).pack(side="left")

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
            mode="folder", on_change=lambda _: None)
        self._folder_sel.grid(row=1, column=0, sticky="ew", pady=(0, 5))

        sort_row = ctk.CTkFrame(c, fg_color="transparent")
        sort_row.grid(row=2, column=0, sticky="ew", pady=(0, 3))
        ctk.CTkLabel(sort_row, text="Ordenar por:",
                     font=FONT_SM,
                     text_color=C["subtext"]).pack(side="left")
        self._sort_var = ctk.StringVar(value="Orden numérico")
        _make_option_menu(
            sort_row, self._sort_var, list(SORT_OPTIONS.keys()),
            width=185
        ).pack(side="left", padx=8)

        i1 = ctk.CTkFrame(c, fg_color="transparent")
        i1.grid(row=3, column=0, sticky="ew", pady=(0, 12))
        self._badge_folder = StatusBadge(i1); self._badge_folder.pack(side="left")
        self._lbl_folder = ctk.CTkLabel(i1, text="Selecciona una carpeta",
                                        font=FONT_SM,
                                        text_color=C["subtext"]); self._lbl_folder.pack(side="left", padx=7)
        ctk.CTkButton(i1, text="Cargar fotos", width=108, height=26,
                      font=FONT_SM,
                      **BTN_SECONDARY,
                      command=self._ctrl.on_load_photos).pack(side="right")

        self._div(c, 4)

        # § 2 excel
        self._sec(c, "2 · Archivo Excel", 5)
        self._excel_sel = PathSelector(
            c, placeholder="/home/usuario/nombres.xlsx  o  D:\\nombres.xlsx",
            mode="file", filetypes=[".xlsx", ".csv", ".tsv", ".txt"], on_change=lambda _: None)
        self._excel_sel.grid(row=6, column=0, sticky="ew", pady=(0, 5))

        i2 = ctk.CTkFrame(c, fg_color="transparent")
        i2.grid(row=7, column=0, sticky="ew", pady=(0, 6))
        self._badge_excel = StatusBadge(i2); self._badge_excel.pack(side="left")
        self._lbl_excel = ctk.CTkLabel(i2, text="Selecciona el archivo Excel",
                                       font=FONT_SM,
                                       text_color=C["subtext"]); self._lbl_excel.pack(side="left", padx=7)
        ctk.CTkButton(i2, text="Cargar Excel", width=108, height=26,
                      font=FONT_SM,
                      **BTN_SECONDARY,
                      command=self._ctrl.on_load_excel).pack(side="right")

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

        self._div(c, 10)

        # § 3 vista previa
        self._sec(c, "3 · Vista previa", 11)

        filter_row = ctk.CTkFrame(c, fg_color="transparent")
        filter_row.grid(row=12, column=0, sticky="ew", pady=(0, 5))
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

        self._preview = PreviewTable(c, on_name_change=self._ctrl.on_preview_name_changed)
        self._preview.grid(row=13, column=0, sticky="ew", pady=(2, 12))

        self._div(c, 14)

        # § 4 renombrar
        self._sec(c, "4 · Renombrar", 15)
        self._progress = ctk.CTkProgressBar(c, height=9,
                                             fg_color=C["surface"],
                                             progress_color=C["accent"])
        self._progress.grid(row=16, column=0, sticky="ew", pady=(3, 3))
        self._progress.set(0)

        act = ctk.CTkFrame(c, fg_color="transparent")
        act.grid(row=17, column=0, sticky="ew", pady=(0, 12))

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

        self._btn_cancel = ctk.CTkButton(
            act, text="✖ Cancelar", width=100, height=36,
            font=FONT_SM,
            **BTN_DANGER,
            command=self._ctrl.on_cancel, state="disabled")
        self._btn_cancel.pack(side="right", padx=(0, 6))

        self._btn_log = ctk.CTkButton(
            act, text="💾 Log", width=80, height=36,
            font=FONT_SM,
            fg_color=C["surface2"], hover_color=C["accent2"],
            command=self._ctrl.on_export_log, state="disabled")
        self._btn_log.pack(side="right", padx=(0, 6))

        self._btn_csv = ctk.CTkButton(
            act, text="📤 CSV", width=80, height=36,
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
        self._filter_var.set(state["filter"])

    # ── API para el Controller ─────────────────────────────────────────────
    def get_folder_path(self) -> str:   return self._folder_sel.get()
    def get_excel_path(self) -> str:    return self._excel_sel.get()
    def get_column(self) -> str:        return self._col_var.get()
    def get_sort_mode(self) -> str:     return SORT_OPTIONS.get(self._sort_var.get(), "natural")
    def get_copy_mode(self) -> bool:    return self._copy_var.get()
    def get_matching_mode(self) -> bool: return self._match_var.get()
    def get_edit_mode(self) -> bool:    return self._edit_var.get()

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

    def render_preview(self, pairs: list[tuple[str, str, Optional[Path], bool, str]]) -> None:
        self._preview.render(pairs)

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

    def toast(self, msg: str, kind: str = "ok") -> None:
        Toast(self, msg, kind)

    def confirm(self, title: str, msg: str) -> bool:
        return ConfirmDialog.ask(self, title, msg)

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
                                font=ctk.CTkFont("Courier New", 10),
                                fg_color=C["surface"], text_color=C["subtext"])
        tb_box.pack(padx=18, pady=(0, 10), fill="both", expand=True)
        tb_box.insert("end", details[:3000])
        tb_box.configure(state="disabled")

        btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_row.pack(pady=(0, 16))

        def _copy():
            self.clipboard_clear(); self.clipboard_append(details)
            self.toast("Detalles copiados al portapapeles.")

        ctk.CTkButton(btn_row, text="📋 Copiar detalles", width=140,
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
        raw = self._view.get_folder_path()
        if self._guard(not raw, "Selecciona primero una carpeta."): return
        path = Path(raw)
        if not path.is_dir():
            self._view.set_folder_status("error", "La ruta no existe."); return

        self._model.folder_path = path
        self._model.sort_mode   = self._view.get_sort_mode()
        _save_state({"last_folder": str(path), "sort": self._model.sort_mode})
        self._view.set_folder_status("loading", "Cargando fotos...")
        threading.Thread(target=self._load_photos_bg, daemon=True).start()

    def _load_photos_bg(self) -> None:
        def on_done(n, err):
            if err:
                self._view.set_folder_status("error", err)
                return
            if n == 0:
                self._view.set_folder_status("warn", "No se encontraron imágenes.")
            else:
                sort_label = self._view._sort_var.get()
                self._view.set_folder_status("ok", f"{n} imagen{'es' if n!=1 else ''}  ·  {sort_label}")
            self._refresh_preview()
        def _work():
            try:
                n = self._model.load_photos()
                self._view.after(0, on_done, n, None)
            except Exception as exc:
                self._view.after(0, on_done, 0, str(exc))
        self._run_safely(_work)

    def on_load_excel(self) -> None:
        raw = self._view.get_excel_path()
        if self._guard(not raw, "Selecciona primero un archivo Excel."): return
        path = Path(raw)
        valid_exts = {".xlsx", ".csv", ".tsv", ".txt"}
        if not path.is_file() or path.suffix.lower() not in valid_exts:
            self._view.set_excel_status("error", f"Formato no soportado: {path.suffix}"); return

        self._model.excel_path = path
        _save_state({"last_excel": str(path)})
        sheets = self._call_model(self._model.load_sheets,
                                  lambda s, m: self._view.set_excel_status(s, f"Error leyendo Excel: {m}"))
        if sheets is None: return

        if len(sheets) > 1:
            self._view.set_excel_status("ok", f"{len(sheets)} hojas — elige una.")
            self._view.show_sheets(sheets)
            self._model.sheet_name = sheets[0]
        else:
            self._model.sheet_name = sheets[0] if sheets else None
            self._view.hide_sheets()
        self._load_columns()

    def on_sheet_selected(self, sheet: str) -> None:
        self._model.sheet_name = sheet
        self._load_columns()

    def on_column_selected(self, column: str) -> None:
        self._model.column_name = column
        self._load_names_and_preview()

    def on_filter_change(self, query: str) -> None:
        self._view.filter_preview(query)

    def on_rename(self) -> None:
        if self._guard(not (self._model.photos and self._model.names),
                       "Carga las fotos y el Excel primero."): return

        n_ph = len(self._model.photos)
        n_nm = len(self._model.names)
        will = min(n_ph, n_nm)
        msg  = f"Se renombrarán {will} foto{'s' if will!=1 else ''}."
        if n_nm < n_ph:
            msg += f"\n⚠ Hay {n_ph - n_nm} foto(s) sin nombre — solo se renombrarán las primeras {will}."

        if not self._view.confirm("¿Confirmar renombramiento?", msg):
            return

        self._cancel_ev = threading.Event()
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
            self._view.show_columns(cols)
            self._view.set_excel_status("ok", f"{len(cols)} columnas — elige cuál tiene los nombres.")

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

        self._refresh_preview()

    def _refresh_preview(self) -> None:
        if not (self._model.photos and self._model.names):
            return
        n_ph, n_nm = len(self._model.photos), len(self._model.names)

        # FIX #6: verificación de desajuste (solo modo posicional: ahí sí el
        # conteo debe cuadrar). En matching seguro cada nombre busca SU foto.
        if not self._model.matching_mode and n_ph != n_nm:
            diff = abs(n_ph - n_nm)
            if n_ph > n_nm:
                log.warning("Hay %d foto(s) sin nombre en el Excel.", diff)
                self._view.toast(
                    f"⚠ {n_ph} fotos pero solo {n_nm} nombres en el Excel.\n"
                    f"Se renombrarán las primeras {min(n_ph,n_nm)}.", "warn")
            else:
                log.warning("Hay %d nombre(s) en el Excel sin foto correspondiente.", diff)
                self._view.toast(
                    f"⚠ {n_nm} nombres pero solo {n_ph} fotos.\n"
                    f"Se usarán los primeros {min(n_ph,n_nm)} nombres.", "warn")

        pairs = self._model.build_preview()
        self._last_pairs = pairs
        self._view.render_preview(pairs)
        self._view.set_btn_rename(True, "▶  Renombrar todo")
        self._view.set_btn_log(True)
        self._view.set_btn_csv(True)

        dup_count = sum(1 for _, _, _, is_dup, _ in pairs if is_dup)
        if dup_count > 0:
            self._view.toast(
                f"⚠ {dup_count} nombre{'s' if dup_count != 1 else ''} duplicado{'s' if dup_count != 1 else ''} "
                f"— se omitirán al renombrar.", "warn")

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
        self._run_async(self._model.rename_all, "", self._finish_rename,
                        cancel_ev=self._cancel_ev, copy_mode=self._view.get_copy_mode(),
                        plan=plan)

    def _do_undo(self) -> None:
        self._run_async(self._model.undo_last, "Revirtiendo ", self._finish_undo)

    def on_edit_mode_change(self, enabled: bool) -> None:
        self._view.set_edit_mode(enabled)

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
        if self._model.photos and self._model.names:
            self._refresh_preview()

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
        def btn_setup():
            self._view.set_btn_rename(True, "▶  Renombrar todo")
            self._view.set_btn_cancel(False)
            self._view.set_btn_undo(True, len(self._model._undo_stack))
        def toast_fn():
            if errors:
                self._view.toast(f"⚠ {ok} OK · {len(errors)} con error.", "warn")
            elif folder:
                self._view.toast(f"✓ {ok} foto{'s' if ok!=1 else ''} renombradas.", "ok")
                self._open_folder(folder)
            else:
                self._view.toast(f"✓ {ok} foto{'s' if ok!=1 else ''} renombradas.", "ok")
        self._finish_operation(1.0, f"Completado · {ok} renombradas.", btn_setup, toast_fn)

    def _finish_undo(self, ok: int, errors: list[str]) -> None:
        def btn_setup():
            count = len(self._model._undo_stack) if self._model.has_undo else 0
            self._view.set_btn_undo(self._model.has_undo, count)
        def toast_fn():
            if errors:
                self._view.toast(f"↩ {ok} OK · {len(errors)} con error.", "warn")
            else:
                self._view.toast(f"↩ {ok} revertida{'s' if ok!=1 else ''}.", "ok")
        self._finish_operation(0, f"Deshacer completo · {ok} revertidas.", btn_setup, toast_fn)


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