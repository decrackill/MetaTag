"""
metatag_matching.py — Motor de emparejamiento seguro (puro, sin Tkinter).

Port FIEL del algoritmo ``_find_image_ex`` de ``metatag_v8.MetaTagApp``
(validado contra el dataset real: 269 nombres, 267 correspondencias y
distribución 18 nombre-exacto / 184 stem-exacto / 21 normalize /
44 id-suffix, 2 not_found conocidas, 0 ambigüedades). NO introduce un
método de semejanza nuevo: cada nombre busca SU fotografía siguiendo
exactamente la misma jerarquía de pasos que MetaTag.

Empareja cada nombre de un listado (Excel/CSV) con SU fotografía en una
carpeta. Es "seguro" porque NUNCA elige un candidato ambiguo: si varios
archivos compiten por la misma clave de fallback, devuelve 'ambiguous'
con la lista de candidatos para que la interfaz decida.

Uso (desde MetaTag o el Renombrador standalone):
    matcher = ImageMatcher()
    path, status, candidates = matcher.find_image_ex("Juan Pérez", "/fotos")
    # status ∈ {"ok", "not_found", "ambiguous"}

Este módulo NO importa Tkinter ni dependencias opcionales (PIL); puede
ejecutarse en hilos o en tests. Todas las decisiones son deterministas:
el índice interno se construye ordenando las rutas, y la jerarquía de
fallback está fijada (no depende del orden del sistema de archivos).
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Tuple

# Extensiones consideradas "de imagen" para INDEXAR (mismo conjunto que
# metatag_v8.IMG_EXTS, incluido .bmp).
IMG_EXTS = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp")
# Extensiones que _safe_stem despoja del nombre (mismo conjunto que
# MetaTagApp._IMG_EXTS_LOWER; se mantiene la divergencia histórica con
# IMG_EXTS para que el port sea IDÉNTICO al algoritmo original).
_STEM_EXTS = frozenset({".jpg", ".jpeg", ".png", ".tif", ".tiff", ".webp"})


def _safe_stem(s: str) -> str:
    """Quita extensiones de imagen (incluso dobles) y marcadores '(N)'.

    '0006_UM_C4_IX_00034_P.jpg (1).JPG' -> '0006_UM_C4_IX_00034_P'
    """
    p = Path(s)
    changed = True
    while changed:
        changed = False
        if p.suffix.lower() in _STEM_EXTS:
            p = p.with_suffix("")
            changed = True
            continue
        m = re.match(r"^(.*?)\s*\(\d+\)$", p.name)
        if m and m.group(1):
            p = p.with_name(m.group(1))
            changed = True
    return p.name


def _full_stem(s: str) -> str:
    return _safe_stem(s)


def _normalize_numbers(s: str) -> str:
    """'0006_UM_C4_IX_00034_P' -> '6_UM_C4_IX_34_P' (quita ceros)."""
    return re.sub(r"\d+", lambda m: str(int(m.group())), s)


def _extract_id_suffix(s: str) -> Optional[Tuple[str, str]]:
    """
    Extrae (numero_pieza, sufijo_vista) de un nombre, sin importar cuántos
    campos haya en medio NI si trae extensión (incluso doble, ej. '.jpg.JPG').
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


def _clean_stem(stem_key: str) -> str:
    return re.sub(r"^[#\s\-_]+|[#\s\-_]+$", "", stem_key)


class ImageMatcher:
    """Emparejador de nombres de archivo con imágenes en una carpeta.

    Réplica pura y determinista de ``MetaTagApp._find_image_ex``.
    """

    def __init__(self) -> None:
        self._index_cache: dict[str, List[Path]] = {}

    def _invalidate(self, folder: str) -> None:
        """Descarta el índice en memoria de una carpeta (si cambió en disco)."""
        self._index_cache.pop(folder, None)

    def _index_folder(self, folder: str) -> dict:
        """
        Construye (una sola vez) el índice {stem_lower: path} de la carpeta.

        Réplica de la construcción del caché de MetaTag, pero iterando las
        rutas ordenadas para que el ganador de una colisión de stems sea
        determinista. Barrido RECURSIVO (mismo comportamiento que rglob).
        """
        cached = self._index_cache.get(folder)
        if cached is not None:
            return cached
        paths = []
        try:
            f = Path(folder)
            if f.is_dir():
                paths = [p for p in f.rglob("*")
                         if p.is_file() and p.suffix.lower() in IMG_EXTS]
        except OSError:
            paths = []
        paths.sort(key=lambda p: p.as_posix())
        index = {_full_stem(p.name).lower(): p for p in paths}
        self._index_cache[folder] = index
        return index

    def find_image(self, name: str, folder: str) -> Optional[str]:
        """Devuelve la ruta emparejada o None (nunca elige ambiguos)."""
        return self.find_image_ex(name, folder)[0]

    def find_image_ex(self, name: str, folder: str) -> Tuple[Optional[str], str, List[str]]:
        """
        Busca la fotografía que corresponde a 'name' dentro de 'folder'.

        Devuelve (path, status, candidatos):
          - path: ruta elegida (o None).
          - status: "ok" | "not_found" | "ambiguous".
          - candidatos: lista de rutas que compitieron (para el mensaje).

        Jerarquía de pasos (idéntica a _find_image_ex de MetaTag):
        1. ruta directa (folder/name)
        2. nombre de archivo exacto (sin importar dónde esté)
        3. stem exacto
        4. stem "limpio" (sin '#'/espacios/guiones a los bordes)
        5. números normalizados (0006 == 6)
        6. id = (numero_pieza, sufijo_vista) — si hay >1, es AMBIGUO
        7. subcadena — si hay >1, es AMBIGUO
        """
        name = name.strip()
        if not name:
            return None, "not_found", []

        folder_path = Path(folder)
        index = self._index_folder(folder)

        p = folder_path / name
        if p.exists():
            return str(p), "ok", [str(p)]

        name_lower = name.lower()
        name_stem = _full_stem(name).lower()

        for _, fpath in index.items():
            if fpath.name.lower() == name_lower:
                return str(fpath), "ok", [str(fpath)]
        if name_stem in index:
            fpath = index[name_stem]
            return str(fpath), "ok", [str(fpath)]

        name_clean = _clean_stem(name_stem)
        for stem_key, fpath in index.items():
            if _clean_stem(stem_key) == name_clean:
                return str(fpath), "ok", [str(fpath)]

        name_normalized = _normalize_numbers(name_clean)
        for stem_key, fpath in index.items():
            stem_clean = _clean_stem(stem_key)
            if _normalize_numbers(stem_clean) == name_normalized:
                return str(fpath), "ok", [str(fpath)]

        # Último recurso y más robusto: comparar solo número + sufijo.
        # Se recogen TODOS los candidatos: si hay más de uno con la misma
        # clave, la correspondencia es ambigua y NO se elige ninguno.
        id_excel = _extract_id_suffix(name)
        if id_excel:
            candidates = []
            for stem_key, fpath in index.items():
                id_archivo = _extract_id_suffix(stem_key)
                if id_archivo and id_archivo == id_excel:
                    candidates.append(str(fpath))
            if len(candidates) == 1:
                return candidates[0], "ok", candidates
            if len(candidates) > 1:
                return None, "ambiguous", candidates

        substring_candidates = []
        for stem_key, fpath in index.items():
            if name_stem in stem_key or stem_key in name_stem:
                substring_candidates.append(str(fpath))
        if len(substring_candidates) == 1:
            return substring_candidates[0], "ok", substring_candidates
        if len(substring_candidates) > 1:
            return None, "ambiguous", substring_candidates

        return None, "not_found", []

    def find_image_ex_with_method(
            self, name: str, folder: str
    ) -> Tuple[Optional[str], str, List[str], Optional[str]]:
        """
        Igual que find_image_ex pero además devuelve el paso que resolvió
        (para auditoría/tests): "direct", "nombre-exacto", "stem-exacto",
        "clean", "normalize", "id-suffix", "substring" o None si not_found.
        No aporta ninguna decisión nueva: el resultado es el mismo.
        """
        name = name.strip()
        if not name:
            return None, "not_found", [], None

        folder_path = Path(folder)
        index = self._index_folder(folder)

        p = folder_path / name
        if p.exists():
            return str(p), "ok", [str(p)], "direct"

        name_lower = name.lower()
        name_stem = _full_stem(name).lower()

        for _, fpath in index.items():
            if fpath.name.lower() == name_lower:
                return str(fpath), "ok", [str(fpath)], "nombre-exacto"
        if name_stem in index:
            return str(index[name_stem]), "ok", [str(index[name_stem])], "stem-exacto"

        name_clean = _clean_stem(name_stem)
        for stem_key, fpath in index.items():
            if _clean_stem(stem_key) == name_clean:
                return str(fpath), "ok", [str(fpath)], "clean"

        name_normalized = _normalize_numbers(name_clean)
        for stem_key, fpath in index.items():
            stem_clean = _clean_stem(stem_key)
            if _normalize_numbers(stem_clean) == name_normalized:
                return str(fpath), "ok", [str(fpath)], "normalize"

        id_excel = _extract_id_suffix(name)
        if id_excel:
            candidates = []
            for stem_key, fpath in index.items():
                id_archivo = _extract_id_suffix(stem_key)
                if id_archivo and id_archivo == id_excel:
                    candidates.append(str(fpath))
            if len(candidates) == 1:
                return candidates[0], "ok", candidates, "id-suffix"
            if len(candidates) > 1:
                return None, "ambiguous", candidates, "id-suffix"

        substring_candidates = []
        for stem_key, fpath in index.items():
            if name_stem in stem_key or stem_key in name_stem:
                substring_candidates.append(str(fpath))
        if len(substring_candidates) == 1:
            return substring_candidates[0], "ok", substring_candidates, "substring"
        if len(substring_candidates) > 1:
            return None, "ambiguous", substring_candidates, "substring"

        return None, "not_found", [], None


# ── helper para tests / CLI ────────────────────────────────────────────────

def match_name_to_photo(name: str, folder: str) -> Optional[str]:
    return ImageMatcher().find_image(name, folder)
