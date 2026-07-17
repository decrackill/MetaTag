"""
MetaTag v8.9 — Módulo de escritura de metadatos en imágenes.
Funciones puras sin dependencias de tkinter, testeables unitariamente.
"""

import json
from pathlib import Path

try:
    from PIL import Image
    import piexif, piexif.helper
    PIL_OK = True
except ImportError:
    PIL_OK = False

META_GROUPS = {
    "Ubicacion":   ["Sitio", "Corte", "Cuadrante", "Unidad", "Nivel", "Profundidad Cm"],
    "Descripcion": ["Vista", "Parte", "Perfil", "Labio"],
    "Tecnica":     ["Tratamiento", "Tecnica", "Motivo"],
    "Notas":       ["Observaciones", "Excluido"],
}
META_GROUP_ORDER = ["Ubicacion", "Descripcion", "Tecnica", "Notas"]
_TIFF_JSON_PREFIX = "METATAG_JSON:"


def formatear_metadatos(meta: dict, organizado: bool = True) -> str:
    if not organizado:
        return "\n".join(f"{k}: {v}" for k, v in meta.items())
    partes, restantes = dict(meta), []
    for grupo in META_GROUP_ORDER:
        items = {k: v for k, v in meta.items() if k in META_GROUPS[grupo]}
        if not items:
            continue
        restantes.append(f"[{grupo}]")
        for k, v in items.items():
            restantes.append(f"  {k}: {v}")
        for k in items:
            partes.pop(k, None)
    if partes:
        restantes.append("[Otros]")
        for k, v in partes.items():
            restantes.append(f"  {k}: {v}")
    return "\n".join(restantes)


def read_existing_metadata(path: str) -> dict:
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
            elif ext in (".tif", ".tiff"):
                description = img.tag_v2.get(270, "")
                if isinstance(description, bytes):
                    description = description.decode("utf-8", errors="replace")
                if isinstance(description, str) and description.startswith(_TIFF_JSON_PREFIX):
                    return json.loads(description[len(_TIFF_JSON_PREFIX):])
    except Exception:
        pass
    return {}


def check_metadata_divergence(path, expected_meta):
    """
    Compara los metadatos ya escritos en la imagen contra los valores
    que el Excel dice que DEBERÍAN estar ahí (expected_meta).
    Devuelve la lista de campos que divergen.
    """
    existing = read_existing_metadata(path)
    if not existing:
        return []
    diffs = []
    for k, v_new in expected_meta.items():
        v_old = str(existing.get(k, "")).strip()
        if v_old and v_old != str(v_new).strip():
            diffs.append(f"{k}: '{v_old}' → '{v_new}'")
    return diffs


def write_jpeg(path: str, meta: dict, organizado: bool = True):
    with Image.open(path) as _img_src:
        if getattr(_img_src, "n_frames", 1) != 1:
            raise RuntimeError("JPEG animado o multipágina no soportado.")
        _img_src.load()
        source_info = dict(_img_src.info)
        img = _img_src.copy()
    texto_organizado = formatear_metadatos(meta, organizado)
    as_json          = json.dumps(meta, ensure_ascii=False)
    keywords         = ";".join(v for v in meta.values() if v.strip())
    try:
        exif = piexif.load(img.info.get("exif", b""))
    except Exception:
        exif = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}}
    exif["0th"][piexif.ImageIFD.ImageDescription] = texto_organizado.encode("utf-8")
    exif["Exif"][piexif.ExifIFD.UserComment] = piexif.helper.UserComment.dump(
        as_json, encoding="unicode")
    exif["0th"][40092] = (texto_organizado + "\x00").encode("utf-16-le")
    exif["0th"][40094] = (keywords         + "\x00").encode("utf-16-le")
    save_args = {"exif": piexif.dump(exif), "quality": 95}
    for key in ("icc_profile", "dpi"):
        if source_info.get(key) is not None:
            save_args[key] = source_info[key]
    img.save(path, "jpeg", **save_args)


def write_png(path: str, meta: dict, organizado: bool = True):
    from PIL import PngImagePlugin
    with Image.open(path) as _img_src:
        if getattr(_img_src, "n_frames", 1) != 1:
            raise RuntimeError("PNG animado no soportado para evitar perder fotogramas.")
        _img_src.load()
        source_info = dict(_img_src.info)
        source_text = dict(getattr(_img_src, "text", {}))
        img = _img_src.copy()
    info = PngImagePlugin.PngInfo()
    for k, v in source_text.items():
        if k not in {"Description", "Comment", *meta.keys()}:
            info.add_text(str(k), str(v))
    texto_organizado = formatear_metadatos(meta, organizado)
    for k, v in meta.items():
        info.add_text(str(k), str(v))
    info.add_text("Description", texto_organizado)
    info.add_text("Comment",     json.dumps(meta, ensure_ascii=False))
    save_args = {"pnginfo": info}
    for key in ("icc_profile", "dpi", "exif"):
        if source_info.get(key) is not None:
            save_args[key] = source_info[key]
    img.save(path, "PNG", **save_args)


def write_tiff(path: str, meta: dict, organizado: bool = True):
    from PIL import TiffImagePlugin
    with Image.open(path) as _img_src:
        if getattr(_img_src, "n_frames", 1) != 1:
            raise RuntimeError("TIFF multipágina no soportado para evitar perder páginas.")
        _img_src.load()
        source_info = dict(_img_src.info)
        tiffinfo = TiffImagePlugin.ImageFileDirectory_v2()
        for tag, value in _img_src.tag_v2.items():
            tiffinfo[tag] = value
        img = _img_src.copy()
    # TIFF no conserva de forma fiable el EXIF UserComment con Pillow. Guardamos
    # el JSON en ImageDescription con un prefijo inequívoco y legible al reabrir.
    tiffinfo[270] = _TIFF_JSON_PREFIX + json.dumps(meta, ensure_ascii=False)
    save_args = {"format": "TIFF", "tiffinfo": tiffinfo}
    for key in ("icc_profile", "dpi", "compression"):
        if source_info.get(key) is not None:
            save_args[key] = source_info[key]
    img.save(path, **save_args)


def write_meta(path: str, meta: dict, organizado: bool = False):
    """Escribe el diccionario `meta` en los metadatos del archivo en `path`."""
    if not PIL_OK:
        raise RuntimeError("Pillow / piexif no instalados.")
    ext = Path(path).suffix.lower()
    if ext in (".jpg", ".jpeg"):
        write_jpeg(path, meta, organizado)
    elif ext == ".png":
        write_png(path, meta, organizado)
    elif ext in (".tif", ".tiff"):
        write_tiff(path, meta, organizado)
    else:
        raise RuntimeError(f"Formato no soportado para escritura de metadatos: {ext}")
