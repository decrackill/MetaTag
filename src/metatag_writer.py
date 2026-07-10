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
    img.save(path, "jpeg", exif=piexif.dump(exif), quality=95)


def write_png(path: str, meta: dict, organizado: bool = True):
    from PIL import PngImagePlugin
    with Image.open(path) as _img_src:
        img = _img_src.copy()
    info = PngImagePlugin.PngInfo()
    texto_organizado = formatear_metadatos(meta, organizado)
    for k, v in meta.items():
        info.add_text(str(k), str(v))
    info.add_text("Description", texto_organizado)
    info.add_text("Comment",     json.dumps(meta, ensure_ascii=False))
    img.save(path, "PNG", pnginfo=info)


def write_tiff(path: str, meta: dict, organizado: bool = True):
    with Image.open(path) as _img_src:
        img = _img_src.copy()
    texto_organizado = formatear_metadatos(meta, organizado)
    try:
        exif = piexif.load(img.info.get("exif", b""))
    except Exception:
        exif = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}}
    exif["0th"][piexif.ImageIFD.ImageDescription] = texto_organizado.encode("utf-8")
    img.save(path, exif=piexif.dump(exif))


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
        try:
            write_jpeg(path, meta, organizado)
        except Exception:
            raise RuntimeError(f"Formato no soportado: {ext}")
