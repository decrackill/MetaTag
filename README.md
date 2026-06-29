# MetaTag v7 — Escritor de Metadatos Arqueológicos

Herramienta de escritorio para escribir metadatos arqueológicos directamente en los archivos de imagen (JPG, PNG, TIFF), a partir de una tabla Excel o CSV.

---

## Requisitos

- Windows 10 o superior
- Python 3.8 o superior → https://www.python.org/downloads/
  - **Importante:** al instalar Python, marca la casilla **"Add Python to PATH"**

---

## Cómo usar

1. Coloca los archivos `metatag_v7.py` e `instalar_y_abrir.bat` en la misma carpeta.
2. Haz doble clic en **`instalar_y_abrir.bat`**.
3. El programa instalará las dependencias automáticamente y abrirá la interfaz.

---

## Archivos incluidos

| Archivo | Descripción |
|---|---|
| `metatag_v7.py` | Programa principal |
| `instalar_y_abrir.bat` | Lanzador / instalador de dependencias |
| `INSTRUCCIONES.docx` | Manual completo de uso con capturas |
| `README.md` | Este archivo |

---

## Dependencias (se instalan automáticamente)

- `pillow` — procesamiento de imágenes
- `piexif` — escritura de metadatos EXIF en JPG/TIFF
- `pandas` — lectura de archivos Excel y CSV
- `openpyxl` — soporte de formato .xlsx

---

## Carpeta de salida

Las imágenes con metadatos escritos se guardan en:
```
(misma carpeta del programa)/Metadatos_Escritos/
```

---

## Novedades v4 (respecto a v3)

- **Búsqueda de imágenes mejorada:** ahora encuentra archivos aunque tengan diferencias en mayúsculas/minúsculas, extensión (`.jpg` vs `.JPG`) o caracteres especiales como `#` al inicio del nombre.
- **Match automático imagen ↔ fila:** al seleccionar una imagen en el explorador lateral, la fila correspondiente en la tabla se resalta automáticamente.
- **Selector de temas:** 4 temas visuales — Arqueológico, Oscuro, Verde Bosque, Azul Marino.

---

## Soporte de formatos de imagen

JPG · JPEG · PNG · TIF · TIFF · BMP · GIF · WEBP · HEIC · HEIF · AVIF · ICO · PPM · PGM

---

## Notas

- El programa **no modifica** el archivo Excel/CSV original.
- Las imágenes originales **no se modifican**; se crea una copia con los metadatos en la carpeta de salida.
- Si una imagen no se encuentra, aparece un error `✗ No encontrada` en el registro. Verifica que el nombre en el Excel coincida con el nombre real del archivo.
