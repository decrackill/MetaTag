  # CONTEXTO_PROYECTO — MetaTag v8.9

  > Memoria técnica persistente del proyecto.
  > Leer este archivo antes de modificar código. El **código actual es la fuente de verdad**; si hay contradicción entre este documento y el código, verificar el código y actualizar esta bitácora.

  ---

  ## 1. Resumen del proyecto

  MetaTag v8.9 es un **escritor de metadatos arqueológicos** en Python (Tkinter + pandas). Permite inyectar metadatos descriptivos (procedencia, morfología, tecnología, tratamiento, etc.) en fotografías de fragmentos cerámicos a partir de registros cargados desde un archivo Excel/CSV.

  - **Qué resuelve:** asociar de forma fiable y masiva el registro de un fragmento (una fila del Excel) con su fotografía, y escribir esos datos en los metadatos de la imagen (EXIF/IPTC), sin modificar los archivos originales.
  - **Flujo general de trabajo:**
    1. Cargar un Excel/CSV con los registros.
    2. Seleccionar la carpeta con las fotografías y la columna de imagen.
    3. Emparejar filas del Excel con imágenes (matching inteligente).
    4. Escribir los metadatos (manual o por lote) sobre **copias** en una carpeta de salida.
    5. Verificar/visualizar resultados (visor, estadísticas, comparador, PDF).

  ---

  ## 2. Estructura del proyecto

  ```
  MetaTag_v8.9/
  ├── src/
  │   ├── metatag_v8.py             # Aplicación principal (~3212 líneas)
  │   ├── metatag_writer.py         # Escritura EXIF/IPTC pura (~145 líneas)
  │   ├── metatag_graficas.py       # Estadísticas y gráficas (~694 líneas)
  │   ├── metatag_widgets.py        # ExcelGrid con viewport culling (~418 líneas)
  │   ├── Visor.py                  # Visor de metadatos / comparador / PDF (~2468 líneas)
  │   ├── renombrar_fotos_gui.py    # Renombrador de Fotos v4.0 (standalone, CustomTkinter)
  │   ├── metatag_theme.py          # Fuente de verdad técnica de temas/fuentes (puro, sin toolkit)
  │   └── editor_casillas_backup.py # Respaldo del editor de casillas (~116 líneas)
  ├── tests/
  │   ├── test_renombrador_pytest.py# Suite pytest original del renombrador (49 tests)
  │   └── test_metatag_theme.py     # Paridad de temas/fuentes contra metatag_v8 (53 tests)
  ├── data/
  │   ├── metatag_config.json       # Configuración persistente de la sesión
  │   └── metatag_debug.log         # Log de errores (logging, nivel ERROR)
  ├── CeramicIA_Dataset_Piloto_103Fragmentos_v1.1_2026-06-26.xlsx  # Dataset de prueba
  ├── Finales 1 a 103/              # 269 imágenes de prueba
  ├── instalar_y_abrir.sh / .bat    # Instalación de dependencias + lanzamiento
  ├── requirements-renombrador.txt  # Dependencias del renombrador (customtkinter aislada)
  └── .venv/                        # Entorno virtual (Python 3.12.3)
  ```

  ### Módulos y responsabilidades

  | Archivo | Responsabilidad | Componentes principales | Estado |
  |---|---|---|---|
  | `src/metatag_v8.py` | Aplicación principal: UI, carga de datos, emparejamiento, escritura, Image Sync | `ImageBrowser`, integración `ExcelGrid`, inyección manual/lote, `_sync_excel_to_images`, `_sync_images_to_excel`, `_find_image`, temas, lupa, búsqueda | Implementado y activo |
  | `src/metatag_writer.py` | Escritura de metadatos **pura y testeable** (sin Tkinter) | `META_GROUPS`, `META_GROUP_ORDER`, `formatear_metadatos`, `read_existing_metadata`, `check_metadata_divergence`, `write_jpeg`, `write_png`, `write_tiff`, `write_meta` | Implementado |
  | `src/metatag_matching.py` | Motor de emparejamiento seguro **puro** (sin Tkinter/PIL), port fiel de `_find_image_ex` | `ImageMatcher` (`find_image_ex`, `find_image`, `find_image_ex_with_method`, `_index_folder`), `_safe_stem`, `_normalize_numbers`, `_extract_id_suffix`, `match_name_to_photo` | Implementado (FASE A) |
  | `src/metatag_graficas.py` | Estadísticas, selector de gráficas, generación y exportación | `show_stats`, `make_selector`, `update_chart`, `export_chart`, `on_hover` | Implementado; mejoras futuras planeadas |
  | `src/metatag_widgets.py` | Grid Excel sobre Canvas con viewport culling | Clase `ExcelGrid`: `load`, `redraw`, selección de filas/columnas, `get_row_metadata`, `get_selected_metadata` | Implementado |
  | `src/Visor.py` | Visor independiente de metadatos EXIF/JSON/GPS, comparador, zoom y export PDF | `VisorApp`, `_extract_exif`, `_extract_gps`, `_extract_json`, `_open_image_comparison`, `_export_pdf`, `_generate_pdf_document` | Implementado |
  | `src/editor_casillas_backup.py` | Respaldo de funciones del editor de casillas | `_open_editor`, `_populate_editor`, `toggle_lock`, `update_df` | Backup; no es el editor activo |
  | `src/renombrar_fotos_gui.py` | Renombrador de Fotos v4.0: herramienta independiente (CustomTkinter) lanzada como subproceso desde MetaTag | `RenameModel`, `MainView`, `AppController`, `_build_plan`/`_build_plan_matching` (modo "matching seguro" con `ImageMatcher`), `rename_all`, `undo_last` | Implementado y verificado (FASE A); integrado como lanzador en `metatag_v8.py` |

  **Relaciones:** `metatag_v8.py` importa `show_stats` desde `metatag_graficas`, `ExcelGrid` desde `metatag_widgets`, las funciones de escritura/divergencia desde `metatag_writer`, y los tokens de tema/fuentes desde `metatag_theme`. `metatag_v8.py` puede lanzar `Visor.py` (`launch_visor`) y `renombrar_fotos_gui.py` (`_launch_renombrador`, subproceso con `sys.executable`, sin withdraw ni callbacks). `src/renombrar_fotos_gui.py` importa `ImageMatcher` desde `metatag_matching` (mismo directorio; el bloque inserta `src/` en `sys.path`) y consume los temas de MetaTag vía `metatag_theme.CustomTkinterThemeAdapter`.

  ---

  ## 3. Arquitectura de metatag_v8

  Archivo único de ~3068 líneas que concentra la mayor parte de la lógica y la UI.

  ### Componentes principales

  - **`ImageBrowser` (clase, ~línea 204):** panel de miniaturas; `load_folder`, `_filter` (filtro por nombre), `_on_select`, `highlight`.
  - **Grid Excel:** utiliza `ExcelGrid` de `metatag_widgets.py` (Canvas con culling); integrado en el panel central con filtros por columna (`_apply_filter`) y sincronización de scroll.
  - **Inyección manual (`_inject_manual`, ~2599):** copia la imagen actual a la carpeta de salida y escribe los metadatos de la fila seleccionada en la copia.
  - **Inyección por lote (`_batch_write_by_order` ~1937, `_batch_worker` ~1869):** escribe por posición, con selección de columnas, orden de fotos, reanudación y ventana de progreso con cancelación.
  - **Image Sync bidireccional:** `_sync_excel_to_images` (~1466) y `_sync_images_to_excel` (~1565). Ver sección 5.
  - **Matching inteligente:** `_find_image` (~2697), `_safe_stem` (~2643), `_full_stem`, `_normalize_numbers`, `_extract_id_suffix` (~2667).
  - **Divergencia de metadatos:** `_check_metadata_divergence` (~2754) delega en `metatag_writer`.
  - **Búsqueda:** diálogo de atajos y búsqueda (`_show_shortcuts`, `_focus_search`); autocompletado sobre el grid (`_apply_autocompletion`).
  - **Lupa:** `_show_loupe_window` (~1098) y renderizado de la lupa (`_render_loupe_img`), con zoom.
  - **Temas:** fuente de verdad técnica en `src/metatag_theme.py` (módulo puro, sin toolkit): diccionario `THEMES` con **3 temas** canónicos — `Arqueológico (Oscuro Refinado)` (por defecto), `Noche Total`, `Carbón` —, `THEME_ORDER`, `THEME_ICONS`, motor de fuentes dinámico (`compute_font_scale` = clamp(sw/1920, 0.82, 1.35); `font_specs` replica byte a byte `set_font_scale`) y adaptadores (`TkThemeAdapter` pasa los tokens canónicos tal cual; `CustomTkinterThemeAdapter` traduce al esquema del Renombrador derivando solo lo que el canónico no define). `metatag_v8.py` importa `THEMES`, `DEFAULT_THEME`, `THEME_ICONS`, `compute_font_scale`, `font_specs` y conserva `CURRENT_THEME`/`C`/`FONTS`/`set_font_scale` como API pública.
  - **Estadísticas:** `_show_stats` (~514) invoca `show_stats` externo.
  - **Persistencia:** `_save_config`, `_load_config_pre_build`, `_load_config_post_build` leen/escriben `data/metatag_config.json`.
  - **Lanzador del Renombrador (`_launch_renombrador`, ~988):** botón "🖼 Renombrador de Fotos" en la sección HERRAMIENTAS AVANZADAS del panel izquierdo; lanza `src/renombrar_fotos_gui.py` como subproceso independiente (`subprocess.Popen([sys.executable, ...])` con ruta construida desde `Path(__file__).resolve().parent`). **No** hace `withdraw()` ni polling: MetaTag sigue usable mientras el renombrador está abierto y no quedan callbacks huérfanos. Errores → `messagebox` + `logging` sin crashear.
  - **Logging:** `_log`/`_log_safe` hacia el panel de log de la UI; `logging.basicConfig` escribe `data/metatag_debug.log` (nivel ERROR).

  ### Flujo de carga

  `_browse_csv` → `_load_file` (~1226): lee `.csv` con `pd.read_csv` o `.xlsx` con `pd.read_excel` (`dtype=str, keep_default_na=False, na_values=[]`), mapea vacíos/NaN a `""` y **elimina espacios en blanco de los nombres de columna**.

  ---

  ## 4. Escritura de metadatos

  ### Flujo de datos

  ```
  Excel/CSV (filas) → matching de imágenes → copia (copy2) a carpeta de salida → escritura de metadatos en la copia
  ```

  **IMPORTANTE (confirmado en el código): los originales NO se modifican.** Tanto la inyección manual como la por lote copian la imagen a `output_folder` con `shutil.copy2` y escriben los metadatos sobre la copia.

  ### `metatag_writer.py` — funciones puras

  - Sin dependencias de Tkinter; depende de `PIL`/`piexif` (flag `PIL_OK`).
  - **Grupos de metadatos:** `META_GROUPS` y `META_GROUP_ORDER` (categorías IPTC/EXIF).
  - `formatear_metadatos`: estructura los valores por grupos.
  - `read_existing_metadata`: lee metadatos actuales de la imagen.
  - `check_metadata_divergence`: compara lo existente con lo que se va a escribir.
  - `write_jpeg`, `write_png`, `write_tiff`: escritura por formato.
  - `write_meta`: despacho según el formato del archivo.
  - Formatos soportados: **JPEG, PNG, TIFF**.

  ### Procesamiento por lote

  - `_batch_write_by_order`: permite usar la tabla ya cargada o seleccionar Excel nuevo; pide columna(s) de metadatos (`_batch_pick_columns`), orden de fotos (`_batch_pick_sort`, `_sort_images`); `total = min(len(img_files), len(batch_df))`; avisa si los conteos difieren.
  - `_batch_worker`: para cada posición construye `meta` desde las columnas elegidas (+ columnas bloqueadas), salta filas vacías (con `omit_empty`), detecta divergencias, copia el original a la salida y escribe. Escribe un archivo de progreso cada 10 imágenes.
  - **Reanudación (`_batch_progress.json`):** si existe un progreso previo con el mismo `total` y el mismo Excel, pregunta si continuar desde la imagen siguiente; al finalizar, elimina el archivo de progreso.

  ---

  ## 5. Image Sync

  Existe **sincronización bidireccional de orden** (no renombra archivos):

  - **`_sync_excel_to_images` (~1466):** reordena las **filas del Excel** para que sigan el orden de los archivos de la carpeta de imágenes. Usa `get_clean_name` (elimina extensiones) e índice de archivos; hace matching por nombre o prefijo. **No modifica los archivos.**
  - **`_sync_images_to_excel` (~1565):** reordena la **lista de imágenes del explorador** para que siga EXACTAMENTE el orden de filas del Excel cargado (sin reordenar el Excel). Detalle del proceso:
    1. `_build_file_index`: índices `by_exact_name` (nombre en minúsculas) y `by_stem` (stem).
  2. `_match_rows_to_files`: por fila intenta coincidencia **exacta → por stem → aproximada** (`_find_image_ex`); devuelve `(ordered_files, matches_aprox, no_encontradas, used_files, matches_ambiguas)`.
  3. Las filas AMBIGUAS (varios archivos comparten la clave de fallback) van a `matches_ambiguas` y NO se emparejan ni caen en `no_encontradas`; se muestran en el log con todos sus candidatos.
  4. Las imágenes sin fila correspondiente quedan al final marcadas como huérfanas con su razón (`_orphan_files`, `_orphan_reasons`).
  5. Actualiza el explorador (`browser._filter()`), contadores (`_sync_excel_count`, `_sync_orphan_count`) y registra en el log las coincidencias aproximadas, ambigüedades, huérfanas y valores sin imagen.

  ### Matching inteligente (helpers)

  - `_safe_stem`: elimina de forma iterativa extensiones de imagen y marcadores de duplicado tipo ` (1)` (ej.: `"0006_..._P.jpg (1).JPG"` → stem limpio).
  - `_full_stem`: alias de `_safe_stem`.
  - `_normalize_numbers`: normaliza secuencias de dígitos (con `int()`; `'0006'`→`'6'`).
  - `_extract_id_suffix`: extrae `(numero_pieza, sufijo_vista)`. Ej. `'0001_UM_C4_UE18_00006_F.jpg'` → `('1', 'F')`; `'0061_EC_C4_III_046'` → `('61', '')`.
  - `_find_image_ex`: implementación ÚNICA del matching. Devuelve `(path, status, candidates)` con status `"ok"`, `"not_found"` o `"ambiguous"`. Orden de pasos idéntico al original: direct → nombre-exacto → stem-exacto → clean → normalize → **id-suffix (recoge TODOS los candidatos)** → substring (idem). Si la clave de fallback (id-suffix o substring) tiene 2+ candidatos → `status="ambiguous"`, `path=None`, y `candidates` lista todos; NO elige uno arbitrario (evita el bug del primero-de-la-caché).
  - `_find_image`: envuelve `_find_image_ex` y devuelve solo el path (para llamadores de preview/legacy que no necesitan el diagnóstico).

  ---

  ## 6. Módulo de estadísticas y gráficas

  `src/metatag_graficas.py` (~694 líneas, flag `MATPLOTLIB_OK`).

  - **`show_stats` (~107):** ventana de estadísticas; recibe el DataFrame y columnas desde la app principal; invocada desde `_show_stats` de `metatag_v8`.
  - **Selector de gráficas (`make_selector`, ~163):** diálogo para elegir tipo/parámetros de la gráfica (`on_var_change`, `on_style_change`).
  - **Generación (`update_chart`, ~358):** construye la gráfica con matplotlib; utilidades de colocación de etiquetas (`_place_labels_clean`) y hover (`on_hover`).
  - **Exportación (`export_chart`, ~281):** guarda la gráfica mediante diálogo nativo de guardado.
  - **Relación con la app principal:** importada como `show_stats` en `metatag_v8`; módulo separado de la UI.
  - **Limitación arquitectónica detectada:** la generación/exportación depende de matplotlib embebido en Tkinter; el módulo concentra UI + lógica de gráficas.
  - **Intención futura (NO implementada):** seguir desarrollando el módulo y posiblemente separarlo aún más si el crecimiento lo justifica. No hay separación adicional implementada hoy.

  ---

  ## 7. Datos del proyecto

  Datos de prueba/dataset actual (no son una limitación del programa):

  - **Excel:** `CeramicIA_Dataset_Piloto_103Fragmentos_v1.1_2026-06-26.xlsx` en la raíz.
    - Hojas: `Principal` y `Categorías`.
    - Leído con pandas: **269 filas × 19 columnas**.
    - Columnas confirmadas: `Fragmento_ID`, `ID Imagen`, `Sitio`, `Corte`, `Unidad`, `Nivel`, `Profundidad Cm`, `Vista`, `Parte`, `Perfil`, `Labio`, `Tratamiento`, `Técnica`, `Motivo`, `Observaciones`, `Ob. rasgos morfológicos especiales`, `Ob. Carácter tecnológicas esp.`, `Conservación`, `Excluido`.
    - Particularidad: en el Excel crudo varias cabeceras llevan espacios finales (ej. `Sitio `); el código los elimina con `.strip()`.
    - La hoja `Categorías` contiene las categorías posibles (Vista, Parte, Perfil, Labio, Tratamiento, Técnica, Motivo, rasgos morfológicos, rasgos tecnológicos, Conservación).
  - **Imágenes:** **269 imágenes** en la carpeta `Finales 1 a 103/`. Nombres típicos: `0001_UM_C4_UE18_00006_F.jpg` (con sufijos `F`/`R`/`P` de vista).
  - Nota observada: `data/metatag_config.json` apunta a copias en `~/Escritorio/` (fuera de la raíz del proyecto); la raíz también contiene su propia copia del Excel y de la carpeta de imágenes.

  ---

  ## 8. Configuración y entorno

  ### `data/metatag_config.json`

  Contenido verificado:

  ```json
  {
    "csv_path": "/home/deivis/Escritorio/CeramicIA_Dataset_Piloto_103Fragmentos_v1.1_2026-06-26.xlsx",
    "img_folder": "/home/deivis/Escritorio/Finales 1 a 103",
    "theme": "Noche Total",
    "process_mode": "Inteligente",
    "left_ratio": 0.165
  }
  ```

  ### Entorno virtual `.venv`

  - Python **3.12.3** (de `pyvenv.cfg`).
  - Dependencias confirmadas (vía `pip list`): **pandas 3.0.3**, openpyxl, pillow, piexif, matplotlib, numpy, reportlab.
  - Verificación no destructiva ejecutada: `py_compile` OK en los 6 módulos; importación de `metatag_writer`, `metatag_widgets` y `metatag_graficas` OK (`PIL_OK` y `MATPLOTLIB_OK` verdaderos); lectura del Excel OK (269×19).

  ### Scripts de instalación/lanzamiento

  - `instalar_y_abrir.sh` y `instalar_y_abrir.bat`: auto-organizan `.py` en `src/`, crean `data/`, verifican/instalan dependencias (pandas, openpyxl, pillow, piexif, matplotlib, numpy, reportlab) y lanzan `src/metatag_v8.py` (el `.bat` con `pythonw` si está disponible).
  - **Dependencia aislada del Renombrador:** `customtkinter` NO es dependencia de la UI principal de MetaTag (que es Tkinter puro). Se gestiona aparte con `requirements-renombrador.txt` (customtkinter, pandas, openpyxl, pillow) en la raíz; los launchers antiguos de `tools/renombrador/` (que creaban un `.venv` local obsoleto) se eliminaron con la migración (2026-08-10). El lanzamiento desde MetaTag usa el Python del `.venv` del proyecto (`sys.executable`); si falta CustomTkinter, el renombrador muestra su propio diálogo de error sin afectar a MetaTag.

  ---

  ## 9. Funcionalidades implementadas

  | Funcionalidad | Estado |
  |---|---|
  | Navegador de imágenes (miniaturas) | IMPLEMENTADO (`ImageBrowser`) |
  | Grid Excel con viewport culling | IMPLEMENTADO (`ExcelGrid`) |
  | Escritura manual de metadatos (inyección sobre copia) | IMPLEMENTADO |
  | Escritura por lote (por posición) | IMPLEMENTADO |
  | Reanudación del lote (`_batch_progress.json`) | IMPLEMENTADO |
  | Matching inteligente (`_find_image`, `_safe_stem`, `_extract_id_suffix`) | IMPLEMENTADO |
  | Image Sync (reordenar Excel según fotos / fotos según Excel) | IMPLEMENTADO (solo reordenación de orden; NO renombra) |
  | Búsqueda y atajos de teclado | IMPLEMENTADO |
  | Lupa con zoom | IMPLEMENTADO |
  | Temas de color (3 en metatag_v8, 3 en Visor) | IMPLEMENTADO |
  | Estadísticas y gráficas | IMPLEMENTADO |
  | Visor de metadatos (EXIF/JSON/GPS) | IMPLEMENTADO |
  | Comparador de imágenes | IMPLEMENTADO |
  | Zoom/pan en visor | IMPLEMENTADO |
  | Exportación PDF (visor/comparador) | IMPLEMENTADO |
  | Respaldo del editor de casillas | IMPLEMENTADO (archivo backup) |
  | Divergencia de metadatos (aviso previo) | IMPLEMENTADO |
  | Renombrador de Fotos (herramienta standalone en `src/renombrar_fotos_gui.py`) | IMPLEMENTADO y VERIFICADO (46/46 pytest + 11 unittest); lanzado desde MetaTag como subproceso independiente (botón HERRAMIENTAS AVANZADAS) |
  | Exportación CSV | NO IMPLEMENTADO (solo lectura de CSV/Excel) |
  | Modo simulación / dry-run | NO IMPLEMENTADO |
  | Sistema de backup/respaldo para renombrado | NO IMPLEMENTADO |

  ---

  ## 10. Funcionalidades pendientes

  Confirmadas como NO implementadas en el código actual:

  - **Exportación CSV** del módulo correspondiente (no hay `to_csv`/export; solo carga de `.csv`).
  - **Dry-run / modo simulación**, donde aplique (no encontrado en el proceso por lote).
  - **Sistema de backup/respaldo para renombrado**, donde aplique.
  - **Mejoras futuras del módulo de estadísticas** (incluida una posible mayor separación).

  > **Importante:** no confundir las funcionalidades existentes de Image Sync (reordenación del orden) con las del nuevo módulo de renombramiento que está en desarrollo (sección 11).

  ---

  ## 11. Herramienta: Renombrador de Fotos

  Estado actual: **INTEGRADO COMO LANZADOR** (2026-08-10). La herramienta vive en
  `src/renombrar_fotos_gui.py` (migrada desde `tools/renombrador/`, ya eliminado),
  usa CustomTkinter (dependencia **aislada** de la herramienta, ver sección 8) y se
  lanza desde MetaTag mediante el botón **"🖼 Renombrador de Fotos"** de la sección
  HERRAMIENTAS AVANZADAS del panel izquierdo. Verificado: **97/97 tests de proyecto
  + 46/46 pytest del renombrador** (`tests/test_renombrador_pytest.py`) verdes;
  smoke real con display: MetaTag permanece `normal` (nunca se retira), el
  subproceso del renombrador arranca, MetaTag sigue responsivo y sin callbacks
  `after` huérfanos al cerrar el renombrador; funciona desde cualquier cwd.

  **Método de lanzamiento:** `subprocess.Popen([sys.executable, str(Path(__file__).resolve().parent / "renombrar_fotos_gui.py")])`
  (`_launch_renombrador` en `metatag_v8.py`). A diferencia del patrón `launch_visor`
  (que hace `withdraw()` + polling), el lanzador del renombrador **NO** retira MetaTag
  ni registra callbacks; cada lanzamiento crea un proceso independiente cuyo ciclo de
  vida no afecta a la app principal. Si el proceso no puede arrancar (p. ej. falta
  CustomTkinter), se muestra `messagebox` + log sin crashear. La ruta se construye
  desde `__file__` (independiente del cwd de invocación).

  Nota de nomenclatura: la herramienta se auto-titula **"Renombrador de Fotos v4.0"**
  (no "Image Sync"). En modo copia crea la subcarpeta `Renombradas/`. El flujo por
  defecto es **posicional** (foto 1 ↔ fila 1 de la columna del Excel); con el modo
  "matching seguro" ON cada nombre busca SU foto y el fallback posicional queda
  PROHIBIDO (estado `error`, FASE A). Procesa `min(fotos, nombres)` parejas y
  conserva la extensión original.

  **Concepto funcional:**

  ```
  Excel        → nombres objetivo
  Fotografías  → nombres actuales
  ```

  La herramienta establece correspondencias entre fotografías y registros del Excel y, posteriormente, **renombra** las imágenes de acuerdo con una columna del Excel.

  **Interfaz:** se auto-presenta como **"Renombrador de Fotos"** con subtítulo **"Renombra fotos desde Excel"**.

  **Estado de funcionalidades (distinguir claramente):**

  - IMPLEMENTADO EN `src/renombrar_fotos_gui.py` (2026-08-10): selección de carpeta de fotografías, carga de Excel/CSV con selección de hoja y columna, ordenamiento de fotografías, vista previa con detección de duplicados, renombramiento real (con reemplazo de destino existente), deshacer de la última operación, modo copia a `Renombradas/` y **modo "matching seguro"** (cada nombre busca SU foto). FASE A (2026-08-10): el motor es un port puro del `_find_image_ex` validado (`src/metatag_matching.py`), sin fallback posicional silencioso (estado `error` si no hay motor/carpeta) y con 7 estados coherentes `ok / ya_correcto / conflicto / duplicado / not_found / ambiguo / error`.
  - IMPLEMENTADO (migración, 2026-08-10): ubicación definitiva `src/renombrar_fotos_gui.py`, import de `ImageMatcher` con `sys.path` apuntando al propio `src/`, botón lanzador en `metatag_v8.py`, suite pytest en `tests/test_renombrador_pytest.py`, dependencia documentada en `requirements-renombrador.txt` (customtkinter/pandas/openpyxl/pillow), `tools/renombrador/` eliminado.
  - PENDIENTE (integración de configuración dentro de MetaTag): diálogo de configuración del renombrador embebido en la app (hoy se lanza la ventana completa de la herramienta), generación de correspondencias por matching desde la UI de MetaTag, simulación/dry-run, registro de operación, y refinar la detección de conflictos a nivel de interfaz.

  ---

  ## 12. Sincronización Excel ↔ Programa (idea futura)

  **NO implementado.** Idea:

  ```
  Excel ↔ MetaTag
  ```

  De modo que si el usuario modifica y guarda el Excel externamente, MetaTag pueda detectar el cambio y actualizar sus datos.

  Antes de implementarlo será necesario estudiar: detección de cambios del archivo, recarga, conflictos, cambios simultáneos, preservación de modificaciones y dirección de sincronización.

  ---

  ## 13. Problemas y warnings

  ### Ambigüedad en el matching por `id-suffix` (resuelto 2026-08-10)

  El fallback final de `_find_image` comparaba solo `(numero_pieza, sufijo_F/R/P)` y devolvía el **primer** archivo de la caché cuya clave coincidiera. Con `0053_EC_RS_372_F.jpg.JPG` y `0053_EC_C7_XII_372_F.jpg.JPG` la clave `(0053, F)` es compartida → podía mapear una fila al archivo equivocado. Resuelto con `_find_image_ex` (detección explícita de ambigüedad).

  La ambigüedad `(0053, F)` es un **riesgo estructural real** del matching: siempre que dos archivos compartan `(numero_pieza, sufijo)` y una fila del Excel solo llegue a ese fallback (sin match previo por stem), se rechaza correctamente con `status="ambiguous"` en lugar de elegir el primero de la caché. El dataset actual NO activa ese riesgo porque esos archivos encuentran antes una coincidencia por stem-exacto; queda cubierto con fixture de prueba.

  ### Línea base del dataset (269 IDs) — corrección de línea base

  NOTA: el reparto 18/184/21/44 es una **corrección de la línea base** (medición con las funciones reales del módulo), no un problema del algoritmo. Un informe previo estimó "65 id-suffix" usando una replicación con `_normalize_numbers` distinto; la función real convierte cada secuencia de dígitos con `int()`, y por eso 21 de esos 65 caen en `normalize`.

  Línea base real verificada:

  - 18 → nombre exacto
  - 184 → stem exacto
  - 21 → normalize
  - 44 → id-suffix
  - 267 → matches totales
  - 0 → ambigüedades en el dataset actual
  - 0 → reusos
  - 2 → IDs sin archivo (`0053_EC_C7_XII_372_R.jpg`, `0055_EC_C7_VI_146_P.jpg`)
  - 2 → archivos huérfanos (`0053_EC_RS_372_F.jpg.JPG`, `0059_EC_RS_109_P.jpg.JPG`)

  ### Complejidad de `ExcelGrid.redraw` (resuelto 2026-08-10)

  - **BUG de complejidad**: `redraw()` (en `src/metatag_widgets.py`) llamaba `self._col_fully_selected(ci)` **dentro del bucle de celdas** (filas × columnas visibles). `_col_fully_selected` recorre todas las filas (`all((r, ci) in selected_cells …)`), así que el coste por redraw era del orden de `filas_visibles × columnas × filas_totales`.
  - **Causa**: el estado "columna totalmente seleccionada" es una propiedad de la columna, pero se recalculaba por cada celda en lugar de una vez por columna.
  - **Solución**: se precalcula `col_sel_map = {ci: self._col_fully_selected(ci) for ci, _, _ in vis_cols}` una sola vez por `redraw()` (antes del bucle de celdas) y se usa `col_sel_map[ci]` en el encabezado y en cada celda. `_col_fully_selected` se invoca ahora **como máximo una vez por columna visible por redraw**.
  - **Archivos modificados**: `src/metatag_widgets.py` (solo `redraw`).
  - **Tests creados**: `tests/test_grid.py` — equivalencia lógica naive vs. mapa (8 estados de selección), comportamiento real de `_col_fully_selected` (7 casos de selección), conteo de llamadas con Tk real (1× por columna visible; con columnas ocultas solo las visibles) y smoke test real de ExcelGrid.
  - **Resultado**: 23/23 pruebas verdes (11 Bloque 1 + 12 Bloque 2); `py_compile` OK.
  - **Smoke test Tk**: SÍ fue posible (display `:0` funcional); se ejecutó de verdad y quedó cubierto por `ExcelGridTkTestCase` (con `skipUnless` para entornos sin display).

  ### Robustez del procesamiento en segundo plano (Bloque 3, resuelto 2026-08-10)

  El proceso por lote por celdas (`_start_processing` → `_process_all` → `_process_queue`) no manejaba correctamente el ciclo de vida del worker en segundo plano:

  - **Acceso a Tk desde el hilo worker:** `_process_all` leía `self.omit_empty_var`/`self.meta_mode_organized` (variables Tk) directamente desde el hilo secundario. Resuelto con **snapshot en el hilo principal**: `_start_processing` captura `df`, `meta_by_row`, `organizado`, `omit_empty`, `empty_cnt` antes de arrancar el hilo; el worker ya no toca la UI.
  - **Terminación inesperada no detectada:** si el worker moría sin emitir `done`/`error`, la cola seguía en polling infinito. Ahora `_process_queue` comprueba `thread.is_alive()`; si el hilo murió sin señal terminal, se registra y restaura la UI (`_proc_finish_ui("unexpected", …)`), sin reprogramar `after`.
  - **Cancelación a mitad de camino:** nueva señal cooperativa `_proc_cancel` (Event) + botón "Cancelar" (`_cancel_btn`). El worker la comprueba por fila y emite `("cancelled", …)` con el conteo parcial; `_cancel_processing` la activa desde la UI.
  - **Excepción no controlada en el worker:** `_process_all` ahora envuelve todo el bucle en `try/except` y emite `("error", …)` con `traceback` en el log, en lugar de morir silenciosamente.
  - **Doble procesamiento:** guard en `_start_processing` si `_proc_thread` está vivo.
  - **Restauración de UI centralizada:** `_proc_finish_ui(msg_type, message)` restaura cursor, botón escribir y oculta el botón de cancelar para `done`/`error`/`cancelled`/`unexpected`; solo `done` y `error` muestran messagebox.

  ### `invalid command name` de Tk al destruir ExcelGrid (resuelto 2026-08-10)

  `ExcelGrid._schedule_redraw` programaba `after_idle(_deferred_redraw)` sin guardar su id; si el grid se destruía antes de que el evento idle se ejecutara, Tk emitía `invalid command name "..._deferred_redraw"`. Resuelto guardando `_redraw_after_id` y cancelándolo en un nuevo `ExcelGrid.destroy()`.

  ### Responsividad de la interfaz (Bloque 4, resuelto 2026-08-10)

  Auditoría completa de la UI frente a resoluciones/tamaños de ventana/DPI y reparación mínima y localizada (sin refactor masivo):

  - **Exportación de gráficas desacoplada del tamaño de ventana** (`src/metatag_graficas.py`): la figura se estiraba con el contenedor (matplotlib 3.11) y `savefig(dpi=300)` usaba ese `figsize`, así que el PNG dependía de la ventana. Ahora `export_chart` fija temporalmente la figura a `EXPORT_FIG_SIZE=(12,8)"` a `EXPORT_DPI=200` (desconectando el bind `<Configure>` del canvas), guarda y restaura. Verificado: **PNG idéntico (2283×1464)** exportado con la ventana a 900×600 y a 1920×1080.
  - **Popup selector limitado en X e Y** (mismo archivo): antes solo había clamp en X; el popup podía salir por el borde inferior. Ahora también clamp en Y.
  - **Insights con scrollbar**: `insight_text` (lista de categorías) ahora vive en un frame con `ttk.Scrollbar` vertical; el contenido largo ya no se corta.
  - **`_place_labels_clean` con fuente adaptativa**: `fontsize` se reduce con el número de categorías (`max(6, min(9, 90//n))`) para evitar solapamiento con hasta 39 categorías.
  - **`show_stats.minsize` escalado** por `current_scale` (antes 800×500 fijo).
  - **Ventanas secundarias centradas y limitadas a pantalla** (`src/metatag_v8.py`): nuevo helper `_clamp_toplevel(win, parent, w, h)` aplicado a atajos, lupa, columnas de orden, columnas de metadatos, orden de fotos, diálogo Sí/No y las 3 ventanas de progreso (antes podían abrirse fuera de pantalla). El popup de tema también se limita al área visible.
  - **Visor** (`src/Visor.py`): `minsize` bajado de 960×640 a **860×520** (a 1024×768 el 90% de la pantalla son 921×691; el minsize de 960 forzaba salirse por la derecha); el diálogo del comparador ahora es redimensionable y limitado a la pantalla; añadido `SetProcessDpiAwareness` portable para Windows (el Visor se lanza como subproceso y no heredaba el del launcher → se veía borroso en HiDPI).
  - **ExcelGrid**: revisado, **sin cambios necesarios** (ya tiene canvas escalable + scrollbars + culling). El `fill="x"` del `ImageBrowser` es intencional: su listbox interno ya expande con scrollbar propio y la vista previa absorbe el espacio extra; convertirlo a `expand=True` quitaría espacio al preview.
  - **DPI**: no se toca el scaling global; la figura usa `_dpi_screen` real y las fuentes escalan con `current_scale` (basado solo en ancho — limitación documentada, no corregida por no estar justificada por las pruebas).
  - **Tests nuevos** (`tests/test_responsive.py`, 12): `_clamp_toplevel` nunca deja ventanas fuera de pantalla (6 resoluciones × 5 tamaños), centrado sobre padre, ventana grande limitada a pantalla, fuente adaptativa de `_place_labels_clean`, y smoke real de Tk: `show_stats` se redimensiona dentro de la pantalla, minsize escalado, scrollbar de insights presente, **exportación idéntica a 2 tamaños de ventana**, ventana principal dentro de la pantalla (incluido minsize 860×520) y ventanas secundarias dentro de la pantalla.
  - **Resultado: 52/52 tests OK** (40 previos + 12 nuevos); `py_compile` OK en los módulos tocados.
  - **Rango soportado documentado**: **≥1024×768** (soporte oficial y adaptativo). 900×600 y 800×600 = degradación controlada / best effort (el minsize 860×520 de la ventana principal limita 800×600 a 860×600 sin romper nada). Solo se probó físicamente en 1920×1080; el resto se simuló con `geometry()`.

  ### Salida corrupta de terminal

  Durante la sesión de auditoría (2026-08-10) se detectó que ciertos comandos largos producían **salida corrupta/duplicada** en la terminal (grep/lecturas extensas devolvían contenido repetido). Esto **no significa necesariamente que el código esté corrupto**. La verificación se realizó escribiendo la salida a archivos temporales y leyéndolos con la herramienta Read, confirmando el contenido real de los módulos. Estado: investigado / workaround aplicado. No se ha establecido causa definitiva.

  ### Warning de openpyxl

  Al leer el `.xlsx` con `read_excel`, openpyxl emite:

  ```
  UserWarning: Data Validation extension is not supported and will be removed
  ```

  Confirmado; no se ha verificado ninguna consecuencia sobre los datos leídos (los datos cargaron correctamente: 269×19).

  ### Otros detalles observados

  - Cabeceras del Excel con espacios finales (`Sitio `, etc.); el código los elimina con `.strip()`.
  - El proyecto usa nombres de módulo prefijados `metatag_*`; no hay `requirements.txt` (las dependencias se gestionan vía los scripts de instalación).

  ### Renombrador standalone: comportamientos revisados (2026-08-10)

  Observaciones de la fase de integración sobre `src/renombrar_fotos_gui.py` (antes `tools/renombrador/renombrar_fotos_gui.py`), todas resueltas:

  - **"Nombres ya correctos" se reporta como conflicto** → RESUELTO: si `dest == origen` (mismo inodo vía `_same_file`) se cuenta como `success` sin tocar el archivo (estado `ya_correcto`). Verificado por `tests/test_renombrador.py` y `test_conflicto_vs_ya_correcto_sin_falsa_alarma`.
  - **Bloque `if dest.exists(): dest.unlink()` duplicado** → RESUELTO: el `unlink()` solo queda en `undo_last` (modo copia), que es su uso legítimo.
  - **Flujo posicional**: `zip(fotos, nombres)` sigue siendo el modo por defecto (compatibilidad standalone); con el modo "matching seguro" ON cada nombre busca SU foto y el fallback posicional silencioso queda PROHIBIDO (FASE A).

  ### Fallback posicional silencioso y estado `ambiguous` (resuelto 2026-08-10)

  - **Fallback posicional silencioso (FASE A)**: `_build_plan` y `_build_plan_matching` del renombrador degradaban a posicional si el motor de matching no estaba disponible o faltaba la carpeta, con matching seguro activo → renombraba por posición fotos que el usuario esperaba emparejadas por nombre. Resuelto con `_error_plan` (estado `error`, `src=None`); `rename_all` lo omite sin tocar disco.
  - **Estado `ambiguous` (inglés) vs `ambiguo` (canónico de la UI)**: el modelo emitía `"ambiguous"` y la vista espera `"ambiguo"` en `PLAN_STATES`; las filas ambiguas se veían sin etiqueta ni color. Unificado a `"ambiguo"`.
  - **`pHash` muerto en `src/metatag_matching.py`**: `_pixel_hash`/`_hamming`/`_hash_cache` nunca participaban en las decisiones (un informe previo los describió como desambiguación por contenido). Eliminados junto con los imports de PIL/hashlib/difflib; el módulo quedó como port puro del algoritmo validado.

  ### Artefacto de pruebas Tk: `_default_root` huérfano (2026-08-10)

  En los tests, crear un primer `tk.Tk()` (p. ej. una comprobación de display) deja esa raíz como `tkinter._default_root`; los `tk.BooleanVar()` sin master del picker se ligan a ese intérprete y `select()`/`get()` leen variables Tcl distintas (resultado vacío en `Aceptar`). En producción esto NO ocurre (la app es el primer `Tk`). Los tests lo evitan destruyendo la raíz de comprobación y reseteando `tk._default_root = None`. No requiere cambio en `metatag_v8.py`.

  ### Migración del renombrador a `src/` (resuelto 2026-08-10)

  - **Ruta del import de `ImageMatcher`:** el bloque que insertaba `Path(__file__).resolve().parents[2] / "src"` en `sys.path` asumía vivir en `tools/renombrador/`; ahora el script está en `src/`, junto a `metatag_matching.py`, y usa `Path(__file__).resolve().parent`. No depende del cwd de invocación (verificado lanzando desde la raíz y desde `/tmp`).
  - **Archivo de estado:** `_STATE_FILE` sigue siendo `Path(__file__).parent / ".renombrador_state.json"`; con la migración su ruta pasó de `tools/renombrador/` a `src/`. Es estado local de la herramienta (preferencias/última carpeta), sin impacto funcional; no se versiona.
  - **Launchers `.sh`/`.bat` de `tools/renombrador/`:** se eliminaron (creaban un `.venv` local obsoleto); el lanzamiento oficial es el botón de MetaTag. Lanzamiento directo: `.venv/bin/python src/renombrar_fotos_gui.py`.

  ---

  ## 14. Decisiones arquitectónicas

  | Decisión | Motivo |
  |---|---|
  | Funciones puras de escritura en `metatag_writer.py` | Testeable sin Tkinter; separa lógica crítica de la UI |
  | UI principal en `metatag_v8.py` | App monolítica heredada; concentra flujo de trabajo |
  | Módulo de gráficas separado (`metatag_graficas.py`) | Aislar matplotlib del flujo principal |
  | `ExcelGrid` propio con viewport culling (`metatag_widgets.py`) | Rendimiento con datasets grandes; renderiza solo lo visible |
  | Uso de DataFrame (pandas) como modelo de datos | Facilita carga, filtro, orden y mapeo |
  | Procesamiento por copias (`shutil.copy2` + escritura en la copia) | Nunca se modifican los originales (seguridad de datos) |
  | Reanudación mediante `_batch_progress.json` | Permitir retomar lotes interrumpidos sin reprocesar |
  | `CONTEXTO_PROYECTO.md` como memoria persistente | Que una sesión futura pueda continuar sin re-descubrir la arquitectura |

  ---

  ## 15. Convenciones del proyecto

  - Interfaz en **español**.
  - Funciones puras cuando corresponde (escritura de metadatos).
  - UI separada de lógica donde ya existe separación (writer, graficas, widgets).
  - **Evitar modificar originales** durante el procesamiento por lote (trabajar con copias).
  - El código real es la fuente de verdad (este documento se actualiza contra el código).

  ---

  ## 16. Historial de sesiones

  ### 2026-08-10 — Auditoría y consolidación de contexto

  - Auditoría completa del proyecto (estructura, módulos, líneas, dependencias, datos).
  - Verificación del código real mediante archivos temporales + Read (workaround por salida corrupta de terminal).
  - Identificación de la arquitectura y del flujo de escritura (copias, reanudación).
  - Revisión de Image Sync (bidireccional, solo orden) y del módulo de estadísticas.
  - Identificación de problemas (salida corrupta, warning de openpyxl) y pendientes (export CSV, dry-run, backup, mejoras de estadísticas).
  - Creación de `CONTEXTO_PROYECTO.md` como memoria técnica canónica.

  ### 2026-08-10 — Matching seguro: detección de ambigüedades (Bloque 1)

  - Línea base REAL del dataset 269 (con funciones reales): 267 matches — **18 nombre-exacto · 184 stem-exacto · 21 normalize · 44 id-suffix**, 2 missing (`0053_EC_C7_XII_372_R.jpg`, `0055_EC_C7_VI_146_P.jpg`), 2 huérfanas (`0053_EC_RS_372_F.jpg.JPG`, `0059_EC_RS_109_P.jpg.JPG`), 0 reusos.
  - Corrección de línea base (NO un problema del algoritmo): un informe previo estimó "65 id-suffix" con una replicación de `_normalize_numbers` distinta; el real convierte dígitos con `int()`, así que 21 de esos 65 caen en `normalize` y 44 en `id-suffix`.
  - `_find_image_ex`: implementación ÚNICA del matching; devuelve `(path, status, candidates)` con `"ok" | "not_found" | "ambiguous"`. Los fallbacks id-suffix y substring recogen TODOS los candidatos; con 2+ devuelven `ambiguous` (sin elegir el primero arbitrariamente). `_find_image` ahora delega.
  - `_match_rows_to_files` devuelve 5-tupla con `matches_ambiguas`; `_sync_images_to_excel` loguea las filas ambiguas con sus candidatos y NO las mete en `no_encontradas`.
  - `_process_all`: en ambigua registra `⛔ Ambigua: …` con los nombres candidatos, incrementa `amb` y lo muestra en el resumen (`· N ambigüedades`).
  - Tests: `tests/test_matching.py` (7 unitarios con fixtures temp, incluido el caso 0053 ambiguo) y `tests/test_dataset_269.py` (4 de integración: correspondencias idénticas al original, distribución de métodos 18/184/21/44, missing y huérfanas esperadas, 0 reusos). Todos verdes.

  ---

  ### 2026-08-10 — Optimización de `ExcelGrid.redraw` (Bloque 2)

  - Bug de complejidad en `ExcelGrid.redraw`: `_col_fully_selected(ci)` se llamaba por cada celda; se precalculó `col_sel_map` una vez por columna visible.
  - Sin cambios de comportamiento observable (mismo modelo de selección, mismos colores/estilos, mismo viewport culling, mismas columnas).
  - `tests/test_grid.py`: equivalencia naive vs. mapa, selección, conteo de llamadas con Tk real y smoke test (display `:0` disponible).
  - Resultado: 23/23 verdes; commit `perf: optimizar redibujado del ExcelGrid`.

  ---

  ### 2026-08-10 — Robustez del procesamiento en segundo plano y cancelación (Bloque 3)

  - **17 tests nuevos en `tests/test_queue.py`**, todos verdes:
    - `WorkerTestCase` (7): worker exitoso; archivo inexistente; matching ambiguo; excepción interna (emite `error`, no muere); cancelación antes de empezar; cancelación a mitad (`1 escritas`); y prueba de que el worker NO accede a Tk (FakeApp con "trampas" que lanzan `AssertionError` si el hilo toca `grid`/variables Tk).
    - `QueueTestCase` (7): restauración de UI tras `done`/`error`/`cancelled`; polling continuo con worker vivo; detección de muerte inesperada; sin pollings infinitos tras la muerte; guard que impide lanzar un segundo procesamiento.
    - `ProcessingSmokeTestCase` (3): **smoke tests reales de Tk** (display `:0` funcional): procesamiento normal con restauración de UI, cancelación a mitad de camino determinista, y error interno con restauración. Usan app real con messagebox/save-config parcheados y datos temporales.
  - **Bugs resueltos en la infraestructura de test (no en la lógica):** los mensajes `("log", …)` del worker solo se vuelcan a la UI vía `_process_queue`, así que los tests de worker aplican los no-terminales con el helper `apply_msgs`; el `messagebox` real abría un modal y colgaba las pruebas → sustituido por registro en `setUp`.
  - **Cambio adicional aprobado en `src/metatag_widgets.py`:** `ExcelGrid.destroy()` cancela el `after_idle(_deferred_redraw)` pendiente (evita el `invalid command name` de Tk al destruir el widget; también aplica a la app real al cerrar justo después de cargar datos).
  - **Resultado final: 40/40 tests OK** (23 Bloques 1–2 + 17 Bloque 3); `py_compile` OK.
  - **Limitaciones:** los smoke tests de Tk requieren display (`skipUnless`); en entornos headless quedan pendientes de validación manual. El tiempo de espera de `_wait_finished` es 15 s (suficiente para el dataset de prueba).

  ---

  ### 2026-08-10 — Responsividad de la interfaz (Bloque 4)

  - Auditoría completa de la UI (`metatag_v8.py`, `metatag_graficas.py`, `metatag_widgets.py`, `Visor.py`) frente a resoluciones (1024×768 a 2560×1440 + 900×600/800×600 de degradación), tamaños de ventana y DPI.
  - Reparaciones mínimas y localizadas: export de gráficas independiente de la ventana, popup selector con clamp X/Y, insights con scrollbar, fuente adaptativa de etiquetas, minsize escalado de `show_stats`, helper `_clamp_toplevel` para todas las Toplevels, popup de tema limitado a pantalla, Visor (minsize 860×520, comparador redimensionable, DPI Windows portable).
  - `tests/test_responsive.py`: **12 tests** (headless de clamp y etiquetas + smoke real de Tk).
  - Resultado: **52/52 OK**; `py_compile` OK; commit `fix: mejorar responsividad de la interfaz`.

  ---

  ### 2026-08-10 — Fix scroll ventanas de selección + tests renombrador (Fases 4–6)

  **Dependencias:** se instaló `customtkinter` en el `.venv` (dependencia del renombrador). `smoke_renombrador.py` real: MainView v4 construida en 1920×1080, ventana 1100×864, OK.

  **Fase 4 — scroll de las ventanas de selección de columnas (`src/metatag_v8.py`):**
  - Se detectó que `_batch_pick_columns` y `_pick_sort_columns` enlazaban la rueda del ratón **por fila** (`<MouseWheel>`, `<Button-4>`, `<Button-5>` sobre cada Checkbutton) y además con `bind_all`: duplicación de handlers y scroll global (cualquier ventana de MetaTag con rueda desplazaba el picker). El `bind_all` de la ventana nueva además podía colisionar con un cierre en curso.
  - Fix mínimo sin tocar la lógica de selección: **nuevo helper `_build_scroll_picker_window`** (Toplevel con canvas + inner + scrollbar; rueda enlazada UNA vez en la Toplevel → funciona sobre canvas, labels, Checkbuttons, scrollbar y huecos vía bindtags; sin `bind_all`; bindings mueren con la ventana). Ambas ventanas de selección (columnas de metadatos y de orden) se refactorizaron para usarlo. `_batch_pick_columns` conserva Todas/Ninguna/Invertir, contador `N / M` y preservación de la columna de imagen al invertir. El helper acepta `on_cancel_hook` para conservar la limpieza de trazas write de Tk que hacía el código original al cancelar (el sort no tiene trazas).
  - `py_compile` OK y **52/52 tests previos siguen verdes** (sin regresiones).

  **Fase 5 — matriz de resoluciones del picker:** test `ClampMatrixTestCase` (8 resoluciones: 640×480 → 2560×1440) que aplica la fórmula de dimensionado real del picker y verifica que la ventana siempre cabe en pantalla (x, y ≥ 0; x+ancho ≤ pantalla; ≥24px). Complementa los smoke scroll de `ColumnPickerScrollTestCase`.

  **Fase 6 — tests nuevos (30):**
  - `tests/test_renombrador.py` (**11, sin display**): nombres simples, extensiones distintas, ceros, archivos faltantes (min), archivos adicionales, duplicados (se saltan), conflicto con destino existente, nombres ya correctos (documentado: se reporta conflicto), cancelación cooperativa, `undo_last` y `build_preview` (marca duplicados). Comportamiento verificado por ejecución real con directorios temporales.
  - `tests/test_column_picker.py` (**19, smoke Tk real**): apertura con 19 columnas y contador `18 / 19`, Todas/Ninguna/Invertir (preserva columna imagen), toggle manual, Aceptar devuelve la selección (18), Cancelar None **y Cancelar limpia las trazas write de Tk**, scroll por rueda sobre canvas/texto/Checkbutton/scrollbar, scrollbar anclada y visible con desbordamiento, selección estable tras desplazar, redimensionado sin overflow horizontal, volver arriba, orden por selección y matriz de resoluciones.
  - Para los smoke se parchea `wait_window` (pump que espera el cierre real) y se respeta la advertencia `_default_root` (sección 13).

  **Resultado: 82/82 tests OK** (52 previos + 30 nuevos). Commit `feat: fijar scroll de selección de columnas y añadir tests del renombrador`.

  ---

  ### 2026-08-10 — FASE A: seguridad del renombrador + matching seguro como port puro

  **Hallazgos de la auditoría del matching seguro (src/metatag_matching.py):**

  - **El "pHash" reportado era código muerto**: `_pixel_hash`/`_hamming`/`_hash_cache`/`_image_hash` existían pero NO participaban en ninguna decisión de `find_image_ex` (que solo usaba umbral de `SequenceMatcher`). Un informe previo lo describió como "para desambiguar" sin ser cierto. Se eliminó por completo (también los imports de `PIL`, `hashlib`, `difflib`, `unicodedata`).
  - **El matcher era DISTINTO al algoritmo validado**: usaba tokenización + `SequenceMatcher` con umbral difuso, mientras la línea base validada (`_find_image_ex` de `metatag_v8.py`) usa la jerarquía direct → nombre-exacto → stem-exacto → clean → normalize → id-suffix → substring. Habría cambiado las correspondencias del dataset (p. ej. el caso 0053 hubiera pasado de stem-exacto a ambiguo).

  **Solución aplicada:** `src/metatag_matching.py` se reescribió como **port FIEL y determinista** de `_find_image_ex` (mismo orden de pasos, misma construcción de caché con `rglob` recursivo, mismas claves), sin Tkinter/PIL/difflib, con índice ordenado por ruta para que la decisión no dependa del orden del sistema de archivos. API: `ImageMatcher.find_image_ex(name, folder) → (path, status, candidates)` (`"ok" | "not_found" | "ambiguous"`) + `find_image_ex_with_method` para etiquetar el paso (auditoría/tests).

  **Verificación contra el dataset real (Finales 1 a 103, 269 IDs):** el port reproduce EXACTAMENTE la línea base — **267 ok / 0 ambiguas / 2 not_found** (las mismas `0053_EC_C7_XII_372_R.jpg`, `0055_EC_C7_VI_146_P.jpg`), **distribución 18/184/21/44**, **0 reusos**, **2 huérfanas** (`0053_EC_RS_372_F.jpg.JPG`, `0059_EC_RS_109_P.jpg.JPG`), **0 discrepancias** vs la réplica de referencia. Coste: índice 3,5 ms + 96 ms para recorrer los 269 (~0,36 ms/nombre).

  **Bugs corregidos en el renombrador (`tools/renombrador/renombrar_fotos_gui.py`):**

  - **Fallback posicional silencioso (criticidad alta)**: con matching seguro activo, si el motor no estaba disponible o faltaba la carpeta, `_build_plan`/`_build_plan_matching` degradaban silenciosamente a posicional → renombraban por posición fotos que el usuario esperaba emparejadas por nombre. Ahora un plan de **ERROR** (`_error_plan`, estado `error`, `src=None`) y `rename_all` lo omite sin tocar disco. **Queda prohibido el fallback posicional cuando matching está ON.**
  - **Estado `ambiguous` vs `ambiguo` (criticidad media)**: el modelo emitía el estado en inglés (`"ambiguous"`) pero la vista define `PLAN_STATES` con `"ambiguo"`; las filas ambiguas se renderizaban sin etiqueta ni color. Unificado a `"ambiguo"` en plan, `_skip_text`, `rename_all` y docstrings (el matcher sigue devolviendo `"ambiguous"` como estado interno de matching).

  **Tests:** `tests/test_metatag_matching.py` (14: pureza sin tkinter/PIL/difflib, ambigüedad id-suffix y substring, normalize, not_found, determinismo, jerarquía de métodos, equivalencia dataset 269). En `tools/renombrador/test_renombrador.py` +4: motor no disponible → error (nada se renombra), sin carpeta → error, reuso → duplicado, ambiguo → `"ambiguo"`. Resultado: **97/97 tests de proyecto + 46/46 pytest del renombrador, todos verdes**.

  **FASE B — PENDIENTE / PROPUESTA FUTURA (NO implementada, sin cambios de código):** contenido visual/perceptual para desambiguar (hash perceptivo real de píxeles, comparación por contenidos con PIL) como último recurso SOLO cuando la clave textual sea ambigua y las candidatas sean imágenes reales. Implica reabrir `ImageMatcher` con dependencia opcional (PIL) y umbrales de similitud — exactamente lo que la FASE A eliminó por no estar validado. Requiere su propio análisis de impacto sobre la línea base (267/0/2) antes de activarse; mientras tanto, la ambigüedad textual se reporta y NO se elige candidato (comportamiento seguro actual).

  ---

  ### 2026-08-10 — Migración e integración del Renombrador en MetaTag (FASE 7)

  **Objetivo:** unificar el renombrador en el proyecto sin copias duplicadas, lanzable desde MetaTag como proceso independiente, y eliminar `tools/renombrador/`.

  **Cambios de código:**
  - `tools/renombrador/renombrar_fotos_gui.py` → `src/renombrar_fotos_gui.py` (`git mv`, sin duplicar). Se corrigió el bloque de import de `ImageMatcher`: `_PROJECT_SRC` pasa de `parents[2] / "src"` a `Path(__file__).resolve().parent` (ahora el script vive junto a `metatag_matching.py`; el bloque inserta `src/` en `sys.path` y funciona desde cualquier cwd).
  - `src/metatag_v8.py`: botón **"🖼 Renombrador de Fotos"** en la sección HERRAMIENTAS AVANZADAS del panel izquierdo (`_build_control_panel`) + método `_launch_renombrador` (~988): `subprocess.Popen([sys.executable, str(Path(__file__).resolve().parent / "renombrar_fotos_gui.py")])`, **sin** `withdraw()` ni polling (a diferencia de `launch_visor`), try/except con `messagebox` + `logging`. MetaTag sigue funcionando mientras el renombrador está abierto y no deja callbacks huérfanos.
  - Tests: `tools/renombrador/test_renombrador.py` → `tests/test_renombrador_pytest.py` (suite pytest original, 46 tests; `sys.path` → `../src`). `tests/test_renombrador.py` (unittest, 11) actualizado a `../src`.
  - `tools/renombrador/` eliminado por completo: `requirements.txt` → `requirements-renombrador.txt` (raíz), launchers `run_renombrador.sh/.bat` eliminados (creaban `.venv` local obsoleto). El lanzamiento oficial es el botón de MetaTag; lanzamiento directo documentado.

  **Verificación:**
  - `py_compile` OK en `src/renombrar_fotos_gui.py`, `src/metatag_v8.py` y ambos test files.
  - **97/97 tests de proyecto** (`python -m unittest discover -s tests`) + **46/46 pytest** del renombrador, todos verdes.
  - Smoke real con display `:0` (script `/tmp/opencode/smoke_btn.py`): MetaTag pasa de `normal` a `normal` (nunca `withdrawn`); el subproceso del renombrador nace; MetaTag sigue responsivo; al cerrar el renombrador quedan **0 callbacks `after` pendientes** en MetaTag.
  - Portabilidad: el botón funciona lanzando MetaTag desde la raíz y desde `/tmp` (ruta construida desde `__file__`). El subproceso importa `ImageMatcher` correctamente (no es `None`).
  - Estado de `_STATE_FILE` relocado a `src/.renombrador_state.json` (estado local, no versionado).

  **Resultado:** commit `feat: integrar renombrador de fotografías en MetaTag`. Sin regresiones sobre matching validado (`src/metatag_matching.py` no se tocó), `ExcelGrid.redraw`, `process_queue` ni responsividad. PENDIENTE (fuera de esta fase): diálogo de configuración embebido, dry-run, registro de operación y refinar conflictos a nivel de interfaz (sección 11).

  ---

  ### 2026-08-11 — Fuente de verdad técnica de temas + Renombrador sobre temas de MetaTag

  **Objetivo:** que MetaTag y el Renombrador compartan un único sistema de temas
  (los 3 canónicos de MetaTag), con `metatag_theme.py` como **fuente de verdad
  técnica** y la regla estricta de "cero diferencia visual" al migrar.

  **Decisiones (aprobadas por el usuario):**
  - Nuevo `src/metatag_theme.py` (módulo puro, sin tkinter/customtkinter): `THEMES`
    (3 temas, valores **verbatim** de `metatag_v8.py`), `THEME_ORDER`,
    `DEFAULT_THEME` = `Arqueológico (Oscuro Refinado)`, `THEME_ICONS`,
    `ACCENT_TEXT` = `#FFF5E8`, motor de fuentes (misma referencia 1920, mismo rango
    `(0.82, 1.35)`, misma fórmula `max(floor, int(base*scale))`), helpers de color
    (`mix`, `relative_luminance`), `fit_to_screen`, y adaptadores:
    `TkThemeAdapter` (tokens canónicos **sin transformación**) y
    `CustomTkinterThemeAdapter` (traduce al esquema del Renombrador: mapeos 1:1
    `subtext→text2`, `green→ok`, `red→err`, `yellow→warn`, `accent2→accent_hover`,
    `surface2→btn_ghost_bg`; derivaciones deterministas documentadas `surface3`,
    `dup_bg`, `accent_text`, `state_bg`/`state_fg`; nada se inventa).
  - `metatag_v8.py`: importa los tokens desde `metatag_theme`; conserva
    `CURRENT_THEME`/`C`/`FONTS`/`set_font_scale` como API pública. Se demostró
    paridad **byte a byte** contra el archivo original de git HEAD (THEMES y
    `set_font_scale` en escalas 0.82–1.35) y con tests de paridad.
  - Renombrador: elimina `PALETTES`/`_current_palette`/modo claro/bypass de alto
    contraste; dropdown con los **3 temas de MetaTag**; `_apply_theme` reconstruye
    la vista completa (`MainView.rebuild_theme`) preservando estado del usuario
    (rutas, opciones, badges, botones, filtro, preview) vía snapshot/restore;
    fuentes con **bases propias del Renombrador (9–18) escaladas con la misma
    lógica de MetaTag**; colores de estado derivados de `ok`/`err`/`warn`; títulos
    "MetaTag v8.9 — Renombrador de Fotos" y subtítulo "desde Excel · integrado en
    MetaTag v8.9". Tema antiguo del estado (`"dark"`) se normaliza al default.
  - Tests: `tests/test_metatag_theme.py` (53: valores canónicos exactos, paridad de
    fuentes contra `metatag_v8`, adaptadores deterministas/completos, migración sin
    diferencias) y `tests/test_renombrador_pytest.py` actualizado (+3, ahora 49,
    con `TestThemeChange`: cambio de tema reconstruye y preserva selecciones).

  **Verificación:** 199 tests pytest + 277 subtests verdes; smoke GUI real (con
  display) de arranque con estado antiguo, rebuild con datos/preview/filtro y
  preservación de selecciones al cambiar entre los 3 temas + fallback; paridad
  exacta de THEMES y motor de fuentes contra el original de git HEAD.

  ---

  ## 17. Estado actual del proyecto

  ### Estado general
  MetaTag v8.9.

  ### Trabajo actual
  **FASE 3B.1 (2026-08-13) — Neutralización de XIM↔iBus por-proceso** ✅ terminado: auditoría de rendimiento (FASE 3A-R.1/2/3, informes en `docs/`) identificó como causa raíz del lag de arranque/cambio de tema la asociación XIM de Tk con iBus (`XMODIFIERS=@im=ibus`, ~20-90 ms por widget). Nuevo `src/metatag_xim.py` con `neutralize_xim_for_tk()` que, solo cuando el IM es iBus y solo dentro del proceso, cambia `XMODIFIERS` a `@im=none` ANTES del primer `Tk()`; se invoca en los entry points de `metatag_v8.py` y `renombrar_fotos_gui.py` (el Renombrador se beneficia al lanzarse desde MetaTag por herencia de entorno y también en ejecución independiente). Sin cambios globales del sistema, sin tocar iBus, layout ni configuración. Medición en el equipo real: startup **20.78 s → 0.27 s** (~77×), cambio de tema **20.30 s → 0.123 s** (~165×). Verificación: **207 pytest + 277 subtests** verdes; smoke GUI real (ventana mapeada en <4 s); pendiente verificación manual de tipeo LATAM (teclas muertas). Commits: `perf: ...` (3B.1).

  **FASE 8 (2026-08-11) — Fuente de verdad técnica de temas + Renombrador sobre temas de MetaTag** ✅ terminado: nuevo `src/metatag_theme.py` (módulo puro) con los 3 temas canónicos verbatim, motor de fuentes idéntico al de `metatag_v8` (paridad byte a byte contra git HEAD) y adaptadores Tk/CustomTkinter sin inventar colores. `metatag_v8.py` importa los tokens desde ahí (cero cambio visual). El Renombrador abandonó `PALETTES`/modo claro/bypass y usa los 3 temas de MetaTag con reconstrucción completa de la UI preservando estado; fuentes con bases propias (9–18) escaladas con la lógica de MetaTag; branding "MetaTag v8.9 — Renombrador de Fotos". Verificación: **199 pytest + 277 subtests** verdes; smoke GUI real (arranque con estado antiguo `"dark"`, rebuild con preview/filtro, preservación de selecciones entre los 3 temas + fallback).

  ### Trabajo próximo
  Completar la integración de configuración del renombrador dentro de `metatag_v8.py`: diálogo de configuración embebido (hoy se lanza la ventana completa de la herramienta), generación de correspondencias por matching desde la UI, simulación/dry-run, registro de operación y refinar conflictos a nivel de interfaz — sección 11.

  ### Trabajo posterior
  Mejorar/separar el módulo de estadísticas.

  ### Idea futura
  Sincronización externa Excel ↔ MetaTag (no implementada). **FASE B — desambiguación por contenido visual/pHash REAL (PENDIENTE/PROPUESTA FUTURA, NO implementada):** usar hash perceptivo de píxeles como último recurso SOLO ante claves textuales ambiguas; requiere análisis de impacto sobre la línea base (267/0/2) antes de activarse (ver entrada de FASE A en la sección 16).

  ### Bloqueadores
  No se han confirmado bloqueadores críticos en esta sesión.

  ### Nota de estado (2026-08-10)
  Matching seguro implementado y verificado (Bloque 1), `ExcelGrid.redraw` optimizado (Bloque 2, `col_sel_map` precalculado), procesamiento en segundo plano robustecido con cancelación (Bloque 3, `_process_all` con snapshot + `_process_queue` con detección de muerte inesperada + `_proc_finish_ui`), responsividad de la interfaz corregida (Bloque 4, export de gráficas desacoplado + `_clamp_toplevel` + Visor), scroll de las ventanas de selección de columnas fijado sin `bind_all` (Fase 4), **FASE A (seguridad del renombrador + matching seguro como port puro)**: `src/metatag_matching.py` reescrito como port fiel/determinista del `_find_image_ex` validado (sin pHash muerto ni umbral difuso), fallback posicional silencioso eliminado (estado `error`), estados unificados a `ambiguo`, verificación real del dataset 269 (267/0/2, 18/184/21/44), y **FASE 7 (migración e integración del renombrador, 2026-08-10)**: el renombrador ahora vive en `src/renombrar_fotos_gui.py` (única copia; `tools/renombrador/` eliminado), se lanza desde MetaTag como subproceso independiente con el botón "🖼 Renombrador de Fotos" (sin withdraw, sin callbacks huérfanos), suite pytest en `tests/test_renombrador_pytest.py` (46/46) y unittest en `tests/test_renombrador.py` (11); 97/97 proyecto + 46/46 pytest verdes y smoke real con display OK.

  ---

  ## 18. Próximo paso recomendado

  1. ✅ **Bloque 4 — responsividad de la UI** (completado, 2026-08-10): auditoría de toda la interfaz frente a resoluciones/tamaños de ventana/DPI; visualización y exportación de gráficas desacopladas; ventanas secundarias centradas/limitadas; Visor arreglado. Tests `tests/test_responsive.py` (12), **52/52 verdes**. Commit `fix: mejorar responsividad de la interfaz`.
  2. ✅ **Fases 4–6 — scroll de selección de columnas + tests del renombrador** (completado, 2026-08-10): helper `_build_scroll_picker_window` sin `bind_all`; tests de renombrador (11) y pickers (19, incl. matriz de resoluciones y limpieza de trazas). **82/82 verdes**. Commit `feat: fijar scroll de selección de columnas y añadir tests del renombrador`.
  3. ✅ **FASE A — seguridad del renombrador + matching seguro como port puro** (completado, 2026-08-10): `src/metatag_matching.py` port fiel de `_find_image_ex`; fallback posicional silencioso eliminado; estados unificados a `ambiguo`; verificación dataset 269 (267/0/2, 18/184/21/44). 97/97 + 46/46 verdes. Commit `fix: harden renamer safety and safe matching`.
  4. ✅ **FASE 7 — migración e integración del renombrador en MetaTag** (completado, 2026-08-10): `src/renombrar_fotos_gui.py` como ubicación única (`tools/renombrador/` eliminado), botón lanzador en HERRAMIENTAS AVANZADAS (subproceso independiente, sin withdraw ni callbacks), suite pytest en `tests/test_renombrador_pytest.py`, `requirements-renombrador.txt`. Commit `feat: integrar renombrador de fotografías en MetaTag`.
  5. **Siguiente: completar la integración de configuración del renombrador en `metatag_v8.py`** (diálogo de configuración embebido, generación de correspondencias por matching desde la UI de MetaTag, simulación/dry-run, registro de operación, refinar conflictos a nivel de interfaz) — sección 11.
  6. Comprobar que el **emparejamiento** sea seguro (evitar falsas coincidencias). ✅ Matching seguro implementado y probado (Bloque 1 + FASE A, 2026-08-10): `_find_image_ex` con detección de ambigüedades; `src/metatag_matching.py` es su port puro y fiel; tests `tests/test_matching.py`, `tests/test_dataset_269.py` y `tests/test_metatag_matching.py`.
  7. Revisar **conflictos y casos límite** (duplicados, nombres vacíos, extensiones dobles, marcadores `(1)`).
  8. Probar con el **dataset de 269 imágenes**. ✅ Correspondencias idénticas al original (267 ok, 2 missing, 2 huérfanas, 0 reusos).
  9. ✅ Rendimiento de `ExcelGrid.redraw` optimizado (Bloque 2, 2026-08-10): `col_sel_map` precalculado una vez por columna visible; tests `tests/test_grid.py` (12), 23/23 verdes.
  10. ✅ Procesamiento en segundo plano robustecido (Bloque 3, 2026-08-10): cancelación, detección de muerte inesperada, snapshot de Tk; tests `tests/test_queue.py` (17), 40/40 verdes.
  11. Después, continuar con las mejoras del **módulo de estadísticas**.

  ---

  ## Instrucciones para futuras sesiones

  Cualquier nueva sesión de OpenCode debe:

  1. Leer `CONTEXTO_PROYECTO.md` antes de modificar código.
  2. Comparar el contexto con el código actual.
  3. Detectar contradicciones.
  4. Actualizar el contexto si está desactualizado.
  5. Registrar cambios importantes.
  6. Registrar decisiones.
  7. Registrar bugs.
  8. Registrar funcionalidades implementadas.
  9. Registrar funcionalidades pendientes.
  10. Actualizar el estado actual.
  11. Actualizar el próximo paso.
  12. No eliminar historial anterior.
  13. No inventar información.
  14. Diferenciar hechos, inferencias, hipótesis y propuestas.
  15. Utilizar el código actual como fuente de verdad.
