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
  │   ├── metatag_v8.py             # Aplicación principal (~3068 líneas)
  │   ├── metatag_writer.py         # Escritura EXIF/IPTC pura (~145 líneas)
  │   ├── metatag_graficas.py       # Estadísticas y gráficas (~694 líneas)
  │   ├── metatag_widgets.py        # ExcelGrid con viewport culling (~418 líneas)
  │   ├── Visor.py                  # Visor de metadatos / comparador / PDF (~2468 líneas)
  │   └── editor_casillas_backup.py # Respaldo del editor de casillas (~116 líneas)
  ├── data/
  │   ├── metatag_config.json       # Configuración persistente de la sesión
  │   └── metatag_debug.log         # Log de errores (logging, nivel ERROR)
  ├── CeramicIA_Dataset_Piloto_103Fragmentos_v1.1_2026-06-26.xlsx  # Dataset de prueba
  ├── Finales 1 a 103/              # 269 imágenes de prueba
  ├── instalar_y_abrir.sh / .bat    # Instalación de dependencias + lanzamiento
  └── .venv/                        # Entorno virtual (Python 3.12.3)
  ```

  ### Módulos y responsabilidades

  | Archivo | Responsabilidad | Componentes principales | Estado |
  |---|---|---|---|
  | `src/metatag_v8.py` | Aplicación principal: UI, carga de datos, emparejamiento, escritura, Image Sync | `ImageBrowser`, integración `ExcelGrid`, inyección manual/lote, `_sync_excel_to_images`, `_sync_images_to_excel`, `_find_image`, temas, lupa, búsqueda | Implementado y activo |
  | `src/metatag_writer.py` | Escritura de metadatos **pura y testeable** (sin Tkinter) | `META_GROUPS`, `META_GROUP_ORDER`, `formatear_metadatos`, `read_existing_metadata`, `check_metadata_divergence`, `write_jpeg`, `write_png`, `write_tiff`, `write_meta` | Implementado |
  | `src/metatag_graficas.py` | Estadísticas, selector de gráficas, generación y exportación | `show_stats`, `make_selector`, `update_chart`, `export_chart`, `on_hover` | Implementado; mejoras futuras planeadas |
  | `src/metatag_widgets.py` | Grid Excel sobre Canvas con viewport culling | Clase `ExcelGrid`: `load`, `redraw`, selección de filas/columnas, `get_row_metadata`, `get_selected_metadata` | Implementado |
  | `src/Visor.py` | Visor independiente de metadatos EXIF/JSON/GPS, comparador, zoom y export PDF | `VisorApp`, `_extract_exif`, `_extract_gps`, `_extract_json`, `_open_image_comparison`, `_export_pdf`, `_generate_pdf_document` | Implementado |
  | `src/editor_casillas_backup.py` | Respaldo de funciones del editor de casillas | `_open_editor`, `_populate_editor`, `toggle_lock`, `update_df` | Backup; no es el editor activo |

  **Relaciones:** `metatag_v8.py` importa `show_stats` desde `metatag_graficas`, `ExcelGrid` desde `metatag_widgets`, y las funciones de escritura/divergencia desde `metatag_writer`. `metatag_v8.py` puede lanzar `Visor.py` (`launch_visor`).

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
  - **Temas:** diccionario `THEMES` con **3 temas** confirmados: `Arqueológico (Oscuro Refinado)` (por defecto), `Noche Total`, `Carbón` (~líneas 63–85). Motor de fuentes dinámico `set_font_scale` y `FONTS`.
  - **Estadísticas:** `_show_stats` (~514) invoca `show_stats` externo.
  - **Persistencia:** `_save_config`, `_load_config_pre_build`, `_load_config_post_build` leen/escriben `data/metatag_config.json`.
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

  ## 11. Nueva herramienta: Image Sync / Renombrador

  Estado actual: **EN DESARROLLO**. No es funcionalidad cerrada.

  **Concepto funcional:**

  ```
  Excel        → nombres objetivo
  Fotografías  → nombres actuales
  ```

  La herramienta debe establecer correspondencias entre fotografías y registros del Excel y, posteriormente, **renombrar** las imágenes de acuerdo con una columna del Excel.

  **Interfaz:** utiliza actualmente el nombre **"Image Sync"** con el subtítulo **"Renombrador de fotos desde Excel"**.

  **Ideas funcionales discutidas** (distinguir claramente estado):

  - IMPLEMENTADO (base reutilizable): matching inteligente (`_find_image`, `_safe_stem`, `_extract_id_suffix`), detección de coincidencias aproximadas, huérfanas y valores sin imagen (usado hoy por Image Sync de orden).
  - PLANEADO: selección de carpeta de fotografías, carga/uso del Excel, selección de hoja, selección de columna, ordenamiento de fotografías, generación de correspondencias, vista previa, detección de conflictos, detección de duplicados, detección de nombres vacíos, simulación, renombramiento real, registro de operación, posible deshacer futuro, posible modo copia/backup futuro.

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

  ## 17. Estado actual del proyecto

  ### Estado general
  MetaTag v8.9.

  ### Trabajo actual
  Desarrollo/adaptación de **Image Sync** como herramienta de correspondencia y **renombramiento de fotografías** (nombres de fotos ↔ columna del Excel).

  ### Trabajo próximo
  Validar y mejorar el renombrador (emparejamiento seguro, conflictos, duplicados, nombres vacíos, simulación).

  ### Trabajo posterior
  Mejorar/separar el módulo de estadísticas.

  ### Idea futura
  Sincronización externa Excel ↔ MetaTag (no implementada).

  ### Bloqueadores
  No se han confirmado bloqueadores críticos en esta sesión.

  ### Nota de estado (2026-08-10)
  Matching seguro implementado y verificado (Bloque 1) y `ExcelGrid.redraw` optimizado (Bloque 2, `col_sel_map` precalculado). El proceso de escritura por lote (`_process_all`) todavía lee `self.omit_empty_var`/`self.meta_mode_organized` desde el hilo worker; el renombrado por Grid sigue pendiente (Bloque 3).

  ---

  ## 18. Próximo paso recomendado

  1. Terminar de validar **Image Sync / renombrador**.
  2. Comprobar que el **emparejamiento** sea seguro (evitar falsas coincidencias). ✅ Matching seguro implementado y probado (Bloque 1, 2026-08-10): `_find_image_ex` con detección de ambigüedades; tests `tests/test_matching.py` y `tests/test_dataset_269.py`.
  3. Revisar **conflictos y casos límite** (duplicados, nombres vacíos, extensiones dobles, marcadores `(1)`).
  4. Probar con el **dataset de 269 imágenes**. ✅ Correspondencias idénticas al original (267 ok, 2 missing, 2 huérfanas, 0 reusos).
  5. ✅ Rendimiento de `ExcelGrid.redraw` optimizado (Bloque 2, 2026-08-10): `col_sel_map` precalculado una vez por columna visible; tests `tests/test_grid.py` (12), 23/23 verdes.
  6. Después, continuar con las mejoras del **módulo de estadísticas**.

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
