# MetaTag v8.9 — Guia Completa de Uso y Desarrollo

---

## PARTE 1: USO PRACTICO

### Que es MetaTag

MetaTag es una herramienta para arqueologos que escribe metadatos descriptivos (sitio, corte, unidad, vista, tecnica, etc.) en fotografias de fragmentos ceramicos. Trabaja sobre **copias** de las imagenes, nunca modifica los originales.

### Programas incluidos

| Programa | Archivo | Que hace |
|---|---|---|
| **MetaTag** (principal) | `src/metatag_v8.py` | Carga un Excel/CSV + carpeta de fotos, escribe metadatos EXIF/IPTC en copias |
| **Image Sync** (renombrador) | `src/renombrar_fotos_gui.py` | Renombra fotos para que coincidan con los nombres del Excel |
| **Visor** | `src/Visor.py` | Visor de imagenes con zoom, comparacion y exportacion a PDF |

### Como iniciar

```bash
# Linux (instalador automatico):
./instalar_y_abrir.sh

# Windows:
instalar_y_abrir.bat

# Manual:
.venv/bin/python src/metatag_v8.py
```

Image Sync y Visor se lanzan desde el programa principal o de forma independiente:
```bash
.venv/bin/python src/renombrar_fotos_gui.py
.venv/bin/python src/Visor.py [ruta_imagen_opcional]
```

---

## MetaTag Principal

### Flujo de trabajo

1. **Cargar datos**: Abri un archivo Excel/CSV con los registros arqueologicos
2. **Abrir imagenes**: Selecciona la carpeta que contiene las fotografias
3. **Elegir columna**: Selecciona que columna del Excel contiene los nombres de archivo
4. **Seleccionar celdas/filas**: Hace clic en las filas o celdas que quieras procesar
5. **Escribir metadatos**: Usa el modo "Inteligente" (lote) o "Libre" (individual)

### Panel izquierdo — Herramientas

#### Archivo de datos
- Campo de ruta + boton `...` para seleccionar el Excel/CSV
- Boton **"Cargar archivo"** (`Ctrl+O`)

#### Carpetas de imagenes
- Campo de ruta + boton `...` para seleccionar la carpeta
- Boton **"Abrir carpeta"** (`Ctrl+B`)

#### Seleccion de datos
- **Selector de columna de imagenes**: Que columna del Excel contiene los nombres de archivo (se auto-detecta: `id`, `imagen`, `image`, `file`, `archivo`, `nombre`, `name`, `foto`, `photo`)
- **Toggle "Omitir celdas vacias"**: Excluye filas sin texto en la columna seleccionada
- **"Seleccionar fila activa"**: Selecciona la fila que coincide con la imagen actual
- **"Limpiar seleccion"** (`Ctrl+L`)

#### Herramientas avanzadas

| Herramienta | Atajo | Que hace |
|---|---|---|
| **Ver Estadisticas** | `Ctrl+G` | Muestra graficos de distribucion de metadatos |
| **Lote por Orden (Excel->Fotos)** | `Ctrl+Shift+B` | Mapea filas del Excel a fotos por posicion (Fila 1 = Foto 1) |
| **Image Sync** | — | Abre el renombrador como subproceso independiente |

#### Sincronizacion de orden
- **"Reordenar imagenes segun Excel"** (`Ctrl+R`): Reordena la lista del navegador de imagenes para que coincida con el orden de las filas del Excel. No toca los archivos, solo la vista.

#### Verificacion de integridad
- **"Verificar imagenes originales"**: Revisa si las fotos originales ya tienen metadatos de MetaTag. Si las encuentra, ofrece limpiarlas antes de escribir nuevos.

#### Modo de trabajo
- **"Inteligente"** (`Ctrl+E`): Modo lote — busca la foto que corresponde a cada fila seleccionada usando matching por nombre, la copia a la carpeta de salida y escribe los metadatos
- **"Libre"** (`Ctrl+I`): Modo individual — escribe los metadatos de la fila activa en la imagen que esta viendo actualmente

#### Carpeta de salida
- Muestra la ruta relativa `/Metadatos_Escritos`
- Boton **"Abrir carpeta de salida"**

#### Atajos de teclado
- Boton **"Ver atajos disponibles"** muestra la lista completa

### Panel central — Tabla de datos

- **ExcelGrid**: Tabla virtualizada que muestra el contenido del Excel/CSV
- **Barra de busqueda**: Filtra filas por texto en cualquier columna o en una columna especifica
- **Contador de seleccion**: Muestra cuantas celdas/filas estan seleccionadas
- **Registro de actividad**: Log de operaciones con colores (ok en verde, error en rojo, info en gris)

### Panel derecho — Explorador y vista previa

- **Explorador de imagenes**: Lista de archivos de imagen de la carpeta seleccionada. Archivos sin fila asociada se muestran en gris con `[sin fila]`
- **Vista previa**: Canvas con zoom (rueda del mouse) y arrastre (clic y arrastra). Doble clic para resetear zoom
- **"Visor Pro"**: Abre el Visor completo (`Visor.py`)
- **Metadatos a escribir**: Vista previa del texto que se va a escribir en la imagen
- **Formato "Organizado"**: Checkbox para alternar entre formato con secciones `[Ubicacion]`, `[Descripcion]`, etc. o formato plano

### Campos de metadatos

Los 16 campos se organizan en 4 grupos:

**Ubicacion:**
`Sitio`, `Corte`, `Cuadrante`, `Unidad`, `Nivel`, `Profundidad Cm`

**Descripcion:**
`Vista`, `Parte`, `Perfil`, `Labio`

**Tecnica:**
`Tratamiento`, `Tecnica`, `Motivo`

**Notas:**
`Observaciones`, `Excluido`

Columnas adicionales que no estan en estos grupos se escriben bajo el grupo `[Otros]`.

### Temas visuales

| Tema | Icono | Color de acento |
|---|---|---|
| Arqueologico (Oscuro Refinado) | `jar` | Marron warm `#A67C52` |
| Noche Total | `moon` | Morado `#BB86FC` |
| Carbon | `black` | Azul `#569CD6` |

### Atajos de teclado (resumen)

| Atajo | Accion |
|---|---|
| `Ctrl+O` | Abrir Excel/CSV |
| `Ctrl+B` | Abrir carpeta de imagenes |
| `Ctrl+E` | Escribir metadatos (modo Inteligente) |
| `Ctrl+I` | Inyectar en foto actual (modo Libre) |
| `Ctrl+G` | Ver estadisticas |
| `Ctrl+R` | Reordenar imagenes segun Excel |
| `Ctrl+Shift+B` | Lote por orden (Excel->Fotos) |
| `Ctrl+F` | Enfocar busqueda |
| `Ctrl+L` | Limpiar seleccion |
| `Ctrl+Q` | Salir (guarda config) |
| `Escape` | Cerrar lupa / ventana / limpiar seleccion |
| `Up`/`Down` | Navegar entre filas |

---

## Image Sync (Renombrador de Fotos)

### Que hace

Sincroniza los nombres de las fotografias con los registros del Excel. Puede funcionar en dos modos:
- **Posicional**: Foto 1 = Registro 1, Foto 2 = Registro 2, etc.
- **Matching seguro**: Cada nombre del Excel busca SU propia foto usando un algoritmo de matching de 7 pasos

### Flujo de 5 pasos

Los pasos se muestran como indicadores en la barra superior:

| Paso | Icono | Que ocurre |
|---|---|---|
| 1. Emparejar fotografias | circled-1 | Se cargan las fotos de la carpeta seleccionada |
| 2. Validar correspondencias | circled-2 | Se calcula el plan de renombrado (en background) |
| 3. Vista previa | circled-3 | Se muestra la tabla con pares original->nuevo nombre |
| 4. Renombrar | circled-4 | Se ejecuta la operacion (con confirmacion) |
| 5. Resultado | circled-5 | Se muestra el resumen final |

### Secciones de la interfaz

#### 1. Carpeta de fotos
- Selector de carpeta con busqueda customizada
- **Ordenar por**: 9 opciones de ordenamiento (ver abajo)
- Boton **"Cargar fotos"** (`Ctrl+O`)

#### 2. Archivo Excel
- Selector de archivo (.xlsx, .csv, .tsv, .txt)
- **Selector de hoja** (si el Excel tiene multiples hojas)
- **Selector de columna** con los nombres (se auto-detecta)
- Boton **"Cargar Excel"** (`Ctrl+E`)

#### 3. Vista previa
- **Barra de busqueda**: Filtra filas por texto (debounced 100ms)
- **"Modo editar"**: Checkbox que permite editar inline los nuevos nombres. Los duplicados se recalculan automaticamente (debounced 150ms)
- **Tabla virtualizada**: Columnas `#`, `Estado`, `Original`, `Nuevo nombre`. Solo renderiza las filas visibles (~20-30 widgets para miles de filas)
- **Tooltip con miniatura**: Al pasar el cursor sobre una fila, aparece una miniatura (180x180) despues de 300ms

#### 4. Opciones
- **"Mantener extension original"** (por defecto ON): Agrega la extension original al nombre nuevo
- **"Crear registro/backup de la operacion"** (por defecto ON): Guarda un archivo `.metatag_backup_*.json` con el historial
- **"Abrir carpeta al finalizar"** (por defecto ON): Abre la carpeta en el explorador de archivos

#### 5. Renombrar
- **"Matching seguro"**: Activa el matching por nombre en vez del posicional
- **"Modo copiar"**: Copia las fotos a una subcarpeta `Renombradas/` en vez de renombrar
- **"Simular"**: Calcula el plan sin tocar archivos (muestra toast de confirmacion)
- **"Renombrar todo"** (`Ctrl+Enter`): Ejecuta el renombrado
- **"Cancelar"** (`Escape`): Detiene la operacion en curso
- **"Guardar log"**: Exporta el registro a un archivo .txt
- **"CSV"**: Exporta la vista previa a un archivo .csv

#### 6. Registro
- Log de la operacion con detalles por archivo
- Boton **"Limpiar"** para borrar el registro

### Estados del plan de renombrado

| Estado | Icono | Significado | Bloquea? |
|---|---|---|---|
| `ok` | checkmark | Se puede renombrar sin problemas | No |
| `ya_correcto` | checkmark | Ya tiene ese nombre (mismo archivo, mismo nombre) | No |
| `existe` | warning | El destino es un archivo EXTERNO que ya existe (nunca se sobreescribe) | Si |
| `conflicto` | warning | Colision interna no resoluble en el lote | Si |
| `duplicado` | warning | Dos filas compiten por el mismo nombre destino | Si |
| `not_found` | x | No se encontro foto para ese registro (solo matching) | Si (posicional) / No (matching) |
| `sin_foto` | x | Mas registros que fotos (solo posicional) | Si |
| `ambiguo` | warning | Multiples archivos compiten por la misma clave | Si |
| `error` | x | Fallo del motor de matching | Si |

### Opciones de ordenamiento

| Etiqueta | Clave | Criterio |
|---|---|---|
| Orden numerico | `natural` | Orden natural: foto2 antes de foto10 |
| Nombre (A -> Z) | `name_asc` | Alfabetico ascendente |
| Nombre (Z -> A) | `name_desc` | Alfabetico descendente |
| Fecha modificacion up | `mtime_asc` | Mas antigua primero |
| Fecha modificacion down | `mtime_desc` | Mas reciente primero |
| Fecha creacion up | `ctime_asc` | Mas antigua primero |
| Fecha creacion down | `ctime_desc` | Mas reciente primero |
| Fecha foto up | `exif_asc` | Fecha EXIF mas antigua primero |
| Fecha foto down | `exif_desc` | Fecha EXIF mas reciente primero |

### Atajos de teclado

| Atajo | Accion |
|---|---|
| `Ctrl+Z` | Deshacer ultimo lote de renombrado |
| `Ctrl+O` | Cargar fotos |
| `Ctrl+E` | Cargar Excel |
| `Escape` | Cancelar operacion en curso |
| `Ctrl+Enter` | Ejecutar renombrado |

### Sistema de deshacer

- Cada renombrado exitoso se guarda en una pila
- El boton "Deshacer" muestra cuantos lotes hay pendientes
- En modo renombrado: mueve los archivos de vuelta a su nombre original
- En modo copia: elimina las copias creadas
- No sobreescribe archivos existentes (detecta conflictos)

### Notificaciones (toast)

Aparecen en la esquina inferior derecha y desaparecen despues de 3.5 segundos. Tres tipos:
- **ok** (verde): Operacion exitosa
- **error** (rojo): Error o accion bloqueada
- **warn** (amarillo): Advertencia o informacion

---

## Visor de Metadatos

### Que hace

Visor de imagenes con zoom, paneo, comparacion lado a lado y exportacion a PDF con metadatos.

### Como iniciar

- Desde MetaTag: boton "Visor Pro" en el panel derecho
- Independiente: `python src/Visor.py [ruta_imagen]`

---

## PARTE 2: ARQUITECTURA

### Estructura del proyecto

```
MetaTag_v8.9/
  src/
    metatag_v8.py              # Aplicacion principal (3224 lineas)
    renombrar_fotos_gui.py     # Image Sync / Renombrador (4025 lineas)
    Visor.py                   # Visor de metadatos (2498 lineas)
    metatag_matching.py        # Motor de matching de imagenes (296 lineas)
    metatag_writer.py          # Funciones de escritura EXIF/IPTC (145 lineas)
    metatag_theme.py           # Tokens de tema y motor de fuentes (220 lineas)
    metatag_widgets.py         # ExcelGrid virtualizado (433 lineas)
    metatag_responsive.py      # Deteccion de perfil de pantalla (100 lineas)
    metatag_graficas.py        # Estadisticas y graficos (724 lineas)
    metatag_xim.py             # Neutralizacion XIM/iBus (37 lineas)
  tests/                       # 20 archivos de test (~500+ tests)
  docs/                        # Informes tecnicos
  data/                        # Configuracion persistente y logs
  Metadatos_Escritos/          # Carpeta de salida (donde se escriben las copias)
```

### Patron de diseno: MVC en Image Sync

Image Sync sigue un patron **Model-View-Controller** estricto:

```
RenameModel (logica pura, sin UI)
    |
    v
AppController (orquesta Model <-> View)
    |
    v
MainView (interfaz CustomTkinter)
```

- **RenameModel**: Toda la logica de negocio. No conoce Tkinter. Maneja: cargar fotos, cargar Excel, construir planes, ejecutar renombrados, deshacer, exportar logs.
- **AppController**: Punto medio. Recibe eventos de la vista, llama al modelo, actualiza la vista. Maneja: guards de precondicion, operaciones asincronas en threads, buffering de logs, flush de UI.
- **MainView**: Solo construye widgets y delega eventos al controller. Nunca toca archivos ni datos directamente.

### Flujo de datos en MetaTag Principal

```
Excel/CSV --> pandas DataFrame --> ExcelGrid (tabla virtualizada)
                                        |
Photos --> ImageBrowser (lista) --> _find_image (matching) --> write_meta (EXIF/IPTC)
                                        |                            |
                                        v                            v
                                 Metadatos_Escritos/        Copias con metadatos
```

### Flujo de datos en Image Sync

```
Carpeta de fotos --> RenameModel.load_photos() --> self._photos (list[Path])
Excel/CSV       --> RenameModel.load_names()   --> self._names (list[str])
                              |
                              v
                    RenameModel.build_preview()
                    (positional o matching mode)
                              |
                              v
                    list[(original, nuevo, path, is_dup, state)]
                              |
                              v
                    PreviewTable.render() --> Slots virtualizados en Canvas
                              |
                              v
                    AppController.on_rename() --> RenameModel.rename_all()
                              |
                              v
                    Dos fases: temp files --> rename final
```

### Algoritmo de matching de 7 pasos

`ImageMatcher.find_image_ex(name, folder)` busca una foto para cada nombre usando una jerarquia de 7 pasos. Si un paso encuentra exactamente 1 resultado, lo devuelve. Si encuentra >1, retorna "ambiguo". Si no encuentra, pasa al siguiente paso.

| Paso | Nombre | Que busca |
|---|---|---|
| 1 | `direct` | Si `folder/name` existe como ruta literal |
| 2 | `nombre-exacto` | Nombre de archivo exacto (case-insensitive) |
| 3 | `stem-exacto` | Stem exacto (sin extension ni marcadores `(N)`) |
| 4 | `clean` | Stem limpio (sin separadores en bordes: `#`, `_`, `-`, espacios) |
| 5 | `normalize` | Numeros normalizados (sin ceros a la izquierda) |
| 6 | `id-suffix` | Sufijo de ID (numero_pieza + sufijo_vista: `F`/`R`/`P`) |
| 7 | `substring` | Substring bidireccional |

### Formato de escritura de metadatos

**JPEG** (4 campos EXIF):
- `ImageDescription` (IFD 0th): Texto organizado
- `UserComment` (IFD Exif): JSON completo del diccionario
- `XPComment` (tag 40092): Texto UTF-16-LE
- `XPKeywords` (tag 40094): Valores separados por `;`, UTF-16-LE

**PNG** (chunks de texto):
- `Description`: Texto organizado
- `Comment`: JSON completo

**TIFF**:
- `ImageDescription` (IFD 0th): Texto organizado

### Sistema de temas

Los temas se definen en `metatag_theme.py` con un diccionario completo de colores:
- Colores de fondo, superficie, texto, borde, acento
- Colores de estado (ok, error, warn)
- Colores de seleccion, fila par/impar
- Colores para graficos
- Fuentes escaladas segun ancho de pantalla (82%-135% de escala base)

### Optimizaciones de rendimiento

- **XIM/iBus neutralization** (`metatag_xim.py`): Desactiva la asociacion XIM/iBus para procesos Tk, eliminando 20-90ms de lag por widget
- **ExcelGrid virtualizado** (`metatag_widgets.py`): Canvas con viewport culling — solo renderiza filas visibles
- **PreviewTable virtualizada** (`renombrar_fotos_gui.py`): Pool reutilizable de ~20-30 widgets para miles de filas
- **Throttle de scroll** (`_throttled_sync`): Leading+trailing a ~60fps para no saturar el pipeline de CTk
- **Cuantizacion de scroll**: `self._px` siempre es multiplo de `ROW_H` (igual que `yscrollincrement` de Tk) para evitar drift
- **Carga por chunks** (`FileBrowser`): 70 items por lote con 20ms de delay para evitar congelamiento
- **Buffer de logs**: Lineas de background thread se acumulan y flushan en lotes de 20
- **Debouncing**: Busqueda (100ms), recalcular duplicados (150ms), plan en background thread

---

## PARTE 3: API DE renombrar_fotos_gui.py

### Funciones del modulo

| Funcion | Descripcion |
|---|---|
| `_refresh_button_constants()` | Refresca los diccionarios de colores de botones desde el tema actual |
| `_init_fonts(screen_width)` | Inicializa 9 fuentes CTkFont escaladas segun ancho de pantalla |
| `_load_state()` / `_save_state(patch)` | Carga/guarda estado persistente de la herramienta |
| `_get_thumb(path, size)` | Retorna thumbnail cacheada (LRU, max 150) para una imagen |
| `_natural_key(p)` | Clave de orden natural: "foto2" antes de "foto10" |
| `_get_exif_date(p)` | Extrae fecha EXIF de una foto (0 si no tiene) |
| `_detect_drives()` | Detecta discos/mount points disponibles segun OS |
| `_invalid_name_chars(name)` | Retorna lista de caracteres invalidos en un nombre de archivo |
| `_place_tip_near_pointer(tip, widget, offset)` | Posiciona un Toplevel flotante cerca del cursor |

### Clase `RenameModel` (linea 347)

Logica de negocio, completamente desacoplada de la interfaz.

#### Propiedades

| Propiedad | Tipo | Descripcion |
|---|---|---|
| `matching_available` | `bool` | True si el motor ImageMatcher esta disponible |
| `photos` | `list[Path]` | Lista de fotos cargadas |
| `names` | `list[str]` | Lista de nombres del Excel |
| `has_undo` | `bool` | True si hay operaciones para deshacer |

#### Metodos principales

| Metodo | Parametros | Descripcion |
|---|---|---|
| `load_photos()` | — | Carga fotos de `self.folder_path` (9 formatos). Retorna count |
| `load_sheets()` | — | Retorna nombres de hojas del Excel |
| `load_columns()` | — | Retorna nombres de columnas de la hoja activa |
| `load_names()` | — | Carga nombres de la columna seleccionada. Retorna count |
| `clear_excel_data()` | — | Limpia todos los datos del Excel cargados |
| `set_name(index, name)` | `index: int, name: str` | Edita un nombre en la lista (edicion inline) |
| `build_preview()` | — | Construye pares de renombrado: `[(original, nuevo, path, is_dup, state)]` |
| `rename_blocked(plan)` | `plan: list` | Retorna `(blocked, reason)` — fuente unica de verdad para bloquear renombrado |
| `rename_all(...)` | Ver abajo | Ejecuta la operacion de renombrado |
| `undo_last(...)` | Ver abajo | Deshace el ultimo lote |
| `export_log(pairs, dest)` | `pairs: list, dest: Path` | Exporta registro a archivo .txt |
| `export_preview_csv(pairs, dest)` | `pairs: list, dest: Path` | Exporta vista previa a .csv |
| `write_backup(dest, batch, ...)` | Varios | Escribe archivo `.metatag_backup_*.json` |

#### `rename_all` parametros

```python
rename_all(
    copy_mode=False,        # True=copiar, False=renombrar
    keep_extension=True,    # True=mantener extension original
    on_progress=None,       # callback(progress_float)
    on_log=None,            # callback(linea_texto)
)
```

#### `undo_last` parametros

```python
undo_last(
    on_progress=None,       # callback(progress_float)
    on_log=None,            # callback(linea_texto)
)
```

### Clase `SmoothScroller` (linea 1165)

Scroll con inercia para CTkScrollableFrame. Solo activo cuando el cursor esta dentro del frame.

| Metodo | Descripcion |
|---|---|
| `__init__(sf)` | Vincula `<Enter>`/`<Leave>` para armar/desarmar |
| `_arm()` / `_disarm()` | Vincula/desvincula eventos de rueda a nivel app |
| `_scroll(direction)` | Calcula posicion objetivo del scroll |
| `_animate()` | Loop de animacion a ~60fps con easing lineal |

### Clase `ImageTooltip` (linea 1235)

Tooltip flotante con miniatura al pasar el cursor sobre filas de la tabla.

| Metodo | Descripcion |
|---|---|
| `_close_active()` | Cierra cualquier tooltip activa (singleton) |
| `__init__(widget, path)` | Muestra despues de 300ms; se oculta al salir del widget |

### Clase `ToolTip` (linea 1322)

Tooltip de texto simple para cualquier widget.

### Clase `StatusBadge` (linea 1358)

Pildora circular de estado (idle/ok/warn/error/loading) con icono.

| Metodo | Descripcion |
|---|---|
| `set_state(state)` | Actualiza la apariencia (bg, fg, icono) |

### Clase `Toast` (linea 1382)

Notificacion flotante auto-dismissing en la esquina inferior derecha. Se destruye despues de 3.5 segundos.

### Clase `FileBrowser` (linea 1424)

Explorador de archivos/carpetas customizado con tema oscuro. Reemplaza los dialogs nativos de tkinter.

| Metodo | Descripcion |
|---|---|
| `__init__(master, mode, filetypes, title)` | Abre browser en modo "folder" o "file" |
| `_navigate(path)` | Carga contenido del directorio |
| `_load_chunk()` | Carga archivos incrementalmente (70 por lote) |
| `_select(path)` | Selecciona un archivo |
| `_go_up()` | Navega al directorio padre |
| `_confirm()` | Cierra con el resultado seleccionado |
| `get_result()` | Retorna la ruta seleccionada despues de cerrar |

### Clase `PathSelector` (linea 1641)

Campo de entrada + boton de busqueda para seleccionar ruta de carpeta o archivo Excel.

| Metodo | Descripcion |
|---|---|
| `_browse()` | Abre FileBrowser y actualiza el campo |
| `_notify()` | Ejecuta el callback `on_change` |
| `get()` / `set(v)` | Obtiene/establece la ruta actual |

### Clase `PreviewTable` (linea 1693)

Tabla virtualizada usando `tk.Canvas` + pool de widgets reutilizables. Solo mantiene O(viewport) widgets activos sin importar el total de filas.

| Metodo | Descripcion |
|---|---|
| `__init__(master, on_name_change, on_filter, **kw)` | Configura canvas, scrollbar, pool de slots |
| `render(pairs, edit_mode)` | Renderiza los pares del plan. Recicla slots, sincroniza viewport |
| `filter(query)` | Filtra filas visibles por texto |
| `update_dup_states(pairs)` | Actualiza highlighting de duplicados sin re-render completo |
| `set_edit_mode(enabled)` | Activa/desactiva edicion inline del nombre |
| `_sync_viewport()` | Core: posiciona y rellena slots del rango visible |
| `_rebind_slot(slot, logical)` | Actualiza el contenido de un slot para mostrar la fila `logical` |
| `_hide_slot(slot)` | Oculta un slot fuera de pantalla |
| `_build_pool()` | Crea el pool de widgets O(viewport) |
| `_make_slot(k)` | Crea un slot con labels, entry, badge de estado |
| `_throttled_sync()` | Sincronizacion con throttle leading+trailing a ~60fps |
| `_sync_pool_width()` | Aplica el ancho actual del canvas a todos los slots |
| `_on_wheel(event)` | Handler de scroll (retorna "break" para detener scroll externo) |
| `_on_motion(event)` | Tracking de hover para posicion del tooltip |

### Clase `ConfirmDialog` (linea 2447)

Dialogo de confirmacion con tema oscuro que reemplaza `messagebox.askyesno`.

| Metodo | Descripcion |
|---|---|
| `ask(master, title, message, ok_text)` | Muestra dialogo modal, retorna True/False |

### Clase `MainView` (linea 2504)

Ventana principal. Extiende `CTk`. Delega toda la logica a `AppController`.

#### Metodos publicos (seleccionados)

| Metodo | Descripcion |
|---|---|
| `update_summary(fotos, registros, corr, conflictos, estado)` | Actualiza panel de 5 celdas de resumen |
| `rebuild_theme()` | Reconstruccion completa de la UI preservando estado |
| `render_preview(pairs, edit_mode)` | Delega a PreviewTable |
| `filter_preview(query)` | Delega filtro a PreviewTable |
| `set_progress(value, msg)` | Actualiza barra de progreso |
| `toast(msg, kind)` | Muestra notificacion Toast |
| `confirm(title, msg, ok_text)` | Muestra ConfirmDialog |
| `append_log_line(line)` | Agrega linea al registro |
| `update_steps(done)` | Actualiza indicador de pasos (0-5) |

### Clase `AppController` (linea 3238)

Orquesta Model <-> View. Toda la logica de interaccion vive aqui.

#### Metodos de carga

| Metodo | Descripcion |
|---|---|
| `on_load_photos()` | Carga fotos en thread background |
| `on_load_excel()` | Carga Excel, muestra hojas si hay multiples |
| `on_sheet_selected(sheet)` | Cambia hoja del Excel y recarga |
| `on_column_selected(column)` | Cambia columna y recarga nombres/preview |

#### Metodos de interaccion

| Metodo | Descripcion |
|---|---|
| `on_filter_change(query)` | Aplica filtro al preview |
| `on_sort_change(value)` | Cambia orden y recarga |
| `on_folder_path_changed(raw)` | Valida y guarda ruta de carpeta |
| `on_excel_path_changed(raw)` | Valida y guarda ruta de Excel |
| `on_simulate()` | Calcula plan sin tocar archivos |
| `on_rename()` | Ejecuta renombrado (con guards) |
| `on_undo()` | Deshace ultimo lote |
| `on_cancel()` | Cancela renombrado en curso |
| `on_export_log()` | Exporta registro a .txt |
| `on_export_csv()` | Exporta preview a .csv |
| `on_edit_mode_change(enabled)` | Activa/desactiva edicion inline |
| `on_keep_ext_change(enabled)` | Activa/desactivar mantener extension |
| `on_backup_change(enabled)` | Activa/desactivar backup |
| `on_open_folder_change(enabled)` | Activa/desactivar abrir carpeta al finalizar |
| `on_matching_toggle()` | Activa/desactiva matching seguro |
| `on_theme_change(theme)` | Cambia tema y reconstruye toda la UI |

#### Metodos internos clave

| Metodo | Descripcion |
|---|---|
| `_guard(condition, msg, kind)` | Verificacion de precondicion (ej: "selecciona una carpeta primero") |
| `_update_sync_state(notify)` | Recalcula plan y actualiza UI |
| `_run_async(model_fn, progress_prefix, finish_fn)` | Ejecuta funcion del modelo en background con progreso |
| `_do_rename()` | Orquestacion interna del renombrado con actualizaciones de pasos |
| `_write_backup()` | Escribe el backup JSON de la operacion |
| `_unstick_ui()` | Recuperacion de emergencia de la UI |

---

## PARTE 4: API DE OTROS MODULOS CLAVE

### `metatag_matching.py` — Motor de matching

#### Clase `ImageMatcher`

| Metodo | Parametros | Retorna | Descripcion |
|---|---|---|---|
| `find_image(name, folder)` | `name: str, folder: str` | `Optional[Path]` | Busca foto simple, retorna Path o None |
| `find_image_ex(name, folder)` | `name: str, folder: str` | `(Path, status, candidates)` | Busqueda con 7 pasos, retorna estado y candidatos |
| `find_image_ex_with_method(...)` | Igual + `method` | Igual + metodo usado | Igual pero retorna que paso del algoritmo encontro la coincidencia |

#### Funciones helper

| Funcion | Descripcion |
|---|---|
| `_safe_stem(s)` | Quita extensiones de imagen y marcadores `(N)` |
| `_full_stem(s)` | Alias de `_safe_stem` |
| `_normalize_numbers(s)` | Quita ceros a la izquierda de todos los grupos de digitos |
| `_extract_id_suffix(s)` | Extrae tupla `(numero_pieza, sufijo_vista)` |
| `_clean_stem(stem)` | Quita separadores de bordes (`#`, `_`, `-`, espacios) |

### `metatag_writer.py` — Escritura de metadatos

| Funcion | Parametros | Descripcion |
|---|---|---|
| `formatear_metadatos(data, organized)` | `data: dict, organized: bool` | Formatea metadatos como texto plano o con secciones |
| `read_existing_metadata(path)` | `path: Path` | Lee metadatos existentes de una imagen (JPEG/PNG) |
| `check_metadata_divergence(path, expected)` | `path: Path, expected: dict` | Compara metadatos existentes vs esperados, retorna diferencias |
| `write_meta(path, metadata, organized)` | `path: Path, metadata: dict, organized: bool` | Dispatcher: escribe segun formato (JPEG/PNG/TIFF) |
| `write_jpeg(path, metadata, organized)` | `path: Path, metadata: dict, organized: bool` | Escribe 4 campos EXIF en JPEG |
| `write_png(path, metadata, organized)` | `path: Path, metadata: dict, organized: bool` | Escribe chunks de texto en PNG |
| `write_tiff(path, metadata, organized)` | `path: Path, metadata: dict, organized: bool` | Escribe ImageDescription en TIFF |

### `metatag_theme.py` — Sistema de temas

| Elemento | Descripcion |
|---|---|
| `THEMES` | Diccionario con 3 temas completos (Arqueologico, Noche Total, Carbon) |
| `THEME_ORDER` | Lista ordenada de nombres de tema |
| `ACCENT_TEXT` | Color de texto de acento: `"#FFF5E8"` |
| `compute_font_scale(screen_width)` | Escala de fuentes: `max(0.82, min(1.35, screen_width / 1920))` |
| `font_specs(scale)` | Retorna diccionario de 9 especificaciones de fuente escaladas |
| `TkThemeAdapter` | Aplica tema a widgets de Tkinter standard |
| `CustomTkinterThemeAdapter` | Aplica tema a widgets de CustomTkinter |

### `metatag_widgets.py` — ExcelGrid

| Metodo | Descripcion |
|---|---|
| `ExcelGrid.__init__(master, on_select, on_double_click, **kw)` | Canvas virtualizado con viewport culling |
| `load(dataframe, img_col, selected_rows, selected_cells)` | Carga DataFrame y renderiza solo filas visibles |
| `select_row(row)` / `select_cell(row, col)` | Seleccion programatica |
| `get_selection()` | Retorna `(selected_rows, selected_cells)` |

### `metatag_responsive.py` — Deteccion de pantalla

| Elemento | Descripcion |
|---|---|
| `ScreenProfile` | Clasifica pantallas: `laptop_small`, `laptop_large`, `desktop` |
| `PROFILE` | Singleton global con el perfil de la pantalla actual |
| `PROFILE.width` / `PROFILE.height` / `PROFILE.scale_factor` | Dimensiones y factor de escala |

### `metatag_xim.py` — Neutralizacion XIM

| Funcion | Descripcion |
|---|---|
| `neutralize_xim_for_tk()` | Desactiva la asociacion XIM/iBus para procesos Tk, eliminando 20-90ms de lag por widget |

---

## PARTE 5: TESTS

### Ejecutar tests

```bash
# Todos los tests:
.venv/bin/pytest tests/ -v

# Solo tests de Image Sync:
.venv/bin/pytest tests/test_renombrador_pytest.py -v

# Solo tests de matching:
.venv/bin/pytest tests/test_metatag_matching.py -v
```

### Cobertura de tests

| Archivo de test | Tests | Que cubre |
|---|---|---|
| `test_renombrador_pytest.py` | 70 | Image Sync completo |
| `test_renombrador.py` | 12 | Image Sync (unittest) |
| `test_reconciliacion.py` | 20 | Reconciliacion de contadores |
| `test_sinteticos_reconciliacion.py` | 39 | Escenarios sinteticos |
| `test_rename_real_seguro.py` | 7 | Renombrado real en tmp |
| `test_metatag_theme.py` | 53 | Paridad de temas |
| `test_metatag_matching.py` | 14 | Motor de matching |
| `test_matching.py` | 7 | Unitarias de matching |
| `test_matching_equivalence.py` | — | Equivalencia de regex compiladas |
| `test_dataset_269.py` | 4 | Integracion con dataset real |
| `test_grid.py` | 23 | ExcelGrid virtualizado |
| `test_queue.py` | 17 | Procesamiento en background |
| `test_responsive.py` | 12 | UI responsiva |
| `test_preview_table_virtualized.py` | 28 | Tabla virtualizada |
| `test_column_picker.py` | 19 | Selector de columnas |
| `test_scroll_layout.py` | — | Layout de scroll |
| `test_xim.py` | — | Neutralizacion XIM |

---

## PARTE 6: CONVENCIONES Y GUIA PARA DESARROLLADORES

### Convenciones de codigo

- **Idioma**: Comentarios y strings de UI en espanol
- **Patron**: MVC estricto en Image Sync; MVC informal en MetaTag principal
- **Naming**: snake_case para funciones/variables, PascalCase para clases, UPPER_CASE para constantes
- **Strings de UI**: Definidos como constantes al inicio del modulo o inline en `_build()`

### Convenciones de archivos

- **Nunca modificar originales**: MetaTag siempre trabaja sobre copias
- **Backup JSON**: `.metatag_backup_YYYYMMDD_HHMMSS.json` en la carpeta de fotos
- **Carpeta de salida**: `Metadatos_Escritos/` relativo a la ubicacion del script
- **Temp files**: `.metatag_tmp_<uuid>` para renombrado en dos fases
- **State persistente**: `.renombrador_state.json` para Image Sync, `data/metatag_config.json` para MetaTag

### Convenciones de tests

- **Framework**: pytest como runner principal, unittest para tests legacy
- **Fixtures**: `tests/fixtures/` para datos baseline
- **Tests sinteticos**: Escenarios controlados sin archivos reales
- **Tests reales**: Usan `Finales 1 a 103/` (269 imagenes de prueba)

### Dependencias

**Core (MetaTag):**
- `pandas`, `openpyxl` — manejo de Excel
- `pillow`, `piexif` — manipulacion de imagenes y EXIF
- `matplotlib`, `numpy` — graficos
- `reportlab` — exportacion PDF

**Image Sync (adicional):**
- `customtkinter` — interfaz moderna de Tkinter
