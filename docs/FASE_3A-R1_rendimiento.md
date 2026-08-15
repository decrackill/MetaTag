# FASE 3A-R.1 — Auditoría de rendimiento de MetaTag v8.9

Estado: completada el 2026-08-13. Método: MEDIR → IDENTIFICAR → EXPLICAR → PRIORIZAR.
Esta fase NO ha modificado código de la aplicación (solo scripts de profiling en `/tmp/opencode`, fuera del repo).

Entorno medido: Linux, X11 `DISPLAY=:0`, Cinnamon (compositor activo), X.Org 21.1.11, i915 (Intel Alder Lake), 8 CPUs, 7.4 GiB RAM, Python 3.12.3 en `.venv` (pandas 3.0.3, pillow 12.3.0, openpyxl 3.1.5, customtkinter 6.0.0, matplotlib 3.11.0, numpy 2.5.1).

Dataset sintético: fotos vacías + Excel de 19 columnas (mismo esquema que el real), N = 269 / 500 / 1000 / 5000 / 10000, generado con `/tmp/opencode/gen_data.py`.

Convención de evidencia usada en todo el informe:

- **HECHO**: medido/verificado con instrumentación real.
- **INFERENCIA**: conclusión derivada de las mediciones.
- **HIPÓTESIS**: afirmación que todavía necesita prueba.
- **PROYECCIÓN**: extrapolación, NO medición (el valor real no se ha ejecutado).

---

## 1. HITO: causa raíz del lag principal — iBus/XIM (Tk ↔ método de entrada X)

### 1.1 Mediciones (HECHO)

| Operación | Con XIM (`XMODIFIERS=@im=ibus`) | Sin XIM (`XMODIFIERS=`) | Mejora |
|---|---|---|---|
| Startup MetaTag (build completo, N=269, config limpia) | **10.93 s** | **0.33 s** | ~33× |
| Cambio de tema (`_apply_rebuild`) | **~19.9 s** | **0.13 s** | ~150× |
| Test aislado: 200 labels + `update_idletasks` | **3.72 s** | **0.037 s** | ~100× |
| Test aislado: 300 labels + `update_idletasks` (1200×700) | ~7 s | — | — |

Verificación adicional (HECHO):

- `GTK_IM_MODULE=` vacío → no cambia nada (7.0 s). `QT_IM_MODULE=` vacío → no cambia nada (7.3 s).
- `XMODIFIERS=` vacío **o** `XMODIFIERS=@im=none` → 0.03–0.04 s. La variable que controla el XIM de Tk es **`XMODIFIERS`**.
- Durante el flush, `ibus-daemon` (que corre con `--xim`) alcanza ~100 % de CPU en un núcleo; el proceso Tk ~100 %; `Xorg` y `cinnamon` por debajo del 10 %. `ibus-daemon` acumula 106 min de CPU en 4.5 h de uptime (~38 % promedio).
- CPU del proceso MetaTag durante el build lento: **2.2 %** → el proceso está mayormente esperando (el trabajo lo hacen iBus/XIM fuera del proceso).
- Reproducible en la app real (`profile_uidle.py`), no solo en un test aislado.

### 1.2 Análisis solicitado — iBus/XIM como P0

**¿Cuánto del problema procede realmente del entorno? (INFERENCIA)**
- El ~95 % del coste de startup y del cambio de tema es XIM/iBus, no el código de MetaTag (10.9 s → 0.33 s con solo quitar XIM). El código restante (build_ui, carga de df) está en 0.1–0.3 s.
- La penalización de XIM es **por ventana X con interacción IM** (widget): cada redraw/config de widget cuesta ~20–90 ms con XIM frente a ~0–1 ms sin él. Esto se multiplica por el número de widgets Tk vivos.

**¿`XMODIFIERS=` es seguro? (INFERENCIA)**
- `XMODIFIERS` solo informa al cliente X de **qué servidor de método de entrada usar**. Vacío/`@im=none` = XIM desactivado. No toca iBus como servicio, ni GTK, ni Qt (esos usan `GTK_IM_MODULE`/`QT_IM_MODULE`, que permanecen intactos y siguen activos para el resto del escritorio).
- Teclado activo: `latam` (pc105). Los acentos latinoamericanos se generan con **teclas muertas** (dead keys) a nivel de layout X11, que NO dependen de XIM. Por tanto no se pierde la entrada de acentos/ñ en Tk con `XMODIFIERS=` vacío.
- Riesgo real de `XMODIFIERS=` solo en: métodos de entrada con composición por IM (chino, japonés, coreano, etc.), que no aplican a este usuario (layout latam, sistema en español).

**¿Puede aplicarse únicamente al proceso de MetaTag? (HECHO→diseño)**
- Sí: es una variable de entorno por proceso. Al lanzar MetaTag se puede exportar `XMODIFIERS=` **solo para ese proceso** (env en el `.desktop`/`.sh`), sin cambiar el entorno global. Ver alternativas en 1.3.
- HIPÓTESIS a verificar en 3B.1: `os.environ["XMODIFIERS"] = ""` al inicio de `metatag_v8.py`, **antes** de crear el primer `Tk()` (y en el proceso del Renombrador), basta para desactivar XIM en Tk sin tocar el entorno del sistema.

**¿Afecta entrada de teclado, dead keys, acentos LATAM u otras funciones? (INFERENCIA + HIPÓTESIS)**
- INFERENCIA: no debe afectar (dead keys viven en el layout X11, no en XIM).
- HIPÓTESIS pendiente de verificación en 3B.1 con test interactivo: escribir con acentos y ñ en un `Entry` de Tk con `XMODIFIERS=` vacío, y en la app real.

**¿También afecta al Renombrador? (HECHO)**
- Sí. El Renombrador es CustomTkinter (Tk). `PreviewTable` con 269 filas: **2.4 s sin XIM vs 3.6 s con XIM**. Con `root.withdraw()` la penalización se reduce pero no desaparece. A 1000 filas el render con XIM no completó en 180 s (timeout) → HECHO de que la combinación widget-per-row × XIM es inviable.

**¿Existe alternativa más localizada que modificar el entorno completo? (diseño)**
1. `XMODIFIERS=` solo en el proceso MetaTag/Renombrador (la más simple; env de lanzamiento o `os.environ` en el entrypoint antes de Tk).
2. ~~HIPÓTESIS~~ → **HECHO (medido el 2026-08-13)**: `option add *useInputMethods 0` en Tk 8.6 NO elimina la penalización (200 labels + flush: 3.49 s base vs 7.10 s con la opción). Alternativa por-widget descartada; queda `XMODIFIERS=` por-proceso como única vía efectiva.
3. Config de iBus para ignorar Tk: no existe un mecanismo fiable/estable → se descarta.

**Decisiones de P0 (no implementadas todavía):**
- NO desactivar iBus globalmente ni modificar `im-config`/config permanente del sistema.
- Se evaluará SOLO el enfoque por-proceso/localizado.

---

## 2. P1 — Arquitectura crítica: PreviewTable widget-per-row

### 2.1 Mediciones (HECHO)

`src/renombrar_fotos_gui.py`, `PreviewTable` (~5 widgets Tk/CustomTkinter por fila, render en chunks de 30 vía `after(5)`):

| Filas | Render sin XIM (HECHO) | Render con XIM (HECHO) | RSS (HECHO) |
|---|---|---|---|
| 269 | **2.42 s** | **3.56 s** | — |
| 500 | **4.34 s** | — | — |
| 1000 | **9.4–13.2 s** | timeout > 180 s | **181 MB** (~16 KB × 5 widgets/fila) |

- 10000 filas ≈ 90–130 s y ~+800 MB: **PROYECCIÓN** (no ejecutado; en esta máquina, con ~2.7 GiB libres, probablemente OOM/swap → inviable).
- El render es "chunked" (`after(5)` cada 30 filas) → no bloquea del todo el event loop, pero la ventana queda congelada/borrosa durante minutos a escala, y el scroll sobre miles de widgets Tk es inservible.

### 2.2 Comparativa formal de soluciones (diseño, sin implementar)

| Criterio | Widgets reciclables (lazy per fila visible) | Canvas dibujado directo | ttk.Treeview | Híbrida (canvas + solapas ligeras) |
|---|---|---|---|---|
| Nº de ventanas X | O(visibles) (~30) | 1 | 1 | ~1 + pocas |
| Coste redraw a 10k filas | bajo (viewport) | muy bajo | medio (virtualización nativa) | bajo |
| Scroll suave | sí (viewport) | sí | sí (integrado) | sí |
| Estilo CustomTkinter/temas | se adapta | se dibuja manual | malo (tema nativo del sistema) | se adapta |
| Celdas editables / click por fila | sí (reaprovechando widget en posición) | manual (hit-testing) | integrado pero limitado | sí |
| Ceros/estados de fila (ok/dup/conflicto…) | natural | manual | columnas extra | natural |
| Copia del estilo visual actual | alta | media (replicar padding/fuentes) | baja | alta |
| Complejidad de implementación | media | media-alta | baja | media |

**Veredicto (INFERENCIA):**
- **ttk.Treeview** es la opción de menor esfuerzo pero rompe el aspecto visual y el patrón de tema de MetaTag (decisión de producto → se descarta como primera opción).
- **Widgets reciclables (viewport) sobre un frame desplazable** conserva estilo y es el mismo patrón que ya usa `ExcelGrid` (canvas + culling) en la app principal, que ya está validado: `grid.redraw()` 10k filas = **0.029 s (HECHO)**.
- **Canvas directo** es el más barato en X pero exige reimplementar hit-testing, edición y scroll.
- **Recomendación de diseño (HIPÓTESIS de trabajo, a validar en 3A-R.4):** híbrido/reciclable siguiendo el patrón `ExcelGrid` (solo renderiza filas visibles; mantiene el tema; scroll por canvas). La virtualización es imprescindible, no opcional, a partir de ~500 filas.

---

## 3. P1 — Algoritmo: `_find_image_ex` / `ImageMatcher`

### 3.1 Dónde está la complejidad (análisis)

En `src/metatag_v8.py` (`_find_image_ex`, línea ~2795) y `src/metatag_matching.py` (`ImageMatcher`):

- **HECHO**: sin match exacto, hace ~5 barridos lineales O(n) sobre el índice de la carpeta (exacto → stem → limpio → normalizado → sufijo ID), cada uno con normalización por archivo.
- **HECHO (cProfile, 100 llamadas × 5000 archivos)**: **1 500 200 llamadas a `re.sub`** y **1 500 200 a `re._compile`** — la regex `r"\d+"` se recompila en cada llamada (patrón literal dentro de `re.sub` no precompilado) dentro de `_normalize_numbers`.
- **HECHO**: `_find_image_ex` not-found: 1 llamada = 0.010 s (269), 0.039 s (1000), 0.13 s (5000), **0.27 s (10000)** → O(n) por llamada.
- **HECHO**: 200 llamadas not-found = 0.27 s (269), 2.09 s (1000), 4.57 s (5000), **10.6 s (10000)**.
- **HECHO (Renombrador)**: `build_preview` con nombres inexistentes = **0.26 s (269), 3.66 s (1000)** → comportamiento **O(n²)** (n llamadas × n archivos).
- **PROYECCIÓN**: un pase completo de 10k filas sin match ≈ 45 min; `build_preview` not-found a 10000 ≈ 6 min.

### 3.2 Qué se repite y qué puede precalcularse/cachearse (diseño)

Repeticiones por llamada (cuando no hay match exacto):
1. Recorrer n archivos varias veces (hasta 5 pasadas) con regex por archivo.
2. `re._compile` de la misma regex en cada normalización → precompilar los patrones a nivel de módulo elimina las 1.5 M recompilaciones.
3. Las claves normalizadas de cada archivo se recalculan en cada llamada, aunque son invariantes dentro de una carpeta.

Precomputable/cacheadle (clave invariante por carpeta):
- Por archivo: `stem` limpio, `stem` normalizado (`_normalize_numbers`), sufijo de ID extraído (`_extract_id_suffix`), `_safe_stem`.
- Índice multi-clave por carpeta: `dict[str, Path]` con cada clave normalizada → cada variante de búsqueda pasa de O(n) a **O(1)**.

### 3.3 Estructura de índice propuesta (diseño, SIN implementar)

- Una vez por carpeta (al cargar/`build_index`): construir y cachear `{exacto, stem, limpio, normalizado, sufijo_id} -> Path`.
- Orden de resolución idéntico al actual (exacto → stem → … ) pero cada paso es un lookup de diccionario, no un barrido.
- Caso `not_found` pasa de O(n) por llamada (O(n²) en `build_preview`) a O(1) por llamada (O(n) total).
- **Regla de equivalencia (crítica):** la línea base de matching debe permanecer intacta. Antes de tocar `ImageMatcher` hay que demostrar equivalencia: mismo resultado que el algoritmo validado sobre un set de casos (exacto, doble extensión, espacios, cero-padding, sufijos, ambiguos, duplicados, not-found). El plan es: (1) extraer el algoritmo actual a un oráculo de referencia, (2) generar casos de prueba, (3) implementar índice multi-clave, (4) validar salida 1:1 contra el oráculo, (5) recién entonces activar.

### 3.4 Coste estimado del fix (PROYECCIÓN)

`find_image_ex` not-found a 10000: de 0.27 s → ~µs (lookup fallido) + coste de normalizar el término de búsqueda (una sola vez por llamada). `build_preview` not-found a 1000: de 3.66 s → < 0.1 s. A verificar con benchmark real en 3B.3.

---

## 4. P2 — Filtro `df.apply(lambda row…)`

### 4.1 Hecho

- `_apply_filter` (metatag_v8.py:1425) con modo "Todas las columnas": `df.apply(lambda row: …axis=1)` = **4.1 M llamadas**, 5000 filas → **1.73 s** (cProfile). El 70 % es el bucle `series_generator` con `astype(str)` + `.str.contains` + `_ixs`/`fast_xs` por fila.
- A escala: 0.05 s (269), 0.23 s (1000), 0.98 s (5000), **1.64 s (10000)**.
- Con el debounce de 300 ms, cada pausa de escritura provoca esta congelación.

### 4.2 Alternativas comparadas (diseño)

| Opción | Mecanismo | Semántica equivalente | Coste a 10k (proy.) | Notas |
|---|---|---|---|---|
| `apply` (actual) | bucle row-wise | sí | 1.6 s | lento |
| `df.astype(str)` + OR de `str.contains` por columna | vectorizado por columna | **exacta** si se replican flags `na`/`regex` actuales | ~0.2–0.4 s | columna con dtype object ya es str en su mayoría → se evita el `astype` por fila |
| Máscara acumulada con `np.char` / regex vectorizada | `np.char.find` o listcomp precompilado por columna | equivalente | ~0.1–0.2 s | depende de tipos |
| Cache de columnas como listas de str (una vez por carga) | pre-indexar columnas en listas Python | equivalente (mismo texto de celda) | ~0.05–0.15 s | evita acceso pandas por fila; costo único al cargar |

**Veredicto (INFERENCIA):** el punto débil no es `str.contains` en sí sino el `apply` row-wise que re-materializa cada fila y re-hace `astype(str)` por fila. Convertir a operación **por columna** (máscara OR) preserva la semántica de "buscar en todas las columnas" y es la vía vectorizada natural. Verificación de equivalencia exigida: mismo conjunto de filas resultantes que el `apply` actual en el dataset real y en casos borde (NaN, números, fechas, textos con regex chars).

---

## 5. P2 — Reconstrucción de UI: `_apply_rebuild`

### 5.1 Por qué es caro (HECHO)

- `_apply_rebuild` (metatag_v8.py:709) destruye y reconstruye TODA la UI y además **re-escanea la carpeta de imágenes** (`browser.load_folder`) y re-carga df. En este equipo, cada widget que se crea/muestra cuesta ~20–90 ms con XIM → 19.9 s con XIM; 0.13 s sin XIM. Aun sin XIM, es trabajo innecesario.

### 5.2 Análisis de qué se puede conservar (diseño)

- ¿Por qué re-escanea? INFERENCIA: práctica "reconstrucción total" defensiva, no un requisito funcional. El escaneo es barato (0.002 s en 269) pero con XIM la creación de widgets es lo que cuesta.
- Estado que en realidad cambia con el tema: paleta de colores y fuentes de los widgets.
- Lo que NO cambia: df cargado, filtros activos, selección, scroll, carpeta, matching, contenido del ExcelGrid.
- Opciones de diseño a evaluar en 3A-R.3:
  1. Reestilizar en sitio (recorrer widgets y actualizar colores/fuentes) sin reconstruir, igual que ya hace el Renombrador con `_refresh_button_constants()` + actualización de widgets.
  2. Si el framework fuerza reconstrucción: reconstruir **solo la porción visible/dependiente del tema**, conservando datos y sin re-escaneo.
  3. Eliminar el `update_idletasks()` síncrono del build (`_build_ui:805`) y de `_save_config` (se ejecuta en load, toggle_mode, tema, procesar) cuando la ventana está visible (con XIM es lo que congela).

**Regla**: no asumir que reconstruir es necesario si existe una forma segura de reestilizar; se demostrará con un cambio de tema ANTES→DESPUÉS en 3B.5 conservando estado visible (fila seleccionada, scroll, filtro).

---

## 6. P3 — Optimizaciones menores (agrupadas)

- `_col_fully_selected` (ExcelGrid): O(rows) por columna en cada redraw; mantener selección por rango en vez de recomprobar por fila. Impacto: bajo a 10k (redraw 0.029 s) pero crece con filas×columnas.
- `_sync_excel_to_images` / `file_order`: mezcla O(rows×files) vía `_sort_key`; solo importa con mucha diferencia entre Excel y carpeta.
- `_save_config` con `update_idletasks()` en cada operación (load, toggle_mode, tema, procesar): con XIM, cada llamada puede costar segundos → eliminar/mover a `after_idle` cuando ventana visible.
- `build_styles()` 0.0004 s: no optimizar.
- Import 0.72 s (pandas 0.26, matplotlib 0.31): solo importar matplotlib bajo demanda si alguna vista lo usa; impacto menor.

---

## 7. Tabla de priorización consolidada

Prioridad = impacto real × frecuencia × escalabilidad. Impacto con XIM activo (lo que ve el usuario hoy).

| Problema | Causa | Capa | Impacto | Complejidad sol. | Riesgo regresión | Prioridad | Benchmark actual (HECHO) | Objetivo |
|---|---|---|---|---|---|---|---|---|
| iBus/XIM | Tk abre XIM e interactúa con iBus por cada widget; ~20–90 ms/widget | Entorno/display | **Crítico**: 10.9 s startup, 19.9 s tema (solo esta máquina) | Baja | Baja si es por-proceso y se valida teclado | **P0** | startup 10.93 s; tema 19.9 s | startup <1 s; tema <1 s (por-proceso, sin tocar sistema) |
| PreviewTable widget-per-row | ~5 widgets/fila × n filas, sin virtualización | Arquitectura | **Crítico a escala**: 2.4 s @269, 9–13 s @1000, 181 MB; 10k inviable (proy. ~90–130 s/+800 MB) | Media-Alta | Media (reemplazo de componente) | **P1** | 2.42 s @269; 9.4–13.2 s @1000 | <100 ms @1000; scroll fluido @10k |
| `_find_image_ex`/ImageMatcher O(n) + regex recompilada | 5 barridos lineales por llamada; `re.sub` recompila 1.5 M veces; O(n²) en not_found | Algoritmo | **Alto** en peor caso: 0.27 s/llamada @10k; 45 min por pase (proy.) | Media | **Alto** (matching validado; exige demostración de equivalencia) | **P1** | 10.6 s (200×not-found @10k); 3.66 s build_preview @1000 not-found | <0.05 s (200× @10k); O(n) total |
| Filtro `apply` row-wise | `df.apply(axis=1)` con `astype(str)`+`contains` por fila; 4.1 M llamadas | Algoritmo | Medio-Alto: 1.64 s @10k por pausa de tecleo | Baja | Baja (equivalencia verificable) | **P2** | 1.64 s @10k | <0.2 s @10k |
| `_apply_rebuild` reconstrucción total | Destruye/reconstruye UI + re-escanea + re-carga, con `update_idletasks` visible | Arquitectura/UI | Alto con XIM (19.9 s); bajo sin XIM (0.13 s) | Media | Media (tema) | **P2** | tema 19.9 s (XIM) / 0.13 s (sin XIM) | <1 s, conservando estado |
| `_save_config` con `update_idletasks` | flush síncrono visible en cada operación | UI | Medio con XIM (segundos por operación) | Baja | Baja | **P3** | ~1–2 s/operación (XIM) | <50 ms |
| P3 menores (`_col_fully_selected`, `_sync` O(n×m), import matplotlib) | ineficiencias puntuales | Varias | Bajo | Baja | Baja | **P3** | — | — |

---

## 8. Orden de implementación propuesto (3B) — sujeto a 3A-R.3/R.4

Basado en las mediciones (no en facilidad): el P0 XIM multiplica el coste de TODO lo demás (widgets), así que va primero y desbloquea mediciones limpias del resto.

1. **3B.1** — Reducir/eliminar el impacto de iBus/XIM de forma **localizada** (solo proceso MetaTag/Renombrador; validar teclado LATAM/dead keys; probar también `*useInputMethods` por-widget). → ANTES 10.9 s / DESPUÉS <1 s.
2. **3B.2** — Virtualizar `PreviewTable` (patrón `ExcelGrid`: render solo filas visibles; mantener estilo/temas). → ANTES 2.4 s @269 / DESPUÉS <100 ms @1000.
3. **3B.3** — Optimizar `ImageMatcher` con índice multi-clave **preservando equivalencia** contra oráculo de referencia. → ANTES 10.6 s / DESPUÉS <0.05 s (200×@10k).
4. **3B.4** — Vectorizar filtro por columnas (máscara OR) con equivalencia verificada. → ANTES 1.64 s / DESPUÉS <0.2 s @10k.
5. **3B.5** — Reestilizar en vez de reconstruir en `_apply_rebuild` + quitar `update_idletasks` síncronos visibles. → ANTES 19.9 s / DESPUÉS <1 s conservando estado.
6. **3B.6** — Regresión completa (tests existentes: matching, undo, Excel, temas, Renombrador, filtros) + benchmarks ANTES/DESPUÉS y memoria.

Cada optimización demostrará ANTES → CAMBIO → DESPUÉS → REGRESIÓN, sin tocar matching validado sin prueba de equivalencia.
