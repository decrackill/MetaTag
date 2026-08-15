# FASE 3A-R.2 — Clasificación y priorización de cuellos de botella

Estado: completada el 2026-08-13. Base: FASE 3A-R.1 (`docs/FASE_3A-R1_rendimiento.md`).
Esta fase NO modifica código de la aplicación.

Convención de evidencia (igual que R.1):

- **HECHO**: medido/verificado con instrumentación real.
- **INFERENCIA**: conclusión derivada de las mediciones.
- **HIPÓTESIS**: afirmación que todavía necesita prueba.
- **PROYECCIÓN**: extrapolación, NO medición.

---

## 0. Verificaciones nuevas de esta fase (HECHO)

- **`option add *useInputMethods 0` NO funciona** en Tk 8.6 de Python 3.12: 200 labels + flush = 3.49 s base vs 7.10 s con la opción. Descartada la alternativa "localizada por widget".
- **Layout X11 activo**: `latam,us` (pc105), `terminate:ctrl_alt_bksp`. Los acentos/ñ en `latam` se componen con **teclas muertas a nivel de layout X11**, no por el método de entrada → no dependen de XIM (INFERENCIA). El test interactivo de tipeo con `XMODIFIERS=` no pudo ejecutarse (no hay `xdotool`) → **HIPÓTESIS** pendiente de verificar en 3B.1.
- **`ImageMatcher.find_image_ex` (src/metatag_matching.py:145) es réplica 1:1** de `MetaTagApp._find_image_ex` (src/metatag_v8.py:2795): misma jerarquía de 7 pasos y mismo coste. Ya existen `_index_folder` (índice `{stem_lower: Path}` cacheado por carpeta) y `find_image_ex_with_method` (trazador de paso) → **infraestructura lista para construir el oráculo de equivalencia**.
- Set de regresión disponible: **199 tests** (`test_matching.py`, `test_metatag_matching.py`, `test_grid.py`, `test_renombrador*.py`, `test_theme.py`, `test_responsive.py`, `test_dataset_269.py`, `test_queue.py`, `test_column_picker.py`).

---

## 1. P0 — iBus/XIM (bloqueo externo / entorno)

**HECHO**: 10.93 s startup / ~19.9 s tema con XIM; 0.33 s / 0.13 s sin XIM. La variable que controla el XIM de Tk es `XMODIFIERS`. `ibus-daemon --xim` al 100 % de un núcleo durante el flush.

Análisis de las preguntas planteadas:

- **¿Cuánto es entorno?** ~95 % del coste de startup/tema es XIM/iBus (10.9→0.33 s solo quitando XIM). El código de MetaTag aporta <0.3 s. (INFERENCIA)
- **¿`XMODIFIERS=` es seguro?** Sí para este equipo: solo informa al cliente X del servidor IM; no toca iBus como servicio ni GTK/Qt (usan `GTK_IM_MODULE`/`QT_IM_MODULE`, intactos). No afecta dead keys/acentos LATAM (nivel layout X11). Única excepción real: métodos de entrada con composición (CJK…), que no aplican. (INFERENCIA + HIPÓTESIS pendiente del test de tipeo)
- **¿Por-proceso?** Sí: variable de entorno por proceso (env de lanzamiento o `os.environ["XMODIFIERS"]=""` en el entrypoint **antes** del primer `Tk()`), tanto para MetaTag como para el Renombrador. Sin tocar la config global. (diseño; HIPÓTESIS de que el ajuste pre-Tk basta)
- **¿Afecta al Renombrador?** Sí (CustomTkinter = Tk): 269 filas → 2.42 s sin XIM vs 3.56 s con XIM. (HECHO)
- **¿Alternativa más localizada?** Medida y descartada: `*useInputMethods 0` no elimina la penalización (HECHO). No existe mecanismo iBus confiable para ignorar Tk. → `XMODIFIERS=` por-proceso es la única vía efectiva.

**Decisión**: NO desactivar iBus globalmente ni modificar `im-config`. 3B.1 usará solo el enfoque por-proceso y validará tipeo/acentos.

---

## 2. P1 — Arquitectura crítica: PreviewTable widget-per-row

**HECHO**: 269→2.42 s, 500→4.34 s, 1000→9.4–13.2 s (RSS 181 MB, ~16 KB/widget × 5 widgets/fila).
**PROYECCIÓN**: 10000→~90–130 s y ~+800 MB (NO ejecutado; inviable en esta máquina con ~2.7 GiB libres).

Comparativa (ver tabla completa en R.1 §2.2):

- **Widgets reciclables (viewport)**: O(visibles) widgets; mismo patrón que `ExcelGrid` (redraw 10k = 0.029 s, HECHO); conserva estilo/tema. ← recomendado.
- **Canvas directo**: 1 ventana X, el más barato, pero exige reimplementar hit-testing/edición/scroll a mano.
- **ttk.Treeview**: fácil y virtualizado, pero rompe el aspecto visual y el sistema de temas de MetaTag (decisión de producto).
- **Híbrida**: canvas + pocos solapas; flexible, más complejo.

Veredicto (INFERENCIA): virtualización con **reciclaje de widgets visibles** siguiendo el patrón `ExcelGrid`, manteniendo el tema. Problema arquitectónico, no microbenchmark.

---

## 3. P1 — Algoritmo: `_find_image_ex` / `ImageMatcher`

### 3.1 Desglose exacto de la complejidad (basado en código + cProfile)

Cuando no hay coincidencia directa (`p.exists()`), el algoritmo (idéntico en ambos archivos) ejecuta hasta **5 barridos O(n)** sobre el índice de la carpeta:

| Paso | Operación | Coste por archivo | O() | ¿Caros? |
|---|---|---|---|---|
| 1. `p.exists()` | stat de fs | 1 syscall | O(1) | no |
| 2. nombre exacto | `fpath.name.lower() == name_lower` | comparación str | O(n) | no |
| 3. stem exacto | lookup en dict `_img_cache`/`index` | dict | **O(1)** | no |
| 4. stem "limpio" | `re.sub(r"^[#\s\-_]+|[#\s\-_]+$", …)` por archivo, **patrón inline → `re._compile` por llamada** | 1–2 regex | O(n) | **sí** |
| 5. números normalizados | `_normalize_numbers` = `re.sub(r"\d+", lambda…)` por archivo, **inline → recompila por llamada** | 2 regex + lambda por archivo | O(n) | **sí (el peor)** |
| 6. id-suffix | `_extract_id_suffix(stem_key)` = ~2 `re.match` por archivo (inline) | 2 regex | O(n) | sí |
| 7. subcadena | `name_stem in stem_key or stem_key in name_stem` | substring str | O(n) | no (barato) |

**HECHO (cProfile)**: 100 llamadas × 5000 archivos → **1 500 200 `re.sub` + 1 500 200 `re._compile`** (la regex se recompila en cada llamada porque va literal en `re.sub`, no precompilada) y `_normalize_numbers` con 1.7 M llamadas. Coste medido: 0.27 s/llamada @10k not-found; 200 llamadas = 10.6 s.
**PROYECCIÓN**: un pase completo de 10k filas not-found ≈ 45 min; `build_preview` not-found del Renombrador a 10k ≈ 6 min (O(n²)).

### 3.2 Qué se repite y qué es invariante (diseño)

- Se repiten **dentro de una misma carpeta**: las claves derivadas de cada archivo (`_full_stem`, `_clean_stem`, `_normalize_numbers`, `_extract_id_suffix`) — son invariantes y se pueden calcular **una sola vez por carpeta**.
- Se repite **en cada llamada**: la recompilación de las regex (inline). → `re.compile` a nivel de módulo las elimina.
- El paso 7 (subcadena) es el único no indexable con un dict (contiene `in`); se conserva como barrido O(n) barato (~2–5 ms @10k, PROYECCIÓN) solo como último recurso.

### 3.3 Índice multi-clave propuesto (diseño; NO implementado)

Precalcular en la construcción del índice (una vez por carpeta):

- `by_name_lower` = `{fpath.name.lower(): path}` (primera ocurrencia, mismo orden de iteración)
- `by_stem` = actual `{stem_lower: path}`
- `by_clean` = `{clean_stem(stem): path}` (primera ocurrencia)
- `by_normalized` = `{normalize_numbers(clean_stem(stem)): path}` (primera ocurrencia)
- `by_id_suffix` = `{extract_id_suffix(stem): [paths]}` (lista, para ambigüedad)

Los pasos 2–6 pasan a **O(1)** por lookup. El paso 7 (subcadena) queda como barrido final. not_found pasa de O(n) por llamada a O(1)+O(n)barato → `build_preview` not-found pasa de O(n²) a O(n).

**Requisitos de equivalencia (críticos, no negociables):**
1. La línea base de matching debe permanecer intacta (oráculo).
2. Semántica de "primera coincidencia en el orden de iteración": en MetaTag el orden es el del rglob (sin ordenar); en ImageMatcher, ordenado por `as_posix`. Los dicts multi-clave deben preservar exactamente ese orden ("first-wins") o el resultado puede diferir ante colisiones de stem.
3. El paso 6/7 devuelve la lista COMPLETA de candidatos en orden → el índice debe conservar listas, no solo el primero.
4. Demostración de equivalencia: corpus de casos (exacto, doble extensión, `(1)`, cero-padding, bordes `#/_-`, mayúsculas, sufijos F/R/P, ambiguos, duplicados, not-found, colisiones de stem) comparando `(path, status, candidatos)` 1:1 contra `find_image_ex_with_method` de la versión actual (el trazador ya existe y NO añade decisiones).

Coste estimado del fix (PROYECCIÓN): `find_image_ex` not-found @10k de 0.27 s → ~µs de lookup + 1 normalización del término + barrido subcadena. `build_preview` not-found @1000 de 3.66 s → <0.1 s. A verificar en 3B.3.

**Quick-win sin riesgo (se ejecutará en 3B.1/3B.3):** precompilar las regex de `_normalize_numbers`, `_clean_stem` y `_extract_id_suffix` a nivel de módulo elimina las 1.5 M `re._compile` sin tocar el algoritmo → mejora parcial inmediata con riesgo casi nulo.

---

## 4. P2 — Filtro `df.apply(lambda row…)`

**HECHO**: `_apply_filter` "Todas las columnas" = 4.1 M llamadas; 0.05 s @269, 0.23 s @1000, 0.98 s @5000, **1.64 s @10000**. Con debounce de 300 ms → congelación por pausa de tecleo.

Comparativa (diseño):

| Opción | Mecanismo | Equivalencia | Coste @10k (proy.) |
|---|---|---|---|
| `apply` row-wise (actual) | bucle Python por fila + `astype(str)` + `str.contains` | — | 1.6 s |
| Máscara OR por columna | `df[col].astype(str).str.contains(pat, na=False, regex=True)` en las columnas relevantes, OR acumulado | exacta replicando flags actuales | ~0.2–0.4 s |
| `np.char` / regex vectorizada | operaciones sobre arrays numpy pre-extraídos | equivalente | ~0.1–0.2 s |
| Cache de columnas como listas str (una vez por carga) + `any` | listcomp por columna sobre cadenas Python | equivalente (mismo texto de celda) | ~0.05–0.15 s |

Veredicto (INFERENCIA): el cuello no es `str.contains` sino el `apply` que re-materializa cada fila. La vía natural es **por columna** (máscara OR). Se exigirá equivalencia exacta (mismo conjunto de filas) en casos borde: NaN/NaN-string ("NA"/"N/A"/"nan"), números, fechas, texto con metacharacteres regex.

**Nota de prioridad real (INFERENCIA)**: el dataset actual del usuario es **269 filas** (filtro = 0.05 s, inapreciable). El filtro solo importa si el dataset crece ≥5k. Su prioridad efectiva es menor que la que sugiere el benchmark a 10k.

---

## 5. P2 — Reconstrucción de UI: `_apply_rebuild`

**HECHO**: destruye/reconstruye toda la UI + re-escanea carpeta + re-carga df + `update_idletasks` visible → 19.9 s con XIM / 0.13 s sin XIM.

Análisis:

- **¿Por qué re-escanea?** (INFERENCIA) práctica defensiva de "reconstrucción total"; el escaneo en sí es barato (0.002 s @269) — lo caro es recrear widgets con XIM.
- **¿Qué cambia con el tema?** solo paleta de colores y fuentes.
- **¿Qué NO cambia?** df cargado, filtros activos, selección, scroll, carpeta, matching, contenido de ExcelGrid, estado del Renombrador.
- **¿Qué puede conservarse?** datos y estado visual; reestilizar widgets en sitio (el Renombrador ya lo hace con `_refresh_button_constants()`).
- **¿Es necesaria la reconstrucción total?** No si existe reestilizado seguro; se demostrará con cambio de tema ANTES→DESPUÉS conservando fila seleccionada/scroll/filtro.

**Impacto real (INFERENCIA)**: con 3B.1 (XIM) ya aplicado, `_apply_rebuild` cae a 0.13 s automáticamente → el valor residual de 3B.5 es pulir (evitar re-escaneo, `update_idletasks`, conservar estado), bajo pero barato.

---

## 6. P3 — Optimizaciones menores (agrupadas)

- `_col_fully_selected` O(rows) por columna en cada redraw → selección por rango. (Bajo; redraw 0.029 s @10k)
- `_save_config` con `update_idletasks()` visible en load/toggle/tema/procesar → con XIM cuesta segundos por operación; mover a `after_idle`/omisión. (Medio con XIM; ya mitigado por 3B.1)
- `_sync_excel_to_images`/`file_order` O(rows×files) vía `_sort_key`. (Bajo)
- Import de `matplotlib.pyplot` (0.31 s) solo si la vista lo requiere. (Bajo)

---

## 7. Tabla de priorización consolidada

Prioridad = impacto real × frecuencia × escalabilidad. "Impacto" = lo que ve el usuario HOY (con XIM activo).

| Problema | Causa | Capa | Impacto | Complejidad sol. | Riesgo regresión | Prioridad | Benchmark actual | Objetivo |
|---|---|---|---|---|---|---|---|---|
| iBus/XIM | XIM de Tk ↔ iBus, ~20–90 ms/widget | Entorno | Crítico: 10.9 s startup, 19.9 s tema | Baja | Baja (por-proceso; validar tipeo) | **P0** | 10.93 s / 19.9 s (HECHO) | <1 s / <1 s |
| PreviewTable widget-per-row | ~5 widgets/fila sin virtualización | Arquitectura | Crítico: 2.4 s @269; 9–13 s @1000; 181 MB; 10k inviable | Media-Alta | Media | **P1** | 2.42 s @269; 9.4–13.2 s @1000 (HECHO); 90–130 s/+800 MB @10k (**PROYECCIÓN**) | <100 ms @1000; scroll fluido @10k |
| `_find_image_ex`/`ImageMatcher` O(n) + regex recompilada | 5 barridos lineales; `re._compile` inline; O(n²) en not_found | Algoritmo | Alto peor caso: 0.27 s/llamada @10k; 45 min/pase (proy.) | Media | **Alto** (matching validado → oráculo) | **P1** | 10.6 s (200× @10k); 3.66 s build_preview @1000 not-found (HECHO) | <0.05 s (200× @10k); O(n) total |
| Filtro `apply` row-wise | `df.apply(axis=1)` + `astype(str)` por fila; 4.1 M llamadas | Algoritmo | Medio (solo >5k filas; @269 = 0.05 s) | Baja | Baja | **P2** | 1.64 s @10k (HECHO) | <0.2 s @10k |
| `_apply_rebuild` reconstrucción total | Rebuild completo + re-escaneo + `update_idletasks` | Arquitectura/UI | Alto con XIM; casi nulo tras 3B.1 | Media | Media | **P2** | 19.9 s (XIM) / 0.13 s (sin XIM) (HECHO) | <1 s conservando estado |
| `_save_config` + `update_idletasks` | flush síncrono visible | UI | Medio con XIM | Baja | Baja | **P3** | ~1–2 s/operación (XIM) (HECHO) | <50 ms |
| P3 menores | varios | varias | Bajo | Baja | Baja | **P3** | — | — |

---

## 8. Orden de implementación propuesto (3B) — revisado según mediciones

El usuario propuso: 3B.1 XIM → 3B.2 virtualizar → 3B.3 matching → 3B.4 filtro → 3B.5 rebuild → 3B.6 regresión.

Revisión según las mediciones:

- **Mantener 3B.1 primero**: XIM multiplica el coste de TODO widget (medido); desbloquea mediciones limpias del resto y elimina de golpe ~95 % del coste de startup/tema.
- **Mantener 3B.2 antes de 3B.3**: el render del PreviewTable es lento incluso sin XIM (HECHO 2.4 s @269) y es la herramienta central del flujo fotos; el matching en el caso normal ya es rápido (0.04 s @1000). Virtualizar ataca el problema más visible.
- **3B.3 matching**: alto riesgo (exige oráculo) → se hace con garantías. Incluye el quick-win de precompilar regex (riesgo casi nulo, se puede adelantar a 3B.1).
- **Bajar 3B.4 (filtro) en la práctica**: el dataset real es 269 filas (0.05 s). Su prioridad sube si el dataset crece ≥5k. Mantenerlo como 3B.4 (como el usuario pidió) pero documentado como bajo impacto real hoy.
- **3B.5 queda casi resuelto por 3B.1** (19.9→0.13 s automáticamente): el trabajo residual es reestilizar sin reconstruir + quitar `update_idletasks` + conservar estado → pasa a pulido de bajo coste.

**Orden final:**
1. **3B.1** — Neutralizar iBus/XIM por-proceso (solo MetaTag + Renombrador; validar tipeo LATAM; incluir precompilación de regex de matching).
2. **3B.2** — Virtualizar `PreviewTable` (patrón `ExcelGrid`: render solo filas visibles; conservar estilo/temas; scroll fluido).
3. **3B.3** — Índice multi-clave en `ImageMatcher` + `_find_image_ex` con oráculo de equivalencia (la línea base de matching permanece intacta hasta demostrar equivalencia).
4. **3B.4** — Vectorizar filtro por columnas (máscara OR) con equivalencia verificada.
5. **3B.5** — `_apply_rebuild`: reestilizar en sitio sin reconstruir ni re-escanear; quitar `update_idletasks` síncronos visibles; conservar estado.
6. **3B.6** — Regresión (199 tests + casos borde matching) + benchmarks ANTES/DESPUÉS + memoria, en múltiples cwd/resoluciones.

**Regla transversal**: cada optimización demostrará ANTES → CAMBIO → DESPUÉS → REGRESIÓN sobre el escenario real del usuario (no solo el benchmark sintético), conservando matching, datos, Excel, imágenes, undo, filtros, temas, Renombrador, cwd y resoluciones.
