# FASE 3A-R.3 — Diseño técnico detallado de las optimizaciones 3B.1–3B.5

Estado: completada el 2026-08-13. **SOLO DISEÑO — no se modificó código, no se ejecutaron migraciones, no se cambió configuración persistente.**

Base de evidencia: FASE 3A-R.1 (`docs/FASE_3A-R1_rendimiento.md`), FASE 3A-R.2 (`docs/FASE_3A-R2_clasificacion.md`).

Regla del diseño: cada optimización debe tener **problema medido → causa raíz → mecanismo exacto → comportamiento a preservar → riesgo de regresión → rollback → tests → benchmark ANTES/DESPUÉS → criterio de aceptación**. No hay optimizaciones "porque parecen más rápidas".

---

## 0. Auditoría de procesos y recursos del sistema (solo lectura, 2026-08-13)

| Proceso | PID | CPU | RSS | Observación |
|---|---|---|---|---|
| `ibus-daemon --xim` | 1142 | 39.4 % acum (5 h) / **~0 % en reposo** / **100 % de un núcleo durante flush XIM** | 501 MB | **Causa demostrada** de la penalización Tk↔XIM |
| `ibus-x11 --kill-daemon` | 1172 | 0.8 % | 128 MB | correlación (parte del stack iBus) |
| `Xorg` | 928 | ~3.9 % | 24 MB | normal |
| `cinnamon --replace` | 1400 | ~2.7 % | 72 MB | normal |
| VSCode/pylance (dev) | 46776… | variable | — | herramienta de desarrollo; swap ~620 MB |
| Brave | 38027… | variable | — | navegador; swap ~290 MB |
| opencode (esta sesión) | 47159 | picos 170–200 % | — | herramienta de desarrollo |

- **Swap**: 2375 MB usados de 4095 MB (58 %) — consumido por VSCode, Brave, opencode y cinnamon. **Ningún proceso de MetaTag** (no hay instancias corriendo; los benchmarks usan ≤ 135 MB RSS y no pagan swap).
- **Docker**: no hay demonio/containers activos.
- **Procesos huérfanos de MetaTag/Renombrador**: ninguno (verificado).
- **Load average**: 0.9–2.3 (moderado). **CPU**: 8 cores, ocioso salvo picos.

**Clasificación (causa vs correlación vs sospecha):**
- **Causa demostrada**: Tk abre XIM (`XMODIFIERS=@im=ibus`) y cada redraw/creación de widget dispara interacción con `ibus` → ~20–90 ms/widget. El 39 % acumulado de CPU de iBus incluye los episodios de profiling de esta auditoría.
- **Correlación**: `ibus-x11`, `ibus-ui-gtk3` están cargados en memoria pero no muestran CPU relevante fuera de los flush.
- **Sospecha (sin evidencia de acción)**: el compositor Cinnamon y Xorg muestran lecturas altas solo como artefacto de `top` redibujando; no se les imputa el cuello de botella.

**No se realizó ninguna acción sobre el sistema**: sin `kill`, sin `systemctl`, sin `apt remove`, sin cambios de config global. La optimización es **local al programa y reversible**.

---

## 1. Diseño 3B.1 — Neutralización de la penalización iBus/XIM por-proceso

### Problema medido
- Startup MetaTag: 10.93 s con XIM vs 0.33 s sin XIM. Cambio de tema: 19.9 s vs 0.13 s. 200 labels + flush: 3.72 s vs 0.037 s. (HECHO)

### Causa raíz
- Tk crea un input context XIM por conexión X y sincroniza con `ibus` por cada widget; en esta máquina ese coste es ~20–90 ms/widget.

### Verificaciones empíricas nuevas (HECHO, realizadas en esta fase)
1. `os.environ["XMODIFIERS"] = ""` **después de `import tkinter` y ANTES del primer `Tk()`** funciona: 200 labels + flush = **0.259 s** (vs 3.7 s). El XIM se abre en `Tk()`, no en el `import`.
2. `XMODIFIERS="@im=none"` igual de efectivo (0.302 s).
3. **El subproceso hereda el entorno**: `Popen` (sin `env` explícito) lanza el Renombrador con `XMODIFIERS=""` heredado (flush 0.041 s) → **el Renombrador lanzado desde MetaTag se beneficia automáticamente**.
4. `option add *useInputMethods 0` **NO funciona** (3.49 s base vs 7.10 s) → descartada (ya documentado en R.2).

### Mecanismo exacto del cambio
Nuevo módulo pequeño `src/metatag_xim.py`:

```python
import os

def neutralize_xim_for_tk() -> bool:
    """Desactiva la asociación XIM de Tk ÚNICAMENTE cuando apunta a ibus.

    No toca GTK/Qt (usan GTK_IM_MODULE/QT_IM_MODULE), no desactiva iBus
    como servicio y es reversible por proceso.
    """
    cur = os.environ.get("XMODIFIERS", "")
    if cur and "@im=ibus" in cur:
        os.environ["XMODIFIERS"] = "@im=none"
        return True
    return False
```

Se llama **antes de la primera creación de `Tk()`** en los dos entry points:
- `src/metatag_v8.py`: al inicio del bloque `if __name__ == "__main__":` (línea ~3197).
- `src/renombrar_fotos_gui.py`: al inicio del bloque `if __name__ == "__main__":` (línea ~2593).

**Conservadurismo**: solo se neutraliza si el IM es iBus (el único caso demostrado patológico). Si el usuario tuviera fcitx/xim u otro IM, no se toca nada. Reversible: al salir del proceso, el entorno del shell/escritorio no cambia (la asignación es solo del proceso).

### Comportamiento que debe preservarse
- Entrada de teclado (acentos/ñ LATAM por teclas muertas → composición a nivel de layout X11, NO dependiente de XIM).
- Otros IMs configurados (no se sobrescriben).
- El entorno del sistema/escritorio (sin cambios globales).
- Cualquier launcher/cwd: la corrección va en el entry point Python, por lo que cubre ejecución directa (`python src/metatag_v8.py`), cualquier cwd, `instalar_y_abrir.sh` y lanzadores de escritorio.

### Riesgo de regresión y rollback
- Riesgo: bajo. Peor caso: perder composición de texto dependiente de IM en Tk (solo si el usuario realmente usara iBus como IM de composición — no es el caso: layout latam con teclas muertas).
- Rollback: revertir el único cambio (borrar la llamada en los 2 entry points) — no hay cambio persistente que deshacer.

### Tests necesarios
1. `tests/test_xim.py` (nuevo, sin display): unitario de `neutralize_xim_for_tk()` — con `XMODIFIERS=@im=ibus` → cambia a `@im=none`; con fcitx → no toca; con vacío → no toca; idempotente.
2. Benchmark del flush (script de verificación): 200 labels + flush < 0.2 s (sin XIM). 
3. **Verificación manual de tipeo** (no automatizable sin xdotool): arrancar MetaTag con el fix y teclear acentos/ñ en el filtro y en un Entry. Paso obligatorio del bloque.

### Benchmark ANTES → DESPUÉS (métricas)

| Métrica | XIM actual | `XMODIFIERS=@im=none` (proceso) | Aceptación |
|---|---|---|---|
| Startup | 10.93 s | 0.33 s (medido) | < 1.0 s |
| Cambio de tema | ~19.9 s | 0.13 s (medido) | < 0.5 s |
| 200 widgets + flush | 3.72 s | 0.037–0.26 s (medido) | < 0.2 s |
| Input (entry + dead keys) | ok | ok (verificación manual) | sin pérdida de entrada |
| Preview (Renderombrador) | 3.56 s @269 | ~2.4 s @269 (ya medido) | < 3.0 s @269 |

### Criterio de aceptación
Startup < 1 s, tema < 0.5 s, flush < 0.2 s, y **tipeo LATAM verificado manualmente** sin pérdida de caracteres. **Requiere aprobación del usuario** (cambia el comportamiento de lanzamiento de la app).

---

## 2. Diseño 3B.1b — Precompilación de regex del matching

### Problema medido
- cProfile: 100 llamadas not-found × 5000 archivos → **1 500 200 `re.sub` + 1 500 200 `re._compile`**; `_normalize_numbers` 1.7 M llamadas. Los patrones van **inline** dentro de `re.sub`/`re.match`, así que Python los recompila en cada llamada.

### Inventario exacto de regex a precompilar (ambos archivos)

| Función | Patrón | Uso | Archivo |
|---|---|---|---|
| `_normalize_numbers` | `r"\d+"` | `re.sub(p, lambda m: str(int(m.group())), s)` | metatag_v8.py:2763, metatag_matching.py:65 |
| `_clean_stem` | `r"^[#\s\-_]+|[#\s\-_]+$"` | `re.sub` | metatag_matching.py:100 |
| pass "clean" de `_find_image_ex` | `r"^[#\s\-_]+|[#\s\-_]+$"` | `re.sub` (inline, duplicado) | metatag_v8.py:2831,2833,2838 |
| `_safe_stem` dup-marker | `r"^(.*?)\s*\(\d+\)$"` | `re.match` | metatag_v8.py:2753, metatag_matching.py:54 |
| `_extract_id_suffix` dup-marker | `r"^(.*?)\s*\(\d+\)$"` | `re.match` | metatag_v8.py:2782, metatag_matching.py:86 |
| `_extract_id_suffix` número | `r"^0*(\d+)"` | `re.match` | metatag_v8.py:2787, metatag_matching.py:91 |
| `_extract_id_suffix` sufijo | `r"[FRP]$"` | `re.search` | metatag_v8.py:2791, metatag_matching.py:95 |

### Mecanismo exacto
Constantes `re.Pattern` a nivel de módulo en cada archivo (p. ej. `_NUM_RE = re.compile(r"\d+")`, `_CLEAN_RE`, `_DUP_RE`, `_NUM_LEAD_RE`, `_FRP_RE`), y sustitución de cada llamada `re.sub(pat, …)` → `_NUM_RE.sub(…)`, `re.match(pat, s)` → `_DUP_RE.match(s)`, etc. **Mismo patrón, mismo repl callable, mismos flags** → semántica byte a byte idéntica. Cero estado mutable (constantes inmutables a nivel de módulo).

### Comportamiento que debe preservarse
- **El algoritmo de matching NO cambia**: ni orden de pasos, ni prioridades, ni resultados, ni la lista de candidatos. Solo se elimina la recompilación.
- `ImageMatcher` como **oráculo**: la versión precompilada debe producir `resultado_original == resultado_optimizado`.

### Equivalencia (criterio obligatorio)
```
resultado_original == resultado_optimizado
```
sobre: dataset real (269 filas × 19 col + 269 fotos), nombres con doble extensión, `(1)` dup-marker, cero-padding, bordes `#/_-`, mayúsculas, sufijos F/R/P, colisiones de stem (ambiguos), duplicados, inexistentes y cadena vacía. Comparar `(path, status, candidates)` **y el paso de resolución** vía `find_image_ex_with_method`.

### Riesgo de regresión / rollback
- Riesgo: mínimo (cambio puramente mecánico). Test de equivalencia exhaustivo antes de integrar.
- Rollback: revertir el commit (cambio acotado a constantes + sustituciones).

### Tests / benchmark / aceptación
- Nuevo test de equivalencia (`tests/test_matching_equivalence.py`) que corre el oráculo original vs optimizado sobre el corpus.
- 199 tests existentes deben pasar sin cambios.
- Aceptación: `find_image_ex` not-found a 10000: de 0.27 s/llamada a un valor sustancialmente menor (objetivo <0.15 s/llamada, PROVISIONAL hasta medir), **sin ningún caso que cambie**.

---

## 3. Diseño 3B.2 — Virtualización de PreviewTable (la pieza central)

### 3.0 Arquitectura actual verificada (HECHO)
- `PreviewTable` es un `ctk.CTkFrame` embebido en el `CTkScrollableFrame` de la vista izquierda (renombrar_fotos_gui.py:1754,1864). El **panel completo** hace scroll; todas las filas existen como widgets físicos dentro del frame interno del scrollable (CTkScrollableFrame NO virtualiza hijos).
- `render()` destruye TODOS los hijos y recrea filas en chunks de 30 (`after(5)`); cada fila = 1 CTkFrame + 4–5 hijos. Filtro = toggle `visible` + `pack/pack_forget`. `set_edit_mode` re-llama a `render()` (rebuild completo). Tooltip por fila (`<Enter>` → `ImageTooltip`). `_thumb_q`/`_schedule_thumbs` es **código muerto** (la columna de miniaturas fue eliminada; nadie rellena `_thumb_q`).
- Coste: 2.42 s @269, 9.4–13.2 s @1000, 181 MB RSS @1000, `build_preview` aparte. (HECHO)

### A. Modelo de datos (NO se virtualiza el modelo)
```
PLAN COMPLETO   = _all_pairs  (list[dict])  ← se conserva íntegro, fuente de verdad
     ↓
FILTRADO        = _filtered   (list[int])   ← índices de _all_pairs que pasan filtro
     ↓
VIEWPORT        = (first, last)             ← rango visible, derivado del scroll
     ↓
MATERIALIZADO   = pool de N slots           ← O(visible) widgets reutilizables
```
- `_all_pairs` mantiene exactamente su contrato actual (lo usan `MainView`, `AppController`, undo, duplicados…). Los estados `ok / ya_correcto / conflicto / duplicado / not_found / ambiguo / error` viven en `_all_pairs[i]["state"]` / `["is_dup"]` (igual que hoy).
- Edición: el valor vive en `_all_pairs[i]["new"]` (ya es así hoy: `_add_row` recibe el dict por referencia y escribe en él). Con virtualización sigue en el modelo; el widget es solo espejo.

### B. Scroll
- La tabla pasa a un **viewport de altura fija** (un `tk.Canvas` interno + scrollbar vertical propio) en lugar de crecer dentro del CTkScrollableFrame. El panel superior/controles siguen en el CTkScrollableFrame; el preview ocupa la altura restante con `grid(... sticky="nsew")` y peso de fila.
- `scrollregion = (0, 0, W, ROW_H * len(_filtered))` con `ROW_H = 32` (altura actual de fila).
- Del evento de scroll (mousewheel / `<Button-4/5>` / scrollbar) se deriva: `first = max(0, floor(canvas.canvasy(0)/ROW_H) - buffer)`, `last = min(len(_filtered), first + visible_rows + 2*buffer)`.
- Los widgets reciclados se colocan con `canvas.create_window(0, y, window=slot.frame, anchor="nw")` en su posición **virtual** `y = r*ROW_H` → el propio canvas los mueve/clipa al desplazarse. No se recalcula el árbol de widgets por evento de scroll; solo se re-enlazan los slots cuyo índice cambió.

### C. Pool de widgets reutilizables
- `POOL = visible_rows + 2*buffer` (p. ej. 20 + 10 = **30 slots**). Cada slot = 1 CTkFrame `height=32` con 5 hijos (label "#", label "Original", label "→", **label + entry superpuestos** para "Nuevo nombre", label "Estado").
- En cada redraw (scroll/filtro/estado): para cada slot `k`, si `row_index(first+k)` cambió → `slot.rebind(model_row)` que solo hace `configure(text=…)`, `configure(fg_color=…)`, actualiza `slot._index` y re-enlaza el `<Enter>` del tooltip. **Nunca** `destroy()`+`create()` durante scroll.
- Redraw completo de ≤30 slots < pocos ms (proyección; se medirá).

### D. Altura virtual
- La altura virtual la define `scrollregion` (virtual_height = `len(_filtered)*ROW_H`), independiente de cuántos slots existan → **sin saltos de scroll**: el scrollbar se mueve continuamente sobre la altura total aunque solo haya 30 widgets.

### E. Filtros / sort / estado
- `filter(query)`: recalcula `_filtered = [i for i,_ in enumerate(_all_pairs) if not q or q in _all_pairs[i]["orig"].lower() or q in _all_pairs[i]["new"].lower()]` (mismo criterio que hoy, línea 1556), clampea `first`, redraw. Sin crear/destruir widgets (≤30 `configure`).
- `sort`/`duplicate state`: `update_dup_states` muta los dicts del modelo y redibuja solo los slots afectados (ya hace recolor por fila; se adapta a slots).
- `set_edit_mode`: ya NO re-renderiza todo; cada slot tiene label+entry y alterna su visibilidad (`grid`/`grid_remove`).

### F. Edición inline (crítica)
- Un `ctk.StringVar` por slot. En `rebind`: `slot.var.set(model["new"])`, `slot._index = idx`, `slot.var.trace_add("write", handler)` (una sola vez por slot).
- Handler: escribe `_all_pairs[slot._index]["new"] = sanitized` (sanitización `Path(value).name` idéntica a la actual) y notifica `on_name_change(idx, safe, path)`.
- Fila editada → scroll → sale del viewport: el valor ya está **en el modelo** (trace en cada tecla). Fila → vuelve: `rebind` restaura su valor desde el modelo. **No se pierde nada.**

### G. Tooltips / miniaturas
- Tooltip: en el canvas, `<Motion>`/`<Enter>` → calcular fila bajo el cursor (`y//ROW_H`) → leer `photo_path` del **modelo en ese momento** → `ImageTooltip(path)`. El tooltip NO retiene referencias a widgets reciclados; solo una ruta.
- `_schedule_thumbs`/`_thumb_q` (muertos): se eliminan en el refactor. Si en el futuro hubiera miniaturas, será un **cache de rutas→PhotoImage acotado** (LRU) + tooltip; ninguna referencia a widgets.

### H. Estados
- Representación visual derivada **en draw time** de `state`/`is_dup` del modelo, con los mismos mapas `C["state_bg"]`, `C["state_fg"]`, `STATE_LABELS`, `_arrow_color`, `_new_color`. 7 estados conservados idénticos.

### I. Criterio de aceptación (benchmark mínimo)
Matriz a 269 / 1000 / 5000 / 10000, midiendo: render inicial, scroll (1 paso), filtro, RSS, widgets creados (total acumulado), **widgets vivos (debe ser ≈ pool ≈ 30 para TODO N)**, CPU, estabilidad tras 100/500 eventos de scroll (sin drift de posición ni fuga).

| Métrica | Actual (HECHO) | Objetivo | Tipo |
|---|---|---|---|
| Render inicial 269 | 2.42 s (sin XIM) / 3.56 s (XIM) | < 0.5 s | provisional |
| Render inicial 1000 | 9.4–13.2 s | < 1.0 s | provisional |
| Render inicial 10000 | ~90–130 s (PROYECCIÓN) | < 2.0 s | provisional |
| Scroll 1 paso @10000 | n/d (widget-per-row inservible) | < 16 ms | provisional |
| Filtro @10000 | n/d | < 50 ms | provisional |
| RSS @1000 | 181 MB | < 110 MB | provisional |
| RSS @10000 | ~+800 MB (PROYECCIÓN) | crecimiento solo del modelo (< +15 MB) | provisional |
| Widgets vivos | N×5 | ≈ 30 + header (independiente de N) | objetivo duro |

**Demostración clave**: la complejidad de render/scroll/filtro deja de depender linealmente del total de filas.

### Riesgo / rollback / tests
- Riesgo: medio (cambio de componente + layout del preview: pasa a altura fija con scroll propio). Mitigación: mantener la misma paleta/mapas de color y los mismos callbacks.
- Rollback: el commit es autocontenido (PreviewTable + su uso en MainView); revert recupera el comportamiento anterior.
- Tests: los 199 existentes (incl. `test_renombrador*.py`, `test_responsive.py`) + nuevos tests de unidad para `_filtered`/viewport/rebind sin display y pruebas de edición que simulan scroll de ida y vuelta. **Requiere aprobación** (cambio visual del área de preview).

---

## 4. Diseño 3B.3 — Matching indexado con oráculo

### 4.1 Documentación del algoritmo actual (ambos archivos idénticos)
Jerarquía de pasos de `find_image_ex` (metatag_matching.py:145 = metatag_v8.py:2795):
1. `folder/name` existe (stat) → ok direct.
2. `fpath.name.lower() == name.lower()` — primer match en orden de iteración → ok. **O(n).**
3. `name_stem in index` (dict) → ok. **O(1)** ya.
4. `_clean_stem(stem_key) == _clean_stem(name)` — primer match → ok. **O(n) con regex por archivo.**
5. `_normalize_numbers(clean(stem_key)) == …` — primer match → ok. **O(n) con 2 regex por archivo.**
6. `_extract_id_suffix(stem_key) == id_excel` — **todos** los match → 1 → ok; >1 → ambiguous. **O(n) con regex por archivo.**
7. subcadena `name_stem in stem_key or stem_key in name_stem` — todos los match → 1 → ok; >1 → ambiguous. **O(n) substring (barato).**
8. else not_found.

Operaciones O(n): pasos 2, 4, 5, 6, 7. Costosas por regex: 4, 5, 6. Caso not_found ejecuta **todos** los pasos → O(n) por llamada → O(n²) en `build_preview`.

### 4.2 Claves indexables y casos con fallback
Indexable en el build del índice (una vez por carpeta, misma iteración que hoy para preservar orden):
- `by_name_lower`: `{fpath.name.lower(): path}` — primer-wins en orden actual.
- `by_clean`: `{_clean_stem(stem): path}` — primer-wins.
- `by_normalized`: `{_normalize_numbers(_clean_stem(stem)): path}` — primer-wins.
- `by_id`: `{_extract_id_suffix(stem): [paths]}` — lista completa (para ambigüedad), mismo orden.
- `by_stem` (ya existe).

**Fallback obligatorio**: paso 7 (subcadena) no es indexable → barrido O(n) barato (sin regex), solo como último recurso.

**Falsos positivos**: el índice solo produce **candidatos**; la decisión final la toma el **algoritmo original como árbitro**:
```
índice → candidatos (posible colisión) → algoritmo original sobre candidatos → resultado definitivo
```
Si por cualquier borde el conjunto de candidatos no es concluyente (p. ej. colisión de `by_clean` entre 2 archivos y el original habría devuelto el primero), se ejecuta el **barrido completo original** para esa llamada (lento pero correcto y raro). Así, equivalencia garantizada por construcción + verificación.

**Equivalencia**: `original(photo, folder) == optimized(photo, folder)` sobre: dataset real, nombres normalizados, ceros, extensiones, mayúsculas/minúsculas, candidatos múltiples, ambiguos, inexistentes, cadena vacía y **colisiones de stem**. Comparar `(path, status, candidates)` + paso de resolución. Si no se puede demostrar equivalencia → la optimización queda **PENDIENTE**, no se implementa.

### 4.3 Beneficio esperado
- not_found @10k: 0.27 s → ~µs lookup + barrido subcadena (proyección). `build_preview` not-found @1000: 3.66 s → <0.2 s.
- El caso normal (match exacto/stem) ya es O(1) → sin cambio perceptible negativo.
- MetaTag y Renombrador usan la misma estructura; el índice se construye en el build de carpeta (tanto `_img_cache` como `_index_folder`).

### 4.4 Riesgo / rollback / tests / aceptación
- Riesgo: **alto** (matching validado). Mitigación completa: oráculo + árbitro + fallback.
- Rollback: revertir commit; el índice es aditivo (nuevas claves en el dict existente).
- Tests: `test_matching.py`, `test_metatag_matching.py` (incluyen el trazador) + nuevo test de equivalencia exhaustivo sobre corpus real + sintético + colisiones.
- Aceptación: `find_image_ex` not-found @10k < 0.05 s (objetivo), 200× @10k < 0.5 s; 0 cambios de resultado. **Requiere aprobación explícita** (toca matching).

---

## 5. Diseño 3B.4 — Filtro vectorizado → **RECOMENDACIÓN: POSPONER**

### Análisis
- Actual: `df.apply(lambda row: row.astype(str).str.contains(query, case=False, na=False).any(), axis=1)` (metatag_v8.py:1432) = 4.1 M llamadas; 1.64 s @10k.
- Diseño vectorizado equivalente: máscara OR por columna:
  ```python
  mask = pd.concat([df[c].astype(str).str.contains(q, case=False, na=False)
                    for c in df.columns], axis=1).any(axis=1)
  ```
  Semántica idéntica (mismo `astype(str)` por celda, mismo `contains`, mismo `any()`), conservando mayúsculas/minúsculas, columnas, vacíos, selección de columna (`col_b != "Todas"` usa el caso ya vectorizado) y estados.
- **Decisión según mediciones**: el dataset real es **269 filas → 0.046 s actual** (inapreciable). El beneficio solo importa ≥5k filas. Coste añadido: 1 test de equivalencia + riesgo nulo de cambio de comportamiento de filtrado. **Se pospone**; se documenta el diseño por si el dataset crece. Si se implementara: benchmark 269/1000/5000/10000 y aceptación <0.2 s @10k con máscara resultante idéntica.

---

## 6. Diseño 3B.5 — Cambio de tema → **RECOMENDACIÓN: MANTENER REBUILD (opción A)**

### Comparativa de opciones
| Opción | Correctitud | Preservación estado | Seguridad | Mantenibilidad | Rendimiento tras 3B.1 |
|---|---|---|---|---|---|
| **A. Rebuild completo** (actual) | máxima (probada) | parcial (df + carpeta; pierde selección/scroll) | máxima | máxima (código actual) | **0.13 s (HECHO)** |
| B. Reconfigurar colores/fonts de widgets existentes | riesgo: los widgets `tk.Label/Button` capturan colores en creación; no hay registro de qué slot de color usó cada widget | n/d | media | baja (bookkeeping manual) | rápido pero frágil |
| C. Híbrido (estáticos → reconfigurar; complejos → rebuild) | media | media | media | baja | — |

**Veredicto (INFERENCIA)**: con 3B.1 aplicado, el rebuild completo mide **0.13 s** (HECHO). B y C añaden bookkeeping de colores que no existe y riesgo de colores desincronizados sin beneficio medible. **Se mantiene A**, con dos micro-mejoras de bajo riesgo:
1. En `_apply_rebuild` (metatag_v8.py:709): no re-escanear la carpeta si `img_folder` no cambió (conservar `self._img_cache` en vez de `browser.load_folder` re-escaneo — o aceptar el coste actual 0.002 s @269).
2. Conservar selección y scroll del `ExcelGrid` al reconstruir (guardar/restaurar `selected_cells`, posición de scroll). Mejora de UX, no de rendimiento.
- Renombrador: su `_apply_theme` también reconstruye + re-renderiza el preview; tras 3B.2 el re-render será barato → tema rápido automáticamente.

---

## 7. Invariantes de integridad (qué NO puede cambiar)

| Área | Invariante |
|---|---|
| `RenameModel` | funcionalmente equivalente (no se toca en 3B.1–3B.5; 3B.2 solo cambia la representación visual) |
| `ImageMatcher` | resultados exactamente idénticos (oráculo) |
| Plan / estados | `ok, ya_correcto, conflicto, duplicado, not_found, ambiguo, error` intactos |
| Rename | sin aumento de riesgo de sobrescritura, pérdida de archivos, TOCTOU o undo incorrecto |
| Excel | los benchmarks nunca escriben datos reales |
| Archivos | nunca se tocan `Finales 1 a 103` ni fotografías personales; pruebas destructivas en `TemporaryDirectory()` |

---

## 8. Matriz de tests

| Bloque | Tests nuevos | Tests existentes que deben pasar |
|---|---|---|
| 3B.1 | `test_xim.py` (neutralize_xim_for_tk, casos iBus/otro/vacío/idempotencia); verificación manual de tipeo LATAM | todos |
| 3B.1b | `test_matching_equivalence.py` (oráculo vs optimizado) | `test_matching.py`, `test_metatag_matching.py`, todos |
| 3B.2 | unidad: `_filtered`, viewport/first/last, pool rebind, edición con scroll ida/vuelta (sin display) | `test_renombrador*.py`, `test_responsive.py`, todos |
| 3B.3 | equivalencia exhaustiva corpus (incl. colisiones/ambiguos) | `test_matching.py`, `test_metatag_matching.py`, todos |
| 3B.4 | (si se hace) equivalencia de máscara filtro | `test_dataset_269.py`, `test_grid.py` |
| 3B.5 | (si se hace) tema conserva selección/scroll | `test_metatag_theme.py` |

---

## 9. Benchmark baseline consolidado (ANTES → OBJETIVO)

| Operación | Actual (HECHO) | Objetivo | Justificación |
|---|---|---|---|
| MetaTag startup | 10.93 s | < 1.0 s | 3B.1; ya medido 0.33 s |
| Theme change | ~19.9 s | < 0.5 s | 3B.1; ya medido 0.13 s |
| Preview 269 | 2.4–3.6 s | < 0.5 s | 3B.2; **provisional** hasta implementar |
| Preview 1000 | 9–13 s | < 1.0 s | 3B.2; **provisional** |
| Preview 10000 | ~90–130 s / +800 MB (**PROYECCIÓN**) | < 2 s / RSS ~modelo | 3B.2; **provisional** |
| Matching worst-case | 0.27 s/llamada @10k; 3.66 s @1000 not-found | <0.05 s @10k; <0.2 s @1000 | 3B.3; **provisional** |
| RAM 1000 (preview) | 181 MB | < 110 MB | 3B.2; **provisional** |

Ningún objetivo se inventa: los de 3B.1 ya están medidos; el resto son **provisionales** a confirmar con el benchmark tras implementar.

---

## 10. Dependencias entre optimizaciones

```
3B.1 (XIM por-proceso)   ── independiente; prerrequisito de MEDICIÓN para todo lo GUI
   └→ 3B.2 (PreviewTable virtualizado)      (requiere 3B.1 para números limpios)
3B.1b (regex precompiladas) ── independiente; compartido con
   └→ 3B.3 (matching indexado)              (usa las regex ya precompiladas)
3B.4 (filtro) ── independiente; POSPUESTO
3B.5 (tema)   ── depende de 3B.1; casi resuelto por él
3B.6 (regresión + benchmark final) ── después de los anteriores
```

---

## 11. Orden exacto de implementación (commits separados)

Cada bloque: **ANTES → IMPLEMENTACIÓN → TESTS → BENCHMARK → REGRESIÓN → COMMIT**. Commits independientes, no mezclados.

1. **Commit 3B.1** — `src/metatag_xim.py` + llamadas en los 2 entry points + `tests/test_xim.py` + benchmark (startup/tema/flush) + verificación manual de tipeo.
2. **Commit 3B.1b** — regex precompiladas en `metatag_v8.py` y `metatag_matching.py` + `tests/test_matching_equivalence.py` + benchmark not_found.
3. **Commit 3B.2** — PreviewTable virtualizado (pool + canvas interno) + tests de unidad + benchmark 269/1000/5000/10000 (render/scroll/filtro/RSS/widgets vivos/CPU/estabilidad 100-500 eventos).
4. **Commit 3B.3** — índice multi-clave con árbitro+fallback + equivalencia exhaustiva + benchmark.
5. **(3B.4)** — solo si el usuario decide hacerlo (recomendado posponer).
6. **(3B.5)** — micro-mejoras de `_apply_rebuild` (conservar selección/scroll) si se decide.
7. **Commit 3B.6** — regresión completa (199+ tests) + benchmark final consolidado ANTES/DESPUÉS + memoria, en múltiples cwd/resoluciones.

---

## 12. Decisiones que requieren aprobación del usuario

1. **3B.1**: aplicar `XMODIFIERS=@im=none` automáticamente en los entry points cuando el IM sea iBus (cambia el entorno de lanzamiento de la app; reversible, por-proceso, sin tocar el sistema).
2. **3B.2**: cambiar el área de preview a altura fija con scroll propio (cambio visual del componente).
3. **3B.3**: implementar el índice de matching (aun con oráculo + árbitro + fallback) — el usuario pidió detenerse y pedir aprobación antes de tocar matching/renombrado.
4. **3B.4**: confirmar el aplazamiento (o autorizar su implementación).
5. **3B.5**: confirmar mantener rebuild completo (opción A) en lugar de reestilizado.
