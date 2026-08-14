"""
MetaTag XIM — Neutralización local (por-proceso) del método de entrada XIM.

Problema medido (FASE 3A-R.1): cuando el sistema usa iBus como método de
entrada X11 (``XMODIFIERS=@im=ibus``), Tk crea un input context XIM y cada
creación/redraw de widget sincroniza con ``ibus`` (~20-90 ms/widget en el
equipo de desarrollo). Startup de MetaTag: ~10-20 s; cambio de tema: ~20 s.

Solución: desactivar SOLO la asociación XIM de Tk dentro del proceso, y solo
cuando el método de entrada activo es iBus (el único caso demostrado
patológico). No se toca GTK/Qt (usan GTK_IM_MODULE/QT_IM_MODULE), no se
detiene/desconfigura iBus como servicio, no se modifica el layout del teclado
ni configuración global. La composición de acentos LATAM se resuelve a nivel
de layout X11 (teclas muertas), independiente de XIM.

La variable debe quedar asignada ANTES de la primera creación de ``Tk()`` en
el proceso (XIM se abre en ``Tk()``, no en ``import tkinter``). Los procesos
hijo lanzados con ``subprocess.Popen`` heredan el entorno modificado.
"""

import os


def neutralize_xim_for_tk() -> bool:
    """Neutraliza la asociación XIM de Tk si apunta a iBus.

    Solo actúa cuando ``XMODIFIERS`` referencia ``ibus``; en cualquier otro
    caso (otro IM, valor vacío o no configurado) no toca nada y es idempotente.

    Returns:
        True si se modificó el entorno del proceso; False en caso contrario.
    """
    current = os.environ.get("XMODIFIERS", "")
    if current and "ibus" in current.lower():
        os.environ["XMODIFIERS"] = "@im=none"
        return True
    return False
