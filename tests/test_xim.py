"""Tests de regresión del módulo ``metatag_xim`` (FASE 3B.1).

Verifican que:
  - ``neutralize_xim_for_tk`` solo actúa cuando ``XMODIFIERS`` referencia
    iBus (el único caso demostrado patológico de Tk↔XIM).
  - Con cualquier otro valor (otro IM, vacío, sin configurar) no toca nada.
  - Es idempotente y no lanza excepciones ante entornos raros.
  - El ajuste es por-proceso (modifica solo ``os.environ`` del proceso).
"""

import os
from unittest.mock import patch

import pytest

sys_path_done = False


def _import_metatag_xim():
    global sys_path_done
    import sys
    from pathlib import Path
    if not sys_path_done:
        sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
        sys_path_done = True
    import metatag_xim
    return metatag_xim


def test_neutraliza_cuando_es_ibus():
    mod = _import_metatag_xim()
    with patch.dict(os.environ, {"XMODIFIERS": "@im=ibus"}, clear=False):
        assert mod.neutralize_xim_for_tk() is True
        assert os.environ["XMODIFIERS"] == "@im=none"


def test_ignora_otro_im():
    mod = _import_metatag_xim()
    with patch.dict(os.environ, {"XMODIFIERS": "@im=fcitx"}, clear=False):
        assert mod.neutralize_xim_for_tk() is False
        assert os.environ["XMODIFIERS"] == "@im=fcitx"


def test_ignora_vacio():
    mod = _import_metatag_xim()
    with patch.dict(os.environ, {"XMODIFIERS": ""}, clear=False):
        assert mod.neutralize_xim_for_tk() is False
        assert os.environ["XMODIFIERS"] == ""


def test_ignora_ausente():
    mod = _import_metatag_xim()
    with patch.dict(os.environ, {}, clear=True):
        assert mod.neutralize_xim_for_tk() is False
        assert "XMODIFIERS" not in os.environ


def test_idempotente():
    mod = _import_metatag_xim()
    with patch.dict(os.environ, {"XMODIFIERS": "@im=ibus"}, clear=False):
        assert mod.neutralize_xim_for_tk() is True
        assert mod.neutralize_xim_for_tk() is False
        assert os.environ["XMODIFIERS"] == "@im=none"


def test_insensible_a_mayusculas():
    mod = _import_metatag_xim()
    with patch.dict(os.environ, {"XMODIFIERS": "@im=IBUS"}, clear=False):
        assert mod.neutralize_xim_for_tk() is True
        assert os.environ["XMODIFIERS"] == "@im=none"


def test_no_rompe_con_valor_malformado():
    mod = _import_metatag_xim()
    with patch.dict(os.environ, {"XMODIFIERS": "@im=ibus,extra=1"}, clear=False):
        assert mod.neutralize_xim_for_tk() is True
        assert os.environ["XMODIFIERS"] == "@im=none"


def test_cambio_es_local_al_proceso():
    """El entorno original debe restaurarse al terminar el patched block."""
    mod = _import_metatag_xim()
    original = os.environ.get("XMODIFIERS")
    try:
        os.environ["XMODIFIERS"] = "@im=ibus"
        assert mod.neutralize_xim_for_tk() is True
    finally:
        if original is None:
            os.environ.pop("XMODIFIERS", None)
        else:
            os.environ["XMODIFIERS"] = original
    assert os.environ.get("XMODIFIERS") == original
