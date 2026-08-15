"""Tests del Renombrador de Fotos incorporado en src/renombrar_fotos_gui.py.

Capa compatible con el sistema de pruebas del proyecto (unittest).
Se prueban los escenarios de renombrado de RenameModel usando SIEMPRE
directorios temporales (nunca se tocan fotografías reales ni el Excel).

El test original de la herramienta (tests/test_renombrador_pytest.py)
usa pytest y se conserva como test propio de la herramienta; esta capa
cubre los mismos escenarios bajo `python -m unittest discover -s tests`.
"""
import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import renombrar_fotos_gui as mod


class _Callbacks:
    """Recolecta las llamadas on_progress/on_done de RenameModel."""

    def __init__(self):
        self.done_called = None
        self.progress = []

    def on_progress(self, cur, total, name):
        self.progress.append((cur, total, name))

    def on_done(self, success, errors):
        self.done_called = (success, list(errors))


def _make_model(folder, files, names, sort="natural"):
    model = mod.RenameModel()
    model.folder_path = Path(folder)
    for f in files:
        (Path(folder) / f).touch()
    model.load_photos()
    model._names = list(names)
    return model


class RenameModelBaseTestCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.folder = self._tmp.name

    def _run_rename(self, model, **kw):
        cb = _Callbacks()
        model.rename_all(cb.on_progress, cb.on_done, **kw)
        self.assertIsNotNone(cb.done_called)
        return cb.done_called

    def _files(self):
        return sorted(p.name for p in Path(self.folder).iterdir() if p.is_file())


class RenameEscenariosTestCase(RenameModelBaseTestCase):
    def test_nombres_simples(self):
        model = _make_model(self.folder, ["a.jpg", "b.jpg"], ["alfa", "beta"])
        ok, errors = self._run_rename(model)
        self.assertEqual((ok, errors), (2, []))
        self.assertEqual(self._files(), ["alfa.jpg", "beta.jpg"])

    def test_extensiones_diferentes(self):
        model = _make_model(self.folder, ["a.jpg", "b.png", "c.tiff"],
                            ["X", "Y", "Z"])
        ok, errors = self._run_rename(model)
        self.assertEqual((ok, errors), (3, []))
        self.assertEqual(self._files(), ["X.jpg", "Y.png", "Z.tiff"])

    def test_nombres_con_ceros(self):
        model = _make_model(self.folder, ["foto1.jpg", "foto10.jpg"],
                            ["0006", "0007"])
        ok, errors = self._run_rename(model)
        self.assertEqual((ok, errors), (2, []))
        self.assertEqual(self._files(), ["0006.jpg", "0007.jpg"])

    def test_archivos_faltantes(self):
        model = _make_model(self.folder, ["a.jpg", "b.jpg"],
                            ["uno", "dos", "tres"])
        ok, errors = self._run_rename(model)
        self.assertEqual(ok, 2)
        self.assertTrue(any("sin fotografía" in e for e in errors),
                        "el registro sin foto se omite explícitamente (no en silencio)")
        self.assertEqual(self._files(), ["dos.jpg", "uno.jpg"])

    def test_archivos_adicionales(self):
        model = _make_model(self.folder, ["a.jpg", "b.jpg", "c.jpg"],
                            ["uno", "dos"])
        ok, errors = self._run_rename(model)
        self.assertEqual((ok, errors), (2, []))
        self.assertEqual(self._files(), ["c.jpg", "dos.jpg", "uno.jpg"])

    def test_duplicados(self):
        model = _make_model(self.folder, ["a.jpg", "b.jpg"], ["dup", "dup"])
        ok, errors = self._run_rename(model)
        self.assertEqual(ok, 1, "el segundo duplicado se salta")
        self.assertTrue(any("duplicado" in e for e in errors))
        self.assertEqual(self._files(), ["b.jpg", "dup.jpg"])

    def test_conflicto_destino_existente(self):
        # m.jpg quiere llamarse "x" pero x.jpg YA existe (otro archivo) →
        # conflicto real: no se sobreescribe, se salta esa fila.
        model = _make_model(self.folder, ["a.jpg", "b.jpg", "m.jpg", "x.jpg"],
                            ["a1", "a2", "x"])
        ok, errors = self._run_rename(model)
        self.assertEqual(ok, 2)
        self.assertTrue(any("ya existe" in e for e in errors))
        self.assertEqual(self._files(), ["a1.jpg", "a2.jpg", "m.jpg", "x.jpg"])

    def test_nombres_ya_correctos(self):
        # alfa.jpg ya se llama "alfa" → NO es conflicto: es "ya correcto".
        model = _make_model(self.folder, ["alfa.jpg", "beta.jpg"],
                            ["alfa", "beta"])
        ok, errors = self._run_rename(model)
        self.assertEqual(ok, 2, "ya correctos se procesan sin error")
        self.assertEqual(errors, [])
        self.assertEqual(self._files(), ["alfa.jpg", "beta.jpg"])

    def test_conflicto_vs_ya_correcto_sin_falsa_alarma(self):
        # Una foto cuyo nombre coincide con el destino ES su propio archivo;
        # se considera "ya correcto", jamás un conflicto destructivo.
        model = _make_model(self.folder, ["foto.jpg"], ["foto"])
        ok, errors = self._run_rename(model)
        self.assertEqual((ok, errors), (1, []))
        self.assertEqual(self._files(), ["foto.jpg"])

    def test_cancelacion_detiene_lote(self):
        import threading
        model = _make_model(self.folder, ["a.jpg", "b.jpg", "c.jpg"],
                            ["x1", "x2", "x3"])
        cancel_ev = threading.Event()
        cb = _Callbacks()
        cancel_ev.set()
        model.rename_all(cb.on_progress, cb.on_done, cancel_ev=cancel_ev)
        self.assertEqual(cb.done_called[0], 0)
        self.assertTrue(any("Cancelado" in e for e in cb.done_called[1]))

    def test_undo_restaura(self):
        model = _make_model(self.folder, ["a.jpg", "b.jpg"], ["alfa", "beta"])
        self._run_rename(model)
        self.assertTrue(model.has_undo)
        cb = _Callbacks()
        model.undo_last(cb.on_progress, cb.on_done)
        self.assertEqual(cb.done_called, (2, []))
        self.assertFalse(model.has_undo)
        self.assertEqual(self._files(), ["a.jpg", "b.jpg"])

    def test_build_preview_marca_duplicados(self):
        model = _make_model(self.folder, ["a.jpg", "b.jpg"], ["dup", "dup"])
        pairs = model.build_preview()
        self.assertEqual([p[3] for p in pairs], [False, True])


if __name__ == "__main__":
    unittest.main()
