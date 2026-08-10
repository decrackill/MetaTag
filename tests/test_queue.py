"""Bloque 3 — robustez de _process_queue / _process_all.

Prueba la lógica de procesamiento y cola SIN depender de un display real,
usando FakeApp (subclase de MetaTagApp sin Tk). FakeApp instala "trampas"
que lanzan AssertionError si el worker toca estado de Tkinter (grid, vars),
de modo que cualquier acceso a Tk desde el worker hace fallar la prueba.

Casos cubiertos:
  1. worker exitoso;
  2. archivo inexistente;
  3. matching ambiguo;
  4. excepción dentro del worker;
  5. cancelación a mitad del procesamiento;
  6. worker que muere sin emitir done/error (terminación inesperada);
  7. cola vacía mientras el worker sigue vivo (continúa polling);
  8-10. restauración de UI tras éxito / error / cancelación;
  11. el worker NO invoca métodos/variables de Tkinter;
  12. sin pollings infinitos tras la muerte del worker;
  guard. no iniciar un segundo procesamiento mientras uno está activo.

Además, si hay display Tk funcional, un smoke test real de la UI.
"""
import os
import queue
import shutil
import sys
import tempfile
import threading
import time
import unittest
from unittest import mock

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import metatag_v8
from metatag_v8 import MetaTagApp


class Recorder:
    def __init__(self):
        self.values = []

    def set(self, v):
        self.values.append(v)


class FakeBtn:
    def __init__(self):
        self.text = None
        self.state = None
        self.visible = False

    def configure(self, text=None, state=None):
        if text is not None:
            self.text = text
        if state is not None:
            self.state = state

    def pack(self, **kw):
        self.visible = True

    def pack_forget(self):
        self.visible = False


class FakeMB:
    def __init__(self):
        self.info = []
        self.error = []
        self.warn = []

    def showinfo(self, title, msg):
        self.info.append((title, msg))

    def showerror(self, title, msg):
        self.error.append((title, msg))

    def showwarning(self, title, msg):
        self.warn.append((title, msg))


class TrapVar:
    def get(self):
        raise AssertionError("Variable Tk accedida desde el worker")

    def set(self, v):
        raise AssertionError("Variable Tk escrita desde el worker")


class TrapGrid:
    @property
    def selected_cells(self):
        raise AssertionError("grid.selected_cells accedido desde el worker")

    @property
    def df(self):
        raise AssertionError("grid.df accedido desde el worker")


class FakeApp(MetaTagApp):
    """Subclase sin Tk: instala los atributos que usa la cola + trampas."""

    def __init__(self):
        self.progress_queue = queue.Queue()
        self.progress_var = Recorder()
        self.status_var = Recorder()
        self.logs = []
        self.after_calls = []
        self.cursor = None
        self._proc_thread = None
        self._proc_cancel = threading.Event()
        self._write_btn = FakeBtn()
        self._cancel_btn = FakeBtn()
        self._img_cache = {}
        self.output_folder = None
        self.grid = TrapGrid()
        self.img_col_var = TrapVar()
        self.meta_mode_organized = TrapVar()
        self.omit_empty_var = TrapVar()

    def _log_safe(self, msg, tag=""):
        self.logs.append((msg, tag))

    def config(self, **kw):
        self.cursor = kw.get("cursor")

    def after(self, ms, cb):
        self.after_calls.append(cb)
        return 1


def make_file(dirpath, name):
    with open(os.path.join(dirpath, name), "wb") as f:
        f.write(b"fake-image-bytes")
    return os.path.join(dirpath, name)


def drain(app):
    out = []
    while True:
        try:
            out.append(app.progress_queue.get_nowait())
        except queue.Empty:
            return out


def apply_msgs(app, msgs):
    """Aplica los mensajes no terminales como hace _process_queue, para que
    los log/status/progress queden registrados en FakeApp."""
    for mtype, val in msgs:
        if mtype == "status":
            app.status_var.set(val)
        elif mtype == "progress":
            app.progress_var.set(val)
        elif mtype == "log":
            msg, tag = val
            app._log_safe(msg, tag)
    return msgs


def run_worker(app, folder, meta_by_row, df, img_col_idx=0,
               organizado=False, omit_empty=True, empty_cnt=0):
    t = threading.Thread(
        target=app._process_all,
        args=(folder, meta_by_row, df, img_col_idx, organizado, omit_empty, empty_cnt),
        daemon=True)
    t.start()
    t.join()
    return t


def terminals(msgs):
    return [m for m in msgs if m[0] in ("done", "error", "cancelled")]


class WorkerTestCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="metatag_q_")
        self.imgdir = os.path.join(self.tmp, "img")
        os.makedirs(self.imgdir)
        self.outdir = os.path.join(self.tmp, "out")
        os.makedirs(self.outdir)
        self.app = FakeApp()
        self.app.output_folder = __import__("pathlib").Path(self.outdir)
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)

    def _df(self, names):
        return pd.DataFrame({"img": names})

    # ── 1. worker exitoso ──────────────────────────────────────────
    def test_worker_exitoso_emite_done(self):
        make_file(self.imgdir, "0001_A_F.jpg")
        self.app._check_metadata_divergence = lambda *a: []
        self.app._write_meta = lambda *a: None
        run_worker(self.app, self.imgdir, {0: {"note": "x"}}, self._df(["0001_A_F.jpg"]))
        msgs = drain(self.app)
        term = terminals(msgs)
        self.assertEqual(len(term), 1)
        self.assertEqual(term[0][0], "done")
        self.assertIn("1 guardadas", term[0][1])
        self.assertIn("0 errores", term[0][1])
        self.assertTrue(os.path.exists(os.path.join(self.outdir, "0001_A_F.jpg")))

    # ── 2. archivo inexistente ─────────────────────────────────────
    def test_archivo_inexistente(self):
        make_file(self.imgdir, "0001_A_F.jpg")
        self.app._check_metadata_divergence = lambda *a: []
        self.app._write_meta = lambda *a: None
        run_worker(self.app, self.imgdir, {0: {"note": "x"}}, self._df(["9999_NOPE.jpg"]))
        msgs = apply_msgs(self.app, drain(self.app))
        term = terminals(msgs)
        self.assertEqual(term[0][0], "done")
        self.assertIn("0 guardadas · 1 errores", term[0][1])
        self.assertTrue(any("No encontrada" in m for m, _ in self.app.logs))

    # ── 3. matching ambiguo ────────────────────────────────────────
    def test_matching_ambiguo_se_cuenta_como_ambiguedad(self):
        make_file(self.imgdir, "0053_EC_RS_372_F.jpg.JPG")
        make_file(self.imgdir, "0053_EC_C7_XII_372_F.jpg.JPG")
        self.app._check_metadata_divergence = lambda *a: []
        self.app._write_meta = lambda *a: None
        run_worker(self.app, self.imgdir, {0: {"note": "x"}},
                   self._df(["0053_OTRO_CAMPO_1_F.jpg"]))
        msgs = apply_msgs(self.app, drain(self.app))
        term = terminals(msgs)
        self.assertEqual(term[0][0], "done")
        self.assertIn("· 1 ambigüedades", term[0][1])
        self.assertTrue(any("Ambigua" in m for m, _ in self.app.logs))

    # ── 4. excepción dentro del worker ─────────────────────────────
    def test_excepcion_del_worker_emite_error(self):
        self.app._find_img_name_in_row_data = lambda row: (_ for _ in ()).throw(
            RuntimeError("boom interno"))

        def run():
            run_worker(self.app, self.imgdir, {0: {"note": "x"}}, self._df(["0001_A_F.jpg"]))

        # El worker NO debe morir silenciosamente: debe emitir ("error", ...).
        run()
        msgs = apply_msgs(self.app, drain(self.app))
        term = terminals(msgs)
        self.assertEqual(len(term), 1)
        self.assertEqual(term[0][0], "error")
        self.assertNotEqual(term[0][0], "done")
        self.assertTrue(any("Error interno" in m for m, _ in self.app.logs))

    # ── 5. cancelación ─────────────────────────────────────────────
    def test_cancelacion_antes_de_empezar(self):
        self.app._proc_cancel.set()
        run_worker(self.app, self.imgdir, {0: {"note": "x"}}, self._df(["0001_A_F.jpg"]))
        msgs = drain(self.app)
        term = terminals(msgs)
        self.assertEqual(term[0][0], "cancelled")
        self.assertNotIn("done", [m[0] for m in term])

    def test_cancelacion_a_mitad_de_procesamiento(self):
        make_file(self.imgdir, "0001_A_F.jpg")
        make_file(self.imgdir, "0002_A_F.jpg")
        self.app._check_metadata_divergence = lambda *a: []
        self.app._write_meta = lambda path, meta, organizado: self.app._proc_cancel.set()
        run_worker(self.app, self.imgdir,
                   {0: {"note": "x"}, 1: {"note": "y"}},
                   self._df(["0001_A_F.jpg", "0002_A_F.jpg"]))
        msgs = drain(self.app)
        term = terminals(msgs)
        self.assertEqual(term[0][0], "cancelled")
        self.assertIn("1 escritas", term[0][1])  # solo la primera fila

    # ── 11. worker NO toca Tkinter ─────────────────────────────────
    def test_worker_no_accede_a_tkinter(self):
        # FakeApp expone trampas: cualquier acceso a grid/vars Tk lanza
        # AssertionError dentro del worker → caería en "error". Un "done"
        # limpio demuestra que el worker no tocó Tk.
        make_file(self.imgdir, "0001_A_F.jpg")
        self.app._check_metadata_divergence = lambda *a: []
        self.app._write_meta = lambda *a: None
        run_worker(self.app, self.imgdir, {0: {"note": "x"}}, self._df(["0001_A_F.jpg"]))
        msgs = drain(self.app)
        term = terminals(msgs)
        self.assertEqual(term[0][0], "done")


class QueueTestCase(unittest.TestCase):
    def setUp(self):
        self.app = FakeApp()
        # messagebox real abre un modal y colgaría las pruebas que llegan a
        # _proc_finish_ui("done"/"error"); lo sustituimos por un registro.
        self.mb = FakeMB()
        self._mb_patch = mock.patch.object(metatag_v8, "messagebox", self.mb)
        self._mb_patch.start()
        self.addCleanup(self._mb_patch.stop)

    def _drain_terminal(self, terminal_msg_type, terminal_msg):
        # Simula el bucle real: _process_queue drena la cola y termina.
        self.app._process_queue()

    # ── 8. restauración tras éxito ─────────────────────────────────
    def test_restauracion_ui_despues_de_done(self):
        self.app.progress_queue.put(("done", "✔ 1 guardadas · 0 errores"))
        self.app._process_queue()
        self.assertEqual(self.app.status_var.values[-1], "✔ 1 guardadas · 0 errores")
        self.assertEqual(self.app.cursor, "")
        self.assertEqual(self.app._write_btn.text, "▶  Escribir Metadatos")
        self.assertEqual(self.app._write_btn.state, "normal")
        self.assertFalse(self.app._cancel_btn.visible)
        self.assertIsNone(self.app._proc_thread)
        self.assertIsNone(self.app._proc_cancel)
        self.assertEqual(len(self.mb.info), 1)
        self.assertEqual(self.app.after_calls, [])

    # ── 9. restauración tras error ─────────────────────────────────
    def test_restauracion_ui_despues_de_error(self):
        self.app.progress_queue.put(("log", ("detalle\n", "err")))
        self.app.progress_queue.put(("error", "Error interno"))
        self.app._process_queue()
        self.assertEqual(self.app.status_var.values[-1], "Error interno")
        self.assertEqual(self.app.cursor, "")
        self.assertEqual(self.app._write_btn.text, "▶  Escribir Metadatos")
        self.assertEqual(self.app._write_btn.state, "normal")
        self.assertFalse(self.app._cancel_btn.visible)
        self.assertIsNone(self.app._proc_thread)
        self.assertEqual(len(self.mb.error), 1)
        self.assertEqual(self.app.after_calls, [])

    # ── 10. restauración tras cancelación ──────────────────────────
    def test_restauracion_ui_despues_de_cancelacion(self):
        self.app.progress_queue.put(("cancelled", "Procesamiento cancelado"))
        self.app._process_queue()
        self.assertEqual(self.app.status_var.values[-1], "Procesamiento cancelado")
        self.assertEqual(self.app.cursor, "")
        self.assertEqual(self.app._write_btn.text, "▶  Escribir Metadatos")
        self.assertEqual(self.app._write_btn.state, "normal")
        self.assertFalse(self.app._cancel_btn.visible)
        self.assertIsNone(self.app._proc_thread)
        self.assertEqual(len(self.mb.info), 0)  # cancelación: sin messagebox
        self.assertTrue(any("cancelado" in m.lower() for m, _ in self.app.logs))
        self.assertEqual(self.app.after_calls, [])

    # ── 7. cola vacía con worker vivo → continúa polling ───────────
    def test_cola_vacia_con_worker_vivo_continua_polling(self):
        hold = threading.Event()

        def worker():
            hold.wait()
            self.app.progress_queue.put(("done", "✔ listo"))

        self.app._proc_thread = threading.Thread(target=worker, daemon=True)
        self.app._proc_thread.start()

        self.app._process_queue()  # cola vacía, worker vivo
        self.assertEqual(len(self.app.after_calls), 1)  # reprograma

        hold.set()
        self.app._proc_thread.join()

        self.app._process_queue()  # drena done
        self.assertEqual(self.app.status_var.values[-1], "✔ listo")
        self.assertEqual(len(self.app.after_calls), 1)  # NO reprograma

    # ── 6. worker muerto sin done/error → terminación inesperada ───
    def test_worker_muerto_sin_senal_terminal(self):
        hold = threading.Event()

        def worker():
            hold.wait()  # muere sin emitir nada

        self.app._proc_thread = threading.Thread(target=worker, daemon=True)
        self.app._proc_thread.start()

        self.app._process_queue()          # cola vacía + vivo → polling
        self.assertEqual(len(self.app.after_calls), 1)

        hold.set()
        self.app._proc_thread.join()       # ahora está muerto

        self.app._process_queue()          # cola vacía + muerto → detección
        self.assertEqual(self.app.status_var.values[-1],
                         "El proceso terminó de forma inesperada.")
        self.assertEqual(self.app.cursor, "")
        self.assertEqual(self.app._write_btn.text, "▶  Escribir Metadatos")
        self.assertEqual(self.app._write_btn.state, "normal")
        self.assertFalse(self.app._cancel_btn.visible)
        self.assertIsNone(self.app._proc_thread)
        self.assertTrue(any("inesperada" in m for m, _ in self.app.logs))

    # ── 12. sin pollings infinitos tras la muerte del worker ────────
    def test_no_hay_polling_infinito_despues_de_la_muerte(self):
        hold = threading.Event()

        def worker():
            hold.wait()

        self.app._proc_thread = threading.Thread(target=worker, daemon=True)
        self.app._proc_thread.start()

        self.app._process_queue()
        hold.set()
        self.app._proc_thread.join()
        self.app._process_queue()          # detecta muerte
        before = len(self.app.after_calls)
        self.app._process_queue()          # un segundo intento: no debe reprogramar
        self.assertEqual(len(self.app.after_calls), before)
        self.assertEqual(self.app._proc_thread, None)

    # ── guard: no iniciar otro procesamiento mientras uno está vivo ─
    def test_no_inicia_segundo_procesamiento_con_worker_vivo(self):
        hold = threading.Event()

        def worker():
            hold.wait()

        t = threading.Thread(target=worker, daemon=True)
        t.start()
        self.app._proc_thread = t
        self.app._start_processing()
        self.assertEqual(len(self.mb.warn), 1)
        self.assertIs(self.app._proc_thread, t)  # sin reemplazar el thread
        hold.set()
        t.join()


def _tk_ok():
    try:
        import tkinter as tk
        root = tk.Tk()
        root.destroy()
        return True
    except Exception:
        return False


@unittest.skipUnless(_tk_ok(), "sin display Tk funcional — smoke pendiente de validación manual")
class ProcessingSmokeTestCase(unittest.TestCase):
    """Smoke test real de la UI (requiere display). Usa datos temporales."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="metatag_smoke_")

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def _make_app(self):
        mb = FakeMB()
        patchers = [
            mock.patch.object(metatag_v8, "messagebox", mb),
            mock.patch.object(MetaTagApp, "_save_config", lambda self: None),
            mock.patch.object(MetaTagApp, "_load_config_pre_build", lambda self: None),
            mock.patch.object(MetaTagApp, "_load_config_post_build", lambda self: None),
        ]
        for p in patchers:
            p.start()
        self.addCleanup(lambda: [p.stop() for p in patchers])

        app = MetaTagApp()
        app.withdraw()
        self.addCleanup(app.destroy)
        app._mb = mb
        return app

    def _wait_finished(self, app, timeout=15):
        deadline = time.time() + timeout
        while time.time() < deadline:
            app.update()
            if app._proc_thread is None:
                return True
            time.sleep(0.05)
        return False

    def test_smoke_procesamiento_normal_y_restauracion(self):
        imgdir = os.path.join(self.tmp, "img1")
        outdir = os.path.join(self.tmp, "out1")
        os.makedirs(imgdir)
        make_file(imgdir, "0001_A_F.jpg")
        make_file(imgdir, "0002_A_F.jpg")

        app = self._make_app()
        app.output_folder = __import__("pathlib").Path(outdir)
        app.img_folder_var.set(imgdir)
        app.img_col_var.set("img")
        df = pd.DataFrame({"img": ["0001_A_F.jpg", "0002_A_F.jpg"],
                           "nota": ["nota A", "nota B"]})
        app.grid.load(df)
        app.grid.selected_cells = {(0, 1), (1, 1)}
        app._check_metadata_divergence = lambda *a: []
        app._write_meta = lambda *a: None

        app._start_processing()
        # Cancelar debe estar visible mientras procesa.
        self.assertEqual(app._cancel_btn.winfo_manager(), "pack")
        self.assertTrue(self._wait_finished(app))
        self.assertEqual(app._cancel_btn.winfo_manager(), "")
        self.assertEqual(app._write_btn.cget("state"), "normal")
        self.assertEqual(app._write_btn.cget("text"), "▶  Escribir Metadatos")
        self.assertEqual(app.cget("cursor"), "")
        self.assertEqual(len(app._mb.info), 1)
        self.assertIn("2 guardadas", app._mb.info[0][1])

    def test_smoke_cancelacion_y_restauracion(self):
        imgdir = os.path.join(self.tmp, "img2")
        outdir = os.path.join(self.tmp, "out2")
        os.makedirs(imgdir)
        for i in range(6):
            make_file(imgdir, f"000{i}_A_F.jpg")

        app = self._make_app()
        app.output_folder = __import__("pathlib").Path(outdir)
        app.img_folder_var.set(imgdir)
        app.img_col_var.set("img")
        df = pd.DataFrame({"img": [f"000{i}_A_F.jpg" for i in range(6)],
                           "nota": [f"nota {i}" for i in range(6)]})
        app.grid.load(df)
        app.grid.selected_cells = {(r, 1) for r in range(6)}
        app._check_metadata_divergence = lambda *a: []
        # La escritura de la primera fila se bloquea hasta que se cancele:
        # garantiza que la cancelación llega a mitad del procesamiento.
        app._write_meta = lambda path, meta, organizado: app._proc_cancel.wait()

        app._start_processing()
        self.assertEqual(app._cancel_btn.winfo_manager(), "pack")
        app.update()
        app._cancel_processing()
        self.assertTrue(self._wait_finished(app))
        self.assertEqual(app._cancel_btn.winfo_manager(), "")
        self.assertEqual(app._write_btn.cget("state"), "normal")
        self.assertEqual(app._write_btn.cget("text"), "▶  Escribir Metadatos")
        self.assertIn("cancelado", app.status_var.get().lower())
        self.assertEqual(len(app._mb.info), 0)  # cancelación: sin messagebox

    def test_smoke_error_y_restauracion(self):
        imgdir = os.path.join(self.tmp, "img3")
        outdir = os.path.join(self.tmp, "out3")
        os.makedirs(imgdir)
        make_file(imgdir, "0001_A_F.jpg")

        app = self._make_app()
        app.output_folder = __import__("pathlib").Path(outdir)
        app.img_folder_var.set(imgdir)
        app.img_col_var.set("img")
        df = pd.DataFrame({"img": ["0001_A_F.jpg"], "nota": ["nota A"]})
        app.grid.load(df)
        app.grid.selected_cells = {(0, 1)}
        app._check_metadata_divergence = lambda *a: []
        app._write_meta = lambda *a: None
        app._find_image_ex = lambda *a: (_ for _ in ()).throw(RuntimeError("boom"))

        app._start_processing()
        self.assertTrue(self._wait_finished(app))
        self.assertEqual(app._cancel_btn.winfo_manager(), "")
        self.assertEqual(app._write_btn.cget("state"), "normal")
        self.assertEqual(app._write_btn.cget("text"), "▶  Escribir Metadatos")
        self.assertEqual(app.cget("cursor"), "")
        self.assertEqual(len(app._mb.error), 1)
        self.assertIn("Error interno", app.status_var.get())


if __name__ == "__main__":
    unittest.main()
