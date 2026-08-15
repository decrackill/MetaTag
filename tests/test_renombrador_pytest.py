"""
Tests automatizados para Renombrador de Fotos v4.0
Ejecutar: python -m pytest tests/test_renombrador_pytest.py -v
"""
import os
import sys
import shutil
import tempfile
from pathlib import Path
from datetime import datetime

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import renombrar_fotos_gui as mod


@pytest.fixture
def tmp_dir():
    d = tempfile.mkdtemp()
    yield d
    shutil.rmtree(d)


@pytest.fixture
def model():
    return mod.RenameModel()


class TestRenameModel:
    def test_load_photos(self, model, tmp_dir):
        for name in ["a.jpg", "b.png", "c.txt"]:
            Path(tmp_dir, name).touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        n = model.load_photos()
        assert n == 2  # .txt no es imagen válida

    def test_load_photos_empty(self, model, tmp_dir):
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        n = model.load_photos()
        assert n == 0

    def test_build_preview_no_dups(self, model, tmp_dir):
        for name in ["a.jpg", "b.jpg"]:
            Path(tmp_dir, name).touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["X", "Y"]
        pairs = model.build_preview()
        assert len(pairs) == 2
        assert all(not p[3] for p in pairs)  # no duplicates

    def test_build_preview_with_dups(self, model, tmp_dir):
        for name in ["a.jpg", "b.jpg", "c.jpg"]:
            Path(tmp_dir, name).touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["X", "X", "Y"]
        pairs = model.build_preview()
        dups = [p for p in pairs if p[3]]
        assert len(dups) == 1
        assert dups[0][1] == "X.jpg"

    def test_rename_all(self, model, tmp_dir):
        for name in ["a.jpg", "b.jpg", "c.jpg"]:
            Path(tmp_dir, name).touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["Uno", "Dos", "Tres"]
        errors = []
        model.rename_all(lambda c, t, n: None, lambda ok, e: errors.extend(e))
        files = sorted(f.name for f in Path(tmp_dir).iterdir())
        assert files == ["Dos.jpg", "Tres.jpg", "Uno.jpg"]
        assert len(errors) == 0

    def test_rename_skips_dups(self, model, tmp_dir):
        for name in ["a.jpg", "b.jpg"]:
            Path(tmp_dir, name).touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["X", "X"]
        errors = []
        model.rename_all(lambda c, t, n: None, lambda ok, e: errors.extend(e))
        dup_errors = [e for e in errors if "duplicado" in e.lower()]
        assert len(dup_errors) == 1

    def test_rename_conflict(self, model, tmp_dir):
        for name in ["a.jpg"]:
            Path(tmp_dir, name).touch()
        Path(tmp_dir, "X.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["X"]
        errors = []
        model.rename_all(lambda c, t, n: None, lambda ok, e: errors.extend(e))
        assert any("existe" in e.lower() for e in errors)

    def test_undo(self, model, tmp_dir):
        for name in ["a.jpg", "b.jpg"]:
            Path(tmp_dir, name).touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["X", "Y"]
        model.rename_all(lambda c, t, n: None, lambda ok, e: None)
        model.undo_last(lambda c, t, n: None, lambda ok, e: None)
        files = sorted(f.name for f in Path(tmp_dir).iterdir())
        assert "a.jpg" in files
        assert "b.jpg" in files

    def test_set_name(self, model):
        model._names = ["A", "B", "C"]
        model.set_name(1, "Z")
        assert model.names[1] == "Z"

    def test_set_name_out_of_range(self, model):
        model._names = ["A"]
        model.set_name(99, "Z")
        assert model.names == ["A"]

    # ── FASE A: seguridad del modelo ──────────────────────────────────────

    def test_ya_correcto_plan_state(self, model, tmp_dir):
        """Si el destino ES el mismo archivo → estado 'ya_correcto', sin error."""
        Path(tmp_dir, "foto.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.load_photos()
        model._names = ["foto"]
        pairs = model.build_preview()
        assert pairs[0][4] == "ya_correcto"
        errors = []
        model.rename_all(lambda c, t, n: None, lambda ok, e: errors.extend(e))
        assert errors == []
        # el archivo sigue existiendo y con el mismo nombre
        assert (Path(tmp_dir) / "foto.jpg").exists()

    def test_ya_correcto_rename_noop(self, model, tmp_dir):
        """rename_all no toca archivos ya correctos."""
        Path(tmp_dir, "a.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.load_photos()
        model._names = ["a"]
        ok = []
        model.rename_all(lambda c, t, n: None, lambda k, e: ok.append(k))
        assert ok == [1]

    def test_conflict_state_skipped(self, model, tmp_dir):
        """'existe' = el destino es un archivo EXTERNO que ya existe y NO
        pertenece al lote → jamás se sobreescribe (bloquea)."""
        Path(tmp_dir, "a.jpg").touch()
        Path(tmp_dir, "X.jpg").write_bytes(b"otro contenido")
        model.folder_path = Path(tmp_dir)
        model.load_photos()
        model._names = ["X"]
        pairs = model.build_preview()
        assert pairs[0][4] == "existe"
        errors = []
        model.rename_all(lambda c, t, n: None, lambda ok, e: errors.extend(e))
        assert any("existe" in e.lower() for e in errors)
        # el archivo de conflicto NO se sobreescribió ni se borró
        assert (Path(tmp_dir) / "X.jpg").read_bytes() == b"otro contenido"
        assert (Path(tmp_dir) / "a.jpg").exists()

    def test_copy_mode_conflict_does_not_overwrite(self, model, tmp_dir):
        """Modo copia también respeta conflictos (no sobreescribe)."""
        Path(tmp_dir, "a.jpg").touch()
        Path(tmp_dir, "Renombradas").mkdir()
        (Path(tmp_dir) / "Renombradas" / "X.jpg").write_bytes(b"existente")
        model.folder_path = Path(tmp_dir)
        model.load_photos()
        model._names = ["X"]
        errors = []
        model.rename_all(lambda c, t, n: None, lambda ok, e: errors.extend(e),
                         copy_mode=True)
        assert any("existe" in e.lower() for e in errors)
        assert (Path(tmp_dir) / "Renombradas" / "X.jpg").read_bytes() == b"existente"

    def test_duplicate_name_skipped_with_error(self, model, tmp_dir):
        """Dos filas con el mismo nombre destino → la segunda se omite."""
        Path(tmp_dir, "a.jpg").touch()
        Path(tmp_dir, "b.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.load_photos()
        model._names = ["X", "X"]
        errors = []
        model.rename_all(lambda c, t, n: None, lambda ok, e: errors.extend(e))
        assert any("duplicado" in e.lower() for e in errors)
        files = sorted(f.name for f in Path(tmp_dir).iterdir())
        assert files == ["X.jpg", "b.jpg"]

    def test_not_found_never_renames(self, model, tmp_dir):
        """Un nombre sin fotografía en modo matching seguro NO toca nada."""
        Path(tmp_dir, "a.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.load_photos()
        model.matching_mode = True
        model._names = ["inexistente"]
        errors = []
        ok = []
        model.rename_all(lambda c, t, n: None, lambda k, e: (ok.append(k), errors.extend(e)))
        assert ok == [0]
        assert any("no se encontr" in e.lower() for e in errors)
        assert (Path(tmp_dir) / "a.jpg").exists()

    def test_undo_does_not_overwrite_new_file(self, model, tmp_dir):
        """Deshacer NO sobreescribe un archivo creado después del rename."""
        Path(tmp_dir, "a.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.load_photos()
        model._names = ["X"]
        model.rename_all(lambda c, t, n: None, lambda ok, e: None)
        # alguien crea un archivo en la ruta original ANTES de deshacer
        Path(tmp_dir, "a.jpg").write_bytes(b"nuevo")
        errors = []
        ok = []
        model.undo_last(lambda c, t, n: None, lambda k, e: (ok.append(k), errors.extend(e)))
        assert ok == [0]
        assert any("conflicto" in e.lower() or "sobreescribe" in e.lower()
                   for e in errors)
        assert (Path(tmp_dir) / "a.jpg").read_bytes() == b"nuevo"
        assert (Path(tmp_dir) / "X.jpg").exists()

    def test_plan_never_marks_ok_when_conflict(self, model, tmp_dir):
        """El plan no puede mentir: un destino externo que ya existe NO puede
        marcarse "ok" — queda 'existe' (bloqueante) en la preview."""
        Path(tmp_dir, "a.jpg").touch()
        Path(tmp_dir, "X.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.load_photos()
        model._names = ["X"]
        pairs = model.build_preview()
        state = pairs[0][4]
        assert state == "existe"

    def test_matching_mode_plan_states(self, model, tmp_dir):
        """Modo matching seguro: not_found / ya_correcto / ok coherentes."""
        Path(tmp_dir, "juan perez.jpg").touch()
        Path(tmp_dir, "maria garcia.png").touch()
        model.folder_path = Path(tmp_dir)
        model.load_photos()
        model.matching_mode = True
        model._names = ["juan perez", "maria garcia", "desconocido"]
        pairs = model.build_preview()
        states = [p[4] for p in pairs]
        assert states == ["ya_correcto", "ya_correcto", "not_found"]

    def test_matching_mode_reuso_marca_duplicado(self, model, tmp_dir):
        """Dos filas que emparejan la MISMA foto → la 2ª es duplicado (reuso)."""
        Path(tmp_dir, "juan perez.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.load_photos()
        model.matching_mode = True
        model._names = ["juan perez", "juan perez"]
        states = [p[4] for p in model.build_preview()]
        assert states == ["ya_correcto", "duplicado"]

    def test_matching_mode_ambiguo_nunca_elige(self, model, tmp_dir):
        """Clave id-suffix compartida → estado 'ambiguo', no se elige nada."""
        Path(tmp_dir, "0053_EC_RS_372_F.jpg").touch()
        Path(tmp_dir, "0053_EC_C7_XII_372_F.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.load_photos()
        model.matching_mode = True
        model._names = ["0053_EC_XX_999_F"]
        states = [p[4] for p in model.build_preview()]
        assert states == ["ambiguo"]

    def test_matching_mode_motor_no_disponible_es_error(self, model, tmp_dir):
        """Motor ausente + matching ON → ERROR (nunca fallback posicional)."""
        Path(tmp_dir, "a.jpg").touch()
        Path(tmp_dir, "b.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.load_photos()
        model.matching_mode = True
        model._names = ["X", "Y"]
        # simula que metatag_matching no pudo importarse
        saved, mod.ImageMatcher = mod.ImageMatcher, None
        try:
            assert not model.matching_available
            states = [p[4] for p in model.build_preview()]
            assert states == ["error", "error"]
            # y rename_all NO renombra nada por posición
            errors = []
            ok = []
            model.rename_all(lambda c, t, n: None,
                             lambda k, e: (ok.append(k), errors.extend(e)))
            assert ok == [0]
            assert any("error" in e.lower() for e in errors)
            files = sorted(f.name for f in Path(tmp_dir).iterdir())
            assert files == ["a.jpg", "b.jpg"]  # nada se tocó
        finally:
            mod.ImageMatcher = saved

    def test_matching_mode_sin_carpeta_es_error(self, model):
        """Matching ON sin carpeta seleccionada → ERROR, no posicional."""
        model.matching_mode = True
        model._names = ["X"]
        assert model.folder_path is None
        states = [p[4] for p in model.build_preview()]
        assert states == ["error"]


class TestSortOptions:
    def test_natural_sort(self, model, tmp_dir):
        for name in ["foto10.jpg", "foto2.jpg", "foto1.jpg", "foto20.jpg"]:
            Path(tmp_dir, name).touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        n = model.load_photos()
        names = [p.name for p in model.photos]
        assert names == ["foto1.jpg", "foto2.jpg", "foto10.jpg", "foto20.jpg"]

    def test_name_sort(self, model, tmp_dir):
        for name in ["c.jpg", "a.jpg", "b.jpg"]:
            Path(tmp_dir, name).touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "name_asc"
        model.load_photos()
        names = [p.name for p in model.photos]
        assert names == ["a.jpg", "b.jpg", "c.jpg"]

    def test_mtime_sort(self, model, tmp_dir):
        for name in ["a.jpg", "b.jpg", "c.jpg"]:
            p = Path(tmp_dir, name)
            p.touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "mtime_asc"
        model.load_photos()
        assert len(model.photos) == 3


class TestCSVSupport:
    def test_csv_load(self, model, tmp_dir):
        import pandas as pd
        csv_file = Path(tmp_dir, "nombres.csv")
        pd.DataFrame({"nombre": ["A", "B", "C"]}).to_csv(csv_file, index=False)
        model.excel_path = csv_file
        model.column_name = "nombre"
        n = model.load_names()
        assert n == 3
        assert model.names == ["A", "B", "C"]

    def test_tsv_load(self, model, tmp_dir):
        import pandas as pd
        tsv_file = Path(tmp_dir, "nombres.tsv")
        pd.DataFrame({"col": ["X", "Y"]}).to_csv(tsv_file, sep="\t", index=False)
        model.excel_path = tsv_file
        model.column_name = "col"
        n = model.load_names()
        assert n == 2


class TestExport:
    def test_export_csv(self, model, tmp_dir):
        for name in ["a.jpg", "b.jpg"]:
            Path(tmp_dir, name).touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["X", "Y"]
        pairs = model.build_preview()
        csv_path = Path(tmp_dir, "out.csv")
        model.export_preview_csv(pairs, csv_path)
        import csv
        with open(csv_path) as f:
            rows = list(csv.reader(f))
        assert rows[0] == ["indice", "estado", "original", "nuevo_nombre",
                           "ruta_original", "ruta_destino", "resultado"]
        assert len(rows) == 3

    def test_export_log(self, model, tmp_dir):
        for name in ["a.jpg"]:
            Path(tmp_dir, name).touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["X"]
        pairs = model.build_preview()
        log_path = Path(tmp_dir, "out.log")
        model.export_log(pairs, log_path)
        content = log_path.read_text()
        assert "LOG DE RENOMBRAMIENTO" in content


class TestUtilities:
    def test_natural_key(self):
        files = [Path("foto10.jpg"), Path("foto2.jpg"), Path("foto1.jpg")]
        sorted_files = sorted(files, key=mod._natural_key)
        assert [f.name for f in sorted_files] == ["foto1.jpg", "foto2.jpg", "foto10.jpg"]

    def test_safe_cancel_after(self):
        import tkinter as _tk
        root = _tk.Tk()
        root.withdraw()
        job = root.after(10000, lambda: None)
        mod._safe_cancel_after(root, job)
        mod._safe_cancel_after(root, None)
        mod._safe_cancel_after(root, "invalid")
        root.destroy()

    def test_themes_are_canonical_metatag(self):
        # El Renombrador usa los 3 temas canónicos de MetaTag (sin modo claro).
        assert mod.THEME_ORDER == [
            "Arqueológico (Oscuro Refinado)", "Noche Total", "Carbón",
        ]
        assert mod.DEFAULT_THEME == "Arqueológico (Oscuro Refinado)"
        assert set(mod.THEME_ORDER) == set(mod.mt.THEMES)

    def test_C_is_default_theme_palette(self):
        assert mod.CURRENT_THEME == mod.DEFAULT_THEME
        assert mod.C == mod._THEME_ADAPTER.palette(mod.CURRENT_THEME)

    def test_palette_has_ctk_schema(self):
        p = mod._THEME_ADAPTER.palette(mod.DEFAULT_THEME)
        for key in ("bg", "surface", "surface2", "surface3", "accent", "accent2",
                    "green", "red", "yellow", "text", "subtext", "border",
                    "overlay", "accent_text", "dup_bg", "state_bg", "state_fg"):
            assert key in p, f"falta {key}"

    def test_button_constants(self):
        assert "fg_color" in mod.BTN_SECONDARY
        assert "hover_color" in mod.BTN_PRIMARY
        assert "fg_color" in mod.BTN_DANGER
        assert mod.BTN_PRIMARY["fg_color"] == mod.C["accent"]


class TestSortByDate:
    def test_has_creation_date_options(self):
        assert "Fecha creación ↑" in mod.SORT_OPTIONS
        assert "Fecha creación ↓" in mod.SORT_OPTIONS

    def test_has_exif_date_options(self):
        assert "Fecha foto ↑" in mod.SORT_OPTIONS
        assert "Fecha foto ↓" in mod.SORT_OPTIONS

    def test_get_exif_date_returns_float(self, tmp_dir):
        img_path = Path(tmp_dir, "test.jpg")
        img_path.touch()
        result = mod._get_exif_date(img_path)
        assert isinstance(result, float)


class TestThemeChange:
    def _make_ctrl(self):
        import tkinter as _tk
        root = _tk.Tk()
        root.withdraw()
        return root, mod.AppController()

    def _other_theme(self):
        return next(t for t in mod.THEME_ORDER if t != mod.CURRENT_THEME)

    def test_apply_theme_switches_palette_and_rebuilds(self):
        root, ctrl = self._make_ctrl()
        try:
            target = self._other_theme()
            ctrl._apply_theme(target)
            assert mod.CURRENT_THEME == target
            assert mod.C == mod._THEME_ADAPTER.palette(target)
            assert ctrl._view._theme_var.get() == target
            assert mod.BTN_PRIMARY["fg_color"] == mod.C["accent"]
        finally:
            ctrl._view.destroy()
            root.destroy()

    def test_apply_theme_unknown_falls_back_to_default(self):
        root, ctrl = self._make_ctrl()
        try:
            # "highcontrast" era el tema antiguo; ahora se normaliza al default.
            ctrl._apply_theme("highcontrast")
            assert mod.CURRENT_THEME == mod.DEFAULT_THEME
            assert mod.C == mod._THEME_ADAPTER.palette(mod.DEFAULT_THEME)
        finally:
            ctrl._view.destroy()
            root.destroy()

    def test_apply_theme_preserves_preview_after_rebuild(self):
        root, ctrl = self._make_ctrl()
        try:
            ctrl._last_pairs = [("a.jpg", "A.jpg", None, False, "ok")]
            target = self._other_theme()
            ctrl._apply_theme(target)
            assert ctrl._view._preview._rows, "la vista previa debe re-renderizarse"
        finally:
            ctrl._view.destroy()
            root.destroy()

    def test_apply_theme_preserves_user_selections(self):
        root, ctrl = self._make_ctrl()
        try:
            ctrl._view._folder_sel.set("/tmp/ejemplo")
            ctrl._view._sort_var.set("Fecha creación ↑")
            target = self._other_theme()
            ctrl._apply_theme(target)
            assert ctrl._view._folder_sel.get() == "/tmp/ejemplo"
            assert ctrl._view._sort_var.get() == "Fecha creación ↑"
        finally:
            ctrl._view.destroy()
            root.destroy()


class TestDetectDrives:
    def test_never_raises(self):
        result = mod._detect_drives()
        assert isinstance(result, list)

    def test_returns_at_least_one(self):
        result = mod._detect_drives()
        assert len(result) >= 1


class TestExifWithRealImage:
    def test_exif_date_from_image(self, tmp_dir):
        from PIL import Image
        from PIL.ExifTags import Base as ExifBase
        img_path = Path(tmp_dir) / "test.jpg"
        img = Image.new("RGB", (100, 100), "red")
        img.save(img_path)
        result = mod._get_exif_date(img_path)
        assert isinstance(result, float)

    def test_exif_date_no_exif(self, tmp_dir):
        from PIL import Image
        img_path = Path(tmp_dir) / "noexif.jpg"
        img = Image.new("RGB", (100, 100), "blue")
        img.save(img_path)
        result = mod._get_exif_date(img_path)
        assert result == 0.0


class TestUpdateDupStates:
    def test_method_exists(self):
        assert hasattr(mod.PreviewTable, "update_dup_states")

    def test_accepts_pairs_format(self):
        import inspect
        sig = inspect.signature(mod.PreviewTable.update_dup_states)
        params = list(sig.parameters.keys())
        assert "pairs" in params


class TestNormalizeExcelValue:
    def test_none_and_empty(self):
        assert mod.RenameModel._normalize_excel_value(None) is None
        assert mod.RenameModel._normalize_excel_value("") is None
        assert mod.RenameModel._normalize_excel_value("  ") is None

    def test_nan_variants(self):
        assert mod.RenameModel._normalize_excel_value("nan") is None
        assert mod.RenameModel._normalize_excel_value("NAT") is None
        assert mod.RenameModel._normalize_excel_value("none") is None
        assert mod.RenameModel._normalize_excel_value(float("nan")) is None

    def test_floats_never_show_decimal(self):
        assert mod.RenameModel._normalize_excel_value(1.0) == "1"
        assert mod.RenameModel._normalize_excel_value("2.00") == "2"

    def test_keeps_leading_zeros(self):
        assert mod.RenameModel._normalize_excel_value("001") == "001"
        assert mod.RenameModel._normalize_excel_value("0006") == "0006"


class TestLoadNamesNormalization:
    def _make_excel(self, tmp_dir, values):
        import pandas as pd
        df = pd.DataFrame({"Nombre": values})
        path = Path(tmp_dir) / "nombres.xlsx"
        df.to_excel(path, index=False)
        return path

    def test_load_names_keeps_zeros_and_skips_empties(self, model, tmp_dir):
        model.excel_path = self._make_excel(
            tmp_dir, ["001", "", "0006", 1.0, None, "002"])
        model.column_name = "Nombre"
        n = model.load_names()
        assert n == 4
        assert model.names == ["001", "0006", "1", "002"]
        assert model.skipped_rows == [3, 6]

    def test_cache_reads_excel_once(self, model, tmp_dir):
        import pandas as pd
        df = pd.DataFrame({"Nombre": ["a", "b"]})
        path = Path(tmp_dir) / "nombres.xlsx"
        df.to_excel(path, index=False)
        model.excel_path = path
        model.column_name = "Nombre"
        model.load_names()
        assert model._df is not None
        key = (str(model.excel_path), model.sheet_name or 0)
        assert model._df_key == key
        n = model.load_names()
        assert n == 2


class TestSinFotoPreview:
    def test_build_preview_marks_sin_foto(self, model, tmp_dir):
        Path(tmp_dir, "a.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["uno", "dos"]
        pairs = model.build_preview()
        assert len(pairs) == 2
        states = [p[4] for p in pairs]
        assert states[0] == "ok"
        assert states[1] == "sin_foto"
        # la fila sin foto tiene src None y deja la foto intacta
        assert pairs[1][2] is None
        assert (Path(tmp_dir) / "a.jpg").exists()

    def test_rename_skips_sin_foto_without_touching(self, model, tmp_dir):
        Path(tmp_dir, "a.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["uno", "dos"]
        errors = []
        ok = []
        model.rename_all(lambda c, t, n: None, lambda k, e: (ok.append(k), errors.extend(e)))
        assert ok == [1]
        assert any("sin fotografía" in e for e in errors)
        assert sorted(f.name for f in Path(tmp_dir).iterdir()) == ["uno.jpg"]


class TestTwoPhaseSwap:
    def test_swap_a_b_b_c(self, model, tmp_dir):
        """A→B mientras B→C: la FASE 1 (temporales) libera B antes de la FASE 2."""
        Path(tmp_dir, "a.jpg").touch()
        Path(tmp_dir, "b.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["b", "c"]
        pairs = model.build_preview()
        states = [p[4] for p in pairs]
        assert states == ["ok", "ok"], states
        ok = []
        errors = []
        model.rename_all(lambda c, t, n: None, lambda k, e: (ok.append(k), errors.extend(e)))
        assert ok == [2]
        assert errors == [], errors
        files = sorted(f.name for f in Path(tmp_dir).iterdir())
        assert files == ["b.jpg", "c.jpg"]
        # no quedan temporales huérfanos
        assert not [f for f in Path(tmp_dir).iterdir() if f.name.startswith(".metatag_tmp_")]

    def test_swap_undo_restores(self, model, tmp_dir):
        Path(tmp_dir, "a.jpg").touch()
        Path(tmp_dir, "b.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["b", "c"]
        model.rename_all(lambda c, t, n: None, lambda ok, e: None)
        model.undo_last(lambda c, t, n: None, lambda ok, e: None)
        files = sorted(f.name for f in Path(tmp_dir).iterdir())
        assert files == ["a.jpg", "b.jpg"]

    def test_cadena_larga_toda_resoluble(self, model, tmp_dir):
        """A→B, B→C, C→D, D→E con E libre: toda la cadena se resuelve
        en el mismo lote (la FASE 1 libera los destinos en cascada)."""
        for f in ("a", "b", "c", "d", "e"):
            Path(tmp_dir, f"{f}.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["b", "c", "d", "e", "f"]  # f.jpg no existe → libre
        pairs = model.build_preview()
        states = [p[4] for p in pairs]
        # ninguna fila puede quedar bloqueada: es una cadena pura que
        # termina en un destino libre.
        assert states == ["ok"] * 5, states
        ok = []
        model.rename_all(lambda c, t, n: None, lambda k, e: ok.append(k))
        assert ok == [5]
        files = sorted(f.name for f in Path(tmp_dir).iterdir())
        assert files == ["b.jpg", "c.jpg", "d.jpg", "e.jpg", "f.jpg"]

    def test_ciclo_a_b_y_b_a_resoluble(self, model, tmp_dir):
        """A→B y B→A: ciclo puro, ambas son 'ok' (renombrado en dos fases)."""
        Path(tmp_dir, "a.jpg").touch()
        Path(tmp_dir, "b.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["b", "a"]
        pairs = model.build_preview()
        assert [p[4] for p in pairs] == ["ok", "ok"], pairs

    def test_ciclo_a_b_c_a_resoluble(self, model, tmp_dir):
        """A→B, B→C, C→A: ciclo de 3 nodos, todas 'ok'."""
        for f in ("a", "b", "c"):
            Path(tmp_dir, f"{f}.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["b", "c", "a"]
        pairs = model.build_preview()
        assert [p[4] for p in pairs] == ["ok", "ok", "ok"], pairs

    def test_cadena_hacia_externo_queda_existe(self, model, tmp_dir):
        """A→B y B→X donde X es un archivo EXTERNO (no se renombra): la
        cadena NO es resoluble; B→X es 'existe' y A→B queda 'conflicto'
        (su destino B no se moverá). Ninguno se marca 'ok'."""
        Path(tmp_dir, "a.jpg").touch()
        Path(tmp_dir, "b.jpg").touch()
        Path(tmp_dir, "x.jpg").write_bytes(b"externo")
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["b", "x"]
        pairs = model.build_preview()
        assert [p[4] for p in pairs] == ["conflicto", "existe"], pairs
        errors = []
        model.rename_all(lambda c, t, n: None, lambda k, e: errors.extend(e))
        # nada se renombra ni se toca el archivo externo
        assert (Path(tmp_dir) / "x.jpg").read_bytes() == b"externo"
        assert any("ya existe" in e.lower() for e in errors)
        assert any("conflicto" in e.lower() for e in errors)

    def test_destino_externo_aislado_existe(self, model, tmp_dir):
        """Un único destino externo existente → 'existe' (bloquea)."""
        Path(tmp_dir, "a.jpg").touch()
        Path(tmp_dir, "x.jpg").write_bytes(b"externo")
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["x"]
        pairs = model.build_preview()
        assert pairs[0][4] == "existe"

    def test_ya_correcto_nunca_existe(self, model, tmp_dir):
        """src == destino → 'ya_correcto', jamás 'existe' (no es colisión)."""
        Path(tmp_dir, "foto.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["foto"]
        pairs = model.build_preview()
        assert pairs[0][4] == "ya_correcto"


class TestRenameBlocked:
    def test_blocked_when_count_mismatch_posicional(self, model, tmp_dir):
        Path(tmp_dir, "a.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["uno", "dos"]
        blocked, reason = model.rename_blocked()
        assert blocked is True
        assert "conteo" in reason

    def test_blocked_on_conflict(self, model, tmp_dir):
        Path(tmp_dir, "a.jpg").touch()
        Path(tmp_dir, "X.jpg").write_bytes(b"otro")
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model._photos = [Path(tmp_dir, "a.jpg")]   # conteo 1↔1, solo a.jpg
        model._names = ["X"]
        blocked, reason = model.rename_blocked()
        assert blocked is True
        assert "existe" in reason

    def test_not_blocked_when_ok(self, model, tmp_dir):
        Path(tmp_dir, "a.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model._names = ["uno"]
        blocked, reason = model.rename_blocked()
        assert blocked is False
        assert reason == ""

    def test_matching_not_found_does_not_block(self, model, tmp_dir):
        """Matching seguro: sin coincidencia → se omite, no bloquea."""
        Path(tmp_dir, "a.jpg").touch()
        model.folder_path = Path(tmp_dir)
        model.sort_mode = "natural"
        model.load_photos()
        model.matching_mode = True
        model._names = ["inexistente"]
        blocked, reason = model.rename_blocked()
        assert blocked is False
        assert reason == ""

    def test_blocked_when_no_data(self, model):
        blocked, reason = model.rename_blocked()
        assert blocked is True
        assert "fotografías" in reason or "nombres" in reason
