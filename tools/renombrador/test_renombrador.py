"""
Tests automatizados para Renombrador de Fotos v4.0
Ejecutar: python -m pytest test_renombrador.py -v
"""
import os
import shutil
import tempfile
from pathlib import Path
from datetime import datetime

import pytest
import sys
sys.path.insert(0, os.path.dirname(__file__))

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
        assert rows[0] == ["original", "nuevo_nombre", "duplicado"]
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

    def test_themes_exist(self):
        assert "dark" in mod.PALETTES
        assert "light" in mod.PALETTES
        assert "highcontrast" in mod.PALETTES

    def test_button_constants(self):
        assert "fg_color" in mod.BTN_SECONDARY
        assert "hover_color" in mod.BTN_PRIMARY
        assert "fg_color" in mod.BTN_DANGER


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


class TestHighContrastBypass:
    def test_bypass_returns_highcontrast_palette(self):
        mod._high_contrast = True
        try:
            assert mod._current_palette() == mod.PALETTES["highcontrast"]
        finally:
            mod._high_contrast = False

    def test_bypass_disabled_uses_ctk(self):
        mod._high_contrast = False
        palette = mod._current_palette()
        assert palette in [mod.PALETTES["dark"], mod.PALETTES["light"]]

    def test_apply_theme_sets_flag(self):
        import tkinter as _tk
        root = _tk.Tk()
        root.withdraw()
        ctrl = mod.AppController()
        ctrl._apply_theme("highcontrast")
        assert mod._high_contrast == True
        ctrl._apply_theme("dark")
        assert mod._high_contrast == False
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
