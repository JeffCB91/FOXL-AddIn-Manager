"""Tests for core/zip_ops.py"""
import os
import zipfile
import pytest
from unittest.mock import patch

from core.zip_ops import find_xll_in_folder, extract_and_install_zip


class TestFindXllInFolder:
    def test_returns_false_for_file_path(self, tmp_path):
        f = tmp_path / "not_a_folder.txt"
        f.write_text("x")
        success, msg = find_xll_in_folder(str(f))
        assert success is False
        assert "not a valid folder" in msg.lower()

    def test_returns_false_for_nonexistent_path(self, tmp_path):
        missing = str(tmp_path / "missing_dir")
        success, msg = find_xll_in_folder(missing)
        assert success is False

    def test_finds_xll_in_root(self, tmp_path):
        xll = tmp_path / "addin.xll"
        xll.write_bytes(b"\x00")
        success, path = find_xll_in_folder(str(tmp_path))
        assert success is True
        assert path.endswith("addin.xll")

    def test_finds_xll_in_subdirectory(self, tmp_path):
        sub = tmp_path / "sub"
        sub.mkdir()
        xll = sub / "deep.xll"
        xll.write_bytes(b"\x00")
        success, path = find_xll_in_folder(str(tmp_path))
        assert success is True
        assert path.endswith("deep.xll")

    def test_returns_false_when_no_xll(self, tmp_path):
        (tmp_path / "readme.txt").write_text("x")
        (tmp_path / "data.json").write_text("{}")
        success, msg = find_xll_in_folder(str(tmp_path))
        assert success is False
        assert "no .xll" in msg.lower()

    def test_empty_directory(self, tmp_path):
        success, msg = find_xll_in_folder(str(tmp_path))
        assert success is False


class TestExtractAndInstallZip:
    def _make_zip(self, tmp_path, name, contents):
        """Helper: create a zip at tmp_path/name with the given {filename: bytes} contents."""
        zip_path = tmp_path / name
        with zipfile.ZipFile(zip_path, 'w') as zf:
            for fname, data in contents.items():
                zf.writestr(fname, data)
        return str(zip_path)

    def test_success_extracts_xll(self, tmp_path):
        zip_path = self._make_zip(tmp_path, "pkg.zip", {
            "addin.xll": b"xll_data",
            "NinetyOne.ExcelAddIn.config.json": b"{}",
        })
        install_dir = tmp_path / "install"
        with patch("core.zip_ops.BASE_LOCAL_PATH", str(install_dir)):
            success, msg, xll_path = extract_and_install_zip(zip_path, "v1.0")
        assert success is True
        assert xll_path is not None
        assert xll_path.endswith("addin.xll")
        assert os.path.exists(xll_path)

    def test_success_copies_config_json(self, tmp_path):
        zip_path = self._make_zip(tmp_path, "pkg.zip", {
            "addin.xll": b"xll_data",
            "NinetyOne.ExcelAddIn.config.json": b'{"key":"val"}',
        })
        install_dir = tmp_path / "install"
        with patch("core.zip_ops.BASE_LOCAL_PATH", str(install_dir)):
            success, msg, xll_path = extract_and_install_zip(zip_path, "v1.0")
        assert success is True
        config_path = os.path.join(os.path.dirname(xll_path), "NinetyOne.ExcelAddIn.config.json")
        assert os.path.exists(config_path)

    def test_no_json_is_fine(self, tmp_path):
        zip_path = self._make_zip(tmp_path, "pkg.zip", {"addin.xll": b"x"})
        install_dir = tmp_path / "install"
        with patch("core.zip_ops.BASE_LOCAL_PATH", str(install_dir)):
            success, msg, xll_path = extract_and_install_zip(zip_path, "v2.0")
        assert success is True
        assert xll_path.endswith("addin.xll")

    def test_bad_zip_file(self, tmp_path):
        bad_zip = tmp_path / "bad.zip"
        bad_zip.write_bytes(b"this is not a zip")
        install_dir = tmp_path / "install"
        with patch("core.zip_ops.BASE_LOCAL_PATH", str(install_dir)):
            success, msg, xll_path = extract_and_install_zip(str(bad_zip), "v1.0")
        assert success is False
        assert "not a valid" in msg.lower()
        assert xll_path is None

    def test_no_xll_in_zip(self, tmp_path):
        zip_path = self._make_zip(tmp_path, "pkg.zip", {"readme.txt": b"hello"})
        install_dir = tmp_path / "install"
        with patch("core.zip_ops.BASE_LOCAL_PATH", str(install_dir)):
            success, msg, xll_path = extract_and_install_zip(zip_path, "v1.0")
        assert success is False
        assert ".xll" in msg
        assert xll_path is None

    def test_xll_in_subdirectory_of_zip(self, tmp_path):
        zip_path = self._make_zip(tmp_path, "pkg.zip", {
            "subdir/addin.xll": b"xll_data",
        })
        install_dir = tmp_path / "install"
        with patch("core.zip_ops.BASE_LOCAL_PATH", str(install_dir)):
            success, msg, xll_path = extract_and_install_zip(zip_path, "v1.0")
        assert success is True
        assert xll_path.endswith("addin.xll")

    def test_creates_target_subdirectory(self, tmp_path):
        zip_path = self._make_zip(tmp_path, "pkg.zip", {"addin.xll": b"x"})
        install_dir = tmp_path / "install"
        with patch("core.zip_ops.BASE_LOCAL_PATH", str(install_dir)):
            extract_and_install_zip(zip_path, "v3.0")
        assert (install_dir / "v3.0").is_dir()

