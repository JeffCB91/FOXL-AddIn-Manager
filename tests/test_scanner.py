"""Tests for core/scanner.py"""
import os
import pytest
from unittest.mock import patch

from core.scanner import scan_path_sync


class TestScanPathSync:
    def test_path_not_found(self, tmp_path):
        missing = str(tmp_path / "nonexistent")
        success, result = scan_path_sync(missing)
        assert success is False
        assert result == "Path not found"

    def test_empty_directory(self, tmp_path):
        success, versions = scan_path_sync(str(tmp_path))
        assert success is True
        assert versions == []

    def test_returns_only_directories(self, tmp_path):
        (tmp_path / "v1.0").mkdir()
        (tmp_path / "v2.0").mkdir()
        (tmp_path / "file.txt").write_text("x")
        success, versions = scan_path_sync(str(tmp_path))
        assert success is True
        assert "v1.0" in versions
        assert "v2.0" in versions
        assert "file.txt" not in versions

    def test_sorted_in_reverse_order(self, tmp_path):
        (tmp_path / "1.0.0").mkdir()
        (tmp_path / "2.0.0").mkdir()
        (tmp_path / "1.5.0").mkdir()
        success, versions = scan_path_sync(str(tmp_path))
        assert success is True
        assert versions == sorted(["1.0.0", "2.0.0", "1.5.0"], reverse=True)

    def test_single_directory(self, tmp_path):
        (tmp_path / "v3.0").mkdir()
        success, versions = scan_path_sync(str(tmp_path))
        assert success is True
        assert versions == ["v3.0"]

    def test_returns_error_on_exception(self, tmp_path):
        with patch("os.path.exists", return_value=True):
            with patch("os.listdir", side_effect=PermissionError("denied")):
                success, result = scan_path_sync(str(tmp_path))
        assert success is False
        assert result == "Error / Offline"
