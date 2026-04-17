"""Tests for core/system_ops.py"""
import pytest
from unittest.mock import patch, MagicMock

from core.system_ops import open_in_explorer, launch_excel, close_excel, kill_excel


class TestOpenInExplorer:
    def test_returns_false_when_path_missing(self, tmp_path):
        missing = str(tmp_path / "nonexistent")
        result = open_in_explorer(missing)
        assert result is False

    def test_calls_startfile_when_path_exists(self, tmp_path):
        with patch("core.system_ops.os.startfile", create=True) as mock_sf:
            result = open_in_explorer(str(tmp_path))
        assert result is True
        mock_sf.assert_called_once_with(str(tmp_path))


class TestLaunchExcel:
    def test_returns_true_on_success(self):
        with patch("core.system_ops.subprocess.Popen") as mock_popen:
            success, msg = launch_excel()
        assert success is True
        assert msg == ""
        mock_popen.assert_called_once_with("start excel", shell=True)

    def test_returns_false_on_exception(self):
        with patch("core.system_ops.subprocess.Popen", side_effect=OSError("no excel")):
            success, msg = launch_excel()
        assert success is False
        assert "no excel" in msg


class TestCloseExcel:
    def test_returns_true_on_success(self):
        mock_result = MagicMock()
        with patch("core.system_ops.subprocess.run", return_value=mock_result) as mock_run:
            success, msg = close_excel()
        assert success is True
        assert msg == ""
        mock_run.assert_called_once_with(
            ["taskkill", "/im", "excel.exe"], capture_output=True, text=True
        )

    def test_returns_false_on_exception(self):
        with patch("core.system_ops.subprocess.run", side_effect=OSError("taskkill failed")):
            success, msg = close_excel()
        assert success is False
        assert "taskkill failed" in msg


class TestKillExcel:
    def test_returns_true_on_success(self):
        mock_result = MagicMock()
        with patch("core.system_ops.subprocess.run", return_value=mock_result) as mock_run:
            success, msg = kill_excel()
        assert success is True
        assert msg == ""
        assert mock_run.call_count == 2

    def test_returns_false_on_exception(self):
        with patch("core.system_ops.subprocess.run", side_effect=OSError("wmic error")):
            success, msg = kill_excel()
        assert success is False
        assert "wmic error" in msg
