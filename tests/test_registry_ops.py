"""Tests for core/registry_ops.py

winreg is a Windows-only module; it is mocked in sys.modules so these tests
run on any platform.
"""
import sys
import types
import pytest
from unittest.mock import patch, MagicMock, call

# ---------------------------------------------------------------------------
# Build a minimal winreg stub so the module can be imported on Linux
# ---------------------------------------------------------------------------
_winreg_stub = types.ModuleType("winreg")
_winreg_stub.HKEY_CURRENT_USER = 0x80000001
_winreg_stub.KEY_READ = 0x20019
_winreg_stub.KEY_SET_VALUE = 0x0002
_winreg_stub.KEY_WRITE = 0x20006
_winreg_stub.REG_SZ = 1
_winreg_stub.OpenKey = MagicMock()
_winreg_stub.EnumValue = MagicMock()
_winreg_stub.CloseKey = MagicMock()
_winreg_stub.SetValueEx = MagicMock()

sys.modules.setdefault("winreg", _winreg_stub)

from core.registry_ops import (  # noqa: E402 – must come after stub registration
    scan_registry_for_ninetyone,
    open_regedit_at_path,
    point_excel_to_addin,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _enum_values(entries):
    """Return a side_effect callable that simulates winreg.EnumValue iteration.

    *entries* is a list of (name, value, type) tuples. Raises OSError when the
    index goes out of range, matching Windows registry behaviour.
    """
    def _side_effect(key, index):
        if index < len(entries):
            return entries[index]
        raise OSError
    return _side_effect


# ---------------------------------------------------------------------------
# scan_registry_for_ninetyone
# ---------------------------------------------------------------------------

class TestScanRegistryForNinetyone:
    def test_found_entries_with_ninetyone(self):
        entries = [
            ("OPEN", r"C:\ninetyone\addin.xll", 1),
            ("OPEN1", r"C:\other\thing.xll", 1),
        ]
        with patch("core.registry_ops.winreg.OpenKey") as mock_open, \
             patch("core.registry_ops.winreg.EnumValue", side_effect=_enum_values(entries)), \
             patch("core.registry_ops.winreg.CloseKey"):
            mock_open.return_value = MagicMock()
            success, found = scan_registry_for_ninetyone()

        assert success is True
        assert len(found) == 1
        assert "ninetyone" in found[0].lower()

    def test_no_ninetyone_entries(self):
        entries = [("OPEN", r"C:\other\addin.xll", 1)]
        with patch("core.registry_ops.winreg.OpenKey") as mock_open, \
             patch("core.registry_ops.winreg.EnumValue", side_effect=_enum_values(entries)), \
             patch("core.registry_ops.winreg.CloseKey"):
            mock_open.return_value = MagicMock()
            success, found = scan_registry_for_ninetyone()

        assert success is True
        assert found == []

    def test_empty_registry_key(self):
        with patch("core.registry_ops.winreg.OpenKey") as mock_open, \
             patch("core.registry_ops.winreg.EnumValue", side_effect=_enum_values([])), \
             patch("core.registry_ops.winreg.CloseKey"):
            mock_open.return_value = MagicMock()
            success, found = scan_registry_for_ninetyone()

        assert success is True
        assert found == []

    def test_key_not_found(self):
        with patch("core.registry_ops.winreg.OpenKey", side_effect=FileNotFoundError):
            success, msg = scan_registry_for_ninetyone()

        assert success is False
        assert "not found" in msg.lower()

    def test_unexpected_exception(self):
        with patch("core.registry_ops.winreg.OpenKey", side_effect=RuntimeError("boom")):
            success, msg = scan_registry_for_ninetyone()

        assert success is False
        assert "boom" in msg

    def test_case_insensitive_ninetyone_match(self):
        entries = [("OPEN", r"C:\NinetyOne\addin.xll", 1)]
        with patch("core.registry_ops.winreg.OpenKey") as mock_open, \
             patch("core.registry_ops.winreg.EnumValue", side_effect=_enum_values(entries)), \
             patch("core.registry_ops.winreg.CloseKey"):
            mock_open.return_value = MagicMock()
            success, found = scan_registry_for_ninetyone()

        assert success is True
        assert len(found) == 1

    def test_non_string_values_ignored(self):
        entries = [("OPEN", 12345, 4)]  # numeric value
        with patch("core.registry_ops.winreg.OpenKey") as mock_open, \
             patch("core.registry_ops.winreg.EnumValue", side_effect=_enum_values(entries)), \
             patch("core.registry_ops.winreg.CloseKey"):
            mock_open.return_value = MagicMock()
            success, found = scan_registry_for_ninetyone()

        assert success is True
        assert found == []


# ---------------------------------------------------------------------------
# open_regedit_at_path
# ---------------------------------------------------------------------------

class TestOpenRegeditAtPath:
    def test_success(self):
        with patch("core.registry_ops.winreg.OpenKey") as mock_open, \
             patch("core.registry_ops.winreg.SetValueEx"), \
             patch("core.registry_ops.winreg.CloseKey"), \
             patch("core.registry_ops.subprocess.Popen") as mock_popen:
            mock_open.return_value = MagicMock()
            success, msg = open_regedit_at_path()

        assert success is True
        assert msg == ""
        mock_popen.assert_called_once_with(["regedit.exe", "/m"])

    def test_permission_error(self):
        with patch("core.registry_ops.winreg.OpenKey", side_effect=PermissionError):
            success, msg = open_regedit_at_path()

        assert success is False
        assert "administrator" in msg.lower()

    def test_unexpected_exception(self):
        with patch("core.registry_ops.winreg.OpenKey", side_effect=RuntimeError("oops")):
            success, msg = open_regedit_at_path()

        assert success is False
        assert "oops" in msg


# ---------------------------------------------------------------------------
# point_excel_to_addin
# ---------------------------------------------------------------------------

class TestPointExcelToAddin:
    NEW_XLL = r"C:\ExcelAddIn\v1.0\addin.xll"

    def test_overwrites_existing_ninetyone_slot(self):
        entries = [("OPEN", r"C:\ninetyone\old.xll", 1)]
        with patch("core.registry_ops.winreg.OpenKey") as mock_open, \
             patch("core.registry_ops.winreg.EnumValue", side_effect=_enum_values(entries)), \
             patch("core.registry_ops.winreg.SetValueEx") as mock_set, \
             patch("core.registry_ops.winreg.CloseKey"):
            key = MagicMock()
            mock_open.return_value = key
            success, msg = point_excel_to_addin(self.NEW_XLL)

        assert success is True
        name_arg = mock_set.call_args[0][1]
        assert name_arg == "OPEN"
        value_arg = mock_set.call_args[0][4]
        assert self.NEW_XLL in value_arg

    def test_overwrites_investmenttech_slot(self):
        entries = [("OPEN", r"C:\InvestmentTechExcelAddIn\old.xll", 1)]
        with patch("core.registry_ops.winreg.OpenKey") as mock_open, \
             patch("core.registry_ops.winreg.EnumValue", side_effect=_enum_values(entries)), \
             patch("core.registry_ops.winreg.SetValueEx") as mock_set, \
             patch("core.registry_ops.winreg.CloseKey"):
            mock_open.return_value = MagicMock()
            success, msg = point_excel_to_addin(self.NEW_XLL)

        assert success is True

    def test_uses_open_slot_when_no_existing_addin(self):
        entries = [("SOMEOTHER", r"C:\other.xll", 1)]
        with patch("core.registry_ops.winreg.OpenKey") as mock_open, \
             patch("core.registry_ops.winreg.EnumValue", side_effect=_enum_values(entries)), \
             patch("core.registry_ops.winreg.SetValueEx") as mock_set, \
             patch("core.registry_ops.winreg.CloseKey"):
            mock_open.return_value = MagicMock()
            success, _ = point_excel_to_addin(self.NEW_XLL)

        assert success is True
        name_arg = mock_set.call_args[0][1]
        assert name_arg == "OPEN"

    def test_uses_open1_when_open_taken(self):
        entries = [("OPEN", r"C:\other.xll", 1)]
        with patch("core.registry_ops.winreg.OpenKey") as mock_open, \
             patch("core.registry_ops.winreg.EnumValue", side_effect=_enum_values(entries)), \
             patch("core.registry_ops.winreg.SetValueEx") as mock_set, \
             patch("core.registry_ops.winreg.CloseKey"):
            mock_open.return_value = MagicMock()
            # OPEN is in existing_opens but not ninetyone → no target_value_name
            # OPEN exists → should use OPEN1
            success, _ = point_excel_to_addin(self.NEW_XLL)

        assert success is True
        name_arg = mock_set.call_args[0][1]
        assert name_arg == "OPEN1"

    def test_uses_open2_when_open_and_open1_taken(self):
        entries = [
            ("OPEN", r"C:\other.xll", 1),
            ("OPEN1", r"C:\another.xll", 1),
        ]
        with patch("core.registry_ops.winreg.OpenKey") as mock_open, \
             patch("core.registry_ops.winreg.EnumValue", side_effect=_enum_values(entries)), \
             patch("core.registry_ops.winreg.SetValueEx") as mock_set, \
             patch("core.registry_ops.winreg.CloseKey"):
            mock_open.return_value = MagicMock()
            success, _ = point_excel_to_addin(self.NEW_XLL)

        assert success is True
        name_arg = mock_set.call_args[0][1]
        assert name_arg == "OPEN2"

    def test_permission_error(self):
        with patch("core.registry_ops.winreg.OpenKey", side_effect=PermissionError):
            success, msg = point_excel_to_addin(self.NEW_XLL)

        assert success is False
        assert "administrator" in msg.lower()

    def test_unexpected_exception(self):
        with patch("core.registry_ops.winreg.OpenKey", side_effect=RuntimeError("crash")):
            success, msg = point_excel_to_addin(self.NEW_XLL)

        assert success is False
        assert "crash" in msg

    def test_registry_value_has_r_flag(self):
        entries = [("OPEN", r"C:\ninetyone\old.xll", 1)]
        with patch("core.registry_ops.winreg.OpenKey") as mock_open, \
             patch("core.registry_ops.winreg.EnumValue", side_effect=_enum_values(entries)), \
             patch("core.registry_ops.winreg.SetValueEx") as mock_set, \
             patch("core.registry_ops.winreg.CloseKey"):
            mock_open.return_value = MagicMock()
            point_excel_to_addin(self.NEW_XLL)

        value_arg = mock_set.call_args[0][4]
        assert value_arg.startswith('/R "')
