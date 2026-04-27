"""Tests for core/log_ops.py"""
import os
import zipfile
import pytest
from unittest.mock import patch

from core.log_ops import (
    extract_service,
    format_service_name,
    get_services,
    get_log_files,
    generate_unified_timeline,
    export_logs_to_zip,
)


class TestExtractService:
    def test_standard_log_filename(self):
        assert extract_service("AddIn.EXCEL.12345.log") == "AddIn"

    def test_single_segment(self):
        assert extract_service("LoaderLog.log") == "LoaderLog"

    def test_no_extension(self):
        assert extract_service("ServiceName") == "ServiceName"

    def test_multiple_dots(self):
        assert extract_service("A.B.C.D.log") == "A"


class TestFormatServiceName:
    def test_explicit_rename(self):
        assert format_service_name("LoaderLog") == "Loader"

    def test_suffix_priority_longest_first(self):
        # 'ExcelAddInService' should be stripped (longer) before 'Service'
        assert format_service_name("MyExcelAddInService") == "My"

    def test_no_match_returns_raw(self):
        assert format_service_name("UnknownModule") == "UnknownModule"


class TestGetServices:
    def test_empty_file_list(self):
        assert get_services([]) == []

    def test_single_file(self):
        result = get_services(["/logs/AddIn.EXCEL.1.log"])
        assert result == [("AddIn", "AddIn")]

    def test_deduplicates_same_service(self):
        files = ["/logs/AddIn.1.log", "/logs/AddIn.2.log"]
        result = get_services(files)
        assert len(result) == 1
        assert result[0][0] == "AddIn"

    def test_multiple_services_sorted(self):
        files = ["/logs/ZService.1.log", "/logs/LoaderLog.1.log"]
        result = get_services(files)
        raws = [r[0] for r in result]
        assert raws == sorted(raws)

    def test_rename_applied(self):
        files = ["/logs/LoaderLog.1.log"]
        result = get_services(files)
        assert result == [("LoaderLog", "Loader")]

    def test_suffix_stripping_applied(self):
        files = ["/logs/FooExcelAddInService.1.log"]
        result = get_services(files)
        assert result == [("FooExcelAddInService", "Foo")]


class TestGetLogFiles:
    def test_nonexistent_directory(self, tmp_path):
        result = get_log_files(str(tmp_path / "missing"))
        assert result == []

    def test_empty_directory(self, tmp_path):
        result = get_log_files(str(tmp_path))
        assert result == []

    def test_returns_log_files(self, tmp_path):
        (tmp_path / "a.log").write_text("x")
        (tmp_path / "b.txt").write_text("x")
        (tmp_path / "c.csv").write_text("x")
        result = get_log_files(str(tmp_path))
        basenames = [os.path.basename(f) for f in result]
        assert "a.log" in basenames
        assert "b.txt" in basenames
        assert "c.csv" not in basenames

    def test_sorted_by_mtime(self, tmp_path):
        import time
        first = tmp_path / "first.log"
        first.write_text("x")
        time.sleep(0.05)
        second = tmp_path / "second.log"
        second.write_text("x")
        result = get_log_files(str(tmp_path))
        assert os.path.basename(result[0]) == "first.log"
        assert os.path.basename(result[1]) == "second.log"


class TestGenerateUnifiedTimeline:
    def test_empty_list(self):
        entries, errors = generate_unified_timeline([])
        assert entries == []
        assert errors == []

    def test_single_log_file(self, tmp_path):
        log = tmp_path / "svc.log"
        log.write_text(
            "2024-01-01 10:00:00 INFO starting\n"
            "2024-01-01 10:00:01 INFO done\n"
        )
        entries, errors = generate_unified_timeline([str(log)])
        assert len(entries) == 2
        assert errors == []
        assert entries[0][0] == "2024-01-01 10:00:00"
        assert entries[1][0] == "2024-01-01 10:00:01"

    def test_entries_sorted_chronologically(self, tmp_path):
        log1 = tmp_path / "a.log"
        log2 = tmp_path / "b.log"
        log1.write_text("2024-01-01 10:00:05 event_a\n")
        log2.write_text("2024-01-01 10:00:01 event_b\n")
        entries, errors = generate_unified_timeline([str(log1), str(log2)])
        assert entries[0][0] == "2024-01-01 10:00:01"
        assert entries[1][0] == "2024-01-01 10:00:05"

    def test_reverse_order(self, tmp_path):
        log = tmp_path / "svc.log"
        log.write_text(
            "2024-01-01 09:00:00 first\n"
            "2024-01-01 10:00:00 second\n"
        )
        entries, errors = generate_unified_timeline([str(log)], reverse=True)
        assert entries[0][0] == "2024-01-01 10:00:00"
        assert entries[1][0] == "2024-01-01 09:00:00"

    def test_continuation_lines_attached_to_block(self, tmp_path):
        log = tmp_path / "svc.log"
        log.write_text(
            "2024-01-01 10:00:00 ERROR something went wrong\n"
            "  at some.stack.frame()\n"
            "  at another.frame()\n"
            "2024-01-01 10:00:01 INFO recovered\n"
        )
        entries, errors = generate_unified_timeline([str(log)])
        assert len(entries) == 2
        assert "at some.stack.frame()" in entries[0][1]
        assert "at another.frame()" in entries[0][1]

    def test_lines_before_first_timestamp_discarded(self, tmp_path):
        log = tmp_path / "svc.log"
        log.write_text(
            "header line without timestamp\n"
            "another header\n"
            "2024-01-01 10:00:00 first real entry\n"
        )
        entries, errors = generate_unified_timeline([str(log)])
        assert len(entries) == 1
        assert "header line" not in entries[0][1]

    def test_unreadable_file_recorded_in_errors(self, tmp_path):
        missing = str(tmp_path / "ghost.log")
        entries, errors = generate_unified_timeline([missing])
        assert entries == []
        assert len(errors) == 1
        assert errors[0][0] == "ghost.log"

    def test_filename_prefix_in_entry_text(self, tmp_path):
        log = tmp_path / "MyService.EXCEL.log"
        log.write_text("2024-01-01 10:00:00 INFO something\n")
        entries, errors = generate_unified_timeline([str(log)])
        assert "[MyService.EXCEL.log]" in entries[0][1]

    def test_multiple_files_errors_and_entries_mixed(self, tmp_path):
        good = tmp_path / "good.log"
        good.write_text("2024-01-01 10:00:00 INFO ok\n")
        bad = str(tmp_path / "bad.log")  # does not exist
        entries, errors = generate_unified_timeline([str(good), bad])
        assert len(entries) == 1
        assert len(errors) == 1


class TestExportLogsToZip:
    def test_creates_zip_with_log_files(self, tmp_path):
        log1 = tmp_path / "a.log"
        log2 = tmp_path / "b.log"
        log1.write_text("log content a")
        log2.write_text("log content b")
        dest = str(tmp_path / "output.zip")
        success, msg = export_logs_to_zip([str(log1), str(log2)], dest)
        assert success is True
        assert msg == ""
        with zipfile.ZipFile(dest, 'r') as zf:
            names = zf.namelist()
        assert "a.log" in names
        assert "b.log" in names

    def test_zip_contains_correct_content(self, tmp_path):
        log = tmp_path / "svc.log"
        log.write_text("hello world")
        dest = str(tmp_path / "out.zip")
        export_logs_to_zip([str(log)], dest)
        with zipfile.ZipFile(dest, 'r') as zf:
            content = zf.read("svc.log").decode()
        assert content == "hello world"

    def test_returns_false_on_unwritable_destination(self, tmp_path):
        log = tmp_path / "svc.log"
        log.write_text("x")
        bad_dest = "/nonexistent_dir/output.zip"
        success, msg = export_logs_to_zip([str(log)], bad_dest)
        assert success is False
        assert msg != ""

    def test_empty_file_list_creates_empty_zip(self, tmp_path):
        dest = str(tmp_path / "empty.zip")
        success, msg = export_logs_to_zip([], dest)
        assert success is True
        with zipfile.ZipFile(dest, 'r') as zf:
            assert zf.namelist() == []
