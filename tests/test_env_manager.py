"""Tests for core/env_manager.py"""
import os
import pytest
from unittest.mock import patch, mock_open, MagicMock

# Patch ENV_FILE_PATH before importing the module under test
with patch.dict('os.environ', {}):
    from core.env_manager import read_env, update_env_param


class TestReadEnv:
    def test_returns_not_set_when_file_missing(self, tmp_path):
        missing = str(tmp_path / "nonexistent_env_file")
        with patch("core.env_manager.ENV_FILE_PATH", missing):
            env, version = read_env()
        assert env == "Not set"
        assert version == "Not set"

    def test_reads_env_and_version(self, tmp_path):
        env_file = tmp_path / "env_file"
        env_file.write_text("ENV=UAT\nVERSION=1.2.3\n")
        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            env, version = read_env()
        assert env == "UAT"
        assert version == "1.2.3"

    def test_reads_only_env_line(self, tmp_path):
        env_file = tmp_path / "env_file"
        env_file.write_text("ENV=PROD\n")
        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            env, version = read_env()
        assert env == "PROD"
        assert version == "Not set"

    def test_reads_only_version_line(self, tmp_path):
        env_file = tmp_path / "env_file"
        env_file.write_text("VERSION=2.0.0\n")
        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            env, version = read_env()
        assert env == "Not set"
        assert version == "2.0.0"

    def test_ignores_unrelated_lines(self, tmp_path):
        env_file = tmp_path / "env_file"
        env_file.write_text("OTHER=value\nENV=DEV\nFOO=bar\nVERSION=0.1\n")
        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            env, version = read_env()
        assert env == "DEV"
        assert version == "0.1"

    def test_returns_not_set_on_read_exception(self, tmp_path):
        env_file = tmp_path / "env_file"
        env_file.write_text("ENV=UAT\nVERSION=1.0\n")
        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            with patch("builtins.open", side_effect=IOError("disk error")):
                env, version = read_env()
        assert env == "Not set"
        assert version == "Not set"

    def test_empty_file(self, tmp_path):
        env_file = tmp_path / "env_file"
        env_file.write_text("")
        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            env, version = read_env()
        assert env == "Not set"
        assert version == "Not set"


class TestUpdateEnvParam:
    @pytest.mark.parametrize("bad_value", [
        "", None, "Checking...", "No versions found", "N/A",
        "Checking something", "No versions available", "Has N/A in it",
    ])
    def test_rejects_invalid_values(self, bad_value):
        success, msg = update_env_param("ENV", bad_value)
        assert success is False
        assert msg == "Invalid value selected."

    def test_updates_existing_key(self, tmp_path):
        env_file = tmp_path / "env_file"
        env_file.write_text("ENV=UAT\nVERSION=1.0\n")
        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            success, msg = update_env_param("ENV", "PROD")
        assert success is True
        assert msg == ""
        content = env_file.read_text()
        assert "ENV=PROD\n" in content
        assert "VERSION=1.0\n" in content

    def test_updates_existing_version_key(self, tmp_path):
        env_file = tmp_path / "env_file"
        env_file.write_text("ENV=UAT\nVERSION=1.0\n")
        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            success, msg = update_env_param("VERSION", "2.5.0")
        assert success is True
        content = env_file.read_text()
        assert "VERSION=2.5.0\n" in content
        assert "ENV=UAT\n" in content

    def test_inserts_env_at_start_when_missing(self, tmp_path):
        env_file = tmp_path / "env_file"
        env_file.write_text("VERSION=1.0\n")
        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            success, msg = update_env_param("ENV", "DEV")
        assert success is True
        lines = env_file.read_text().splitlines()
        assert lines[0] == "ENV=DEV"

    def test_inserts_version_at_position_1_when_missing(self, tmp_path):
        env_file = tmp_path / "env_file"
        env_file.write_text("ENV=UAT\n")
        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            success, msg = update_env_param("VERSION", "3.0.0")
        assert success is True
        lines = env_file.read_text().splitlines()
        assert "VERSION=3.0.0" in lines

    def test_creates_file_when_not_exists(self, tmp_path):
        env_file = tmp_path / "subdir" / "env_file"
        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            success, msg = update_env_param("ENV", "STAGING")
        assert success is True
        assert env_file.exists()
        assert "ENV=STAGING" in env_file.read_text()

    def test_returns_false_on_write_exception(self, tmp_path):
        env_file = tmp_path / "env_file"
        env_file.write_text("ENV=UAT\n")

        real_open = open

        def _open_side_effect(path, mode="r", *args, **kwargs):
            if "w" in mode:
                raise IOError("write error")
            return real_open(path, mode, *args, **kwargs)

        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            with patch("builtins.open", side_effect=_open_side_effect):
                success, msg = update_env_param("ENV", "PROD")
        assert success is False
        assert "write error" in msg

    def test_value_containing_equals_sign(self, tmp_path):
        """A value that contains '=' should be written verbatim."""
        env_file = tmp_path / "env_file"
        env_file.write_text("ENV=UAT\n")
        with patch("core.env_manager.ENV_FILE_PATH", str(env_file)):
            success, _ = update_env_param("VERSION", "1.0.0+build=42")
        assert success is True
        content = env_file.read_text()
        assert "VERSION=1.0.0+build=42\n" in content
