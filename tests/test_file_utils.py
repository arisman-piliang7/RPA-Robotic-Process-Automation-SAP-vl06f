"""
tests/test_file_utils.py
------------------------
Unit tests for utils/file_utils.py.
"""

import os
import tempfile
from datetime import datetime

import pytest

from utils.file_utils import resolve_path, ensure_dir, build_output_path


FIXED_DT = datetime(2024, 3, 15, 8, 5, 9)


class TestResolvePath:
    def test_all_placeholders(self):
        template = "{YYYY}-{MM}-{DD}_{HH}{MIN}{SS}"
        result = resolve_path(template, FIXED_DT)
        assert result == "2024-03-15_080509"

    def test_nested_directory_template(self):
        template = "output/{YYYY}/{MM}/report_{DD}.xlsx"
        result = resolve_path(template, FIXED_DT)
        assert result == "output/2024/03/report_15.xlsx"

    def test_no_placeholders(self):
        template = "plain_file.txt"
        result = resolve_path(template, FIXED_DT)
        assert result == "plain_file.txt"

    def test_defaults_to_now_when_no_dt_given(self):
        template = "{YYYY}"
        result = resolve_path(template)
        assert result == datetime.now().strftime("%Y")

    def test_two_digit_zero_padding(self):
        dt = datetime(2024, 1, 5, 9, 3, 7)
        result = resolve_path("{MM}-{DD}-{HH}-{MIN}-{SS}", dt)
        assert result == "01-05-09-03-07"


class TestEnsureDir:
    def test_creates_directory(self, tmp_path):
        new_dir = str(tmp_path / "a" / "b" / "c")
        returned = ensure_dir(new_dir)
        assert os.path.isdir(new_dir)
        assert returned == new_dir

    def test_existing_directory_is_not_an_error(self, tmp_path):
        existing = str(tmp_path)
        ensure_dir(existing)  # should not raise
        assert os.path.isdir(existing)


class TestBuildOutputPath:
    def test_path_is_under_resolved_dir(self, tmp_path):
        dir_tpl  = str(tmp_path / "{YYYY}" / "{MM}")
        file_tpl = "VL06F_{YYYY}{MM}{DD}.xlsx"
        result = build_output_path(dir_tpl, file_tpl, FIXED_DT)

        expected_dir  = str(tmp_path / "2024" / "03")
        expected_file = os.path.join(expected_dir, "VL06F_20240315.xlsx")

        assert result == expected_file
        assert os.path.isdir(expected_dir)

    def test_directory_is_created(self, tmp_path):
        dir_tpl  = str(tmp_path / "out" / "{YYYY}")
        file_tpl = "report.xlsx"
        build_output_path(dir_tpl, file_tpl, FIXED_DT)
        assert os.path.isdir(str(tmp_path / "out" / "2024"))
