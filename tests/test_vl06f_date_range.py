"""
tests/test_vl06f_date_range.py
-------------------------------
Unit tests for VL06FDownloader._resolve_date_range().

These tests use a fake session and config so that SAP GUI is never required.
"""

import configparser
from datetime import date, timedelta
from unittest.mock import MagicMock

import pytest

from sap.vl06f import VL06FDownloader


def _make_downloader(vl06f_overrides: dict | None = None) -> VL06FDownloader:
    """Build a VL06FDownloader with a mock session and a minimal config."""
    cfg = configparser.ConfigParser()
    cfg["SAP"] = {"screen_wait_timeout": "10"}
    cfg["OUTPUT"] = {"output_dir": "output", "file_name": "report.xlsx"}
    cfg["VL06F"] = {
        "date_range_mode": "today",
        "last_n_days":     "7",
        "date_from":       "01.01.2024",
        "date_to":         "31.01.2024",
        "shipping_point":  "SP01",
        "layout_variant":  "",
    }
    if vl06f_overrides:
        cfg["VL06F"].update(vl06f_overrides)

    fake_session = MagicMock()
    fake_session.Busy = False
    return VL06FDownloader(fake_session, cfg)


class TestResolveDateRange:
    def test_today(self):
        dl = _make_downloader({"date_range_mode": "today"})
        d_from, d_to = dl._resolve_date_range()
        assert d_from == date.today()
        assert d_to   == date.today()

    def test_yesterday(self):
        dl = _make_downloader({"date_range_mode": "yesterday"})
        d_from, d_to = dl._resolve_date_range()
        yesterday = date.today() - timedelta(days=1)
        assert d_from == yesterday
        assert d_to   == yesterday

    def test_last_n_days(self):
        dl = _make_downloader({"date_range_mode": "last_n_days", "last_n_days": "14"})
        d_from, d_to = dl._resolve_date_range()
        expected_from = date.today() - timedelta(days=13)
        assert d_from == expected_from
        assert d_to   == date.today()

    def test_custom(self):
        dl = _make_downloader({
            "date_range_mode": "custom",
            "date_from": "05.03.2024",
            "date_to":   "20.03.2024",
        })
        d_from, d_to = dl._resolve_date_range()
        assert d_from == date(2024, 3, 5)
        assert d_to   == date(2024, 3, 20)

    def test_unknown_mode_falls_back_to_today(self):
        dl = _make_downloader({"date_range_mode": "nonsense"})
        d_from, d_to = dl._resolve_date_range()
        assert d_from == date.today()
        assert d_to   == date.today()

    def test_last_n_days_single_day(self):
        dl = _make_downloader({"date_range_mode": "last_n_days", "last_n_days": "1"})
        d_from, d_to = dl._resolve_date_range()
        assert d_from == date.today()
        assert d_to   == date.today()
