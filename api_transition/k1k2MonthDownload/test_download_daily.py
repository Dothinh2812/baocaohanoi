# -*- coding: utf-8 -*-

from datetime import date

import download_daily


def test_default_target_date_is_yesterday():
    assert download_daily.resolve_target_date(today=date(2026, 5, 1)) == date(2026, 4, 30)


def test_selected_month_uses_yesterday_day_number():
    assert download_daily.resolve_target_date(
        today=date(2026, 5, 16),
        report_year=2026,
        report_month=4,
    ) == date(2026, 4, 15)


def test_explicit_day_overrides_yesterday_day_number():
    assert download_daily.resolve_target_date(
        today=date(2026, 5, 16),
        report_year=2026,
        report_month=4,
        report_day=23,
    ) == date(2026, 4, 23)
