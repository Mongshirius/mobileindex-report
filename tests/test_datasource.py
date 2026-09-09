from datetime import date

import pandas as pd

from report_pipeline.datasource import APPS, COLUMNS, LOOKBACK_DAYS, fetch


def test_fetch_returns_expected_schema(settings):
    df = fetch(settings, date(2026, 9, 9))
    assert list(df.columns) == COLUMNS
    assert not df.empty
    assert len(df) == LOOKBACK_DAYS * len(APPS)


def test_fetch_is_deterministic(settings):
    first = fetch(settings, date(2026, 9, 9))
    second = fetch(settings, date(2026, 9, 9))
    pd.testing.assert_frame_equal(first, second)


def test_fetch_covers_lookback_window_ending_on_run_date(settings):
    df = fetch(settings, date(2026, 9, 9))
    assert df["date"].max() == "2026-09-09"
    assert df["date"].min() == "2026-09-03"  # 7일 창


def test_fetch_metric_columns_are_numeric(settings):
    df = fetch(settings, date(2026, 9, 9))
    for column in ("dau", "installs", "revenue"):
        assert pd.api.types.is_numeric_dtype(df[column])
