from datetime import date

import pandas as pd

from report_pipeline.transform import ReportData, transform


def _sample_raw() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"date": "2026-09-08", "app_name": "앱A", "dau": 100, "installs": 10, "revenue": 1.0},
            {"date": "2026-09-09", "app_name": "앱A", "dau": 200, "installs": 20, "revenue": 2.0},
            {"date": "2026-09-09", "app_name": "앱B", "dau": 50, "installs": 5, "revenue": 0.5},
        ],
        columns=["date", "app_name", "dau", "installs", "revenue"],
    )


def test_transform_returns_report_data_with_two_sheets():
    data = transform(_sample_raw(), date(2026, 9, 9))
    assert isinstance(data, ReportData)
    assert list(data.sheets.keys()) == ["원본", "요약"]


def test_raw_sheet_is_unmodified_copy():
    raw = _sample_raw()
    data = transform(raw, date(2026, 9, 9))
    pd.testing.assert_frame_equal(data.sheets["원본"], raw)
    assert data.sheets["원본"] is not raw


def test_summary_sheet_aggregates_metrics_by_app():
    data = transform(_sample_raw(), date(2026, 9, 9))
    summary_sheet = data.sheets["요약"].set_index("app_name")
    assert summary_sheet.loc["앱A", "dau"] == 300
    assert summary_sheet.loc["앱A", "installs"] == 30
    assert summary_sheet.loc["앱B", "revenue"] == 0.5


def test_summary_metadata_values():
    data = transform(_sample_raw(), date(2026, 9, 9))
    assert data.summary["row_count"] == 3
    assert data.summary["app_count"] == 2
    assert data.summary["date_range"] == ["2026-09-08", "2026-09-09"]
    assert data.summary["run_date"] == "2026-09-09"
    assert data.summary["generated_at"].startswith("20")
