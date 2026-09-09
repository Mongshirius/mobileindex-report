from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from zoneinfo import ZoneInfo

import pandas as pd

METRIC_COLUMNS = ["dau", "installs", "revenue"]
_REQUIRED_COLUMNS = {"date", "app_name", *METRIC_COLUMNS}


@dataclass
class ReportData:
    sheets: dict[str, pd.DataFrame]
    summary: dict[str, object]


def transform(raw: pd.DataFrame, run_date: date, tz: str = "Asia/Seoul") -> ReportData:
    missing = _REQUIRED_COLUMNS - set(raw.columns)
    if missing:
        raise ValueError(f"입력 데이터에 필수 컬럼이 없습니다: {sorted(missing)}")
    summary_sheet = (
        raw.groupby("app_name", as_index=False)[METRIC_COLUMNS].sum()
    )
    sheets = {"원본": raw.copy(), "요약": summary_sheet}
    summary = {
        "row_count": int(len(raw)),
        "app_count": int(raw["app_name"].nunique()),
        "date_range": [str(raw["date"].min()), str(raw["date"].max())],
        "run_date": run_date.isoformat(),
        "generated_at": datetime.now(ZoneInfo(tz)).isoformat(),
    }
    return ReportData(sheets=sheets, summary=summary)
