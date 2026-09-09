"""데이터 소스.

1단계에서는 결정적 더미 데이터를 반환한다.

2단계에서 이 함수 내부를 모바일인덱스 연동으로 교체한다:
  - 경로 A: 공식 API / 원격 MCP 엔드포인트 호출 + 인증
  - 경로 B: 웹 로그인 자동화 + CSV 다운로드·파싱
시그니처와 반환 스키마(COLUMNS)는 유지한다.
"""

from __future__ import annotations

from datetime import date, timedelta

import pandas as pd

from .config import Settings

COLUMNS = ["date", "app_name", "dau", "installs", "revenue"]
APPS = ["앱A", "앱B", "앱C"]
LOOKBACK_DAYS = 7


def fetch(settings: Settings, run_date: date) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for offset in range(LOOKBACK_DAYS):
        day = run_date - timedelta(days=LOOKBACK_DAYS - 1 - offset)
        for app_index, app_name in enumerate(APPS):
            seed = day.toordinal() + app_index * 1000
            rows.append(
                {
                    "date": day.isoformat(),
                    "app_name": app_name,
                    "dau": 10_000 + seed % 5_000,
                    "installs": 500 + seed % 300,
                    "revenue": round((seed % 900) * 1.5, 2),
                }
            )
    return pd.DataFrame(rows, columns=COLUMNS)
