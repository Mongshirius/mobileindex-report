# 모바일인덱스 리포트 파이프라인 (1단계) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 더미 데이터를 엑셀 리포트로 만들어 이메일로 발송하는 파이프라인을 GitHub Actions에서 무인 실행 가능하게 구축한다.

**Architecture:** 순수 함수 4개 모듈(`datasource` → `transform` → `report` → `notify`)과 얇은 오케스트레이터(`pipeline.run`), CLI 진입점(`__main__`). 데이터 소스는 `fetch(settings, run_date) -> DataFrame` 시그니처 뒤에 격리하여, 2단계에서 모바일인덱스 실연동 시 이 함수 몸통만 교체한다. 실행은 GitHub Actions의 `workflow_dispatch` + (미정인) `schedule` 트리거.

**Tech Stack:** Python 3.11+, pandas, openpyxl, 표준 라이브러리 `smtplib`/`email`, pytest, GitHub Actions.

**Spec:** `docs/superpowers/specs/2026-09-09-mobileindex-report-pipeline-design.md`

## Global Constraints

- Python `>=3.11` (표준 라이브러리 `zoneinfo` 사용).
- 런타임 의존성은 `pandas>=2.0`, `openpyxl>=3.1` 두 개로 제한. 그 외는 개발 의존성(`pytest`, `python-dotenv`).
- 패키지명·import 경로: `report_pipeline` (위치: `src/report_pipeline/`). pytest는 `pythonpath = ["src"]`.
- 수신자는 단일 값 `MAIL_TO`. 다중 수신자 로직을 만들지 않는다.
- 엑셀 파일명 형식(정확히): `mobileindex_report_{YYYY-MM-DD}.xlsx` (기준일 기반).
- 이메일 제목 형식(정확히): `[모바일인덱스 리포트] {YYYY-MM-DD}`.
- 엑셀 시트 이름(정확히, 순서 포함): `["원본", "요약"]`.
- 필수 환경변수 7개(정확히): `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`, `MAIL_FROM`, `MAIL_TO`, `REPORT_TZ`. 누락 시 즉시 실패하며 누락된 변수명을 메시지에 포함.
- `.env` 계열 값이 `"..."` 또는 `'...'`로 감싸져 있으면 따옴표를 제거하고 읽는다.
- `.env.local`은 절대 커밋하지 않는다(`.gitignore` 등록).
- 워크북은 저장소 루트의 `build/` 디렉터리에 생성(`.gitignore` 등록).
- SMTP 전송은 STARTTLS 사용, 실패 시 2회 지수 백오프(2s, 4s) 재시도 후 예외 전파.
- GitHub Actions `schedule` cron은 UTC 기준. 발송 주기는 **미정** — `report.yml`에서 주석 처리한 채로 둔다.

---

## 파일 구조

| 파일 | 책임 | 태스크 |
|---|---|---|
| `pyproject.toml` | 패키지 메타·의존성·pytest 설정 | 1 |
| `.gitignore` | 비밀·빌드 산출물 제외 | 1 |
| `.env.example` | 필수 환경변수 키 목록(값 없음) | 1 |
| `src/report_pipeline/__init__.py` | 패키지 마커 | 1 |
| `tests/conftest.py` | 공용 `settings` 픽스처 | 2 |
| `src/report_pipeline/config.py` | 환경변수 로딩·검증, `Settings`, `ConfigError` | 2 |
| `src/report_pipeline/datasource.py` | `fetch()` — 더미 DataFrame (2단계 교체 지점) | 3 |
| `src/report_pipeline/transform.py` | `transform()` — `ReportData` 생성 | 4 |
| `src/report_pipeline/report.py` | `build_workbook()` — xlsx 작성·서식 | 5 |
| `src/report_pipeline/notify.py` | `send_email()` — SMTP 발송·재시도 | 6 |
| `src/report_pipeline/pipeline.py` | `run()` — 4모듈 오케스트레이션 | 7 |
| `src/report_pipeline/__main__.py` | CLI 진입점, `--date` 파싱, dotenv 로딩 | 8 |
| `.github/workflows/ci.yml` | push/PR 시 pytest | 9 |
| `.github/workflows/report.yml` | 수동/스케줄 리포트 실행 | 9 |
| `README.md` | 셋업·시크릿·cron 안내 | 9 |

---

## Task 1: 프로젝트 스캐폴딩

**Files:**
- Create: `pyproject.toml`
- Create: `src/report_pipeline/__init__.py`
- Create: `.env.example`
- Modify: `.gitignore` (이미 존재 — 항목 확인/보강)
- Test: `tests/test_smoke.py`

**Interfaces:**
- Consumes: 없음
- Produces: 설치 가능한 패키지 `report_pipeline`; `pytest`가 `pythonpath=["src"]`로 동작.

- [ ] **Step 1: 스모크 테스트 작성 (실패)**

`tests/test_smoke.py`:

```python
def test_package_importable():
    import report_pipeline  # noqa: F401
```

- [ ] **Step 2: 테스트 실행 → 실패 확인**

Run: `pytest tests/test_smoke.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'report_pipeline'` (또는 pytest가 `pyproject.toml`의 설정을 못 찾음)

- [ ] **Step 3: `pyproject.toml` 작성**

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[project]
name = "mobileindex-report"
version = "0.1.0"
description = "모바일인덱스 데이터를 엑셀 리포트로 만들어 이메일 발송하는 자동화 파이프라인"
requires-python = ">=3.11"
dependencies = [
    "pandas>=2.0",
    "openpyxl>=3.1",
]

[project.optional-dependencies]
dev = [
    "pytest>=8.0",
    "python-dotenv>=1.0",
]

[project.scripts]
mobileindex-report = "report_pipeline.__main__:main"

[tool.hatch.build.targets.wheel]
packages = ["src/report_pipeline"]

[tool.pytest.ini_options]
pythonpath = ["src"]
testpaths = ["tests"]
```

- [ ] **Step 4: 패키지 마커와 `.env.example` 작성**

`src/report_pipeline/__init__.py`:

```python
"""모바일인덱스 리포트 파이프라인."""
```

`.env.example`:

```dotenv
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=
SMTP_PASSWORD=
MAIL_FROM=
MAIL_TO=
REPORT_TZ=Asia/Seoul
```

- [ ] **Step 5: `.gitignore` 확인/보강**

파일에 아래 항목이 모두 있는지 확인하고 없으면 추가:

```gitignore
.env.local
build/
dist/
*.egg-info/
__pycache__/
.pytest_cache/
*.xlsx
.venv/
```

- [ ] **Step 6: 개발 설치 후 테스트 통과 확인**

Run: `pip install -e ".[dev]"` 그다음 `pytest -v`
Expected: PASS — `test_package_importable`

- [ ] **Step 7: 커밋**

```bash
git add pyproject.toml .env.example .gitignore src/report_pipeline/__init__.py tests/test_smoke.py
git commit -m "chore: 프로젝트 스캐폴딩 (pyproject, 패키지, pytest 설정)"
```

---

## Task 2: 설정 로딩 (`config.py`)

**Files:**
- Create: `src/report_pipeline/config.py`
- Create: `tests/conftest.py`
- Test: `tests/test_config.py`

**Interfaces:**
- Consumes: 없음
- Produces:
  - `Settings` (frozen dataclass): `smtp_host: str`, `smtp_port: int`, `smtp_user: str`, `smtp_password: str`, `mail_from: str`, `mail_to: str`, `report_tz: str`
  - `load_settings(environ: dict[str, str] | None = None) -> Settings`
  - `ConfigError(RuntimeError)`
  - pytest 픽스처 `settings` (conftest) — 채워진 `Settings` 인스턴스

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_config.py`:

```python
import pytest

from report_pipeline.config import ConfigError, load_settings

_ENV = {
    "SMTP_HOST": "smtp.gmail.com",
    "SMTP_PORT": "587",
    "SMTP_USER": "sender@gmail.com",
    "SMTP_PASSWORD": "app-password-1234",
    "MAIL_FROM": "sender@gmail.com",
    "MAIL_TO": "me@company.example",
    "REPORT_TZ": "Asia/Seoul",
}


def test_load_settings_reads_all_fields():
    settings = load_settings(dict(_ENV))
    assert settings.smtp_port == 587
    assert settings.smtp_host == "smtp.gmail.com"
    assert settings.mail_to == "me@company.example"
    assert settings.report_tz == "Asia/Seoul"


def test_load_settings_missing_lists_every_missing_name():
    env = dict(_ENV)
    del env["SMTP_USER"]
    del env["MAIL_TO"]
    with pytest.raises(ConfigError) as exc:
        load_settings(env)
    assert "SMTP_USER" in str(exc.value)
    assert "MAIL_TO" in str(exc.value)


def test_load_settings_treats_blank_as_missing():
    env = dict(_ENV)
    env["SMTP_PASSWORD"] = "   "
    with pytest.raises(ConfigError):
        load_settings(env)


def test_load_settings_strips_wrapping_quotes():
    env = dict(_ENV)
    env["SMTP_PASSWORD"] = '"quoted-secret"'
    env["MAIL_TO"] = "'me@company.example'"
    settings = load_settings(env)
    assert settings.smtp_password == "quoted-secret"
    assert settings.mail_to == "me@company.example"


def test_load_settings_rejects_non_integer_port():
    env = dict(_ENV)
    env["SMTP_PORT"] = "abc"
    with pytest.raises(ConfigError):
        load_settings(env)
```

- [ ] **Step 2: 테스트 실행 → 실패 확인**

Run: `pytest tests/test_config.py -v`
Expected: FAIL — `ImportError: cannot import name 'ConfigError'`

- [ ] **Step 3: `config.py` 구현**

```python
from __future__ import annotations

import os
from dataclasses import dataclass

_REQUIRED = [
    "SMTP_HOST",
    "SMTP_PORT",
    "SMTP_USER",
    "SMTP_PASSWORD",
    "MAIL_FROM",
    "MAIL_TO",
    "REPORT_TZ",
]


class ConfigError(RuntimeError):
    """필수 환경변수가 없거나 형식이 잘못된 경우."""


@dataclass(frozen=True)
class Settings:
    smtp_host: str
    smtp_port: int
    smtp_user: str
    smtp_password: str
    mail_from: str
    mail_to: str
    report_tz: str


def _unquote(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
        return value[1:-1]
    return value


def load_settings(environ: dict[str, str] | None = None) -> Settings:
    env = os.environ if environ is None else environ
    present = {
        key: _unquote(env[key])
        for key in _REQUIRED
        if env.get(key, "").strip()
    }
    missing = [key for key in _REQUIRED if key not in present]
    if missing:
        raise ConfigError(f"필수 환경변수 누락: {', '.join(missing)}")

    try:
        port = int(present["SMTP_PORT"])
    except ValueError as error:
        raise ConfigError(
            f"SMTP_PORT는 정수여야 합니다: {present['SMTP_PORT']!r}"
        ) from error

    return Settings(
        smtp_host=present["SMTP_HOST"],
        smtp_port=port,
        smtp_user=present["SMTP_USER"],
        smtp_password=present["SMTP_PASSWORD"],
        mail_from=present["MAIL_FROM"],
        mail_to=present["MAIL_TO"],
        report_tz=present["REPORT_TZ"],
    )
```

- [ ] **Step 4: 공용 픽스처 작성**

`tests/conftest.py`:

```python
import pytest

from report_pipeline.config import Settings


@pytest.fixture
def settings() -> Settings:
    return Settings(
        smtp_host="smtp.example.com",
        smtp_port=587,
        smtp_user="sender@example.com",
        smtp_password="secret",
        mail_from="sender@example.com",
        mail_to="me@company.example",
        report_tz="Asia/Seoul",
    )
```

- [ ] **Step 5: 테스트 실행 → 통과 확인**

Run: `pytest tests/test_config.py -v`
Expected: PASS — 5개 테스트

- [ ] **Step 6: 커밋**

```bash
git add src/report_pipeline/config.py tests/conftest.py tests/test_config.py
git commit -m "feat: 환경변수 설정 로딩 (Settings, load_settings, 따옴표 제거)"
```

---

## Task 3: 더미 데이터 소스 (`datasource.py`)

**Files:**
- Create: `src/report_pipeline/datasource.py`
- Test: `tests/test_datasource.py`

**Interfaces:**
- Consumes: `report_pipeline.config.Settings`
- Produces:
  - `COLUMNS: list[str]` = `["date", "app_name", "dau", "installs", "revenue"]`
  - `APPS: list[str]`, `LOOKBACK_DAYS: int`
  - `fetch(settings: Settings, run_date: datetime.date) -> pandas.DataFrame` — 행 수 `LOOKBACK_DAYS * len(APPS)`, 결정적

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_datasource.py`:

```python
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
```

- [ ] **Step 2: 테스트 실행 → 실패 확인**

Run: `pytest tests/test_datasource.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'report_pipeline.datasource'`

- [ ] **Step 3: `datasource.py` 구현**

```python
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
```

- [ ] **Step 4: 테스트 실행 → 통과 확인**

Run: `pytest tests/test_datasource.py -v`
Expected: PASS — 4개 테스트

- [ ] **Step 5: 커밋**

```bash
git add src/report_pipeline/datasource.py tests/test_datasource.py
git commit -m "feat: 결정적 더미 데이터 소스 (2단계 교체 지점)"
```

---

## Task 4: 데이터 가공 (`transform.py`)

**Files:**
- Create: `src/report_pipeline/transform.py`
- Test: `tests/test_transform.py`

**Interfaces:**
- Consumes: `datasource.COLUMNS` 스키마의 `pandas.DataFrame`
- Produces:
  - `ReportData` (dataclass): `sheets: dict[str, pandas.DataFrame]`, `summary: dict[str, object]`
  - `METRIC_COLUMNS: list[str]` = `["dau", "installs", "revenue"]`
  - `transform(raw: pandas.DataFrame, run_date: datetime.date, tz: str = "Asia/Seoul") -> ReportData`
  - `summary` 키(정확히): `row_count: int`, `app_count: int`, `date_range: list[str]` (`[min, max]`), `run_date: str` (ISO), `generated_at: str` (ISO, tz-aware)
  - `sheets` 키(정확히, 순서): `"원본"`, `"요약"`

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_transform.py`:

```python
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
```

- [ ] **Step 2: 테스트 실행 → 실패 확인**

Run: `pytest tests/test_transform.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'report_pipeline.transform'`

- [ ] **Step 3: `transform.py` 구현**

```python
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from zoneinfo import ZoneInfo

import pandas as pd

METRIC_COLUMNS = ["dau", "installs", "revenue"]


@dataclass
class ReportData:
    sheets: dict[str, pd.DataFrame]
    summary: dict[str, object]


def transform(raw: pd.DataFrame, run_date: date, tz: str = "Asia/Seoul") -> ReportData:
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
```

- [ ] **Step 4: 테스트 실행 → 통과 확인**

Run: `pytest tests/test_transform.py -v`
Expected: PASS — 4개 테스트

- [ ] **Step 5: 커밋**

```bash
git add src/report_pipeline/transform.py tests/test_transform.py
git commit -m "feat: 데이터 가공 — 원본/요약 시트 + 메타데이터"
```

---

## Task 5: 엑셀 리포트 생성 (`report.py`)

**Files:**
- Create: `src/report_pipeline/report.py`
- Test: `tests/test_report.py`

**Interfaces:**
- Consumes: `report_pipeline.transform.ReportData`
- Produces:
  - `build_workbook(data: ReportData, out_dir: pathlib.Path, run_date: datetime.date) -> pathlib.Path`
  - 반환 경로의 파일명: `mobileindex_report_{run_date:%Y-%m-%d}.xlsx`
  - 각 `data.sheets` 항목이 동일 이름의 시트로 기록되고, 헤더 행 볼드 + 상단 틀고정(`A2`)

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_report.py`:

```python
from datetime import date

import pandas as pd
from openpyxl import load_workbook

from report_pipeline.report import build_workbook
from report_pipeline.transform import ReportData


def _sample_data() -> ReportData:
    raw = pd.DataFrame(
        [
            {"date": "2026-09-09", "app_name": "앱A", "dau": 200, "installs": 20, "revenue": 2.0},
            {"date": "2026-09-09", "app_name": "앱B", "dau": 50, "installs": 5, "revenue": 0.5},
        ],
        columns=["date", "app_name", "dau", "installs", "revenue"],
    )
    summary_sheet = pd.DataFrame(
        [
            {"app_name": "앱A", "dau": 200, "installs": 20, "revenue": 2.0},
            {"app_name": "앱B", "dau": 50, "installs": 5, "revenue": 0.5},
        ]
    )
    return ReportData(sheets={"원본": raw, "요약": summary_sheet}, summary={"row_count": 2})


def test_build_workbook_uses_date_in_filename(tmp_path):
    path = build_workbook(_sample_data(), tmp_path, date(2026, 9, 9))
    assert path.name == "mobileindex_report_2026-09-09.xlsx"
    assert path.exists()


def test_build_workbook_creates_out_dir_if_missing(tmp_path):
    target = tmp_path / "nested" / "build"
    path = build_workbook(_sample_data(), target, date(2026, 9, 9))
    assert path.parent == target
    assert path.exists()


def test_workbook_sheets_and_header_formatting(tmp_path):
    path = build_workbook(_sample_data(), tmp_path, date(2026, 9, 9))
    workbook = load_workbook(path)
    assert workbook.sheetnames == ["원본", "요약"]
    sheet = workbook["원본"]
    assert sheet["A1"].value == "date"
    assert sheet["A1"].font.bold is True
    assert sheet.freeze_panes == "A2"


def test_workbook_row_values_roundtrip(tmp_path):
    path = build_workbook(_sample_data(), tmp_path, date(2026, 9, 9))
    workbook = load_workbook(path)
    sheet = workbook["요약"]
    assert sheet["A2"].value == "앱A"
    assert sheet["B2"].value == 200
```

- [ ] **Step 2: 테스트 실행 → 실패 확인**

Run: `pytest tests/test_report.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'report_pipeline.report'`

- [ ] **Step 3: `report.py` 구현**

```python
from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet

from .transform import ReportData

_MAX_COLUMN_WIDTH = 50


def build_workbook(data: ReportData, out_dir: Path, run_date: date) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"mobileindex_report_{run_date:%Y-%m-%d}.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, frame in data.sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
            _format_sheet(writer.sheets[sheet_name], frame)
    return path


def _format_sheet(worksheet: Worksheet, frame: pd.DataFrame) -> None:
    worksheet.freeze_panes = "A2"
    for column_index, column_name in enumerate(frame.columns, start=1):
        header_cell = worksheet.cell(row=1, column=column_index)
        header_cell.font = Font(bold=True)
        values = frame.iloc[:, column_index - 1].tolist()
        width = max([len(str(column_name))] + [len(str(value)) for value in values])
        worksheet.column_dimensions[header_cell.column_letter].width = min(
            width + 2, _MAX_COLUMN_WIDTH
        )
```

- [ ] **Step 4: 테스트 실행 → 통과 확인**

Run: `pytest tests/test_report.py -v`
Expected: PASS — 4개 테스트

- [ ] **Step 5: 커밋**

```bash
git add src/report_pipeline/report.py tests/test_report.py
git commit -m "feat: 엑셀 워크북 생성 — 시트별 기록 + 헤더 서식"
```

---

## Task 6: 이메일 발송 (`notify.py`)

**Files:**
- Create: `src/report_pipeline/notify.py`
- Test: `tests/test_notify.py`

**Interfaces:**
- Consumes: `report_pipeline.config.Settings`
- Produces:
  - `send_email(settings: Settings, subject: str, body: str, attachment: pathlib.Path) -> None`
  - `MAX_ATTEMPTS: int` = `3`
  - 내부적으로 `smtplib.SMTP` 사용(모듈 경로 `report_pipeline.notify.smtplib` 로 패치 가능), `time.sleep`으로 백오프(패치 가능)
  - 실패 시 `RuntimeError` 전파

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_notify.py`:

```python
from datetime import date
from unittest.mock import patch

import pytest

from report_pipeline.notify import MAX_ATTEMPTS, send_email


def _make_attachment(tmp_path):
    path = tmp_path / "mobileindex_report_2026-09-09.xlsx"
    path.write_bytes(b"fake-xlsx-bytes")
    return path


def test_send_email_delivers_message_with_attachment(tmp_path, settings):
    attachment = _make_attachment(tmp_path)
    with patch("report_pipeline.notify.smtplib.SMTP") as smtp_class:
        server = smtp_class.return_value.__enter__.return_value
        send_email(settings, "제목", "본문", attachment)

    server.starttls.assert_called_once()
    server.login.assert_called_once_with(settings.smtp_user, settings.smtp_password)
    sent_message = server.send_message.call_args[0][0]
    assert sent_message["Subject"] == "제목"
    assert sent_message["From"] == settings.mail_from
    assert sent_message["To"] == settings.mail_to
    parts = list(sent_message.iter_attachments())
    assert len(parts) == 1
    assert parts[0].get_filename() == "mobileindex_report_2026-09-09.xlsx"


def test_send_email_retries_twice_then_raises(tmp_path, settings):
    attachment = _make_attachment(tmp_path)
    with patch("report_pipeline.notify.smtplib.SMTP", side_effect=OSError("boom")), patch(
        "report_pipeline.notify.time.sleep"
    ) as sleep:
        with pytest.raises(RuntimeError):
            send_email(settings, "제목", "본문", attachment)
    assert sleep.call_count == MAX_ATTEMPTS - 1


def test_send_email_succeeds_after_transient_failure(tmp_path, settings):
    attachment = _make_attachment(tmp_path)
    calls = {"n": 0}

    def flaky(*args, **kwargs):
        calls["n"] += 1
        if calls["n"] == 1:
            raise OSError("transient")
        from unittest.mock import MagicMock

        return MagicMock()

    with patch("report_pipeline.notify.smtplib.SMTP", side_effect=flaky), patch(
        "report_pipeline.notify.time.sleep"
    ):
        send_email(settings, "제목", "본문", attachment)
    assert calls["n"] == 2
```

- [ ] **Step 2: 테스트 실행 → 실패 확인**

Run: `pytest tests/test_notify.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'report_pipeline.notify'`

- [ ] **Step 3: `notify.py` 구현**

```python
from __future__ import annotations

import smtplib
import time
from email.message import EmailMessage
from pathlib import Path

from .config import Settings

MAX_ATTEMPTS = 3
_XLSX_MAINTYPE = "application"
_XLSX_SUBTYPE = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def send_email(settings: Settings, subject: str, body: str, attachment: Path) -> None:
    message = _build_message(settings, subject, body, Path(attachment))
    last_error: Exception | None = None
    for attempt in range(MAX_ATTEMPTS):
        try:
            _deliver(settings, message)
            return
        except Exception as error:  # 재시도 목적의 광범위 캐치
            last_error = error
            if attempt == MAX_ATTEMPTS - 1:
                break
            time.sleep(2 ** (attempt + 1))
    raise RuntimeError(f"이메일 전송 실패 ({MAX_ATTEMPTS}회 시도)") from last_error


def _build_message(settings: Settings, subject: str, body: str, attachment: Path) -> EmailMessage:
    message = EmailMessage()
    message["From"] = settings.mail_from
    message["To"] = settings.mail_to
    message["Subject"] = subject
    message.set_content(body)
    message.add_attachment(
        attachment.read_bytes(),
        maintype=_XLSX_MAINTYPE,
        subtype=_XLSX_SUBTYPE,
        filename=attachment.name,
    )
    return message


def _deliver(settings: Settings, message: EmailMessage) -> None:
    with smtplib.SMTP(settings.smtp_host, settings.smtp_port) as server:
        server.starttls()
        server.login(settings.smtp_user, settings.smtp_password)
        server.send_message(message)
```

- [ ] **Step 4: 테스트 실행 → 통과 확인**

Run: `pytest tests/test_notify.py -v`
Expected: PASS — 3개 테스트

- [ ] **Step 5: 커밋**

```bash
git add src/report_pipeline/notify.py tests/test_notify.py
git commit -m "feat: SMTP 이메일 발송 — STARTTLS + 첨부 + 재시도"
```

---

## Task 7: 오케스트레이션 (`pipeline.py`)

**Files:**
- Create: `src/report_pipeline/pipeline.py`
- Test: `tests/test_pipeline.py`

**Interfaces:**
- Consumes: `config.Settings`, `datasource.fetch`, `transform.transform`, `report.build_workbook`, `notify.send_email`
- Produces:
  - `BUILD_DIR: pathlib.Path` = `Path("build")`
  - `run(settings: Settings, run_date: datetime.date) -> dict` — 키: `attachment: str`, `sent_to: str`, `subject: str`, `row_count: int`
  - 제목: `[모바일인덱스 리포트] {run_date:%Y-%m-%d}`

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_pipeline.py`:

```python
from datetime import date
from pathlib import Path

from report_pipeline.pipeline import run


def test_run_wires_all_stages(settings, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    captured = {}

    def fake_send_email(_settings, subject, body, attachment):
        captured["subject"] = subject
        captured["body"] = body
        captured["attachment"] = attachment

    monkeypatch.setattr(
        "report_pipeline.pipeline.notify.send_email", fake_send_email
    )

    result = run(settings, date(2026, 9, 9))

    assert result["subject"] == "[모바일인덱스 리포트] 2026-09-09"
    assert Path(result["attachment"]).name == "mobileindex_report_2026-09-09.xlsx"
    assert Path(result["attachment"]).exists()
    assert result["sent_to"] == settings.mail_to
    assert result["row_count"] == 21  # datasource: 7일 * 3앱
    assert "행 수: 21" in captured["body"]
    assert captured["attachment"] == Path(result["attachment"])


def test_run_writes_into_build_dir(settings, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(
        "report_pipeline.pipeline.notify.send_email",
        lambda *args, **kwargs: None,
    )
    result = run(settings, date(2026, 9, 9))
    assert Path(result["attachment"]).parent == Path("build")
```

- [ ] **Step 2: 테스트 실행 → 실패 확인**

Run: `pytest tests/test_pipeline.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'report_pipeline.pipeline'`

- [ ] **Step 3: `pipeline.py` 구현**

```python
from __future__ import annotations

import logging
from datetime import date
from pathlib import Path

from . import datasource, notify, report, transform
from .config import Settings

logger = logging.getLogger(__name__)

BUILD_DIR = Path("build")


def run(settings: Settings, run_date: date) -> dict:
    logger.info("데이터 취득 시작 run_date=%s", run_date)
    raw = datasource.fetch(settings, run_date)

    logger.info("가공 시작 rows=%d", len(raw))
    data = transform.transform(raw, run_date, settings.report_tz)

    logger.info("엑셀 생성 시작")
    attachment = report.build_workbook(data, BUILD_DIR, run_date)

    subject = f"[모바일인덱스 리포트] {run_date:%Y-%m-%d}"
    body = _build_body(run_date, data.summary)

    logger.info("이메일 발송 시작 to=%s", settings.mail_to)
    notify.send_email(settings, subject, body, attachment)

    logger.info("완료")
    return {
        "attachment": str(attachment),
        "sent_to": settings.mail_to,
        "subject": subject,
        "row_count": int(data.summary["row_count"]),
    }


def _build_body(run_date: date, summary: dict) -> str:
    start, end = summary["date_range"]
    return (
        f"{run_date:%Y-%m-%d} 모바일인덱스 리포트입니다.\n\n"
        f"- 데이터 기간: {start} ~ {end}\n"
        f"- 행 수: {summary['row_count']}\n"
        f"- 앱 수: {summary['app_count']}\n"
        f"- 생성 시각: {summary['generated_at']}\n"
    )
```

- [ ] **Step 4: 테스트 실행 → 통과 확인**

Run: `pytest tests/test_pipeline.py -v`
Expected: PASS — 2개 테스트

- [ ] **Step 5: 전체 테스트 실행**

Run: `pytest -v`
Expected: PASS — 지금까지의 모든 테스트

- [ ] **Step 6: 커밋**

```bash
git add src/report_pipeline/pipeline.py tests/test_pipeline.py
git commit -m "feat: 파이프라인 오케스트레이션 (fetch→transform→build→send)"
```

---

## Task 8: CLI 진입점 (`__main__.py`)

**Files:**
- Create: `src/report_pipeline/__main__.py`
- Test: `tests/test_main.py`

**Interfaces:**
- Consumes: `config.load_settings`, `pipeline.run`
- Produces:
  - `main(argv: list[str] | None = None) -> int`
  - CLI: `python -m report_pipeline [--date YYYY-MM-DD]`
  - `--date` 미지정 시 `settings.report_tz` 기준 오늘
  - 잘못된 날짜 형식 → `argparse`가 `SystemExit(2)`
  - 성공 시 요약 dict를 JSON 한 줄로 stdout 출력, 반환 `0`

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_main.py`:

```python
import json

import pytest

from report_pipeline.__main__ import main

_ENV = {
    "SMTP_HOST": "smtp.gmail.com",
    "SMTP_PORT": "587",
    "SMTP_USER": "sender@gmail.com",
    "SMTP_PASSWORD": "app-password-1234",
    "MAIL_FROM": "sender@gmail.com",
    "MAIL_TO": "me@company.example",
    "REPORT_TZ": "Asia/Seoul",
}


def test_main_rejects_bad_date_format():
    with pytest.raises(SystemExit) as exc:
        main(["--date", "2026/09/09"])
    assert exc.value.code == 2


def test_main_runs_with_explicit_date(monkeypatch, capsys):
    for key, value in _ENV.items():
        monkeypatch.setenv(key, value)
    monkeypatch.setattr(
        "report_pipeline.__main__.run",
        lambda settings, run_date: {"run_date": run_date.isoformat(), "ok": True},
    )
    exit_code = main(["--date", "2026-09-09"])
    assert exit_code == 0
    printed = json.loads(capsys.readouterr().out)
    assert printed["run_date"] == "2026-09-09"


def test_main_defaults_date_to_today_in_report_tz(monkeypatch):
    for key, value in _ENV.items():
        monkeypatch.setenv(key, value)
    seen = {}
    monkeypatch.setattr(
        "report_pipeline.__main__.run",
        lambda settings, run_date: seen.setdefault("run_date", run_date) and {} or {},
    )
    main([])
    assert "run_date" in seen


def test_main_raises_config_error_when_env_missing(monkeypatch):
    for key in _ENV:
        monkeypatch.delenv(key, raising=False)
    monkeypatch.setattr("report_pipeline.__main__._load_dotenv", lambda: None)
    from report_pipeline.config import ConfigError

    with pytest.raises(ConfigError):
        main(["--date", "2026-09-09"])
```

- [ ] **Step 2: 테스트 실행 → 실패 확인**

Run: `pytest tests/test_main.py -v`
Expected: FAIL — `ModuleNotFoundError` 또는 `ImportError`

- [ ] **Step 3: `__main__.py` 구현**

```python
from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime
from zoneinfo import ZoneInfo

from .config import load_settings
from .pipeline import run


def _parse_date(value: str) -> date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as error:
        raise argparse.ArgumentTypeError(
            f"날짜 형식은 YYYY-MM-DD 여야 합니다: {value!r}"
        ) from error


def _load_dotenv() -> None:
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    load_dotenv(".env.local")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="report_pipeline")
    parser.add_argument(
        "--date",
        dest="run_date",
        type=_parse_date,
        default=None,
        help="리포트 기준일 (YYYY-MM-DD). 생략 시 REPORT_TZ 기준 오늘.",
    )
    args = parser.parse_args(argv)

    _load_dotenv()
    settings = load_settings()
    run_date = args.run_date or datetime.now(ZoneInfo(settings.report_tz)).date()

    summary = run(settings, run_date)
    print(json.dumps(summary, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: 테스트 실행 → 통과 확인**

Run: `pytest tests/test_main.py -v`
Expected: PASS — 4개 테스트

- [ ] **Step 5: CLI 수동 확인**

Run: `python -m report_pipeline --date 2026-13-01`
Expected: 비정상 종료, stderr에 "날짜 형식은 YYYY-MM-DD" 메시지

- [ ] **Step 6: 커밋**

```bash
git add src/report_pipeline/__main__.py tests/test_main.py
git commit -m "feat: CLI 진입점 (--date, dotenv 로딩, JSON 요약 출력)"
```

---

## Task 9: GitHub Actions 워크플로 + README

**Files:**
- Create: `.github/workflows/ci.yml`
- Create: `.github/workflows/report.yml`
- Create: `README.md`

**Interfaces:**
- Consumes: `python -m report_pipeline` CLI, 7개 Secrets
- Produces: 없음 (배포 산출물)

- [ ] **Step 1: `ci.yml` 작성**

```yaml
name: CI

on:
  push:
  pull_request:

jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.11"
      - run: pip install -e ".[dev]"
      - run: pytest -v
```

- [ ] **Step 2: `report.yml` 작성**

```yaml
name: Report

on:
  workflow_dispatch:
    inputs:
      date:
        description: "리포트 기준일 (YYYY-MM-DD, 비우면 오늘)"
        required: false
        type: string
  # 발송 주기 미정 — 데이터 특성 확인 후 아래 cron 활성화 (UTC 기준).
  # KST 08:00 = 전날 23:00 UTC → "0 23 * * *"
  # schedule:
  #   - cron: "0 23 * * *"

jobs:
  send-report:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.11"
      - run: pip install .
      - name: Run pipeline
        env:
          SMTP_HOST: ${{ secrets.SMTP_HOST }}
          SMTP_PORT: ${{ secrets.SMTP_PORT }}
          SMTP_USER: ${{ secrets.SMTP_USER }}
          SMTP_PASSWORD: ${{ secrets.SMTP_PASSWORD }}
          MAIL_FROM: ${{ secrets.MAIL_FROM }}
          MAIL_TO: ${{ secrets.MAIL_TO }}
          REPORT_TZ: ${{ secrets.REPORT_TZ }}
        run: |
          if [ -n "${{ inputs.date }}" ]; then
            python -m report_pipeline --date "${{ inputs.date }}"
          else
            python -m report_pipeline
          fi
      - name: Upload workbook on failure
        if: failure()
        uses: actions/upload-artifact@v4
        with:
          name: report-xlsx
          path: build/*.xlsx
          if-no-files-found: ignore
```

- [ ] **Step 3: `README.md` 작성**

```markdown
# 모바일인덱스 리포트 자동화

외부 SaaS(모바일인덱스) 데이터를 엑셀 리포트로 만들어 이메일로 주기 발송하는
파이프라인. GitHub Actions에서 무인 실행된다.

## 현재 상태 (1단계)

데이터 소스는 더미 스텁이다. 리포트 생성·이메일 발송·스케줄 실행이 동작한다.
2단계에서 `src/report_pipeline/datasource.py`의 `fetch()`를 모바일인덱스
실연동(API 또는 CSV)으로 교체한다.

## 구조

`datasource → transform → report → notify` 4개 모듈을 `pipeline.run()`이
순서대로 호출한다. CLI 진입점은 `python -m report_pipeline`.

## 로컬 실행

```bash
pip install -e ".[dev]"
cp .env.example .env.local   # 값은 직접 입력, 커밋 금지
python -m report_pipeline --date 2026-09-09
pytest
```

## GitHub Actions 설정

저장소 Settings → Secrets and variables → Actions 에 등록:

| Secret | 예시 |
|---|---|
| `SMTP_HOST` | `smtp.gmail.com` |
| `SMTP_PORT` | `587` |
| `SMTP_USER` | Gmail 주소 |
| `SMTP_PASSWORD` | Gmail 앱 비밀번호(16자, 2단계 인증 필요) |
| `MAIL_FROM` | 발신 주소 |
| `MAIL_TO` | 수신 주소 (내부망 업무메일) |
| `REPORT_TZ` | `Asia/Seoul` |

- 수동 실행: Actions 탭 → Report → Run workflow (원하면 `date` 입력).
- 주기 실행: 발송 주기 확정 후 `.github/workflows/report.yml`의 `schedule`
  주석을 해제한다. **cron은 UTC 기준** — KST 08:00은 `0 23 * * *`.
- 실행이 실패하면 GitHub이 저장소 소유자에게 알림 메일을 보낸다.
```

- [ ] **Step 4: 전체 테스트 실행**

Run: `pytest -v`
Expected: PASS — 전체

- [ ] **Step 5: 커밋**

```bash
git add .github/workflows/ci.yml .github/workflows/report.yml README.md
git commit -m "ci: GitHub Actions 워크플로(CI + 리포트) 및 README"
```

---

## Self-Review

**1. Spec coverage:**

| 스펙 항목 | 구현 태스크 |
|---|---|
| GitHub Actions 스케줄 + `workflow_dispatch` | 9 (`report.yml`) |
| 4모듈 파이프라인 + 오케스트레이터 | 3–7 |
| `config` — env 로딩·검증·따옴표 제거 | 2 |
| `datasource.fetch()` 더미 + 2단계 교체 주석 | 3 |
| `transform` — 원본/요약 시트 + summary | 4 |
| `report.build_workbook()` — 파일명·서식·틀고정 | 5 |
| `notify.send_email()` — STARTTLS·첨부·재시도 | 6 |
| `pipeline.run()` — 로그·제목·본문 | 7 |
| CLI `--date` + 오늘 기본값 + dotenv | 8 |
| 에러 처리(설정 누락 즉시 실패 / SMTP 재시도 / 실패 시 아티팩트) | 2, 6, 9 |
| 테스트 6개 파일 (config/datasource/transform/report/notify/pipeline) + main | 2–8 |
| `ci.yml` pytest | 9 |
| 의존성 pandas/openpyxl (+dev pytest/dotenv) | 1 |
| `.gitignore` — `.env.local`, `build/` | 1 |
| README — UTC↔KST 주의 | 9 |

스펙의 진입점 `python -m src.main` → **정제**: `python -m report_pipeline`
(`src/report_pipeline/__main__.py`). 표준 패키지 관용구이며 스펙의 의도
(무인 실행 가능한 단일 진입점)를 그대로 만족한다. 스펙 §3.7, §6, §10의
경로 표기는 이 계획 기준으로 읽는다.

**2. Placeholder scan:** `report.yml`의 주석 처리된 `schedule` cron은
스펙이 명시적으로 "주기 미정, 주석 처리"를 요구한 항목이므로 의도된 것.
그 외 "TBD/TODO/적절히 처리" 없음.

**3. Type consistency:** 확인 완료.
- `fetch(settings, run_date)` — Task 3 정의, Task 7 호출 일치.
- `transform(raw, run_date, tz)` — Task 4 정의, Task 7이 `settings.report_tz` 전달.
- `ReportData.sheets`/`.summary` — Task 4 정의, Task 5·7 소비 일치.
- `build_workbook(data, out_dir, run_date) -> Path` — Task 5 정의, Task 7 호출 일치.
- `send_email(settings, subject, body, attachment)` — Task 6 정의, Task 7 호출 일치.
- `run(settings, run_date) -> dict` (키 `attachment/sent_to/subject/row_count`) — Task 7 정의, Task 8 소비 일치.
- `summary` 키 이름(`row_count`, `app_count`, `date_range`, `run_date`, `generated_at`) — Task 4 정의, Task 7 `_build_body`에서 사용 일치.
