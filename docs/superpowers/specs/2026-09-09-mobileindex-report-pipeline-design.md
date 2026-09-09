# 모바일인덱스 리포트 자동화 — 1단계: 이메일 파이프라인 설계

- 작성일: 2026-09-09
- 상태: 승인됨 (사용자 리뷰 대기)
- 범위: 1단계만. 2단계(모바일인덱스 실데이터 취득)는 별도 스펙.

## 1. 배경과 목표

외부 SaaS(모바일인덱스)에서 데이터를 내려받아 가공하고, 엑셀 리포트로 만들어
이메일로 주기적으로 자동 발송한다. 이 작업은 사용자 PC가 꺼져 있어도 항상
클라우드에서 수행되어야 한다.

전체 프로젝트를 두 단계로 분리한다.

- **1단계 (이 스펙):** 입력 데이터가 주어졌을 때 → 엑셀 리포트 생성 → 이메일 발송.
  클라우드에서 스케줄 실행. 데이터 취득부는 더미 데이터 스텁으로 대체하고
  인터페이스만 확정한다.
- **2단계 (별도 스펙):** 모바일인덱스 데이터 취득. 두 경로가 존재한다 —
  (A) 공식 API 또는 원격 MCP 엔드포인트 + API 키, (B) 웹 대시보드 CSV 내보내기.
  제공 데이터 범위가 달라 나중에 검증 후 확정한다. 확정되면 1단계의
  `datasource.fetch()` 몸통만 교체한다.

### 1단계 성공 기준

1. GitHub Actions에서 수동 트리거(`workflow_dispatch`) 시, 더미 데이터로 만든
   엑셀 파일이 지정된 수신자 1명에게 이메일로 도착한다.
2. 스케줄(cron) 설정만 채우면 사용자 개입 없이 주기 실행된다.
3. 파이프라인 중간 실패 시 저장소 소유자가 자동으로 통보받는다.
4. 모든 모듈이 단위 테스트로 커버되고 CI에서 통과한다.

### 비목표 (1단계에서 하지 않음)

- 모바일인덱스 실제 API/CSV 연동
- 발송 주기 확정 (데이터 특성 확인 후 결정)
- 엑셀 차트/피벗 (실데이터 형태 확정 후)
- 다중 수신자 관리 UI
- 리포트 본문 템플릿 엔진

## 2. 아키텍처

### 실행 환경

- **GitHub Actions 스케줄 워크플로.** cron 트리거(UTC 기준) + `workflow_dispatch`
  수동 트리거. 서버 관리 불필요, 실행 로그·시크릿 관리 내장. 2단계에서
  Playwright 기반 CSV 로그인 자동화도 동일 러너에서 가능.
- 회당 실행 1~2분. GitHub 무료 한도 내에서 충분.

### 구조 원칙

순수 함수 4개 모듈 + 얇은 오케스트레이터. 프레임워크 없음.
데이터 소스는 함수 시그니처(`fetch(settings) -> DataFrame`) 뒤에 격리하여
2단계 교체 시 하위 모듈을 건드리지 않는다.

### 저장소 레이아웃

```
mobileindex-report/
├─ .github/workflows/
│  ├─ report.yml        # 스케줄 + 수동 실행 (본 파이프라인)
│  └─ ci.yml            # push/PR 시 pytest
├─ src/report_pipeline/
│  ├─ __init__.py
│  ├─ config.py         # 환경변수 로딩·검증 (Settings)
│  ├─ datasource.py     # fetch(settings, run_date) -> DataFrame   ← 2단계 교체 지점
│  ├─ transform.py      # transform(raw) -> ReportData
│  ├─ report.py         # build_workbook(data, out_dir) -> Path
│  ├─ notify.py         # send_email(settings, subject, body, attachment)
│  └─ pipeline.py       # run(settings, run_date) -> summary
├─ src/main.py          # python -m 진입점, --date 옵션
├─ tests/
│  ├─ test_config.py
│  ├─ test_datasource.py
│  ├─ test_transform.py
│  ├─ test_report.py
│  ├─ test_notify.py
│  └─ test_pipeline.py
├─ pyproject.toml
├─ .env.example
├─ .gitignore
└─ README.md
```

## 3. 모듈 계약

### 3.1 `config.py`

```python
@dataclass(frozen=True)
class Settings:
    smtp_host: str          # 예: smtp.gmail.com
    smtp_port: int           # 예: 587
    smtp_user: str
    smtp_password: str       # Gmail 앱 비밀번호 (16자)
    mail_from: str
    mail_to: str             # 수신자 1명
    report_tz: str           # 예: Asia/Seoul, 리포트 날짜 계산용

def load_settings() -> Settings
```

- 값은 환경변수에서 읽는다. 로컬 개발은 `.env.local`(gitignore됨),
  CI는 GitHub Secrets.
- 필수 변수 누락 시 `SystemExit` 또는 명시적 예외로 즉시 실패하고,
  **어떤 변수가 빠졌는지** 메시지에 포함한다.
- `.env` 계열 값이 따옴표로 감싸져 있으면 따옴표를 제거하고 읽는다.

### 3.2 `datasource.py`

```python
def fetch(settings: Settings, run_date: date) -> pandas.DataFrame
```

- **1단계:** `run_date` 기준 결정적(deterministic) 더미 DataFrame 반환.
  예: 최근 7일 날짜 인덱스 + 가짜 지표 컬럼 (`app_name`, `dau`, `installs`, `revenue`).
  랜덤이 아니라 `run_date`에서 파생된 고정 값 → 테스트 재현 가능.
- 모듈 docstring에 "2단계에서 이 함수 내부를 모바일인덱스 API(A) 또는
  CSV 파싱(B)으로 교체한다. 시그니처와 반환 스키마는 유지한다"를 명시.
- 반환 스키마(컬럼명·dtype)를 모듈 상수로 문서화하여 `transform`이 의존.

### 3.3 `transform.py`

```python
@dataclass
class ReportData:
    sheets: dict[str, pandas.DataFrame]   # 시트명 -> 데이터
    summary: dict[str, object]            # 이메일 본문·파일명에 쓸 메타

def transform(raw: pandas.DataFrame, run_date: date) -> ReportData
```

- **1단계:** 가벼운 가공.
  - `sheets["원본"]` = raw 그대로
  - `sheets["요약"]` = 지표별 합계/평균 집계표
  - `summary` = `{"row_count": ..., "date_range": ..., "generated_at": ...}`
- 2단계에서 실데이터 컬럼이 정해지면 이 모듈이 주 변경 대상.

### 3.4 `report.py`

```python
def build_workbook(data: ReportData, out_dir: Path, run_date: date) -> Path
```

- `pandas.ExcelWriter`(engine=openpyxl)로 `data.sheets`의 각 항목을
  별도 시트로 기록.
- 파일명: `mobileindex_report_{run_date:%Y-%m-%d}.xlsx`
- 기본 서식: 헤더 행 볼드, 컬럼 너비 자동(내용 길이 기반), 상단 행 틀고정.
- 반환: 생성된 파일의 `Path`.

### 3.5 `notify.py`

```python
def send_email(
    settings: Settings,
    subject: str,
    body: str,
    attachment: Path,
) -> None
```

- 표준 라이브러리 `smtplib` + `email.message.EmailMessage`.
- STARTTLS 사용. `settings.smtp_user` / `smtp_password`로 로그인.
- 첨부: xlsx MIME 타입
  (`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`).
- 본문은 평문(또는 단순 HTML). 템플릿 엔진 없음.
- 전송 실패 시 2회까지 지수 백오프(예: 2s, 4s) 재시도 후 예외 전파.

### 3.6 `pipeline.py`

```python
def run(settings: Settings, run_date: date) -> dict
```

- 순서: `fetch` → `transform` → `build_workbook` → `send_email`.
- 각 단계 시작·완료를 `logging`으로 기록.
- 이메일 제목/본문을 `ReportData.summary`에서 구성.
  예: 제목 `"[모바일인덱스 리포트] 2026-09-09"`.
- 반환: 실행 요약 dict (`{"attachment": <경로>, "sent_to": ..., "row_count": ...}`).
- 워크북은 저장소 루트 하위 `build/` 디렉터리에 생성 (gitignore됨).

### 3.7 `main.py`

```
python -m src.main [--date YYYY-MM-DD]
```

- `--date` 미지정 시 `settings.report_tz` 기준 오늘.
- `load_settings()` → `run(settings, run_date)` → 요약을 stdout에 출력.
- 예외 발생 시 비정상 종료 코드(→ GitHub Actions 실패 처리).

## 4. 데이터 흐름

```
GitHub Secrets ─► config.load_settings() ─► Settings
   ─► datasource.fetch(settings, run_date) ─► DataFrame
   ─► transform(df, run_date) ─► ReportData
   ─► report.build_workbook(data, out_dir, run_date) ─► report.xlsx
   ─► notify.send_email(settings, subject, body, attachment=report.xlsx)
   ─► Gmail SMTP ─► 수신자(내부망 업무메일) ─► (사용자측 자동 전달 규칙)
```

## 5. 에러 처리

| 상황 | 동작 |
|---|---|
| 필수 환경변수 누락 | 즉시 실패, 누락 변수명 출력 |
| `fetch`/`transform`/`build` 예외 | 전파 → GitHub Actions 실행 실패 → 저장소 소유자에게 자동 알림 메일 |
| SMTP 전송 실패 | 2회 지수 백오프 재시도 후 예외 전파 |
| 워크북 생성됨 but 이후 실패 | 워크플로가 xlsx를 아티팩트로 업로드(디버깅용) |

별도의 상태 저장소/알림 채널은 두지 않는다. GitHub Actions의 기본 실패
알림을 1차 안전망으로 사용한다.

## 6. 스케줄링과 설정

### `report.yml`

- 트리거:
  - `schedule: - cron: "<TBD>"` — **주기 미정.** 데이터 특성 확인 후 확정.
    확정 전까지 주석 처리하거나 매우 낮은 빈도(예: 매주 월요일)로 둔다.
  - `workflow_dispatch:` — 수동 실행. 선택 입력 `date`(YYYY-MM-DD).
- 스텝: checkout → setup-python → `pip install .` → `python -m src.main`
  → (실패 시) xlsx 아티팩트 업로드.
- Secrets: `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`,
  `MAIL_FROM`, `MAIL_TO`, `REPORT_TZ`.
- README에 UTC ↔ KST 환산 주의: KST 08:00 = cron `0 23 * * *`(전날 UTC).

### `.env.example`

모든 필요한 키를 값 없이 나열. 실제 값은 각자 `.env.local`에 입력하며
대화창에 붙여넣지 않는다.

## 7. 테스트 (TDD)

| 파일 | 검증 내용 |
|---|---|
| `test_config.py` | 정상 env → `Settings` 채워짐; 누락 시 명확한 오류; 따옴표 제거 |
| `test_datasource.py` | 더미가 기대 컬럼·dtype의 비어있지 않은 df를 결정적으로 반환 |
| `test_transform.py` | 알려진 df 입력 → `sheets` 키·`summary` 값 정확 |
| `test_report.py` | `build_workbook`가 파일 생성 → openpyxl로 재오픈, 시트명·셀 값 확인; 파일명이 날짜와 일치 |
| `test_notify.py` | `smtplib.SMTP` mock → 헤더(From/To/Subject)·첨부 존재·STARTTLS 호출·재시도 동작 |
| `test_pipeline.py` | `datasource`·`smtplib` mock → end-to-end, 첨부 파일명이 `run_date`와 일치, summary 반환 |

- 실 SMTP 통합 테스트는 두지 않음(또는 env 플래그 뒤에 두고 기본 skip).
- `ci.yml`: push/PR 시 `pytest` 실행.

## 8. 의존성

- 런타임: `pandas`, `openpyxl`
- 개발: `pytest`, `python-dotenv`
- Python 3.11+ (`zoneinfo` 표준 라이브러리 사용)

## 9. 로컬 개발

```
cp .env.example .env.local   # 값은 직접 입력
python -m src.main --date 2026-09-09
pytest
```

`.gitignore`: `.env.local`, `build/`, `__pycache__/`, `*.xlsx`, `.pytest_cache/`

## 10. 2단계 연결 지점 (참고)

2단계에서 바뀌는 것은 원칙적으로 `datasource.py` 하나.
경로 A(API/MCP)면 HTTP 클라이언트 + 인증, 경로 B(CSV)면 로그인 자동화 +
CSV 파싱이 `fetch()` 내부에 들어간다. 반환 스키마가 달라지면
`transform.py`가 그에 맞춰 갱신되고, 나머지(`report`, `notify`, `pipeline`)는
변경 없음.
