# 작업 로그

프로젝트 진행 기록. 최신 항목이 위.

---

## 2026-09-09 ~ 09-10 — 1단계 구현 + GitHub/Apps Script 배포

### 한 일

**1단계 파이프라인 구현 (완료, master 병합)**
- 브레인스토밍 → 스펙(`docs/superpowers/specs/2026-09-09-*`) → 계획(`docs/superpowers/plans/2026-09-09-*`)
- SDD(subagent-driven-development)로 9개 태스크 TDD 구현:
  스캐폴딩 → `config` → `datasource` → `transform` → `report` → `notify` → `pipeline` → `__main__`(CLI) → GitHub Actions/README
- 최종 whole-branch 리뷰(opus) Important 5건 반영: 로깅 설정, SMTP 타임아웃,
  transform 컬럼 가드, 워크플로 입력 하드닝, tz 테스트
- 테스트 30개 통과. `feat/phase1-email-pipeline` → `master` `--no-ff` 병합

**GitHub 배포 (완료)**
- 저장소: `github.com/Mongshirius/mobileindex-report` (public)
- Secrets 7개 등록 (`gh secret set -f .env.local`): `SMTP_HOST/PORT/USER/PASSWORD`, `MAIL_FROM/TO`, `REPORT_TZ`
- `report.yml`: `workflow_dispatch`(수동) + `schedule: "0 23 * * *"`(매일 KST 08:00)
- 검증: 수동 실행(`workflow_dispatch`)으로 GitHub 러너에서 전체 체인
  (`pip install` → 파이프라인 → xlsx → 이메일) 성공. run 34358348456

**Apps Script 버전 (코드 완료, 배포는 사용자)**
- `apps-script/{Code.gs, appsscript.json, README.md}`
- Python 파이프라인의 GAS 포팅. `MailApp`(SMTP 불필요) + 시간 트리거
- 순수 로직(더미데이터·집계·컬럼가드·결정성) Node로 검증 통과
- GAS API 의존부는 편집기에서 실행해봐야 검증됨

**발송 검증 (2026-09-09)**
- 로컬: `python -m report_pipeline` → `kaka_moka@naver.com` 수신 OK
- GitHub 러너: 수동 실행 → 동일 주소 수신 OK

### 내린 결정 (Rulings)

1. **Python 인터프리터**: 시스템 `py` 런처엔 3.6/3.8만 있음. 스펙이 3.11+(`zoneinfo`)
   요구 → `C:\conda\python.exe`(3.13.11) 기반 `.venv/` 사용. 모든 명령은
   `.venv\Scripts\python.exe -m ...`.
2. **CLI 진입점**: 스펙의 `python -m src.main` → `python -m report_pipeline`
   (`src/report_pipeline/__main__.py`)로 정제. 표준 패키지 관용구.
3. **config 따옴표/공백 검사 순서**: 계획서가 지정한 코드와 다르게, 따옴표 제거를
   먼저 하고 공백 검사하도록 수정(스펙 §3.1 의도). `KEY=""` → "누락"으로 처리.
4. **작업 격리**: 신규 전용 저장소라 별도 worktree 없이 피처 브랜치만 사용.
5. **저장소 공개 범위**: public (Actions 무제한). 코드에 비밀 없음(Secrets 별도).
6. **발송 주기**: 일 1회 KST 08:00. (원래 미정 → `*/5` 테스트가 GitHub throttling에
   막혀 실용성 없음이 확인되어 일간으로 확정)

### 알려진 문제 / 함정

- **GitHub cron `*/5` throttling**: 새 public 저장소의 고빈도 cron을 GitHub이
  남용 방지로 심하게 지연/스킵함. 90분간 스케줄 실행 0건. **일간 cron은 정상.**
- **비ASCII 워크플로 파일**: `report.yml`에 한글 `description`/`→` 기호가 있으면
  GitHub이 에러 없이 워크플로 자체를 무시함("This workflow does not exist").
  → `.github/workflows/*.yml`은 ASCII만. push 후 `gh api .../actions/workflows`로
  등록 확인 필수. (`.gitattributes`로 `*.yml eol=lf` 고정)
- **Node 20 deprecation 경고**: `actions/checkout@v4`, `actions/setup-python@v5`가
  Node 20 사용 → 경고만, 동작엔 무관. 나중에 `@v5`/`@v6`로 상향 가능.
- **Apps Script 할당량**: 일반 Gmail 하루 수신자 100통. 5분 간격(288/일)은 초과.
  일/주 단위는 무관.

### 남은 일 (deferred / 후속)

- **2단계**: 모바일인덱스 실연동. 경로 A(API/MCP) 또는 B(CSV 로그인). 검증 후
  `datasource.fetch()` (Python) / `datasource_()` (GAS) 몸통만 교체. 스키마 바뀌면
  `transform` 도 갱신. 경로 B는 Apps Script 불가(헤드리스 브라우저 없음).
- **Apps Script 배포**: `script.google.com`에 붙여넣기 → `MAIL_TO` 스크립트 속성
  → `runPipeline` → `installTrigger`. (`apps-script/README.md`)
- **deferred minors** (병합됨, 비차단): `test_transform`의 `generated_at` 약한 검증,
  `test_notify` 안 쓰는 import, `build_workbook` 빈 시트 가드 없음, README venv 단계,
  `test_main` 격리 2건, F5 테스트가 `pipeline.run` 경로 미검증.
- **플랫폼 후속 고려**: `REPORT_TZ`를 Secret 대신 repo Variable로; 실패 아티팩트가
  실데이터 노출한다는 README 경고(2단계 후).
- 정확한 스케줄이 중요해지면: GCP Cloud Run Jobs + Cloud Scheduler, 또는 Apps Script.

### 로컬 상태 메모

- `.env.local` — 로컬에만 존재(gitignore), 실제 SMTP 값 채워져 있음.
  발신 `mongshil0422@gmail.com`(앱 비번), 수신 `kaka_moka@naver.com`.
