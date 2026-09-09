# Apps Script 버전

`report_pipeline` (Python) 파이프라인을 Google Apps Script로 포팅한 것.
스케줄 실행이 GitHub Actions cron보다 안정적이고, SMTP/앱 비밀번호가 필요 없다
(MailApp이 실행 계정 권한으로 발송).

- `Code.gs` — 파이프라인 전체
- `appsscript.json` — 매니페스트 (OAuth 스코프)

## 빠른 시작 (복사·붙여넣기)

1. https://script.google.com → **새 프로젝트**.
2. 기본 `Code.gs` 내용을 지우고 이 폴더의 `Code.gs` 를 붙여넣기.
3. 왼쪽 톱니바퀴(**프로젝트 설정**) → **"appsscript.json" 매니페스트 파일 표시** 체크
   → 편집기에 나타난 `appsscript.json` 을 이 폴더 내용으로 교체.
4. **프로젝트 설정 → 스크립트 속성 → 속성 추가**:
   - `MAIL_TO` = 받을 주소 (예: `kaka_moka@naver.com`)
   - `REPORT_TZ` = `Asia/Seoul` (선택 — 없으면 프로젝트 시간대)
5. 함수 선택 드롭다운에서 **`runPipeline`** → **실행**.
   첫 실행 시 권한 승인 창이 뜬다 ("고급 → 안전하지 않은 페이지로 이동" → 허용).
   → `MAIL_TO` 로 더미 리포트 xlsx가 첨부된 메일이 온다.
6. 스케줄 등록: 함수 드롭다운에서 **`installTrigger`** → **실행** (한 번만).
   - 매일 발송: `Code.gs` 상단 `SCHEDULE_MODE = 'DAILY'`, `DAILY_HOUR = 8`
   - 5분 테스트: `SCHEDULE_MODE = 'TEST_5MIN'` 으로 바꾸고 `installTrigger` 재실행
7. 스케줄 해제: **`removeTrigger`** 실행. (또는 왼쪽 ⏰ **트리거** 메뉴에서 삭제)

실행 로그는 왼쪽 **실행수(Executions)** 메뉴에서 확인.

## 할당량 (일반 Gmail 계정)

| 항목 | 한도 |
|---|---|
| 하루 메일 수신자 | **100** (Workspace 계정은 1,500) |
| 트리거 총 실행 시간 | 90분/일 |
| 실행당 최대 시간 | 6분 |

→ 5분 간격(하루 288회)은 **100통 한도를 초과**한다. 테스트로 몇 시간은 가능하지만
하루 종일은 안 됨. 매일/매주 발송은 여유롭다.
`MailApp.getRemainingDailyQuota()` 로그로 남은 할당량을 확인할 수 있다.

## clasp 로 git 동기화 (선택)

```bash
npm i -g @google/clasp
clasp login
cd apps-script
clasp create --type standalone --title "mobileindex-report"   # 또는 clasp clone <scriptId>
clasp push
```

`.clasp.json` 이 생성되며 scriptId 를 담는다 (git에 커밋해도 무방, 비밀 아님).

## 2단계 (모바일인덱스 실연동)

`datasource_()` 내부만 교체한다.
- **경로 A (API):** `UrlFetchApp.fetch(apiUrl, { headers: { Authorization: 'Bearer ' + key } })`
  로 받아 `COLUMNS` 스키마의 2차원 배열로 변환. `appsscript.json` 에
  `script.external_request` 스코프는 이미 있음.
- **경로 B (CSV 로그인):** Apps Script는 헤드리스 브라우저가 없어 불가.
  이 경우 Python 버전(GitHub Actions / Cloud Run)을 쓸 것.
