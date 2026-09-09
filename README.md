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
