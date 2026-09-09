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
