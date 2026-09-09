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
