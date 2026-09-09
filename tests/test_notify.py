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
