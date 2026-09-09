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
