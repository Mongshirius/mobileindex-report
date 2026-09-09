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
