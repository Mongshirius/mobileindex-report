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
