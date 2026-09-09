from __future__ import annotations

import os
from dataclasses import dataclass

_REQUIRED = [
    "SMTP_HOST",
    "SMTP_PORT",
    "SMTP_USER",
    "SMTP_PASSWORD",
    "MAIL_FROM",
    "MAIL_TO",
    "REPORT_TZ",
]


class ConfigError(RuntimeError):
    """필수 환경변수가 없거나 형식이 잘못된 경우."""


@dataclass(frozen=True)
class Settings:
    smtp_host: str
    smtp_port: int
    smtp_user: str
    smtp_password: str
    mail_from: str
    mail_to: str
    report_tz: str


def _unquote(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
        return value[1:-1]
    return value


def load_settings(environ: dict[str, str] | None = None) -> Settings:
    env = os.environ if environ is None else environ
    present: dict[str, str] = {}
    for key in _REQUIRED:
        raw = env.get(key)
        if raw is None:
            continue
        value = _unquote(raw)
        if value.strip():
            present[key] = value
    missing = [key for key in _REQUIRED if key not in present]
    if missing:
        raise ConfigError(f"필수 환경변수 누락: {', '.join(missing)}")

    try:
        port = int(present["SMTP_PORT"])
    except ValueError as error:
        raise ConfigError(
            f"SMTP_PORT는 정수여야 합니다: {present['SMTP_PORT']!r}"
        ) from error

    return Settings(
        smtp_host=present["SMTP_HOST"],
        smtp_port=port,
        smtp_user=present["SMTP_USER"],
        smtp_password=present["SMTP_PASSWORD"],
        mail_from=present["MAIL_FROM"],
        mail_to=present["MAIL_TO"],
        report_tz=present["REPORT_TZ"],
    )
