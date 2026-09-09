from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import date, datetime
from zoneinfo import ZoneInfo

from .config import load_settings
from .pipeline import run


def _parse_date(value: str) -> date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as error:
        raise argparse.ArgumentTypeError(
            f"날짜 형식은 YYYY-MM-DD 여야 합니다: {value!r}"
        ) from error


def _load_dotenv() -> None:
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    load_dotenv(".env.local")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="report_pipeline")
    parser.add_argument(
        "--date",
        dest="run_date",
        type=_parse_date,
        default=None,
        help="리포트 기준일 (YYYY-MM-DD). 생략 시 REPORT_TZ 기준 오늘.",
    )
    args = parser.parse_args(argv)

    _load_dotenv()
    settings = load_settings()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )

    run_date = args.run_date or datetime.now(ZoneInfo(settings.report_tz)).date()

    summary = run(settings, run_date)
    print(json.dumps(summary, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
