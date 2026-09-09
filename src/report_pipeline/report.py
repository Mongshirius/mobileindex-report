from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet

from .transform import ReportData

_MAX_COLUMN_WIDTH = 50


def build_workbook(data: ReportData, out_dir: Path, run_date: date) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"mobileindex_report_{run_date:%Y-%m-%d}.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, frame in data.sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
            _format_sheet(writer.sheets[sheet_name], frame)
    return path


def _format_sheet(worksheet: Worksheet, frame: pd.DataFrame) -> None:
    worksheet.freeze_panes = "A2"
    for column_index, column_name in enumerate(frame.columns, start=1):
        header_cell = worksheet.cell(row=1, column=column_index)
        header_cell.font = Font(bold=True)
        values = frame.iloc[:, column_index - 1].tolist()
        width = max([len(str(column_name))] + [len(str(value)) for value in values])
        worksheet.column_dimensions[header_cell.column_letter].width = min(
            width + 2, _MAX_COLUMN_WIDTH
        )
