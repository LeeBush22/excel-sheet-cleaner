# -*- coding: utf-8 -*-
"""Shared helpers for Excel worksheet cleanup and verification scripts."""

from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path


DEFAULT_KEEP_KEYWORD = "\u8fdc\u7a0b\u6f0f\u6d1e"
SUPPORTED_EXTENSIONS = (".xls", ".xlsx")
EXCEL_FILE_FORMATS = {
    ".xls": 56,  # Excel 97-2003 Workbook (*.xls)
    ".xlsx": 51,  # Excel Workbook (*.xlsx)
}

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_SOURCE_DIR = SCRIPT_DIR / "excel"
DEFAULT_OUTPUT_DIR = SCRIPT_DIR / "output"


def supported_extensions_text() -> str:
    return "、".join(SUPPORTED_EXTENSIONS)


def normalize_keyword(keyword: str) -> str:
    normalized = keyword.strip()
    if not normalized:
        raise ValueError("工作表名称关键字不能为空。")
    return normalized


def is_temporary_excel_file(file: Path) -> bool:
    return file.name.startswith("~$")


def is_supported_excel_file(file: Path) -> bool:
    return (
        file.is_file()
        and file.suffix.lower() in SUPPORTED_EXTENSIONS
        and not is_temporary_excel_file(file)
    )


def collect_excel_files(source: Path) -> list[Path]:
    if source.is_file():
        if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError(f"源文件不是支持的 Excel 文件（{supported_extensions_text()}）：{source}")
        if is_temporary_excel_file(source):
            return []
        return [source]

    if not source.is_dir():
        raise ValueError(f"源路径不存在：{source}")

    return sorted(file for file in source.iterdir() if is_supported_excel_file(file))


def collect_excel_files_by_name(directory: Path) -> dict[str, Path]:
    if not directory.is_dir():
        raise ValueError(f"目录不存在：{directory}")

    return {file.name: file for file in collect_excel_files(directory)}


def file_format_for(output_file: Path) -> int:
    suffix = output_file.suffix.lower()
    try:
        return EXCEL_FILE_FORMATS[suffix]
    except KeyError as exc:
        raise ValueError(f"不支持的输出文件格式：{output_file}") from exc


def ensure_not_same_path(input_file: Path, output_file: Path) -> None:
    input_path = str(input_file.resolve()).casefold()
    output_path = str(output_file.resolve()).casefold()
    if input_path == output_path:
        raise ValueError(f"输出文件不能和原文件相同：{output_file}")


def sheet_names(workbook) -> list[str]:
    return [
        workbook.Worksheets.Item(index).Name
        for index in range(1, workbook.Worksheets.Count + 1)
    ]


def matched_sheet_names(names: list[str], keep_keyword: str) -> list[str]:
    return [name for name in names if keep_keyword in name]


def load_excel_com():
    try:
        import win32com.client  # type: ignore
    except ImportError:
        print("缺少 pywin32 组件，无法通过 Python 调用 Excel。")
        print("请先执行：python -m pip install pywin32")
        raise SystemExit(1)

    return win32com.client


def create_excel_app():
    win32com_client = load_excel_com()
    excel = win32com_client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.AutomationSecurity = 3
    return excel


def open_workbook(excel, workbook_file: Path, *, read_only: bool = True):
    return excel.Workbooks.Open(
        str(workbook_file),
        UpdateLinks=0,
        ReadOnly=read_only,
        IgnoreReadOnlyRecommended=True,
    )


def close_workbook(workbook) -> None:
    if workbook is not None:
        workbook.Close(SaveChanges=False)


def join_sheet_names(names: list[str]) -> str:
    return "; ".join(names)


def timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def timestamp_for_file() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def write_csv_log(output_dir: Path, prefix: str, fieldnames: list[str], rows: list[dict[str, str]]) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    log_file = output_dir / f"{prefix}_{timestamp_for_file()}.csv"

    with log_file.open("w", newline="", encoding="utf-8-sig") as fp:
        writer = csv.DictWriter(fp, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    return log_file
