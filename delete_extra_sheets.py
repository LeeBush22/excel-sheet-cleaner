# -*- coding: utf-8 -*-
"""
Delete extra worksheets from .xls/.xlsx files and keep matching sheets.

Default behavior:
    - Read Excel files from the "excel" folder next to this script.
    - Save processed files to the "output" folder next to this script.
    - Keep every worksheet whose name contains the keyword "远程漏洞".
    - Refuse to overwrite existing output files unless --overwrite is set.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from excel_common import (
    DEFAULT_KEEP_KEYWORD,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_SOURCE_DIR,
    close_workbook,
    collect_excel_files,
    create_excel_app,
    ensure_not_same_path,
    file_format_for,
    join_sheet_names,
    matched_sheet_names,
    normalize_keyword,
    open_workbook,
    sheet_names,
    supported_extensions_text,
    timestamp,
    write_csv_log,
)


LOG_FIELDS = [
    "run_time",
    "file",
    "extension",
    "status",
    "dry_run",
    "overwrite",
    "keyword",
    "source_sheet_count",
    "matched_sheet_count",
    "deleted_sheet_count",
    "source_sheets",
    "kept_sheets",
    "deleted_sheets",
    "output_file",
    "message",
]


def build_log_row(
    *,
    input_file: Path,
    keyword: str,
    dry_run: bool,
    overwrite: bool,
    status: str,
    source_sheets: list[str] | None = None,
    kept_sheets: list[str] | None = None,
    deleted_sheets: list[str] | None = None,
    output_file: Path | None = None,
    message: str = "",
) -> dict[str, str]:
    source_sheets = source_sheets or []
    kept_sheets = kept_sheets or []
    deleted_sheets = deleted_sheets or []

    return {
        "run_time": timestamp(),
        "file": str(input_file),
        "extension": input_file.suffix.lower(),
        "status": status,
        "dry_run": str(dry_run),
        "overwrite": str(overwrite),
        "keyword": keyword,
        "source_sheet_count": str(len(source_sheets)),
        "matched_sheet_count": str(len(kept_sheets)),
        "deleted_sheet_count": str(len(deleted_sheets)),
        "source_sheets": join_sheet_names(source_sheets),
        "kept_sheets": join_sheet_names(kept_sheets),
        "deleted_sheets": join_sheet_names(deleted_sheets),
        "output_file": str(output_file) if output_file else "",
        "message": message,
    }


def process_workbook(
    excel,
    input_file: Path,
    output_dir: Path,
    keep_keyword: str,
    *,
    dry_run: bool,
    overwrite: bool,
) -> dict[str, str]:
    output_file = output_dir / input_file.name
    ensure_not_same_path(input_file, output_file)

    workbook = None
    names: list[str] = []
    kept_sheets: list[str] = []
    deleted_sheets: list[str] = []

    try:
        if output_file.exists() and not overwrite and not dry_run:
            return build_log_row(
                input_file=input_file,
                keyword=keep_keyword,
                dry_run=dry_run,
                overwrite=overwrite,
                status="跳过",
                output_file=output_file,
                message="输出文件已存在；如需覆盖请添加 --overwrite",
            )

        workbook = open_workbook(excel, input_file, read_only=True)
        names = sheet_names(workbook)
        kept_sheets = matched_sheet_names(names, keep_keyword)
        deleted_sheets = [name for name in names if name not in kept_sheets]

        if not kept_sheets:
            return build_log_row(
                input_file=input_file,
                keyword=keep_keyword,
                dry_run=dry_run,
                overwrite=overwrite,
                status="跳过",
                source_sheets=names,
                output_file=output_file,
                message=f"未找到名称包含“{keep_keyword}”的工作表",
            )

        if dry_run:
            return build_log_row(
                input_file=input_file,
                keyword=keep_keyword,
                dry_run=dry_run,
                overwrite=overwrite,
                status="预演",
                source_sheets=names,
                kept_sheets=kept_sheets,
                deleted_sheets=deleted_sheets,
                output_file=output_file,
                message="dry-run 模式未修改或保存文件",
            )

        for index in range(workbook.Worksheets.Count, 0, -1):
            sheet = workbook.Worksheets.Item(index)
            if sheet.Name not in kept_sheets:
                sheet.Delete()

        workbook.SaveAs(str(output_file), FileFormat=file_format_for(output_file))

        return build_log_row(
            input_file=input_file,
            keyword=keep_keyword,
            dry_run=dry_run,
            overwrite=overwrite,
            status="成功",
            source_sheets=names,
            kept_sheets=kept_sheets,
            deleted_sheets=deleted_sheets,
            output_file=output_file,
        )
    except Exception as exc:
        return build_log_row(
            input_file=input_file,
            keyword=keep_keyword,
            dry_run=dry_run,
            overwrite=overwrite,
            status="失败",
            source_sheets=names,
            kept_sheets=kept_sheets,
            deleted_sheets=deleted_sheets,
            output_file=output_file,
            message=str(exc),
        )
    finally:
        close_workbook(workbook)


def write_log(output_dir: Path, rows: list[dict[str, str]]) -> Path:
    return write_csv_log(output_dir, "处理日志", LOG_FIELDS, rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=f"删除 Excel 文件中的多余子表，支持 {supported_extensions_text()}，只保留名称包含指定关键字的工作表。"
    )
    parser.add_argument("positional_source", nargs="?", help="源 Excel 文件或源目录")
    parser.add_argument("positional_output", nargs="?", help="输出目录")
    parser.add_argument("-s", "--source", help="源 Excel 文件或源目录")
    parser.add_argument("-o", "--output", help="输出目录")
    parser.add_argument(
        "-k",
        "--keyword",
        default=DEFAULT_KEEP_KEYWORD,
        help="要保留的工作表名称关键字，默认为“远程漏洞”。",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅预览将保留和删除的工作表，不保存输出文件。",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="允许覆盖输出目录中已存在的同名文件。",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    source_text = args.source or args.positional_source
    output_text = args.output or args.positional_output

    source = Path(source_text).expanduser() if source_text else DEFAULT_SOURCE_DIR
    output_dir = Path(output_text).expanduser() if output_text else DEFAULT_OUTPUT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        keyword = normalize_keyword(args.keyword)
        excel_files = collect_excel_files(source)
    except ValueError as exc:
        print(f"错误：{exc}")
        return 1

    if not excel_files:
        print(f"未找到可处理的 Excel 文件（{supported_extensions_text()}）。")
        return 1

    excel = create_excel_app()
    rows: list[dict[str, str]] = []
    try:
        for index, input_file in enumerate(excel_files, start=1):
            print(f"[{index}/{len(excel_files)}] 正在处理：{input_file}")
            row = process_workbook(
                excel,
                input_file,
                output_dir,
                keyword,
                dry_run=args.dry_run,
                overwrite=args.overwrite,
            )
            rows.append(row)
            print(f"    {row['status']}：{row['message'] or row['output_file']}")
    finally:
        excel.Quit()

    log_file = write_log(output_dir, rows)
    success_count = sum(1 for row in rows if row["status"] == "成功")
    preview_count = sum(1 for row in rows if row["status"] == "预演")
    skipped_count = sum(1 for row in rows if row["status"] == "跳过")
    failed_count = sum(1 for row in rows if row["status"] == "失败")

    print("")
    print("处理完成。")
    print(f"成功：{success_count}，预演：{preview_count}，跳过：{skipped_count}，失败：{failed_count}")
    print(f"输出目录：{output_dir}")
    print(f"处理日志：{log_file}")
    return 0 if failed_count == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
