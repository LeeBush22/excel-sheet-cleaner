# -*- coding: utf-8 -*-
"""
Verify that cleaned .xls/.xlsx files kept the expected worksheets unchanged.

The verification rule matches delete_extra_sheets.py:
    - Every worksheet whose name contains the keyword must be kept.
    - Every non-matching worksheet must be removed from the output file.
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
    collect_excel_files_by_name,
    create_excel_app,
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
    "keyword",
    "source_sheet_count",
    "output_sheet_count",
    "matched_sheet_count",
    "source_sheets",
    "expected_kept_sheets",
    "output_sheets",
    "message",
]


def normalize_matrix(value):
    if value is None:
        return tuple()
    if not isinstance(value, tuple):
        return ((value,),)
    if value and not isinstance(value[0], tuple):
        return (value,)
    return value


def read_range_matrix(used_range, attribute: str):
    try:
        return normalize_matrix(getattr(used_range, attribute))
    except Exception:
        row_count = used_range.Rows.Count
        column_count = used_range.Columns.Count
        rows = []
        for row_index in range(1, row_count + 1):
            row_values = []
            for column_index in range(1, column_count + 1):
                cell = used_range.Cells(row_index, column_index)
                row_values.append(getattr(cell, attribute))
            rows.append(tuple(row_values))
        return tuple(rows)


def snapshot_sheet(sheet) -> dict:
    used_range = sheet.UsedRange
    data = {
        "shape": (
            used_range.Row,
            used_range.Column,
            used_range.Rows.Count,
            used_range.Columns.Count,
        ),
        "values": read_range_matrix(used_range, "Value2"),
        "formulas": None,
        "formula_error": "",
    }

    try:
        data["formulas"] = read_range_matrix(used_range, "Formula")
    except Exception as exc:
        data["formula_error"] = str(exc)

    return data


def snapshot_workbook(excel, workbook_file: Path, keep_keyword: str) -> dict:
    workbook = None
    stage = "打开文件"
    try:
        workbook = open_workbook(excel, workbook_file, read_only=True)

        stage = "读取工作表名称"
        names = sheet_names(workbook)
        matched_names = matched_sheet_names(names, keep_keyword)
        snapshots = {}

        for sheet_name in matched_names:
            stage = f"读取工作表 {sheet_name}"
            sheet = workbook.Worksheets.Item(sheet_name)
            snapshots[sheet_name] = snapshot_sheet(sheet)

        return {
            "sheet_count": len(names),
            "sheet_names": names,
            "matched_names": matched_names,
            "snapshots": snapshots,
        }
    except Exception as exc:
        raise RuntimeError(f"{workbook_file.name} 阶段={stage}：{exc}") from exc
    finally:
        close_workbook(workbook)


def compare_sheet_snapshot(sheet_name: str, source_snapshot: dict, output_snapshot: dict) -> tuple[list[str], list[str]]:
    diffs: list[str] = []
    notes: list[str] = []

    if source_snapshot["shape"] != output_snapshot["shape"]:
        diffs.append(
            f"{sheet_name} UsedRange 不一致："
            f"原文件={source_snapshot['shape']}，输出文件={output_snapshot['shape']}"
        )
    if source_snapshot["values"] != output_snapshot["values"]:
        diffs.append(f"{sheet_name} 单元格值不一致")

    if source_snapshot["formula_error"] or output_snapshot["formula_error"]:
        notes.append(
            f"{sheet_name} 公式检查跳过："
            f"原文件错误={source_snapshot['formula_error'] or '无'}，"
            f"输出文件错误={output_snapshot['formula_error'] or '无'}"
        )
    elif source_snapshot["formulas"] != output_snapshot["formulas"]:
        diffs.append(f"{sheet_name} 单元格公式不一致")

    return diffs, notes


def build_log_row(
    *,
    file_name: str,
    extension: str,
    keyword: str,
    status: str,
    source_sheets: list[str] | None = None,
    expected_kept_sheets: list[str] | None = None,
    output_sheets: list[str] | None = None,
    message: str = "",
) -> dict[str, str]:
    source_sheets = source_sheets or []
    expected_kept_sheets = expected_kept_sheets or []
    output_sheets = output_sheets or []

    return {
        "run_time": timestamp(),
        "file": file_name,
        "extension": extension,
        "status": status,
        "keyword": keyword,
        "source_sheet_count": str(len(source_sheets)),
        "output_sheet_count": str(len(output_sheets)),
        "matched_sheet_count": str(len(expected_kept_sheets)),
        "source_sheets": join_sheet_names(source_sheets),
        "expected_kept_sheets": join_sheet_names(expected_kept_sheets),
        "output_sheets": join_sheet_names(output_sheets),
        "message": message,
    }


def compare_workbook_pair(excel, source_file: Path, output_file: Path, keep_keyword: str) -> dict[str, str]:
    try:
        source_data = snapshot_workbook(excel, source_file, keep_keyword)
        output_data = snapshot_workbook(excel, output_file, keep_keyword)

        expected_kept = source_data["matched_names"]
        output_names = output_data["sheet_names"]
        diffs: list[str] = []
        notes: list[str] = []

        if not expected_kept:
            diffs.append(f"原文件未匹配到名称包含“{keep_keyword}”的工作表")

        if output_names != expected_kept:
            diffs.append(
                "输出工作表列表不符合清理规则："
                f"期望={join_sheet_names(expected_kept)}；实际={join_sheet_names(output_names)}"
            )

        for sheet_name in expected_kept:
            if sheet_name not in output_data["snapshots"]:
                diffs.append(f"输出文件缺少工作表：{sheet_name}")
                continue

            sheet_diffs, sheet_notes = compare_sheet_snapshot(
                sheet_name,
                source_data["snapshots"][sheet_name],
                output_data["snapshots"][sheet_name],
            )
            diffs.extend(sheet_diffs)
            notes.extend(sheet_notes)

        status = "一致"
        message = ""
        if diffs:
            status = "异常"
            message = "；".join(diffs)
        elif notes:
            message = "；".join(notes)

        return build_log_row(
            file_name=source_file.name,
            extension=source_file.suffix.lower(),
            keyword=keep_keyword,
            status=status,
            source_sheets=source_data["sheet_names"],
            expected_kept_sheets=expected_kept,
            output_sheets=output_names,
            message=message,
        )
    except Exception as exc:
        return build_log_row(
            file_name=source_file.name,
            extension=source_file.suffix.lower(),
            keyword=keep_keyword,
            status="失败",
            message=str(exc),
        )


def write_log(output_dir: Path, rows: list[dict[str, str]]) -> Path:
    return write_csv_log(output_dir, "比对日志", LOG_FIELDS, rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=f"比对清理前后的 Excel 文件，支持 {supported_extensions_text()}，确认保留工作表内容是否一致。"
    )
    parser.add_argument("-s", "--source", default=str(DEFAULT_SOURCE_DIR), help="原始 Excel 文件目录")
    parser.add_argument("-o", "--output", default=str(DEFAULT_OUTPUT_DIR), help="处理后 Excel 文件目录")
    parser.add_argument(
        "-k",
        "--keyword",
        default=DEFAULT_KEEP_KEYWORD,
        help="要比对的工作表名称关键字，默认为“远程漏洞”。",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source_dir = Path(args.source).expanduser()
    output_dir = Path(args.output).expanduser()

    try:
        keyword = normalize_keyword(args.keyword)
        source_files = collect_excel_files_by_name(source_dir)
        output_files = collect_excel_files_by_name(output_dir)
    except ValueError as exc:
        print(f"错误：{exc}")
        return 1

    missing_outputs = sorted(set(source_files) - set(output_files))
    extra_outputs = sorted(set(output_files) - set(source_files))
    common_files = sorted(set(source_files) & set(output_files))

    rows: list[dict[str, str]] = []
    if common_files:
        excel = create_excel_app()
        try:
            for index, file_name in enumerate(common_files, start=1):
                print(f"[{index}/{len(common_files)}] 正在比对：{file_name}")
                rows.append(
                    compare_workbook_pair(
                        excel,
                        source_files[file_name],
                        output_files[file_name],
                        keyword,
                    )
                )
        finally:
            excel.Quit()

    for file_name in missing_outputs:
        source_file = source_files[file_name]
        rows.append(
            build_log_row(
                file_name=file_name,
                extension=source_file.suffix.lower(),
                keyword=keyword,
                status="异常",
                message="缺少输出文件",
            )
        )

    for file_name in extra_outputs:
        output_file = output_files[file_name]
        rows.append(
            build_log_row(
                file_name=file_name,
                extension=output_file.suffix.lower(),
                keyword=keyword,
                status="异常",
                message="输出目录存在多余文件",
            )
        )

    log_file = write_log(output_dir, rows)
    same_count = sum(1 for row in rows if row["status"] == "一致")
    different_count = sum(1 for row in rows if row["status"] == "异常")
    failed_count = sum(1 for row in rows if row["status"] == "失败")

    print("")
    print("比对完成。")
    print(f"原始文件数：{len(source_files)}")
    print(f"输出文件数：{len(output_files)}")
    print(f"同名文件数：{len(common_files)}")
    print(f"内容一致：{same_count}")
    print(f"内容差异：{different_count}")
    print(f"读取失败：{failed_count}")
    print(f"比对日志：{log_file}")

    problems = [row for row in rows if row["status"] != "一致"]
    if problems:
        print("")
        print("异常详情：")
        for row in problems:
            print(f"  [{row['status']}] {row['file']} - {row['message']}")

    return 0 if not problems else 2


if __name__ == "__main__":
    sys.exit(main())
