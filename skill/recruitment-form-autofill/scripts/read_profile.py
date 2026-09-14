"""Read recruitment-profile records from the approved Excel database."""

from __future__ import annotations

import argparse
import json
import os
import sys
import zipfile
from datetime import date, datetime
from pathlib import Path
from typing import Iterable
from xml.etree.ElementTree import ParseError

import openpyxl
from openpyxl.utils.exceptions import InvalidFileException


from profile_config import ProfileConfigError, load_config

PERSONAL_SHEET = "个人资料"
PUBLIC_EXPERIENCE_SHEET = "经历清单"
ATTACHMENT_SHEET = "附件清单"
PERSONAL_HEADERS = (
    "姓名", "手机号码", "邮箱", "出生日期", "籍贯", "现居城市", "学校名称", "专业名称",
    "学历", "学位", "入学时间", "预计毕业时间", "期望岗位方向", "期望工作城市",
)
ATTACHMENT_HEADERS = ("文件类型", "本地文件路径", "说明")
SUPPORTED_CATEGORIES = (
    "教育经历", "实习经历", "工作经历", "项目经历", "语言能力", "专业证书",
    "竞赛获奖", "荣誉称号", "校园活动", "科研成果",
)
GENERAL_SHEET = "资料总表"
EXPERIENCE_SHEET = "经历与成果"
GENERAL_HEADERS = (
    "序号", "分类", "字段", "当前值", "状态", "填写提示", "来源", "备注", "使用场景"
)
GENERAL_KEYS = (
    "index", "category", "field", "value", "status", "hint", "source", "note", "usage"
)
EXPERIENCE_HEADERS = (
    "类别", "项目/单位", "角色/级别", "地点", "开始/考试时间", "结束/颁发时间",
    "成绩/编号", "详细内容", "来源", "备注"
)
EXPERIENCE_KEYS = (
    "category", "organization", "role", "location", "start", "end", "result", "details", "source", "note"
)
PUBLIC_EXPERIENCE_HEADERS = EXPERIENCE_HEADERS[:8] + ("备注",)
PUBLIC_EXPERIENCE_KEYS = EXPERIENCE_KEYS[:8] + ("note",)


class ProfileError(Exception):
    """A database problem that can be reported without exposing profile data."""


class SchemaError(ProfileError):
    """The workbook does not match the fixed, supported schema."""


MALFORMED_WORKBOOK_ERRORS = (OSError, zipfile.BadZipFile, InvalidFileException, ParseError)


def _rows_from_sheet(
    path: Path,
    sheet_name: str,
    header_row: int,
    headers: tuple[str, ...],
    keys: tuple[str, ...],
) -> list[dict[str, object]]:
    workbook = _open_workbook(path)
    try:
        return _rows_from_workbook(workbook, sheet_name, header_row, headers, keys)
    finally:
        workbook.close()


def _open_workbook(path: Path):
    path = Path(path)
    if not path.is_file():
        raise ProfileError("database not found")
    try:
        return openpyxl.load_workbook(path, read_only=True, data_only=True)
    except MALFORMED_WORKBOOK_ERRORS as exc:
        raise ProfileError("database could not be read") from exc


def _rows_from_workbook(
    workbook,
    sheet_name: str,
    header_row: int,
    headers: tuple[str, ...],
    keys: tuple[str, ...],
) -> list[dict[str, object]]:
    try:
        if sheet_name not in workbook.sheetnames:
            raise ProfileError(f"required sheet missing: {sheet_name}")
        sheet = workbook[sheet_name]
        actual_headers = tuple(
            sheet.cell(row=header_row, column=column).value
            for column in range(1, len(headers) + 1)
        )
        trailing_headers = tuple(
            sheet.cell(row=header_row, column=column).value
            for column in range(len(headers) + 1, sheet.max_column + 1)
        )
        if actual_headers != headers or any(value is not None for value in trailing_headers):
            raise SchemaError(f"unexpected headers in sheet: {sheet_name}")
        records: list[dict[str, object]] = []
        for values in sheet.iter_rows(min_row=header_row + 1, max_col=len(headers), values_only=True):
            if not any(value is not None for value in values):
                continue
            records.append(dict(zip(keys, values)))
        return records
    except MALFORMED_WORKBOOK_ERRORS as exc:
        raise ProfileError("database could not be read") from exc


def _detect_workbook_schema(workbook) -> str:
    if PERSONAL_SHEET in workbook.sheetnames:
        return "public"
    if GENERAL_SHEET in workbook.sheetnames:
        return "legacy"
    raise SchemaError("unsupported workbook schema")


def detect_schema(path: Path) -> str:
    """Identify a supported workbook layout without exposing any values."""
    workbook = _open_workbook(path)
    try:
        return _detect_workbook_schema(workbook)
    finally:
        workbook.close()


def _general_from_workbook(workbook):
    if _detect_workbook_schema(workbook) == "legacy":
        return _rows_from_workbook(workbook, GENERAL_SHEET, 8, GENERAL_HEADERS, GENERAL_KEYS)
    _rows_from_workbook(workbook, PERSONAL_SHEET, 1, PERSONAL_HEADERS, PERSONAL_HEADERS)
    sheet = workbook[PERSONAL_SHEET]
    return [dict(zip(GENERAL_KEYS, (index, PERSONAL_SHEET, field, sheet.cell(2, index).value, None, None, None, None, None)))
            for index, field in enumerate(PERSONAL_HEADERS, 1)]


def _experiences_from_workbook(workbook):
    if _detect_workbook_schema(workbook) == "legacy":
        return _rows_from_workbook(workbook, EXPERIENCE_SHEET, 6, EXPERIENCE_HEADERS, EXPERIENCE_KEYS)
    records = _rows_from_workbook(workbook, PUBLIC_EXPERIENCE_SHEET, 1, PUBLIC_EXPERIENCE_HEADERS, PUBLIC_EXPERIENCE_KEYS)
    return [dict(record, source=None) for record in records]


def load_general_profile(path: Path) -> list[dict[str, object]]:
    """Load public row 2 or legacy field rows into the existing record shape."""
    workbook = _open_workbook(path)
    try:
        return _general_from_workbook(workbook)
    finally:
        workbook.close()


def load_experience_records(path: Path) -> list[dict[str, object]]:
    """Load experience rows using the fixed 经历与成果 schema."""
    workbook = _open_workbook(path)
    try:
        return _experiences_from_workbook(workbook)
    finally:
        workbook.close()


def load_inventory(path: Path) -> dict[str, list[object]]:
    """List profile labels using one workbook open and without returning values."""
    workbook = _open_workbook(path)
    try:
        general = _general_from_workbook(workbook)
        experiences = _experiences_from_workbook(workbook)
    finally:
        workbook.close()
    return _inventory(general, experiences)


def load_bundle(
    path: Path,
    fields: Iterable[str],
    categories: Iterable[str],
) -> dict[str, list[dict[str, object]]]:
    """Read requested records from both sheets using one workbook open."""
    workbook = _open_workbook(path)
    try:
        general = _general_from_workbook(workbook)
        experiences = _experiences_from_workbook(workbook)
    finally:
        workbook.close()

    selected_general: list[dict[str, object]] = []
    for field in fields:
        matches = [record for record in general if record["field"] == field]
        if not matches:
            raise ProfileError("requested field not found")
        if any(not _nonblank(record["value"]) for record in matches):
            raise ProfileError("requested field has a blank value")
        selected_general.extend(matches)

    selected_experiences: list[dict[str, object]] = []
    for category in categories:
        matches = [record for record in experiences if record["category"] == category]
        if not matches:
            raise ProfileError("requested experience category not found")
        selected_experiences.extend(matches)
    return {"general": selected_general, "experiences": selected_experiences}


def _json_default(value: object) -> str:
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return str(value)


def _write_json(payload: object) -> None:
    print(json.dumps(payload, ensure_ascii=False, default=_json_default))


def _nonblank(value: object) -> bool:
    return value is not None and (not isinstance(value, str) or bool(value.strip()))


def _inventory(general: Iterable[dict[str, object]], experiences: Iterable[dict[str, object]]) -> dict[str, list[object]]:
    return {
        "categories": sorted({record["category"] for record in general if _nonblank(record["category"])}),
        "fields": [record["field"] for record in general if _nonblank(record["field"])],
        "experience_categories": sorted(
            {record["category"] for record in experiences if _nonblank(record["category"])}
        ),
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Read the approved recruitment profile workbook.")
    paths = parser.add_mutually_exclusive_group()
    paths.add_argument("--config", type=Path, metavar="PATH")
    paths.add_argument("--database", type=Path, metavar="PATH", help="Explicit workbook override for legacy use and testing.")
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("inventory", help="List profile field and category names only.")
    get_parser = commands.add_parser("get", help="Read explicitly requested general fields.")
    get_parser.add_argument("--field", action="append", required=True, metavar="FIELD")
    experiences_parser = commands.add_parser("experiences", help="Read one experience category.")
    experiences_parser.add_argument("--category", required=True, metavar="CATEGORY")
    bundle_parser = commands.add_parser("bundle", help="Read requested fields and experience categories together.")
    bundle_parser.add_argument("--field", action="append", default=[], metavar="FIELD")
    bundle_parser.add_argument("--category", action="append", default=[], metavar="CATEGORY")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.database is None:
            config_path = args.config or os.environ.get("RECRUITMENT_PROFILE_CONFIG")
            if not config_path:
                raise ProfileError("local configuration required; use --config")
            args.database = load_config(config_path).workbook_path
        if args.command == "inventory":
            _write_json(load_inventory(args.database))
            return 0

        if args.command == "get":
            general = load_general_profile(args.database)
            selected: list[dict[str, object]] = []
            for field in args.field:
                matches = [record for record in general if record["field"] == field]
                if not matches:
                    raise ProfileError("requested field not found")
                if any(not _nonblank(record["value"]) for record in matches):
                    return _error("requested field has a blank value", 3)
                selected.extend(matches)
            _write_json({"records": selected})
            return 0

        if args.command == "bundle":
            try:
                _write_json(load_bundle(args.database, args.field, args.category))
                return 0
            except ProfileError as exc:
                if str(exc) == "requested field has a blank value":
                    return _error(str(exc), 3)
                raise

        experiences = load_experience_records(args.database)
        selected = [record for record in experiences if record["category"] == args.category]
        if not selected:
            raise ProfileError("requested experience category not found")
        _write_json({"records": selected})
        return 0
    except (ProfileError, ProfileConfigError) as exc:
        return _error(str(exc), 2)


def _error(message: str, exit_code: int) -> int:
    print(f"error: {message}", file=sys.stderr)
    return exit_code


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    raise SystemExit(main())
