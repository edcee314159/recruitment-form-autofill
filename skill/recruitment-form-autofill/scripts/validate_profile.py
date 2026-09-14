"""Validate workbook structure without including candidate values in diagnostics."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re

import read_profile as reader


@dataclass(frozen=True)
class ValidationIssue:
    sheet: str
    row: int
    code: str
    message: str


MESSAGES = {
    "headers": "Required sheet headers do not match the supported schema.",
    "sheet_missing": "A required sheet is missing.",
    "workbook": "The workbook could not be read or its schema is unsupported.",
    "unsupported_category": "Choose a supported experience category.",
    "structural_blank": "A populated record is missing a structural field.",
    "invalid_date": "Use an Excel date, YYYY-MM, or YYYY-MM-DD; an end date may be 至今.",
    "date_order": "The end date precedes the start date.",
    "duplicate_identity": "A record repeats an existing identity.",
    "language_pair": "Keep one language credential and its own score in each row.",
    "attachment_missing": "The attachment does not resolve to an existing local file.",
    "extra_personal_row": "Personal data must be entered only in row 2.",
}


def _issue(sheet, row, code):
    return ValidationIssue(sheet, row, code, MESSAGES[code])


def _date(value, *, end=False):
    if not reader._nonblank(value) or (end and value == "至今"):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}(?:-\d{2})?", value):
        return date.fromisoformat(value if len(value) == 10 else value + "-01")
    raise ValueError


def _dates(sheet, row, start, end):
    issues = []
    parsed = []
    for value, is_end in ((start, False), (end, True)):
        try:
            parsed.append(_date(value, end=is_end))
        except (ValueError, TypeError):
            parsed.append(None)
            issues.append(_issue(sheet, row, "invalid_date"))
    if all(parsed) and parsed[0] > parsed[1]:
        issues.append(_issue(sheet, row, "date_order"))
    return issues


def _identity(values):
    return tuple(str(value).strip() if value is not None else "" for value in values)


def _language_valid(name, score):
    if not reader._nonblank(name) or not reader._nonblank(score):
        return False
    credentials = re.findall(r"CET[- ]?[46]|IELTS|TOEFL|四级|六级|雅思|托福", str(name), re.I)
    if len(credentials) > 1:
        return False
    if credentials:
        try:
            number = float(score)
        except (TypeError, ValueError):
            return False
        return number >= 0 and number < float("inf")
    return True


def validate_profile(path: Path) -> list[ValidationIssue]:
    """Return fixed diagnostics with only known sheet names and row coordinates."""
    path = Path(path)
    try:
        workbook = reader._open_workbook(path)
    except reader.ProfileError:
        return [_issue("", 0, "workbook")]
    issues = []
    try:
        try:
            schema = reader._detect_workbook_schema(workbook)
        except reader.ProfileError:
            return [_issue("", 0, "workbook")]
        if schema == "public":
            specifications = ((reader.PERSONAL_SHEET, 1, reader.PERSONAL_HEADERS), (reader.PUBLIC_EXPERIENCE_SHEET, 1, reader.PUBLIC_EXPERIENCE_HEADERS), (reader.ATTACHMENT_SHEET, 1, reader.ATTACHMENT_HEADERS))
        else:
            specifications = ((reader.GENERAL_SHEET, 8, reader.GENERAL_HEADERS), (reader.EXPERIENCE_SHEET, 6, reader.EXPERIENCE_HEADERS))
        for sheet_name, header_row, headers in specifications:
            if sheet_name not in workbook.sheetnames:
                issues.append(_issue(sheet_name, header_row, "sheet_missing"))
                continue
            sheet = workbook[sheet_name]
            actual = tuple(sheet.cell(header_row, col).value for col in range(1, len(headers) + 1))
            trailing = [sheet.cell(header_row, col).value for col in range(len(headers) + 1, sheet.max_column + 1)]
            if actual != headers or any(reader._nonblank(v) for v in trailing):
                issues.append(_issue(sheet_name, header_row, "headers"))
                continue
            seen = set()
            for row, values in enumerate(sheet.iter_rows(min_row=header_row + 1, max_col=len(headers), values_only=True), header_row + 1):
                if not any(reader._nonblank(v) for v in values):
                    continue
                if sheet_name == reader.PERSONAL_SHEET:
                    if row != 2:
                        issues.append(_issue(sheet_name, row, "extra_personal_row"))
                    issues.extend(_dates(sheet_name, row, values[3], None))
                    issues.extend(_dates(sheet_name, row, values[10], values[11]))
                elif sheet_name == reader.GENERAL_SHEET:
                    if not reader._nonblank(values[2]):
                        issues.append(_issue(sheet_name, row, "structural_blank"))
                    identity = _identity(values[2:3])
                    if identity in seen:
                        issues.append(_issue(sheet_name, row, "duplicate_identity"))
                    seen.add(identity)
                elif sheet_name == reader.ATTACHMENT_SHEET:
                    if not all(reader._nonblank(v) for v in values[:2]):
                        issues.append(_issue(sheet_name, row, "structural_blank"))
                    if reader._nonblank(values[1]):
                        try:
                            attachment = Path(str(values[1])).expanduser()
                            if not attachment.is_absolute():
                                attachment = path.resolve().parent / attachment
                            exists = attachment.is_file()
                        except (OSError, ValueError):
                            exists = False
                        if not exists:
                            issues.append(_issue(sheet_name, row, "attachment_missing"))
                else:
                    if not all(reader._nonblank(v) for v in values[:2]):
                        issues.append(_issue(sheet_name, row, "structural_blank"))
                    if values[0] not in reader.SUPPORTED_CATEGORIES:
                        issues.append(_issue(sheet_name, row, "unsupported_category"))
                    issues.extend(_dates(sheet_name, row, values[4], values[5]))
                    identity = _identity((values[0], values[1], values[2], values[4], values[5]))
                    if identity in seen:
                        issues.append(_issue(sheet_name, row, "duplicate_identity"))
                    seen.add(identity)
                    if values[0] == "语言能力" and not _language_valid(values[1], values[6]):
                        issues.append(_issue(sheet_name, row, "language_pair"))
        return issues
    except reader.MALFORMED_WORKBOOK_ERRORS:
        return [_issue("", 0, "workbook")]
    finally:
        workbook.close()
