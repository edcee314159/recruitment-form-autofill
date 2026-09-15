"""Reject private files before a public repository release without printing contents."""

from __future__ import annotations

from pathlib import Path
import sys

from openpyxl import load_workbook


ALLOWED_WORKBOOKS = {
    Path("templates") / "秋招个人资料库模板.xlsx": "资料类型：空白模板",
    Path("examples") / "示例资料库.xlsx": "资料类型：虚构示例",
}
PROHIBITED_EXTENSIONS = {".docx", ".pdf", ".jpg", ".jpeg", ".png"}


def _iter_files(root: Path):
    for path in root.rglob("*"):
        if path.is_file() and not {".git", "node_modules", "__pycache__"}.intersection(path.parts):
            yield path


def _workbook_label_is_expected(path: Path, expected_label: str) -> bool:
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            return "使用说明" in workbook.sheetnames and workbook["使用说明"]["A2"].value == expected_label
        finally:
            workbook.close()
    except (OSError, ValueError, KeyError):
        return False


def check(root: Path) -> list[str]:
    root = Path(root).resolve()
    issues = []
    for path in _iter_files(root):
        relative = path.relative_to(root)
        if path.name == "profile-config.json":
            issues.append("local profile configuration")
        elif path.name.endswith(".personal.xlsx"):
            issues.append("personal workbook")
        elif path.suffix.lower() in PROHIBITED_EXTENSIONS:
            issues.append("personal attachment")
        elif path.suffix.lower() == ".xlsx":
            expected_label = ALLOWED_WORKBOOKS.get(relative)
            if expected_label is None:
                issues.append("workbook outside template/example locations")
            elif not _workbook_label_is_expected(path, expected_label):
                issues.append("template/example workbook label")
    return sorted(set(issues))


def main(argv: list[str] | None = None) -> int:
    arguments = argv if argv is not None else sys.argv[1:]
    if len(arguments) != 1:
        print("usage: prepublish_check.py REPOSITORY", file=sys.stderr)
        return 2
    issues = check(Path(arguments[0]))
    if issues:
        for issue in issues:
            print(f"error: prohibited {issue}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
