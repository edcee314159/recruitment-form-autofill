"""Create and load local recruitment-profile configuration outside this repository."""

from __future__ import annotations

import json
import shutil
from dataclasses import dataclass
from pathlib import Path


REPOSITORY_ROOT = Path(__file__).resolve().parents[3]
BLANK_TEMPLATE = REPOSITORY_ROOT / "templates" / "秋招个人资料库模板.xlsx"


class ProfileConfigError(Exception):
    """Raised when a local profile configuration is invalid or unsafe."""


@dataclass(frozen=True)
class ProfileConfig:
    """The filesystem locations used for one local recruitment profile."""

    workbook_path: Path
    attachment_dir: Path | None = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "workbook_path", Path(self.workbook_path).resolve())
        if self.attachment_dir is not None:
            object.__setattr__(
                self, "attachment_dir", Path(self.attachment_dir).resolve()
            )


def _resolved(path: Path | str) -> Path:
    return Path(path).expanduser().resolve()


def _validate_config_location(
    config_path: Path | str,
    *,
    repository_root: Path | str = REPOSITORY_ROOT,
    allow_repository_config: bool = False,
) -> Path:
    resolved_config = _resolved(config_path)
    resolved_root = _resolved(repository_root)
    if not allow_repository_config:
        try:
            resolved_config.relative_to(resolved_root)
        except ValueError:
            pass
        else:
            raise ProfileConfigError("configuration must not be stored in the repository")
    return resolved_config


def write_config(
    profile: ProfileConfig,
    config_path: Path | str,
    *,
    repository_root: Path | str = REPOSITORY_ROOT,
    allow_repository_config: bool = False,
) -> Path:
    """Persist a profile configuration as UTF-8 JSON at a safe local location."""
    destination = _validate_config_location(
        config_path,
        repository_root=repository_root,
        allow_repository_config=allow_repository_config,
    )
    destination.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "workbook_path": str(profile.workbook_path),
        "attachment_dir": (
            str(profile.attachment_dir) if profile.attachment_dir is not None else None
        ),
    }
    destination.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return destination


def load_config(
    config_path: Path | str,
    *,
    repository_root: Path | str = REPOSITORY_ROOT,
    allow_repository_config: bool = False,
) -> ProfileConfig:
    """Load local configuration and ensure its configured workbook still exists."""
    source = _validate_config_location(
        config_path,
        repository_root=repository_root,
        allow_repository_config=allow_repository_config,
    )
    try:
        payload = json.loads(source.read_text(encoding="utf-8"))
        profile = ProfileConfig(
            workbook_path=payload["workbook_path"],
            attachment_dir=payload.get("attachment_dir"),
        )
    except (OSError, TypeError, ValueError, KeyError, json.JSONDecodeError) as exc:
        raise ProfileConfigError("configuration could not be read") from exc
    if not profile.workbook_path.is_file():
        raise ProfileConfigError("workbook not found")
    return profile


def init_profile(
    workbook_path: Path | str,
    attachment_dir: Path | str | None,
    config_path: Path | str,
    *,
    template_path: Path | str = BLANK_TEMPLATE,
    repository_root: Path | str = REPOSITORY_ROOT,
    allow_repository_config: bool = False,
) -> ProfileConfig:
    """Initialize a local workbook and its config without overwriting an existing file."""
    destination = _validate_config_location(
        config_path,
        repository_root=repository_root,
        allow_repository_config=allow_repository_config,
    )
    workbook = _resolved(workbook_path)
    attachments = _resolved(attachment_dir) if attachment_dir is not None else None
    template = _resolved(template_path)

    if workbook.exists():
        if not workbook.is_file():
            raise ProfileConfigError("workbook path is not a file")
    else:
        if not template.is_file():
            raise ProfileConfigError("blank template not found")
        workbook.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template, workbook)

    if attachments is not None:
        attachments.mkdir(parents=True, exist_ok=True)

    profile = ProfileConfig(workbook_path=workbook, attachment_dir=attachments)
    write_config(
        profile,
        destination,
        repository_root=repository_root,
        allow_repository_config=allow_repository_config,
    )
    return profile
