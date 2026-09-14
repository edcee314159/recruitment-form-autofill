"""Command-line initializer for a local recruitment-profile workbook."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from profile_config import ProfileConfigError, init_profile


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Create a local recruitment profile from the blank public template."
    )
    parser.add_argument("--workbook", type=Path, required=True, metavar="PATH")
    parser.add_argument("--config", type=Path, required=True, metavar="PATH")
    parser.add_argument("--attachments", type=Path, metavar="DIRECTORY")
    parser.add_argument(
        "--allow-repository-config",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        init_profile(
            args.workbook,
            args.attachments,
            args.config,
            allow_repository_config=args.allow_repository_config,
        )
    except ProfileConfigError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
