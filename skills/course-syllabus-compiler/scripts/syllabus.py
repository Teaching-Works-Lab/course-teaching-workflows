from __future__ import annotations

import argparse
from pathlib import Path
import sys
from typing import Sequence

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from inspect_docx import inspect_docx
from render_school_syllabus import render_syllabus
from validate_release import check_markdown


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Course syllabus Web/Core compiler")
    subparsers = parser.add_subparsers(dest="command", required=True)

    check = subparsers.add_parser("check", help="Validate syllabus Markdown")
    check.add_argument("--input", required=True, type=Path)

    render = subparsers.add_parser("render", help="Render school-format DOCX")
    render.add_argument("--input", required=True, type=Path)
    render.add_argument("--output", required=True, type=Path)
    render.add_argument("--master", type=Path)

    inspect = subparsers.add_parser("inspect", help="Inspect final DOCX")
    inspect.add_argument("--input", required=True, type=Path)

    args = parser.parse_args(argv)
    if args.command == "check":
        issues = check_markdown(args.input)
        if issues:
            for issue in issues:
                print(f"{issue['code']}: {issue['message']}")
            return 2
        print("PASS")
        return 0
    if args.command == "render":
        issues = check_markdown(args.input)
        if issues:
            for issue in issues:
                print(f"{issue['code']}: {issue['message']}")
            return 2
        render_syllabus(args.input, args.output, args.master)
        print(args.output)
        return 0
    if args.command == "inspect":
        report = inspect_docx(args.input)
        if report["issues"]:
            for issue in report["issues"]:  # type: ignore[index]
                print(f"{issue['code']}: {issue['message']}")
            return 2
        print("PASS")
        return 0
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
