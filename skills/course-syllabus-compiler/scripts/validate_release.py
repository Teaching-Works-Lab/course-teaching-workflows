from __future__ import annotations

import argparse
from pathlib import Path
import re
from typing import Sequence

COLLEGES = {
    "经济管理学院",
    "机械与动力工程学院",
    "土木建筑工程学院",
    "电子与电气工程学院",
    "计算机与信息技术学院",
    "轨道交通学院",
    "艺术学院",
    "外国语学院",
    "化学与制药工程学院",
}
FORBIDDEN = (
    "to_confirm",
    "draft_for_review",
    "drafted_course_design",
    "source_id",
    "AI生成",
    "由AI生成",
    "ChatGPT生成",
    "Codex生成",
    "[待填写]",
)


def _selected_textbook(text: str) -> str | None:
    match = re.search(r"###\s*（一）选用教材\s*\n\s*\n([^\n]+)", text)
    return match.group(1).strip() if match else None


def _basic_table(text: str) -> dict[str, str]:
    section = re.search(r"##\s+一、课程基本信息([\s\S]*?)(?=^##\s+)", text, re.MULTILINE)
    if not section:
        return {}
    result: dict[str, str] = {}
    for line in section.group(1).splitlines():
        if not line.strip().startswith("|"):
            continue
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        if len(cells) >= 2 and cells[0] not in {"项目", "---"} and not set(cells[0]) <= {"-", ":"}:
            result[cells[0]] = cells[1]
    return result


def _parse_hours(value: str) -> dict[str, int]:
    result: dict[str, int] = {}
    for part in re.split(r"[；;,，]", value):
        match = re.match(r"\s*([^0-9]+?)\s*([0-9]+)\s*(?:学时)?\s*$", part)
        if match:
            result[match.group(1).strip()] = int(match.group(2))
    return result


def check_markdown(path: Path) -> list[dict[str, str]]:
    text = Path(path).read_text(encoding="utf-8")
    issues: list[dict[str, str]] = []

    def add(code: str, message: str) -> None:
        issues.append({"code": code, "message": message})

    if not re.search(r"^#\s+《.+》课程教学大纲\s*$", text, re.MULTILINE):
        add("TITLE_FORMAT", "缺少《课程名称》课程教学大纲格式的一级标题")

    for marker in FORBIDDEN:
        if marker in text:
            add("FORBIDDEN_MARKER", f"终稿包含内部或生成标记：{marker}")

    fields = _basic_table(text)
    college = fields.get("授课学院") or fields.get("开课单位")
    if not college:
        add("COLLEGE_MISSING", "课程基本信息缺少授课学院")
    elif college not in COLLEGES:
        add("COLLEGE_NOT_IN_WHITELIST", f"授课学院必须使用正式全称：{college}")
    if "计信学院" in text:
        add("COLLEGE_ABBREVIATION", "不得使用简称“计信学院”")

    textbook = _selected_textbook(text)
    if not textbook:
        add("TEXTBOOK_MISSING", "缺少选用教材条目")
    else:
        if not re.search(r"ISBN\s+\d{13}$", textbook):
            add("TEXTBOOK_TERMINAL_ISBN", "选用教材必须以13位ISBN数字结束，ISBN后不加句号或附加说明")
        if "课程使用第" in textbook or re.search(r"第\s*\d+\s*[—-]\s*\d+\s*章", textbook):
            add("TEXTBOOK_CHAPTER_SCOPE", "选用教材条目不得附写课程使用章节")

    hours = _parse_hours(fields.get("课程学时及分配", ""))
    if hours:
        total = hours.get("总学时")
        components = sum(hours.get(name, 0) for name in ("理论", "实践", "实验", "上机", "项目式", "项目"))
        if total is not None and total != components:
            add("HOURS_TOTAL_MISMATCH", f"总学时{total}与分项学时合计{components}不一致")

    headings = re.findall(r"^##\s+([一二三四五六七八九十]+)、", text, re.MULTILINE)
    mapping = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
    values = [mapping[h] for h in headings if h in mapping]
    if values and values != list(range(values[0], values[0] + len(values))):
        add("SECTION_NUMBERING", f"一级章节编号不连续：{headings}")

    for required in ("课程目标", "课程评价", "参考书目及学习资料"):
        if required not in text:
            add("SECTION_MISSING", f"缺少章节：{required}")
    return issues


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate a school syllabus Markdown release")
    parser.add_argument("--input", required=True, type=Path)
    args = parser.parse_args(argv)
    issues = check_markdown(args.input)
    if issues:
        for issue in issues:
            print(f"{issue['code']}: {issue['message']}")
        return 2
    print("PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
