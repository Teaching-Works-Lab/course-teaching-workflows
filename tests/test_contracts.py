from collections import defaultdict
from pathlib import Path

import yaml


def front_matter(text: str) -> dict:
    _, raw, _ = text.split("---", 2)
    return yaml.safe_load(raw)


def table_after(text: str, heading: str) -> list[dict[str, str]]:
    section = text.split(heading, 1)[1]
    section = section.split("\n#", 1)[0]
    lines = [line for line in section.splitlines() if line.startswith("|")]
    headers = [cell.strip() for cell in lines[0].strip("|").split("|")]
    rows = []
    for line in lines[2:]:
        values = [cell.strip() for cell in line.strip("|").split("|")]
        rows.append(dict(zip(headers, values, strict=True)))
    return rows


def test_assessment_example_is_a_complete_reusable_contract(plugin_root: Path):
    path = (
        plugin_root
        / "skills"
        / "course-assessment-planner"
        / "assets"
        / "assessment-plan-example.md"
    )
    text = path.read_text(encoding="utf-8")
    metadata = front_matter(text)
    assessments = table_after(text, "# 考核项目")
    coverage = table_after(text, "# 目标与单元覆盖")
    criteria = table_after(text, "# 评分标准")

    assert metadata["document_state"] == "example_only"
    assert metadata["content_hash"]
    assert sum(int(row["权重"]) for row in assessments) == 100

    assessment_ids = {row["assessment_id"] for row in assessments}
    assert {row["assessment_id"] for row in coverage} == assessment_ids
    assert all(row["goal_id"] and row["unit_id"] for row in coverage)

    criteria_totals = defaultdict(int)
    for row in criteria:
        assert row["assessment_id"] in assessment_ids
        criteria_totals[row["assessment_id"]] += int(row["指标权重"])
    assert criteria_totals == {assessment_id: 100 for assessment_id in assessment_ids}
