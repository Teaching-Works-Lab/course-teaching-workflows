import json
from pathlib import Path

import yaml


ROOT = Path(__file__).parents[1]


def test_manifest_exposes_three_skills():
    manifest = json.loads(
        (ROOT / ".codex-plugin" / "plugin.json").read_text(encoding="utf-8")
    )
    assert manifest["name"] == "course-teaching-workflows"
    assert manifest["version"] == "0.1.0"
    assert manifest["skills"] == "./skills/"

    expected = {
        "course-foundation-builder",
        "course-assessment-planner",
        "course-syllabus-compiler",
    }
    actual = {path.name for path in (ROOT / "skills").iterdir() if path.is_dir()}
    assert actual == expected
    assert all((ROOT / "skills" / name / "SKILL.md").exists() for name in expected)


def test_skill_agent_manifests_are_utf8_yaml():
    expected_names = {
        "course-foundation-builder": "课程基座构建器",
        "course-assessment-planner": "课程考核设计器",
        "course-syllabus-compiler": "课程大纲编译器",
    }
    for path in (ROOT / "skills").glob("*/agents/openai.yaml"):
        payload = yaml.safe_load(path.read_text(encoding="utf-8"))
        assert payload["interface"]["display_name"] == expected_names[path.parents[1].name]
        assert payload["interface"]["default_prompt"].startswith("Use $")
