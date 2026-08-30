from pathlib import Path

import yaml


INTENT_FIELDS = {
    "schema_version",
    "intent_version",
    "mode",
    "objective",
    "requested_artifacts",
    "stop_condition",
    "input_scope",
    "official_constraints",
    "logic_chain",
    "unresolved_items",
    "execution_requested",
}


def front_matter(path: Path) -> dict:
    text = path.read_text(encoding="utf-8")
    _, raw, _ = text.split("---", 2)
    return yaml.safe_load(raw)


def intent_template(plugin_root: Path) -> dict:
    path = (
        plugin_root
        / "skills"
        / "course-foundation-builder"
        / "assets"
        / "workflow-intent-template.md"
    )
    return front_matter(path)


def test_intent_template_has_the_complete_semantic_contract(plugin_root: Path):
    metadata = intent_template(plugin_root)

    assert set(metadata) == INTENT_FIELDS
    assert metadata["schema_version"] == 1
    assert metadata["intent_version"] == 1


def test_intent_template_defaults_to_os_independent_plan_only(plugin_root: Path):
    metadata = intent_template(plugin_root)

    assert metadata["mode"] == "plan-only"
    assert metadata["execution_requested"] is False
    assert "operating_system" not in metadata
    assert "shell" not in metadata
    assert metadata["requested_artifacts"] == []
    assert metadata["input_scope"] == []
    assert metadata["official_constraints"] == []
    assert metadata["logic_chain"] == []
    assert metadata["unresolved_items"] == []
