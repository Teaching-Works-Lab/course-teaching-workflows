from pathlib import Path
import subprocess
import sys


def test_invalid_syllabus_reports_objective_blockers(fixtures: Path, syllabus_module):
    data = syllabus_module.read_syllabus(fixtures / "invalid-syllabus.md")
    issues = syllabus_module.check_syllabus(data, level="basic")
    codes = {issue.code for issue in issues}
    assert {
        "MISSING_REQUIRED_FIELD",
        "HOURS_MISMATCH",
        "ASSESSMENT_WEIGHT_MISMATCH",
        "GOAL_NOT_ASSESSED",
    } <= codes


def test_valid_syllabus_has_no_basic_blockers(fixtures: Path, syllabus_module):
    data = syllabus_module.read_syllabus(fixtures / "valid-syllabus.md")
    issues = syllabus_module.check_syllabus(data, level="basic")
    assert [issue for issue in issues if issue.severity == "blocking"] == []


def test_semantic_quality_is_a_warning_not_a_blocker(fixtures: Path, syllabus_module):
    data = syllabus_module.read_syllabus(fixtures / "valid-syllabus.md")
    issues = syllabus_module.check_syllabus(data, level="enhanced")
    semantic = [issue for issue in issues if issue.code.startswith("SEMANTIC_")]
    assert semantic
    assert {issue.severity for issue in semantic} == {"warning"}


def test_unchanged_build_signature_reuses_check_state(
    tmp_path: Path, fixtures: Path, syllabus_module
):
    state = tmp_path / "build-state.json"
    first = syllabus_module.check_file(fixtures / "valid-syllabus.md", "basic", state)
    second = syllabus_module.check_file(fixtures / "valid-syllabus.md", "basic", state)
    assert first["reused"] is False
    assert second["reused"] is True


def test_check_cli_writes_report_and_returns_two_for_blockers(
    tmp_path: Path, fixtures: Path, plugin_root: Path
):
    script = (
        plugin_root
        / "skills"
        / "course-syllabus-compiler"
        / "scripts"
        / "syllabus.py"
    )
    report = tmp_path / "review.md"
    result = subprocess.run(
        [
            sys.executable,
            str(script),
            "check",
            "--input",
            str(fixtures / "invalid-syllabus.md"),
            "--report",
            str(report),
            "--level",
            "basic",
        ],
        text=True,
        encoding="utf-8",
        capture_output=True,
    )
    assert result.returncode == 2
    assert "MISSING_REQUIRED_FIELD" in report.read_text(encoding="utf-8")
