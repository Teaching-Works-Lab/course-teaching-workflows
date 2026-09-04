from pathlib import Path
import re
import yaml

ROOT = Path(__file__).resolve().parents[1]


def read(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8")


def test_package_is_web_core_without_external_router():
    assert (ROOT / "SKILL.md").exists()
    assert (ROOT / "scripts" / "syllabus.py").exists()
    assert (ROOT / "scripts" / "render_school_syllabus.py").exists()
    assert (ROOT / "assets" / "syllabus-master.docx").exists()
    assert not (ROOT / "router.yaml").exists()
    assert not (ROOT / "scripts" / "route.py").exists()
    assert not (ROOT / "model-registry.example.yaml").exists()


def test_skill_is_standalone_and_marks_markdown_as_semantic_not_layout():
    body = read("SKILL.md")
    for phrase in ["完整初稿", "Markdown 只表达语义", "scripts/syllabus.py", "school-layout-contract.yaml"]:
        assert phrase in body
    assert "model-registry" not in body


def test_college_whitelist_contains_exact_nine_formal_names():
    body = read("references/school-format-contract.md")
    colleges = [
        "经济管理学院",
        "机械与动力工程学院",
        "土木建筑工程学院",
        "电子与电气工程学院",
        "计算机与信息技术学院",
        "轨道交通学院",
        "艺术学院",
        "外国语学院",
        "化学与制药工程学院",
    ]
    assert all(name in body for name in colleges)
    assert "不得使用简称" in body


def test_bibliography_contract_ends_selected_textbook_at_isbn():
    body = read("references/bibliography-contract.md")
    good = "戴波. C与C++程序设计（第二版）[M]. 北京：北京大学出版社，2025. ISBN 9787301362648"
    assert good in body
    assert "ISBN 后直接结束" in body
    assert "课程使用第" in body


def test_layout_contract_has_fixed_school_basic_info_shape():
    data = yaml.safe_load(read("references/school-layout-contract.yaml"))
    assert data["page"]["width_twip"] == 11906
    assert data["page"]["height_twip"] == 16838
    assert data["basic_info"]["rows"] == 12
    assert data["basic_info"]["columns"] == 8
    assert len(data["basic_info"]["grid_twip"]) == 8
    assert data["basic_info"]["course_intro_inside_table"] is True
    assert data["basic_info"]["hours_use_two_row_header"] is True


def test_readme_explains_web_and_harness_capability_boundary():
    body = read("README.md")
    assert "普通 Web 对话" in body
    assert "文件系统和 Python" in body
    assert "不自动切换模型" in body
    assert "check" in body and "render" in body and "inspect" in body


def test_examples_selected_textbook_have_terminal_isbn():
    for path in (ROOT / "examples").glob("*.md"):
        text = path.read_text(encoding="utf-8")
        m = re.search(r"### （一）选用教材\s*\n\s*\n([^\n]+)", text)
        assert m, path.name
        line = m.group(1).strip()
        assert re.search(r"ISBN\s+\d{13}$", line), (path.name, line)
