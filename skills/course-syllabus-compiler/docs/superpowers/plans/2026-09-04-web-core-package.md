# Course Syllabus Compiler Web/Core Package Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a self-contained, no-Router course-syllabus Skill ZIP that can author, validate, render, and inspect school-format DOCX files in a Web Harness or Terminal environment with filesystem/script access.

**Architecture:** One model authors the course content; deterministic Python scripts own structural validation, school-layout rendering, and DOCX inspection. The package includes a reusable blank master, machine-readable layout contract, two reviewed course examples, and a `syllabus.py` CLI compatible with `check`, `render`, and `inspect` workflows.

**Tech Stack:** Python 3.11+, python-docx, PyYAML, OOXML ZIP inspection, pytest.

**Spec:** Approved Web/Core design in the project conversation: no external model Router; retain a complete renderer/master/contract/checker package.

## Global Constraints

- No `router.yaml`, `route.py`, model registry, or automatic multi-model dependency.
- Markdown is semantic input; final Word table geometry is determined by renderer code and layout contracts.
- First Word table is the school 12-row × 8-column basic-information form with a two-row hours header and course introduction inside the table.
- Final DOCX uses A4, school fonts, valid twip widths, no comments/revisions/hidden text/tool authorship metadata.
- Selected textbook entries end at the ISBN digits without a trailing period.
- `授课学院` must use one of the nine formal school names.
- Visual QA is required for the packaged Golden DOCX outputs before release.

---

### Task 1: Freeze package contract tests

**Files:** Create `tests/test_web_core_contract.py`, `tests/test_renderer.py`.

- [ ] Write tests that require no Router files, require the renderer/master/contracts, validate the nine-college whitelist and textbook punctuation, and assert rendered DOCX basic-info structure.
- [ ] Run tests and confirm failure because package files do not yet exist.

### Task 2: Implement semantic validation and DOCX inspection

**Files:** Create `scripts/validate_release.py`, `scripts/inspect_docx.py`, `references/*.md`.

- [ ] Implement Markdown hard-rule checks.
- [ ] Implement OOXML/DOCX cleanliness and layout checks.
- [ ] Run focused tests until green.

### Task 3: Implement deterministic school renderer

**Files:** Create `scripts/render_school_syllabus.py`, `scripts/syllabus.py`, `assets/syllabus-master.docx`.

- [ ] Parse the approved school-style Markdown shape.
- [ ] Render school basic-information, goals, teaching, assessment, bibliography, and signature sections.
- [ ] Expose `check`, `render`, and `inspect` through `syllabus.py`.
- [ ] Run renderer tests until green.

### Task 4: Add Skill documentation and reviewed examples

**Files:** Create `SKILL.md`, `README.md`, `assets/syllabus-working-template.md`, `examples/*`, `agents/openai.yaml`.

- [ ] Explain the one-model Web/Core workflow and capability boundary.
- [ ] Include the reviewed AI and intelligent-equipment examples as content/layout references, with a warning not to copy course facts without upstream verification.
- [ ] Run all package tests.

### Task 5: Render, inspect, and seal

- [ ] Render both example Markdown files using the packaged renderer.
- [ ] Inspect every rendered page visually.
- [ ] Run `syllabus.py check/render/inspect`, pytest, ZIP integrity, and SHA-256 checks.
- [ ] Create the final uploadable ZIP with a single top-level `course-syllabus-compiler/` directory.
