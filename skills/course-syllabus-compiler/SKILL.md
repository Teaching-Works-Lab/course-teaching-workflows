---
name: course-syllabus-compiler
description: Use when a university course syllabus needs to be planned, drafted, continued, reviewed, revised, checked, or compiled from Markdown and supporting course materials into a standards-compliant Word document.
---

# Course Syllabus Compiler

Route directly to the artifact the user requests: a generic Markdown template, a course-specific working draft, a review, a revision, a Word compilation, or a Word inspection.

## Conditional routing

- For a generic Markdown request, return the reusable template and stop; do not create a course package or Word file.
- For raw course materials without a valid foundation, use `course-foundation-builder` for only the missing evidence.
- For missing, stale, or newly edited course evaluation, use `course-assessment-planner`, then materialize the confirmed plan into the working syllabus.
- For an already valid course package, run the basic gate and compile Word directly. Reuse unchanged build state.
- Review or inspect only the requested artifact. A warning does not trigger an unrelated full audit.

- Reuse valid `course-foundation.md`, `assessment-plan.md`, Markdown caches, templates, scripts, and prior checks when available.
- Keep the complete assessment text materialized in the working syllabus while treating `assessment-plan.md` as its editable source; block compilation when their versions or hashes differ.
- Separate internal evidence states, gaps, and review comments from formal syllabus content.
- Do not require a fixed workflow and do not generate Word unless requested.
- Formal Word output must contain no AI authorship or tool markers, comments, revisions, hidden prompts, unresolved placeholders, or metadata residue.

## Resources

- Copy [syllabus-working-template.md](assets/syllabus-working-template.md) when the user needs a generic or course-specific Markdown draft. Use [syllabus-working-example.md](assets/syllabus-working-example.md) only to understand the contract; its data is fictional.
- Read [markdown-review-rules.md](references/markdown-review-rules.md) when establishing, revising, or checking Markdown, choosing a review level, or deciding whether a warning blocks Word compilation.
- Run `scripts/syllabus.py check --help` for deterministic structure, total, relationship, placeholder, and build-state checks. Reuse the script; do not generate a course-specific checker.
- Read [word-layout-rules.md](references/word-layout-rules.md) before compiling or visually reviewing Word. The rules use the union of two annotated reference syllabi and retain course-type variants.
- Use [syllabus-master.docx](assets/syllabus-master.docx) as the reusable blank master. Compile with `scripts/syllabus.py render --help`; inspect a final document with `scripts/syllabus.py inspect --help`.
