# Assessment Design Rules

## Choose one mode

- **Design:** use a course foundation, confirmed course goals, teaching units, course type, and official constraints to create `assessment-plan.md`.
- **Extract:** copy the complete course-evaluation content from an existing syllabus into the assessment contract, preserving uncertainty and source location.
- **Review:** inspect an existing plan without replacing confirmed choices unless the user asks for revision.

An assessment plan that states only broad ratios such as “final exam 60%, coursework 40%” is an existing requirement, not a complete design. Report the missing goal coverage, scoring criteria, implementation stage, and evidence instead of inventing them.

## Required invariants

- Assessment item weights total 100.
- Criteria under each assessment item total 100.
- Every confirmed course goal is taught by at least one unit and assessed by at least one item.
- Every assessment item identifies its covered goals and units.
- The assessment method, timing, evidence, and scoring criteria fit the course type and taught content.
- `assessment_id`, `criterion_id`, `goal_id`, and `unit_id` remain stable across revisions.

Weights, totals, missing IDs, and version hashes are objective checks. Quality, difficulty, authenticity, and pedagogical fit are evidence-backed warnings or decisions, not false-precision assertions.

## Ownership and synchronization

`assessment-plan.md` is the only course-specific editable source. The syllabus carries a complete materialized copy so another AI can read it independently, plus `assessment_version` and `assessment_hash`. If a user edits the syllabus copy, reconcile the change into the assessment plan, review it, and then rematerialize it.

Re-review after goals, units, hours, assessment items, weights, or criteria change. Do not invoke assessment design for changes limited to course name, textbook, description, or signatures.

## Evidence boundaries

Mark each constraint or decision as `official`, `provided`, `derived`, `reference_only`, or `to_confirm`. External syllabi and textbooks may inform design but do not override the current training program or confirmed local policy.
