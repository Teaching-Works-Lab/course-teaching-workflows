---
name: course-assessment-planner
description: Use when a university course assessment plan needs to be designed, extracted from an existing syllabus, reviewed, or revised, including assessment weights, learning-goal coverage, scoring criteria, and evidence requirements.
---

# Course Assessment Planner

Choose the mode that matches the request: design a plan, extract it from an existing syllabus, or review an existing plan.

## Conditional routing

- If a valid `assessment-plan.md` exists and its goals, units, hours, and constraints have not changed, reuse it.
- If an existing syllabus already contains course evaluation, extract that content directly before asking for unrelated raw materials.
- If goals, units, or official constraints are missing, request the smallest missing set; use `course-foundation-builder` only when raw sources must be synthesized.
- Return `assessment-plan.md` and stop unless the user also requested syllabus drafting or Word compilation.

- Use available course goals, teaching units, course type, and official constraints; report missing inputs instead of inventing them.
- Treat `assessment-plan.md` as the course-specific editable source for assessment content.
- Check weights, goal and unit coverage, scoring criteria, implementation evidence, and course-type fit in proportion to the task.
- Reuse a valid existing assessment plan. Re-review it when goals, units, hours, or assessment content changes; do not reload this Skill for unrelated edits.
- Do not generate the final syllabus Word document.

## Resources

- Read [assessment-design-rules.md](references/assessment-design-rules.md) when designing, extracting, or reviewing weights, coverage, criteria, or evidence.
- Copy [assessment-plan-template.md](assets/assessment-plan-template.md) for course work. Use [assessment-plan-example.md](assets/assessment-plan-example.md) only to understand the contract; its data is fictional.

Return a gap list with a concrete example when required inputs are missing. Keep internal source states and review notes outside the formal assessment text that will be materialized into the syllabus.
