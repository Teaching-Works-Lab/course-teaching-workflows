---
name: course-foundation-builder
description: Use when course materials, training-program data, reference syllabi, teaching plans, PPT/PPTX, DOC/DOCX, PDF, or Markdown need to be converted, organized, reviewed, or synthesized into a reusable course foundation for syllabi, lesson plans, assessment, or teaching assets.
---

# Course Foundation Builder

Build only the course evidence and foundation needed by the user's current task. Accept mixed source formats and upstream Skills without requiring a common directory layout.

## Conditional routing

- If a valid `course-foundation.md` already covers the requested facts, use it and stop; do not rebuild the foundation.
- If the user provides only a professional/training-program Skill, request the compact course facts defined below and return gaps with examples.
- If raw PPT, PDF, Word,教材,教案, or reference syllabi are provided, convert unchanged sources once, synthesize the foundation, and stop unless another artifact was explicitly requested.
- Downstream syllabus, lesson-plan, presentation, and assessment Skills may consume the foundation independently or in parallel. This Skill does not prescribe their order.

- Preserve original files and distinguish official, provided, derived, reference-only, and unresolved content.
- Reuse an existing valid course foundation or Markdown cache instead of re-reading unchanged sources.
- Record missing facts, conflicts, decisions, and evidence boundaries; never promote reference material into official course facts without confirmation.
- Stop after the requested foundation, source analysis, template, or gap report. Do not generate a syllabus Word document.
- Produce course-specific data in the course package, never in this Skill's reusable instructions or assets.

## Resources

- For source conversion, cache routing, and original-file review boundaries, read [conversion-strategy.md](references/conversion-strategy.md). Run `scripts/update_markdown_cache.ps1` directly when its interface fits; do not rewrite it per course.
- When a conversion fails or the tool environment changes, read [conversion-known-issues.md](references/conversion-known-issues.md) and keep local observations separate from generalized rules.
- Copy [course-foundation-template.md](assets/course-foundation-template.md) for a new foundation. Use [course-foundation-example.md](assets/course-foundation-example.md) only to understand field roles and evidence states.
- Use [course-evidence-request.md](assets/course-evidence-request.md) to ask an upstream professional or course Skill for only the target course's information.
