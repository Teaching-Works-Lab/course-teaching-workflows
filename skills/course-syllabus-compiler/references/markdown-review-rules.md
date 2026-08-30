# Syllabus Markdown and Review Rules

## Working document contract

The working Markdown is optimized for AI review and content collaboration, not for imitating Word layout. It keeps:

- YAML control fields and assessment synchronization hashes;
- source and evidence boundaries outside formal content;
- one field-value table for course facts;
- stable IDs for goals, units, assessments, and scoring criteria;
- the complete materialized assessment text;
- internal gaps and change traces outside formal content.

Use `not_applicable` for a course module that does not apply. Do not create a zero-filled empty table to make every course type look the same.

## Evidence and formal content

Formal Word may include `official`, `provided`, and confirmed `derived` content. It must not include `reference_only`, unresolved `to_confirm`, source-state labels, internal review comments, or change proposals.

The assessment plan remains the editable source. The working syllabus stores `assessment_version` and `assessment_hash` and carries a complete readable copy. A mismatch blocks compilation.

## Review levels

- `basic`: objective formal-delivery gates—required fields, hours, weights, IDs, goal-unit-assessment relationships, assessment synchronization, and unresolved formal placeholders.
- `changed`: reuse an unchanged build signature; when content changed, rerun the cheap objective gates and restrict additional review to affected sections and downstream relationships.
- `enhanced`: use for a new template or environment, unreliable extraction, official-source conflicts, batch generation, major structural change, a failed basic gate, or an explicit comprehensive-review request.

Only objective values are asserted. Pedagogical fit, wording, difficulty, authenticity, ideology integration, and unit naming are warnings or decisions with evidence. A warning does not become a blocker merely because an enhanced review was requested.

## Stop conditions

Do not run equivalent checks again when input, rules, script, configuration, and environment signatures are unchanged. After the requested artifact passes its applicable gate and affected-range review, stop. Do not generate Word for a template-only, analysis-only, or review-only request.
