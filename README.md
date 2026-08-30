# Course Teaching Workflows

A public Codex plugin for building reusable course evidence, assessment plans, and standards-compliant university course syllabi without forcing every request through one fixed pipeline.

The plugin contains three independently callable Skills:

- `course-foundation-builder` converts and synthesizes PPT, Word, PDF, Markdown,教材,教案,培养方案 data, and reference syllabi into a reusable `course-foundation.md`.
- `course-assessment-planner` designs, extracts, or reviews a course-specific `assessment-plan.md` with weights, goal/unit coverage, rubrics, and evidence requirements.
- `course-syllabus-compiler` provides an AI-friendly Markdown contract, deterministic checks, and clean DOCX compilation based on an annotated university syllabus format.

## Install

Install the public repository as a Codex plugin:

```text
https://github.com/Teaching-Works-Lab/course-teaching-workflows
```

One installation exposes all three Skills. They may call one another conditionally, but none requires a fixed Task 1 → Task 2 → Task 3 sequence.

## Typical routes

| Available input / request | Route | Stop point |
|---|---|---|
| PPT,教材,参考大纲,教案 | Course foundation | `course-foundation.md` and gap report |
| 培养方案 Skill only | Compact evidence request | Missing facts with examples |
| Existing syllabus with course evaluation | Assessment extraction/review | `assessment-plan.md` |
| Generic syllabus Markdown requested | Syllabus template | Reusable `.md`; no course package or Word |
| Valid course foundation + assessment plan + working syllabus | Word compilation | Clean `.docx` plus inspection report |

## Course package contract

Course-specific data stays outside this plugin. A typical private course package can contain:

```text
course-package/
├── raw/                       # original PPT, Word, PDF,教材,参考资料
├── markdown-cache/            # source-derived Markdown and manifest
├── course-foundation.md       # reusable course evidence and decisions
├── assessment-plan.md         # editable assessment source
├── syllabus-working.md        # AI-friendly complete syllabus
├── build-state.json           # optional incremental state
└── output/                    # generated Word and reports
```

`assessment-plan.md` is the editable source for assessment design. Its complete content is materialized into `syllabus-working.md`; compilation stops if their version or hash differs.

## Reusable tools

Convert local sources once and reuse unchanged Markdown. Invoke the Python core directly on any supported platform:

```bash
python skills/course-foundation-builder/scripts/update_markdown_cache.py \
  --source-root <source-directory> --cache-root <cache-directory>
```

On Windows, the PowerShell adapter forwards to the same core:

```powershell
./skills/course-foundation-builder/scripts/update_markdown_cache.ps1 `
  -SourceRoot <source-directory> -CacheRoot <cache-directory>
```

On Linux or macOS, use the Shell adapter:

```sh
./skills/course-foundation-builder/scripts/update_markdown_cache.sh \
  --source-root <source-directory> --cache-root <cache-directory>
```

PowerShell is the Windows adapter, not the workflow definition. The Python core defines cache behavior; both adapters only select the platform-native entry point.

Check a working syllabus:

```bash
python skills/course-syllabus-compiler/scripts/syllabus.py check \
  --input syllabus-working.md --report review.md --level basic
```

Compile and inspect Word:

```bash
python skills/course-syllabus-compiler/scripts/syllabus.py render \
  --input syllabus-working.md \
  --assessment assessment-plan.md \
  --template skills/course-syllabus-compiler/assets/syllabus-master.docx \
  --output syllabus.docx \
  --state build-state.json

python skills/course-syllabus-compiler/scripts/syllabus.py inspect \
  --input syllabus.docx --report word-inspection.md
```

The Python compiler requires Python 3.10+, `PyYAML`, and `python-docx`. Source conversion prefers local MarkItDown and Pandoc when available. Legacy `.doc` conversion may require a compatible Windows Word COM provider.

## Review scope

- `basic`: required fields, totals, IDs, goal–unit–assessment closure, synchronization, and unresolved placeholders.
- `changed`: focus on changed sections and their direct dependencies.
- `enhanced`: semantic quality warnings such as generic unit names.

Warnings do not automatically trigger a full audit. Word compilation always runs the basic gate and final DOCX cleanliness inspection.

## Word output boundary

The reusable master contains only page, font, paragraph, table, and page-number styles. It contains no real course content. Final DOCX inspection rejects comments, tracked changes, hidden text, unresolved placeholders, AI/tool authorship markers, and tool metadata. Normal academic phrases such as “人工智能” or “AI 算法” are allowed.

## Tests

```bash
python -m pytest tests -q
```

The test fixtures are fictional. Tests cover cache reuse, real MarkItDown/Pandoc conversions when the tools are present, assessment contracts, syllabus checks, two-course Word compilation, stale-assessment rejection, and final DOCX inspection.

## Privacy

Do not commit course packages, raw teaching materials, conversion caches, build states, generated outputs, student data, local paths, or private review logs to this repository. Public assets and examples must remain fictional and de-identified.

## License

MIT License. See [LICENSE](LICENSE).
