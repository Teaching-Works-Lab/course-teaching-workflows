# Conversion Known Issues

This reference stores only de-identified, reusable experience. Course names, user paths, private files, and raw task logs stay in the course package or `_markdown_cache/conversion-log.jsonl`.

## Status model

- `candidate`: observed but not reliably reproduced or bounded;
- `verified-local`: reproduced and resolved for a named tool, version, and environment only;
- `generalized`: applicability, exclusions, conflicts, and repeatable regression evidence are documented.

One run may append evidence but must not edit `SKILL.md` or promote an issue automatically. Conflicting experiences remain side by side with their applicability conditions. Revalidate an issue after relevant tool or environment versions change.

## KI-001 — Reused Word application fails during legacy DOC batches

- **Status:** `verified-local`
- **Format:** legacy `.doc`
- **Environment:** Windows; Microsoft Word 12.0 COM automation
- **Symptom:** later documents in a batch fail with COM/RPC unavailable or RPC call failure after save or close
- **Verified workaround:** create and close a fresh Word Application instance for each `.doc`, save to a temporary `.docx`, then run MarkItDown
- **Verification:** the same local files converted successfully when isolated per Word Application instance
- **Applicability:** the verified Word 12.0 environment only
- **Exclusions:** no claim is made about newer Word, WPS, LibreOffice, or every `.doc` file
- **Conflict rule:** if a newer environment safely reuses one application, retain both results and route by version
- **Last verified:** 2026-08-30

## KI-002 — COM shutdown reports RPC failure after a valid save

- **Status:** `verified-local`
- **Format:** legacy `.doc` converted to `.docx`
- **Environment:** Windows; local `Word.Application` COM provider
- **Symptom:** `Quit()` or final cleanup raises an RPC unavailable/call-failed exception although the target `.docx` was already written
- **Verified handling:** do not equate a shutdown exception with a failed conversion; first verify that the target exists, is a readable ZIP/DOCX, and can be opened by the next parser, while still logging the cleanup exception
- **Verification:** the emitted target opened successfully with `python-docx` and its structure was inspected after the shutdown exception
- **Applicability:** only the observed local COM provider and files
- **Exclusions:** a present file is not sufficient by itself; unreadable or truncated DOCX output remains a failed conversion
- **Conflict rule:** if the save itself fails or the target cannot be parsed, return failure even when a path exists
- **Last verified:** 2026-08-30
