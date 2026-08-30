# Source Conversion Strategy

Convert once, preserve the original, and reuse Markdown while the source SHA-256 is unchanged. Conversion makes content easier to analyze; it does not make the cache more authoritative than the source or verify the facts it contains.

## Format routes

| Source | Preferred route | Required review boundary |
|---|---|---|
| Markdown | Register and copy without semantic conversion | Compare with the declared source role |
| PDF | MarkItDown | Review the original for scans, columns, complex tables, figures, or formulas |
| DOCX | MarkItDown | Use Pandoc only as an alternate or cross-check when structure differs materially |
| DOC | Word, WPS, or LibreOffice to temporary DOCX; then MarkItDown | Record compatibility tool and version |
| HTML, EPUB | Pandoc to GitHub-flavored Markdown | Check embedded media and footnotes |
| PPTX, XLSX | Appropriate presentation or spreadsheet tooling | Keep slide, sheet, and cell locators |

Pandoc is not a PDF reader. Do not build an unconditional MarkItDown-to-Pandoc fallback chain; choose by source format and structure.

## Cache contract

`scripts/update_markdown_cache.py` is the sole authority for cache semantics, command behavior, manifest fields, and exit codes. It accepts `--source-root`, `--cache-root`, optional tool commands, `--force`, and `--json`.

The checked-in adapters are convenience entry points to that Python core:

- Windows PowerShell: `scripts/update_markdown_cache.ps1`
- Linux/macOS Shell: `scripts/update_markdown_cache.sh`

They locate Python and forward arguments; they do not define a separate cache contract. The core writes:

```text
_markdown_cache/
├── manifest.json
├── conversion-log.jsonl
└── by-source/
```

`manifest.json` records relative source path, SHA-256, timestamp, route, tool and version, status, and cache path. `conversion-log.jsonl` is append-only run evidence. Source files and cached Markdown are never overwritten in place.

Keep duplicate and historical source versions. Prefer the newest version only when its formal status is also clear, and record why it was selected. If formal status is unclear, report the conflict instead of guessing from the filename.

## Exit codes

- `0`: conversion or reuse completed without a required review flag;
- `1`: at least one source could not be read or converted;
- `2`: conversion completed, but one or more sources require artifact or visual review.
