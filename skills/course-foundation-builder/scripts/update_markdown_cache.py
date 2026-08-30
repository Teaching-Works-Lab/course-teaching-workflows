"""Create and maintain a portable Markdown cache for course sources."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import tempfile
import uuid
import zipfile


SUPPORTED_EXTENSIONS = {".md", ".doc", ".docx", ".pdf", ".html", ".htm", ".epub", ".pptx", ".xlsx"}
WORD_COM_HELPER = Path(__file__).with_name("convert_legacy_doc.ps1")


def legacy_doc_strategy(system_name: str) -> str:
    if system_name == "Windows":
        return "word-com"
    if system_name in {"Linux", "Darwin"}:
        return "libreoffice"
    raise ValueError(f"Unsupported platform for legacy DOC conversion: {system_name}")


def run_command(command: str, arguments: list[str]) -> subprocess.CompletedProcess[str]:
    executable = shutil.which(command) or (command if Path(command).exists() else None)
    if executable is None:
        raise RuntimeError(f"Required command is unavailable: {command}")
    return subprocess.run(
        [str(executable), *arguments],
        text=True,
        encoding="utf-8",
        capture_output=True,
    )


def command_version(command: str) -> str:
    result = run_command(command, ["--version"])
    output = (result.stdout or result.stderr).strip()
    return output.splitlines()[0] if output else "unknown"


def require_success(result: subprocess.CompletedProcess[str], description: str) -> None:
    if result.returncode == 0:
        return
    detail = (result.stderr or result.stdout).strip()
    raise RuntimeError(f"{description} failed" + (f": {detail}" if detail else ""))


def word_com_tool_version(result: subprocess.CompletedProcess[str]) -> str:
    try:
        metadata = json.loads(result.stdout)
    except json.JSONDecodeError as error:
        raise RuntimeError("Word COM helper returned invalid JSON metadata") from error
    if not isinstance(metadata, dict):
        raise RuntimeError("Word COM helper metadata must be a JSON object")
    tool = metadata.get("tool")
    version = metadata.get("version")
    if not isinstance(tool, str) or not tool.strip():
        raise RuntimeError("Word COM helper metadata is missing a tool")
    if not isinstance(version, str) or not version.strip():
        raise RuntimeError("Word COM helper metadata is missing a version")
    return f"{tool.strip()} {version.strip()}"


def convert_with_markitdown(source: Path, destination: Path, args: argparse.Namespace) -> str:
    result = run_command(args.markitdown_command, [str(source), "-o", str(destination)])
    require_success(result, "MarkItDown conversion")
    if not destination.is_file():
        raise RuntimeError("MarkItDown conversion did not create the cache output")
    return command_version(args.markitdown_command)


def convert_with_pandoc(source: Path, destination: Path, args: argparse.Namespace) -> str:
    result = run_command(args.pandoc_command, [str(source), "-t", "gfm", "-o", str(destination)])
    require_success(result, "Pandoc conversion")
    if not destination.is_file():
        raise RuntimeError("Pandoc conversion did not create the cache output")
    return command_version(args.pandoc_command)


def convert_legacy_doc(
    source: Path,
    temporary_docx: Path,
    args: argparse.Namespace,
    system_name: str,
) -> tuple[str, str]:
    strategy = legacy_doc_strategy(system_name)
    temporary_docx.parent.mkdir(parents=True, exist_ok=True)
    if strategy == "word-com":
        if not WORD_COM_HELPER.is_file():
            raise RuntimeError(
                "Windows legacy DOC conversion requires the Word COM helper "
                f"at {WORD_COM_HELPER} (provided by Task 3)"
            )
        result = run_command(
            args.powershell_command,
            [
                "-NoProfile",
                "-File",
                str(WORD_COM_HELPER),
                "-InputPath",
                str(source),
                "-OutputPath",
                str(temporary_docx),
                "-Json",
            ],
        )
        require_success(result, "Word COM DOC conversion")
        route = "word-doc-to-docx+markitdown"
        converter_version = word_com_tool_version(result)
    else:
        result = run_command(
            args.libreoffice_command,
            [
                "--headless",
                "--convert-to",
                "docx",
                "--outdir",
                str(temporary_docx.parent),
                str(source),
            ],
        )
        require_success(result, "LibreOffice DOC conversion")
        emitted_docx = temporary_docx.parent / f"{source.stem}.docx"
        if not emitted_docx.is_file():
            raise RuntimeError("LibreOffice DOC conversion did not create a DOCX file")
        if emitted_docx != temporary_docx:
            shutil.move(emitted_docx, temporary_docx)
        route = "libreoffice-doc-to-docx+markitdown"
        converter_version = command_version(args.libreoffice_command)

    if not temporary_docx.is_file():
        raise RuntimeError("Legacy DOC conversion did not create a DOCX file")
    try:
        with zipfile.ZipFile(temporary_docx) as archive:
            bad_member = archive.testzip()
    except zipfile.BadZipFile as error:
        raise RuntimeError("Legacy DOC conversion created an invalid DOCX file") from error
    if bad_member is not None:
        raise RuntimeError(f"Legacy DOC conversion created a corrupt DOCX member: {bad_member}")
    return route, converter_version


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_manifest(path: Path) -> dict[str, dict[str, object]]:
    if not path.exists():
        return {}
    payload = json.loads(path.read_text(encoding="utf-8"))
    return {str(item["source_relative_path"]): item for item in payload.get("files", [])}


def timestamp() -> str:
    return datetime.now(timezone.utc).isoformat()


def relative_cache_path(source_relative_path: Path, extension: str) -> Path:
    filename = source_relative_path.name if extension == ".md" else f"{source_relative_path.name}.md"
    return Path("by-source") / source_relative_path.parent / filename


def update_cache(args: argparse.Namespace, system_name: str) -> dict[str, object]:
    source_root = Path(args.source_root).resolve()
    cache_root = Path(args.cache_root).resolve()
    manifest_path = cache_root / "manifest.json"
    log_path = cache_root / "conversion-log.jsonl"
    if not source_root.is_dir():
        raise ValueError(f"Source directory does not exist: {source_root}")

    cache_root.mkdir(parents=True, exist_ok=True)
    existing_by_source = read_manifest(manifest_path)
    files: list[dict[str, object]] = []
    errors: list[str] = []
    converted = reused = requires_review = 0
    source_files = sorted(
        (
            path
            for path in source_root.rglob("*")
            if path.is_file()
            and path.suffix.lower() in SUPPORTED_EXTENSIONS
            and not path.is_relative_to(cache_root)
        ),
        key=lambda path: str(path),
    )

    with tempfile.TemporaryDirectory(prefix="course-foundation-cache-") as temporary_root:
        for source_path in source_files:
            source_relative_path = source_path.relative_to(source_root)
            source_key = source_relative_path.as_posix()
            extension = source_path.suffix.lower()
            cache_relative_path = relative_cache_path(source_relative_path, extension)
            cache_path = cache_root / cache_relative_path
            sha256 = sha256_file(source_path)
            existing = existing_by_source.get(source_key)
            base = {
                "source_relative_path": source_key,
                "extension": extension,
                "sha256": sha256,
                "source_last_write_time_utc": datetime.fromtimestamp(
                    source_path.stat().st_mtime, timezone.utc
                ).isoformat(),
                "cache_path": cache_relative_path.as_posix(),
            }
            if (
                not args.force
                and existing is not None
                and existing.get("sha256") == sha256
                and cache_path.exists()
            ):
                reused += 1
                files.append(
                    base
                    | {
                        "route": existing["route"],
                        "tool": existing["tool"],
                        "tool_version": existing["tool_version"],
                        "status": "reused",
                        "converted_at_utc": existing["converted_at_utc"],
                    }
                )
                continue

            if extension == ".md":
                cache_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source_path, cache_path)
                converted += 1
                files.append(
                    base
                    | {
                        "route": "direct-markdown",
                        "tool": "existing-markdown",
                        "tool_version": "source",
                        "status": "converted",
                        "converted_at_utc": timestamp(),
                    }
                )
                continue

            try:
                cache_path.parent.mkdir(parents=True, exist_ok=True)
                if extension in {".docx", ".pdf"}:
                    route = f"markitdown-{extension.removeprefix('.')}"
                    tool = "markitdown"
                    tool_version = convert_with_markitdown(source_path, cache_path, args)
                    status = "converted"
                    if extension == ".pdf":
                        status = "converted-review-recommended"
                        requires_review += 1
                elif extension in {".html", ".htm", ".epub"}:
                    route = "pandoc-html" if extension in {".html", ".htm"} else "pandoc-epub"
                    tool = "pandoc"
                    tool_version = convert_with_pandoc(source_path, cache_path, args)
                    status = "converted"
                elif extension == ".doc":
                    temporary_dir = Path(temporary_root) / uuid.uuid4().hex
                    temporary_docx = temporary_dir / f"{source_path.stem}.docx"
                    route, converter_version = convert_legacy_doc(
                        source_path, temporary_docx, args, system_name
                    )
                    markitdown_version = convert_with_markitdown(temporary_docx, cache_path, args)
                    tool = f"{legacy_doc_strategy(system_name)}+markitdown"
                    tool_version = f"{converter_version}; {markitdown_version}"
                    status = "converted"
                else:
                    route = "artifact-tool-required"
                    tool = "document-artifact-tool"
                    tool_version = "not-run"
                    status = "needs-extraction"
                    requires_review += 1
                    cache_path.write_text(
                        f"---\nsource_file: {source_key}\nstatus: needs-extraction\n---\n",
                        encoding="utf-8",
                    )

                if not cache_path.is_file():
                    raise RuntimeError("Expected cache output was not created")
                converted += 1
                files.append(
                    base
                    | {
                        "route": route,
                        "tool": tool,
                        "tool_version": tool_version,
                        "status": status,
                        "converted_at_utc": timestamp(),
                    }
                )
            except (RuntimeError, ValueError) as error:
                message = str(error)
                errors.append(f"{source_key}: {message}")
                files.append(
                    base
                    | {
                        "route": "failed",
                        "tool": "unknown",
                        "tool_version": "unknown",
                        "status": "failed",
                        "cache_path": None,
                        "converted_at_utc": timestamp(),
                        "error": message,
                    }
                )

    manifest = {
        "schema_version": 1,
        "source_root": str(source_root),
        "cache_root": str(cache_root),
        "generated_at_utc": timestamp(),
        "source_count": len(source_files),
        "files": files,
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    summary: dict[str, object] = {
        "run_id": uuid.uuid4().hex,
        "generated_at_utc": timestamp(),
        "source_root": str(source_root),
        "cache_root": str(cache_root),
        "converted": converted,
        "reused": reused,
        "requires_review": requires_review,
        "failed": len(errors),
        "errors": errors,
    }
    with log_path.open("a", encoding="utf-8", newline="\n") as stream:
        stream.write(json.dumps(summary, ensure_ascii=False) + "\n")
    return summary


def parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True)
    parser.add_argument("--cache-root", required=True)
    parser.add_argument("--markitdown-command", default="markitdown")
    parser.add_argument("--pandoc-command", default="pandoc")
    parser.add_argument("--libreoffice-command", default="soffice")
    parser.add_argument("--powershell-command", default="powershell")
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--json", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None, system_name: str | None = None) -> int:
    args = parse_args(argv)
    try:
        summary = update_cache(args, system_name or __import__("platform").system())
    except ValueError as error:
        print(str(error), file=__import__("sys").stderr)
        return 1
    if args.json:
        print(json.dumps(summary, ensure_ascii=False))
    else:
        print(f"Cache root: {summary['cache_root']}")
        print(f"Converted: {summary['converted']}")
        print(f"Reused: {summary['reused']}")
        print(f"Requires review: {summary['requires_review']}")
        print(f"Failed: {summary['failed']}")
    return 1 if summary["failed"] else 2 if summary["requires_review"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
