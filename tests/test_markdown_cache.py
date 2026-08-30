import json
import importlib.util
import os
from pathlib import Path
import platform
import shutil
import subprocess
import sys

from docx import Document
import pytest


def cache_script(plugin_root: Path) -> Path:
    return (
        plugin_root
        / "skills"
        / "course-foundation-builder"
        / "scripts"
        / "update_markdown_cache.py"
    )


@pytest.fixture
def cache_module(plugin_root: Path):
    path = cache_script(plugin_root)
    spec = importlib.util.spec_from_file_location("markdown_cache_script", path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def run_cache(script: Path, source: Path, cache: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [
            sys.executable,
            str(script),
            "--source-root",
            str(source),
            "--cache-root",
            str(cache),
            "--json",
        ],
        text=True,
        encoding="utf-8",
        capture_output=True,
    )


def powershell_command() -> str | None:
    return shutil.which("powershell") or shutil.which("pwsh")


def posix_shell() -> str | None:
    if platform.system() != "Windows":
        return shutil.which("sh") or shutil.which("bash")

    git = shutil.which("git")
    if git is None:
        return None
    git_root = Path(git).resolve().parents[1]
    for candidate in (git_root / "bin" / "bash.exe", git_root / "usr" / "bin" / "sh.exe"):
        if candidate.is_file():
            return str(candidate)
    return None


def posix_path(shell: str, path: Path) -> str:
    if platform.system() != "Windows":
        return str(path)
    git_root = Path(shell).resolve().parents[1]
    cygpath = git_root / "usr" / "bin" / "cygpath.exe"
    result = subprocess.run(
        [str(cygpath), "-u", str(path)],
        text=True,
        encoding="utf-8",
        capture_output=True,
        check=True,
    )
    return result.stdout.strip()


def posix_environment(shell: str) -> dict[str, str]:
    environment = os.environ.copy()
    if platform.system() == "Windows":
        git_root = Path(shell).resolve().parents[1]
        environment["PATH"] = os.pathsep.join(
            [
                str(Path(sys.executable).parent),
                str(git_root / "bin"),
                str(git_root / "usr" / "bin"),
            ]
        )
    return environment


@pytest.mark.skipif(platform.system() != "Windows", reason="Windows wrapper test")
def test_windows_wrapper_runs_markdown_end_to_end(tmp_path: Path, plugin_root: Path):
    powershell = powershell_command()
    assert powershell is not None, "PowerShell is required on Windows"
    scripts = plugin_root / "skills" / "course-foundation-builder" / "scripts"
    wrapper = scripts / "update_markdown_cache.ps1"
    source = tmp_path / "课程源"
    source.mkdir()
    (source / "lesson.md").write_text("# Wrapper fixture\n", encoding="utf-8")
    cache = tmp_path / "课程缓存"
    command = [
        powershell,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(wrapper),
        "-SourceRoot",
        str(source),
        "-CacheRoot",
        str(cache),
        "-MarkItDownPath",
        "unused-markitdown",
        "-PandocPath",
        "unused-pandoc",
        "-LibreOfficePath",
        "unused-libreoffice",
        "-Json",
    ]

    first = subprocess.run(command, text=True, encoding="utf-8", capture_output=True)
    second = subprocess.run(command, text=True, encoding="utf-8", capture_output=True)
    forced = subprocess.run([*command, "-Force"], text=True, encoding="utf-8", capture_output=True)

    assert first.returncode == 0, first.stderr
    assert second.returncode == 0, second.stderr
    assert forced.returncode == 0, forced.stderr
    assert json.loads(first.stdout)["converted"] == 1
    assert json.loads(second.stdout)["reused"] == 1
    assert json.loads(forced.stdout)["converted"] == 1
    assert (cache / "by-source" / "lesson.md").read_text(encoding="utf-8") == (
        "# Wrapper fixture\n"
    )


@pytest.mark.skipif(
    platform.system() != "Windows" or shutil.which("pwsh") is None,
    reason="PowerShell 7 wrapper test",
)
def test_windows_wrapper_propagates_review_required_exit_code(
    tmp_path: Path, plugin_root: Path
):
    pwsh = shutil.which("pwsh")
    assert pwsh is not None
    wrapper = (
        plugin_root
        / "skills"
        / "course-foundation-builder"
        / "scripts"
        / "update_markdown_cache.ps1"
    )
    source = tmp_path / "review-source"
    source.mkdir()
    (source / "slides.pptx").write_bytes(b"review-required fixture")
    cache = tmp_path / "review-cache"
    environment = os.environ.copy()
    environment.update(
        {
            "CACHE_WRAPPER": str(wrapper),
            "CACHE_SOURCE_ROOT": str(source),
            "CACHE_ROOT": str(cache),
        }
    )
    driver = tmp_path / "invoke-wrapper.ps1"
    driver.write_text(
        '$ErrorActionPreference = "Stop"\n'
        "$PSNativeCommandUseErrorActionPreference = $true\n"
        "& $env:CACHE_WRAPPER -SourceRoot $env:CACHE_SOURCE_ROOT "
        "-CacheRoot $env:CACHE_ROOT -Json\n"
        "exit $LASTEXITCODE\n",
        encoding="utf-8",
    )

    result = subprocess.run(
        [
            pwsh,
            "-NoProfile",
            "-File",
            str(driver),
        ],
        text=True,
        encoding="utf-8",
        capture_output=True,
        env=environment,
    )

    assert result.returncode == 2, result.stderr
    assert json.loads(result.stdout)["requires_review"] == 1
    manifest = json.loads((cache / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["files"][0]["status"] == "needs-extraction"
    assert (cache / manifest["files"][0]["cache_path"]).is_file()


@pytest.mark.skipif(posix_shell() is None, reason="No usable POSIX shell is installed")
def test_posix_wrapper_runs_markdown_end_to_end(tmp_path: Path, plugin_root: Path):
    shell = posix_shell()
    assert shell is not None
    wrapper = (
        plugin_root
        / "skills"
        / "course-foundation-builder"
        / "scripts"
        / "update_markdown_cache.sh"
    )
    source = tmp_path / "posix-source"
    source.mkdir()
    (source / "lesson.md").write_text("# POSIX wrapper fixture\n", encoding="utf-8")
    cache = tmp_path / "posix-cache"
    command = [
        shell,
        posix_path(shell, wrapper),
        "--source-root",
        posix_path(shell, source),
        "--cache-root",
        posix_path(shell, cache),
        "--json",
    ]
    environment = posix_environment(shell)

    first = subprocess.run(
        command, text=True, encoding="utf-8", capture_output=True, env=environment
    )
    second = subprocess.run(
        command, text=True, encoding="utf-8", capture_output=True, env=environment
    )

    assert first.returncode == 0, first.stderr
    assert second.returncode == 0, second.stderr
    assert json.loads(first.stdout)["converted"] == 1
    assert json.loads(second.stdout)["reused"] == 1
    assert (cache / "by-source" / "lesson.md").read_text(encoding="utf-8") == (
        "# POSIX wrapper fixture\n"
    )


def test_legacy_doc_strategy_is_platform_specific(cache_module):
    assert cache_module.legacy_doc_strategy("Windows") == "word-com"
    assert cache_module.legacy_doc_strategy("Linux") == "libreoffice"
    assert cache_module.legacy_doc_strategy("Darwin") == "libreoffice"


def test_unknown_platform_rejects_legacy_doc(cache_module):
    with pytest.raises(ValueError, match="Unsupported platform"):
        cache_module.legacy_doc_strategy("Plan9")


def test_windows_legacy_doc_reports_missing_word_com_helper(
    cache_module, monkeypatch, tmp_path: Path
):
    monkeypatch.setattr(cache_module, "WORD_COM_HELPER", tmp_path / "convert_legacy_doc.ps1")

    with pytest.raises(RuntimeError, match=r"convert_legacy_doc\.ps1"):
        cache_module.convert_legacy_doc(
            tmp_path / "legacy.doc",
            tmp_path / "legacy.docx",
            cache_module.parse_args(
                ["--source-root", str(tmp_path), "--cache-root", str(tmp_path / "cache")]
            ),
            "Windows",
        )


def test_windows_legacy_doc_creates_parent_before_invoking_helper(
    cache_module, monkeypatch, tmp_path: Path
):
    source = tmp_path / "legacy.doc"
    temporary_docx = tmp_path / "temporary" / "legacy.docx"
    helper = tmp_path / "convert_legacy_doc.ps1"
    helper.write_text("# test helper", encoding="utf-8")
    monkeypatch.setattr(cache_module, "WORD_COM_HELPER", helper)

    def fake_run_command(command: str, arguments: list[str]) -> subprocess.CompletedProcess[str]:
        assert temporary_docx.parent.is_dir(), "helper ran before its output directory was created"
        document = Document()
        document.add_paragraph("Word helper output")
        document.save(temporary_docx)
        return subprocess.CompletedProcess(
            [command, *arguments],
            0,
            stdout='{"tool": "Microsoft Word", "version": "16.0"}',
            stderr="",
        )

    monkeypatch.setattr(cache_module, "run_command", fake_run_command)

    route, _ = cache_module.convert_legacy_doc(
        source,
        temporary_docx,
        cache_module.parse_args(
            ["--source-root", str(tmp_path), "--cache-root", str(tmp_path / "cache")]
        ),
        "Windows",
    )

    assert route == "word-doc-to-docx+markitdown"


def test_windows_legacy_doc_uses_helper_json_for_converter_version(
    cache_module, monkeypatch, tmp_path: Path
):
    source = tmp_path / "legacy.doc"
    temporary_docx = tmp_path / "temporary" / "legacy.docx"
    temporary_docx.parent.mkdir()
    helper = tmp_path / "convert_legacy_doc.ps1"
    helper.write_text("# test helper", encoding="utf-8")
    monkeypatch.setattr(cache_module, "WORD_COM_HELPER", helper)
    calls: list[list[str]] = []

    def fake_run_command(command: str, arguments: list[str]) -> subprocess.CompletedProcess[str]:
        assert command == "configured-powershell"
        calls.append(arguments)
        if arguments != [
            "-NoProfile",
            "-File",
            str(helper),
            "-InputPath",
            str(source),
            "-OutputPath",
            str(temporary_docx),
            "-Json",
        ]:
            raise AssertionError(f"unexpected command arguments: {arguments}")
        document = Document()
        document.add_paragraph("Word helper output")
        document.save(temporary_docx)
        return subprocess.CompletedProcess(
            [command, *arguments],
            0,
            stdout='{"tool": "Microsoft Word", "version": "16.0"}',
            stderr="",
        )

    monkeypatch.setattr(cache_module, "run_command", fake_run_command)

    _, tool_version = cache_module.convert_legacy_doc(
        source,
        temporary_docx,
        cache_module.parse_args(
            [
                "--source-root",
                str(tmp_path),
                "--cache-root",
                str(tmp_path / "cache"),
                "--powershell-command",
                "configured-powershell",
            ]
        ),
        "Windows",
    )

    assert calls == [
        [
            "-NoProfile",
            "-File",
            str(helper),
            "-InputPath",
            str(source),
            "-OutputPath",
            str(temporary_docx),
            "-Json",
        ]
    ]
    assert tool_version == "Microsoft Word 16.0"


@pytest.mark.skipif(platform.system() != "Windows", reason="Windows helper test")
def test_windows_doc_helper_rejects_missing_input_without_creating_output(
    tmp_path: Path, plugin_root: Path
):
    powershell = powershell_command()
    assert powershell is not None, "PowerShell is required on Windows"
    helper = (
        plugin_root
        / "skills"
        / "course-foundation-builder"
        / "scripts"
        / "convert_legacy_doc.ps1"
    )
    output = tmp_path / "legacy.docx"

    result = subprocess.run(
        [
            powershell,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(helper),
            "-InputPath",
            str(tmp_path / "missing.doc"),
            "-OutputPath",
            str(output),
            "-Json",
        ],
        text=True,
        encoding="utf-8",
        capture_output=True,
    )

    assert result.returncode != 0
    assert "Legacy DOC input does not exist" in result.stderr
    assert not output.exists()


def test_unchanged_markdown_is_reused_and_each_run_is_logged(
    tmp_path: Path, plugin_root: Path, fixtures: Path
):
    source = tmp_path / "source"
    source.mkdir()
    shutil.copy2(fixtures / "source-course.md", source / "course.md")
    cache = tmp_path / "cache"
    script = cache_script(plugin_root)

    first = run_cache(script, source, cache)
    second = run_cache(script, source, cache)

    assert first.returncode == 0, first.stderr
    assert second.returncode == 0, second.stderr
    assert json.loads(first.stdout)["converted"] == 1
    assert json.loads(second.stdout)["reused"] == 1

    manifest = json.loads((cache / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["files"][0]["route"] == "direct-markdown"
    assert manifest["files"][0]["tool"] == "existing-markdown"
    assert (cache / manifest["files"][0]["cache_path"]).exists()

    log_lines = (cache / "conversion-log.jsonl").read_text(encoding="utf-8").splitlines()
    assert len(log_lines) == 2
    assert json.loads(log_lines[1])["reused"] == 1


@pytest.mark.skipif(shutil.which("markitdown") is None, reason="MarkItDown is not installed")
def test_docx_is_converted_with_markitdown(tmp_path: Path, plugin_root: Path):
    source = tmp_path / "source"
    source.mkdir()
    document = Document()
    document.add_heading("Reusable course source", level=1)
    document.add_paragraph("A deterministic conversion fixture.")
    document.save(source / "course.docx")
    cache = tmp_path / "cache"
    script = cache_script(plugin_root)

    result = run_cache(script, source, cache)

    assert result.returncode == 0, result.stderr
    manifest = json.loads((cache / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["files"][0]["route"] == "markitdown-docx"
    cached = cache / manifest["files"][0]["cache_path"]
    assert "Reusable course source" in cached.read_text(encoding="utf-8")


@pytest.mark.skipif(shutil.which("pandoc") is None, reason="Pandoc is not installed")
def test_html_is_converted_with_pandoc(tmp_path: Path, plugin_root: Path):
    source = tmp_path / "source"
    source.mkdir()
    (source / "course.html").write_text(
        "<h1>Course evidence</h1><p>Reusable HTML fixture.</p>", encoding="utf-8"
    )
    cache = tmp_path / "cache"
    script = cache_script(plugin_root)

    result = run_cache(script, source, cache)

    assert result.returncode == 0, result.stderr
    manifest = json.loads((cache / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["files"][0]["route"] == "pandoc-html"
    cached = cache / manifest["files"][0]["cache_path"]
    assert "# Course evidence" in cached.read_text(encoding="utf-8")
