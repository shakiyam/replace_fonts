import re
import tempfile
import zipfile
from pathlib import Path

import pytest
from pptx.exc import PackageNotFoundError

from replace_fonts import main, process_pptx_file


def normalize_log(log_content: str) -> str:
    """
    Normalize log content for comparison.

    - Remove timestamps (YYYY-MM-DD HH:MM:SS)
    - Normalize backup file names (remove numbers in parentheses)
    - Remove temporary directory paths
    """
    lines = log_content.split("\n")
    normalized = []

    for line in lines:
        line = re.sub(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ", "", line)
        line = re.sub(r" - backup \(\d+\)\.pptx", " - backup.pptx", line)
        line = re.sub(r"/tmp/[^/]+/", "", line)
        normalized.append(line)

    return "\n".join(normalized)


@pytest.mark.parametrize(("preserve_code_fonts", "log_suffix"), [
    (True, ""),
    (False, "_nocode"),
])
def test_sample_pptx(
    workspace: tuple[Path, Path],
    preserve_code_fonts: bool,
    log_suffix: str,
) -> None:
    """Test all sample files and verify log output."""
    work_dir, expected_dir = workspace

    for original in sorted(work_dir.glob("sample*.pptx")):
        name = original.stem
        test_pptx_path = work_dir / f"{name}.pptx"
        log_path = work_dir / f"{name}.log"
        expected_log_path = expected_dir / f"{name}{log_suffix}.log"

        process_pptx_file(test_pptx_path, preserve_code_fonts)

        with open(log_path) as actual_log_file:
            actual = normalize_log(actual_log_file.read())

        with open(expected_log_path) as expected_log_file:
            expected = normalize_log(expected_log_file.read())

        assert actual == expected, (
            f"{name}{log_suffix}.log does not match expected output"
        )


def test_nonexistent_pptx() -> None:
    """Test that processing a non-existent PPTX file raises appropriate error."""
    with tempfile.TemporaryDirectory() as tmpdir:
        nonexistent_pptx = Path(tmpdir) / "nonexistent.pptx"

        with pytest.raises(FileNotFoundError):
            process_pptx_file(nonexistent_pptx, preserve_code_fonts=True)


def test_invalid_pptx() -> None:
    """Test that processing an invalid PPTX file raises appropriate error."""
    with tempfile.TemporaryDirectory() as tmpdir:
        invalid_pptx = Path(tmpdir) / "invalid.pptx"
        invalid_pptx.write_text("This is not a valid PPTX file")

        with pytest.raises((PackageNotFoundError, zipfile.BadZipFile)):
            process_pptx_file(invalid_pptx, preserve_code_fonts=True)


def test_multiple_pptx_with_error(
    workspace: tuple[Path, Path], monkeypatch: pytest.MonkeyPatch
) -> None:
    """Test that processing continues when one file fails among multiple files."""
    work_dir, _ = workspace

    valid_pptx = work_dir / "sample1.pptx"
    nonexistent_pptx = work_dir / "nonexistent.pptx"
    another_valid_pptx = work_dir / "sample2.pptx"

    args = [
        "replace_fonts.py",
        str(valid_pptx),
        str(nonexistent_pptx),
        str(another_valid_pptx),
    ]
    monkeypatch.setattr("sys.argv", args)

    exit_code = main()

    assert exit_code != 0, "Should return non-zero exit code when errors occur"
    assert (work_dir / "sample1.log").exists(), "First file should be processed"
    assert (work_dir / "sample2.log").exists(), "Third file should be processed"


def test_no_files_specified(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that the program handles no files gracefully."""
    args = ["replace_fonts.py"]
    monkeypatch.setattr("sys.argv", args)

    exit_code = main()

    assert exit_code == 0, "Should return zero when no files specified"


def test_dry_run_does_not_modify_file(workspace: tuple[Path, Path]) -> None:
    """Test that dry run does not modify the input file."""
    work_dir, _ = workspace
    pptx_path = work_dir / "sample1.pptx"
    original_content = pptx_path.read_bytes()

    process_pptx_file(pptx_path, preserve_code_fonts=True, dry_run=True)

    assert pptx_path.read_bytes() == original_content


def test_dry_run_does_not_create_backup(workspace: tuple[Path, Path]) -> None:
    """Test that dry run does not create a backup file."""
    work_dir, _ = workspace
    pptx_path = work_dir / "sample1.pptx"

    process_pptx_file(pptx_path, preserve_code_fonts=True, dry_run=True)

    backups = list(work_dir.glob("*backup*"))
    assert backups == []


def test_dry_run_creates_log(workspace: tuple[Path, Path]) -> None:
    """Test that dry run creates a .log file with (dry run) marker."""
    work_dir, _ = workspace
    pptx_path = work_dir / "sample1.pptx"

    process_pptx_file(pptx_path, preserve_code_fonts=True, dry_run=True)

    log_path = work_dir / "sample1.log"
    assert log_path.exists()
    log_content = log_path.read_text()
    assert "was opened. (dry run)" in log_content
    assert "was backed up" not in log_content
    assert "was saved." not in log_content


@pytest.mark.parametrize(("preserve_code_fonts", "log_suffix"), [
    (True, ""),
    (False, "_nocode"),
])
def test_dry_run_log_matches_normal_run(
    workspace: tuple[Path, Path],
    preserve_code_fonts: bool,
    log_suffix: str,
) -> None:
    """Test that dry run replacement log lines match normal run."""
    work_dir, expected_dir = workspace

    for original in sorted(work_dir.glob("sample*.pptx")):
        name = original.stem
        pptx_path = work_dir / f"{name}.pptx"
        log_path = work_dir / f"{name}.log"
        expected_log_path = expected_dir / f"{name}{log_suffix}.log"

        process_pptx_file(pptx_path, preserve_code_fonts, dry_run=True)

        with open(log_path) as actual_log_file:
            actual = normalize_log(actual_log_file.read())
        with open(expected_log_path) as expected_log_file:
            expected = normalize_log(expected_log_file.read())

        actual_lines = [
            line for line in actual.splitlines()
            if "was backed up" not in line
            and "was saved." not in line
            and "was opened." not in line
        ]
        expected_lines = [
            line for line in expected.splitlines()
            if "was backed up" not in line
            and "was saved." not in line
            and "was opened." not in line
        ]

        assert actual_lines == expected_lines, (
            f"{name}{log_suffix}.log dry run replacement lines do not match"
        )


def test_dry_run_then_normal_run(workspace: tuple[Path, Path]) -> None:
    """Test that normal run works correctly after dry run on the same file."""
    work_dir, expected_dir = workspace
    pptx_path = work_dir / "sample1.pptx"
    log_path = work_dir / "sample1.log"

    process_pptx_file(pptx_path, preserve_code_fonts=True, dry_run=True)
    log_path.unlink()

    process_pptx_file(pptx_path, preserve_code_fonts=True)

    with open(log_path) as actual_log_file:
        actual = normalize_log(actual_log_file.read())
    with open(expected_dir / "sample1.log") as expected_log_file:
        expected = normalize_log(expected_log_file.read())

    assert actual == expected


def test_dry_run_cli(
    workspace: tuple[Path, Path], monkeypatch: pytest.MonkeyPatch
) -> None:
    """Test that --dry-run CLI option works."""
    work_dir, _ = workspace
    pptx_path = work_dir / "sample1.pptx"
    original_content = pptx_path.read_bytes()

    args = ["replace_fonts.py", "--dry-run", str(pptx_path)]
    monkeypatch.setattr("sys.argv", args)

    exit_code = main()

    assert exit_code == 0
    assert pptx_path.read_bytes() == original_content
    backups = list(work_dir.glob("*backup*"))
    assert backups == []
