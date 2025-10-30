import re
import shutil
import tempfile
import zipfile
from collections.abc import Generator
from pathlib import Path

import pytest
from conftest import phase_report_key
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


@pytest.fixture
def workspace(
    request: pytest.FixtureRequest,
) -> Generator[tuple[Path, Path]]:
    """Create temporary workspace for test execution."""
    test_dir = Path(__file__).parent

    with tempfile.TemporaryDirectory() as tmpdir:
        work_dir = Path(tmpdir)

        for original in sorted((test_dir / "original").glob("sample*.pptx")):
            shutil.copy(original, work_dir / original.name)

        yield work_dir, test_dir / "expected"

        reports = request.node.stash.get(phase_report_key, {})
        if "call" in reports and reports["call"].failed:
            failure_dir = test_dir / "failures"
            failure_dir.mkdir(exist_ok=True)
            for f in work_dir.glob("*.log"):
                shutil.copy(f, failure_dir / f.name)


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
