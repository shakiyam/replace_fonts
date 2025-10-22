import re
import shutil
import tempfile
from pathlib import Path
from typing import Generator

import pytest

from replace_fonts import process_pptx_file


def normalize_log(log_content: str) -> str:
    """
    Normalize log content for comparison.

    - Remove timestamps (YYYY-MM-DD HH:MM:SS)
    - Normalize backup file names (remove numbers in parentheses)
    - Remove temporary directory paths
    """
    lines = log_content.split('\n')
    normalized = []

    for line in lines:
        line = re.sub(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ', '', line)
        line = re.sub(r' - backup \(\d+\)\.pptx', ' - backup.pptx', line)
        line = re.sub(r'/tmp/[^/]+/', '', line)
        normalized.append(line)

    return '\n'.join(normalized)


@pytest.fixture
def test_workspace(
    request: pytest.FixtureRequest,
) -> Generator[tuple[Path, Path], None, None]:
    """Create temporary workspace for test execution."""
    test_dir = Path(__file__).parent

    with tempfile.TemporaryDirectory() as tmpdir:
        work_dir = Path(tmpdir)

        for original in sorted((test_dir / 'original').glob('sample*.pptx')):
            shutil.copy(original, work_dir / original.name)

        yield work_dir, test_dir / 'expected'

        if request.node.rep_call.failed:
            failure_dir = test_dir / 'failures'
            failure_dir.mkdir(exist_ok=True)
            for f in work_dir.glob('*.log'):
                shutil.copy(f, failure_dir / f.name)


def test_sample_files_with_code_option(
    test_workspace: tuple[Path, Path]
) -> None:
    """Test all sample files with --code option and verify log output."""
    work_dir, expected_dir = test_workspace

    for original in sorted(work_dir.glob('sample*.pptx')):
        name = original.stem
        test_pptx_path = str(work_dir / f'{name}.pptx')
        log_path = work_dir / f'{name}.log'
        expected_log_path = expected_dir / f'{name}.log'

        process_pptx_file(test_pptx_path, preserve_code_fonts=True)

        with open(log_path) as actual_log_file:
            actual = normalize_log(actual_log_file.read())

        with open(expected_log_path) as expected_log_file:
            expected = normalize_log(expected_log_file.read())

        assert actual == expected, f'{name}.log does not match expected output'
