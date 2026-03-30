import shutil
import tempfile
from collections.abc import Generator
from pathlib import Path
from typing import Any

import pytest

phase_report_key = pytest.StashKey[dict[str, pytest.TestReport]]()


@pytest.hookimpl(tryfirst=True, hookwrapper=True)
def pytest_runtest_makereport(
    item: pytest.Item, call: pytest.CallInfo[None]
) -> Generator[None, Any]:
    outcome = yield
    rep = outcome.get_result()
    item.stash.setdefault(phase_report_key, {})[rep.when] = rep


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
