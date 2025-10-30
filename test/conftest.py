from collections.abc import Generator
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
