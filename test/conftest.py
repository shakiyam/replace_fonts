from typing import Any, Generator

import pytest


@pytest.hookimpl(tryfirst=True, hookwrapper=True)
def pytest_runtest_makereport(
    item: pytest.Item, call: pytest.CallInfo[None]
) -> Generator[None, Any, None]:
    outcome = yield
    rep = outcome.get_result()
    setattr(item, f'rep_{rep.when}', rep)
