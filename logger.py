from collections.abc import Callable
from datetime import datetime
from typing import TextIO


def log(log_file: TextIO, message: str, element_text: str | None = None) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if element_text is not None:
        message = f"[{element_text}] {message}"
    print(f"{timestamp} {message}", file=log_file)
    print(f"{timestamp} {message}")


LogFn = Callable[[TextIO, str], None]
