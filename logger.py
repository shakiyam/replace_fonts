from datetime import datetime
from typing import TextIO


class Logger:
    def __init__(self, log_file: TextIO) -> None:
        self._log_file = log_file

    def log(self, message: str, element_text: str | None = None) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if element_text is not None:
            message = f"[{element_text}] {message}"
        print(f"{timestamp} {message}", file=self._log_file)
        print(f"{timestamp} {message}")
