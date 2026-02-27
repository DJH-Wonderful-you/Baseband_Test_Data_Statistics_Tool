from __future__ import annotations

from typing import Any


class AppError(Exception):
    def __init__(self, code: str, message: str, detail: Any | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.detail = detail

    def __str__(self) -> str:
        if self.detail is None:
            return f"[{self.code}] {self.message}"
        return f"[{self.code}] {self.message} | detail={self.detail}"
