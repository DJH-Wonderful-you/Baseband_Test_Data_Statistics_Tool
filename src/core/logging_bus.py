from __future__ import annotations

from typing import Callable


LogListener = Callable[[str, str], None]


class LoggingBus:
    def __init__(self) -> None:
        self._listeners: list[LogListener] = []

    def subscribe(self, listener: LogListener) -> None:
        if listener not in self._listeners:
            self._listeners.append(listener)

    def unsubscribe(self, listener: LogListener) -> None:
        if listener in self._listeners:
            self._listeners.remove(listener)

    def emit(self, level: str, message: str) -> None:
        for listener in tuple(self._listeners):
            listener(level, message)
