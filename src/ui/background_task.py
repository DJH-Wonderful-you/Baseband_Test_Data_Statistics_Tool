from __future__ import annotations

import traceback
from typing import Any, Callable

from PySide6.QtCore import QObject, Signal, Slot


class BackgroundTaskWorker(QObject):
    log = Signal(str, str)
    result_ready = Signal(object)
    error = Signal(str)
    finished = Signal()

    def __init__(
        self,
        task: Callable[..., Any],
        *,
        args: tuple[Any, ...] = (),
        kwargs: dict[str, Any] | None = None,
    ) -> None:
        super().__init__()
        self._task = task
        self._args = args
        self._kwargs = kwargs.copy() if kwargs is not None else {}

    @Slot()
    def run(self) -> None:
        try:
            task_kwargs = self._kwargs.copy()
            task_kwargs["logger"] = self._emit_log
            result = self._task(*self._args, **task_kwargs)
            self.result_ready.emit(result)
        except Exception as exc:  # pragma: no cover
            self.error.emit(f"{exc}\n{traceback.format_exc()}")
        finally:
            self.finished.emit()

    def _emit_log(self, level: str, message: str) -> None:
        self.log.emit(level, message)
