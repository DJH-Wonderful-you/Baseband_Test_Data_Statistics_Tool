from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QMessageBox, QWidget


def open_directory(parent: QWidget, raw_path: str, *, label: str = "输出目录") -> bool:
    path_text = raw_path.strip()
    if not path_text:
        QMessageBox.warning(parent, "提示", f"请先设置{label}。")
        return False

    target = Path(path_text).expanduser()
    try:
        target.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        QMessageBox.warning(parent, "提示", f"{label}无法创建或打开：{exc}")
        return False

    if not target.is_dir():
        QMessageBox.warning(parent, "提示", f"{label}不是有效文件夹：{target}")
        return False

    opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(target.resolve())))
    if not opened:
        QMessageBox.warning(parent, "提示", f"无法打开{label}：{target}")
    return opened
