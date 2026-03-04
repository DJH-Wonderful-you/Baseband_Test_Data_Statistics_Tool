from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)


class FileUploadWidget(QWidget):
    """统一的文件上传组件，合并拖拽区域和文件列表"""
    paths_changed = Signal(list)

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        supported_hint: str = "支持 .xlsx / .xls / .csv 格式",
        file_filter: str = "Data Files (*.xlsx *.xls *.csv);;All Files (*.*)",
    ) -> None:
        super().__init__(parent)
        self._paths: list[Path] = []
        self._supported_hint = supported_hint
        self._file_filter = file_filter
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(8)

        # 创建可拖拽的容器
        self.drop_container = QFrame()
        self.drop_container.setObjectName("dropContainer")
        self.drop_container.setAcceptDrops(True)
        self.drop_container.dragEnterEvent = self._drag_enter
        self.drop_container.dragLeaveEvent = self._drag_leave
        self.drop_container.dropEvent = self._drop

        container_layout = QVBoxLayout(self.drop_container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)

        # 堆叠切换：空状态提示 / 文件列表
        self.stack = QStackedWidget()
        self.stack.setObjectName("uploadStack")

        # 空状态页面（带图标和提示）
        empty_page = QWidget()
        empty_page.setObjectName("emptyPage")
        empty_layout = QVBoxLayout(empty_page)
        empty_layout.setContentsMargins(20, 24, 20, 24)
        empty_layout.setSpacing(8)
        empty_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        icon_label = QLabel("📁")
        icon_label.setObjectName("uploadIcon")
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title_label = QLabel("拖拽文件或文件夹到这里")
        title_label.setObjectName("uploadTitle")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        hint_label = QLabel(self._supported_hint)
        hint_label.setObjectName("uploadHint")
        hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        empty_layout.addWidget(icon_label)
        empty_layout.addWidget(title_label)
        empty_layout.addWidget(hint_label)

        # 文件列表页面
        list_page = QWidget()
        list_page.setObjectName("listPage")
        list_layout = QVBoxLayout(list_page)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(0)

        self.list_widget = QListWidget()
        self.list_widget.setObjectName("inputPathList")
        self.list_widget.setAlternatingRowColors(False)
        self.list_widget.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.list_widget.setAcceptDrops(False)  # 由父容器处理拖拽
        list_layout.addWidget(self.list_widget)

        self.stack.addWidget(empty_page)
        self.stack.addWidget(list_page)
        self.stack.setCurrentIndex(0)

        container_layout.addWidget(self.stack)
        root.addWidget(self.drop_container, stretch=1)

        # 按钮行
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        self.btn_add_files = QPushButton("添加文件")
        self.btn_add_folder = QPushButton("添加文件夹")
        self.btn_remove_selected = QPushButton("移除选中")
        self.btn_clear = QPushButton("清空列表")
        self.btn_remove_selected.setProperty("accent", "warn")
        self.btn_clear.setProperty("accent", "subtle")
        btn_row.addWidget(self.btn_add_files)
        btn_row.addWidget(self.btn_add_folder)
        btn_row.addWidget(self.btn_remove_selected)
        btn_row.addWidget(self.btn_clear)
        btn_row.addStretch()
        root.addLayout(btn_row)

        # 连接信号
        self.btn_add_files.clicked.connect(self._on_pick_files)
        self.btn_add_folder.clicked.connect(self._on_pick_folder)
        self.btn_remove_selected.clicked.connect(self._on_remove_selected)
        self.btn_clear.clicked.connect(self.clear)

    def _update_view(self) -> None:
        """根据文件数量切换显示状态"""
        if self._paths:
            self.stack.setCurrentIndex(1)
        else:
            self.stack.setCurrentIndex(0)

    def _drag_enter(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_container.setProperty("dragActive", True)
            self.drop_container.style().unpolish(self.drop_container)
            self.drop_container.style().polish(self.drop_container)
            self.drop_container.update()
        else:
            event.ignore()

    def _drag_leave(self, event: QDragEnterEvent) -> None:
        del event
        self.drop_container.setProperty("dragActive", False)
        self.drop_container.style().unpolish(self.drop_container)
        self.drop_container.style().polish(self.drop_container)
        self.drop_container.update()

    def _drop(self, event: QDropEvent) -> None:
        self.drop_container.setProperty("dragActive", False)
        self.drop_container.style().unpolish(self.drop_container)
        self.drop_container.style().polish(self.drop_container)
        self.drop_container.update()
        if not event.mimeData().hasUrls():
            event.ignore()
            return
        paths: list[Path] = []
        for url in event.mimeData().urls():
            if url.isLocalFile():
                paths.append(Path(url.toLocalFile()))
        if paths:
            self.add_paths(paths)
            event.acceptProposedAction()
        else:
            event.ignore()

    def _emit_changed(self) -> None:
        self.paths_changed.emit(self.get_paths())

    def _on_pick_files(self) -> None:
        selected, _ = QFileDialog.getOpenFileNames(
            self,
            "选择文件",
            str(Path.cwd()),
            self._file_filter,
        )
        if selected:
            self.add_paths([Path(item) for item in selected])

    def _on_pick_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹", str(Path.cwd()))
        if folder:
            self.add_paths([Path(folder)])

    def _on_remove_selected(self) -> None:
        selected = self.list_widget.selectedItems()
        if not selected:
            return
        selected_paths = {
            Path(item.data(Qt.ItemDataRole.UserRole)) for item in selected
        }
        self._paths = [path for path in self._paths if path not in selected_paths]
        for item in selected:
            self.list_widget.takeItem(self.list_widget.row(item))
        self._update_view()
        self._emit_changed()

    def clear(self) -> None:
        self._paths.clear()
        self.list_widget.clear()
        self._update_view()
        self._emit_changed()

    def add_paths(self, raw_paths: list[Path]) -> None:
        existing = {path.resolve() for path in self._paths}
        added = False
        for raw_path in raw_paths:
            path = raw_path.expanduser()
            if not path.exists():
                continue
            resolved = path.resolve()
            if resolved in existing:
                continue
            existing.add(resolved)
            self._paths.append(resolved)
            item = QListWidgetItem(resolved.name)
            item.setData(Qt.ItemDataRole.UserRole, str(resolved))
            item.setToolTip(str(resolved))
            self.list_widget.addItem(item)
            added = True
        if added:
            self._update_view()
            self._emit_changed()

    def get_paths(self) -> list[Path]:
        return self._paths.copy()

    def set_controls_enabled(self, enabled: bool) -> None:
        for widget in (
            self.btn_add_files,
            self.btn_add_folder,
            self.btn_remove_selected,
            self.btn_clear,
            self.drop_container,
            self.list_widget,
        ):
            widget.setEnabled(enabled)
