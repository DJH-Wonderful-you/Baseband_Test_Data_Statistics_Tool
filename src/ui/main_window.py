from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication, QIcon, QShowEvent
from PySide6.QtWidgets import (
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QStackedWidget,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from src.core.logging_bus import LoggingBus
from src.ui.charge_tab import ChargeTab
from src.ui.endurance_tab import EnduranceTab
from src.ui.placeholders import AboutTab, UpdateLogTab, build_placeholder_tab


class MainWindow(QMainWindow):
    DEFAULT_WINDOW_SIZE = (1280, 800)
    BASE_MIN_WINDOW_SIZE = (1100, 700)

    def __init__(self) -> None:
        super().__init__()
        self._startup_geometry_adjusted = False
        self._build_ui()
        self._apply_initial_size()

    def _build_ui(self) -> None:
        self.setWindowTitle("基带测试数据统计工具 V1.2")
        self.setMinimumSize(*self.BASE_MIN_WINDOW_SIZE)
        
        # Set window icon
        icon_path = Path(__file__).parent.parent / "resources" / "app_icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        root_widget = QWidget()
        root_widget.setObjectName("appRoot")
        root_layout = QHBoxLayout(root_widget)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        self.page_stack = QStackedWidget()
        self.page_stack.setObjectName("pageStack")
        self.page_stack.addWidget(ChargeTab(LoggingBus()))
        self.page_stack.addWidget(EnduranceTab(LoggingBus()))
        self.page_stack.addWidget(build_placeholder_tab("待开发功能页，当前暂无功能。"))
        self.page_stack.addWidget(AboutTab())
        self.page_stack.addWidget(UpdateLogTab())

        self.sidebar = self._build_sidebar()
        root_layout.addWidget(self.sidebar)

        content_container = QFrame()
        content_container.setObjectName("contentContainer")
        content_layout = QVBoxLayout(content_container)
        content_layout.setContentsMargins(12, 12, 12, 12)
        content_layout.setSpacing(0)
        content_layout.addWidget(self.page_stack)

        root_layout.addWidget(content_container, 1)
        self.setCentralWidget(root_widget)

        self.nav_group.idClicked.connect(self.page_stack.setCurrentIndex)

    def _build_sidebar(self) -> QWidget:
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(210)
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(14, 16, 14, 16)
        layout.setSpacing(8)

        app_title = QLabel("基带测试\n数据统计工具")
        app_title.setObjectName("sidebarTitle")
        app_title.setAlignment(Qt.AlignCenter)
        layout.addWidget(app_title)

        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        nav_items = [
            ("充电测试", 0),
            ("续航测试", 1),
            ("待开发功能", 2),
            ("关于", 3),
            ("更新日志", 4),
        ]
        for text, index in nav_items:
            button = QPushButton(text)
            button.setCheckable(True)
            button.setCursor(Qt.PointingHandCursor)
            button.setProperty("navButton", True)
            self.nav_group.addButton(button, index)
            layout.addWidget(button)
        first = self.nav_group.button(0)
        if first is not None:
            first.setChecked(True)
        layout.addStretch()
        return sidebar

    def _apply_initial_size(self) -> None:
        width, height = self._get_initial_window_size()
        min_width = min(self.BASE_MIN_WINDOW_SIZE[0], width)
        min_height = min(self.BASE_MIN_WINDOW_SIZE[1], height)
        self.setMinimumSize(min_width, min_height)
        self.resize(width, height)
        self._center_on_screen()

    def showEvent(self, event: QShowEvent) -> None:
        super().showEvent(event)
        if self._startup_geometry_adjusted:
            return
        self._startup_geometry_adjusted = True
        self._fit_and_center_on_screen()

    def _fit_and_center_on_screen(self) -> None:
        screen = self._get_target_screen()
        if screen is None:
            return

        geometry = screen.availableGeometry()
        frame_extra_width = max(0, self.frameGeometry().width() - self.width())
        frame_extra_height = max(0, self.frameGeometry().height() - self.height())
        max_client_width = max(1, geometry.width() - frame_extra_width)
        max_client_height = max(1, geometry.height() - frame_extra_height)

        current_width = self.width()
        current_height = self.height()
        if current_width > max_client_width or current_height > max_client_height:
            scale = min(max_client_width / current_width, max_client_height / current_height)
            width = max(1, int(current_width * scale))
            height = max(1, int(current_height * scale))
            self.setMinimumSize(min(self.minimumWidth(), width), min(self.minimumHeight(), height))
            self.resize(width, height)

        self._center_on_screen()

    def _center_on_screen(self) -> None:
        screen = self._get_target_screen()
        if screen is None:
            return

        geometry = screen.availableGeometry()
        frame = self.frameGeometry()
        x = geometry.x() + (geometry.width() - frame.width()) // 2
        y = geometry.y() + (geometry.height() - frame.height()) // 2
        max_x = geometry.x() + max(0, geometry.width() - frame.width())
        max_y = geometry.y() + max(0, geometry.height() - frame.height())
        self.move(min(max(geometry.x(), x), max_x), min(max(geometry.y(), y), max_y))

    def _get_target_screen(self):
        window_handle = self.windowHandle()
        if window_handle is not None and window_handle.screen() is not None:
            return window_handle.screen()
        return QGuiApplication.primaryScreen()

    def _get_initial_window_size(self) -> tuple[int, int]:
        default_width, default_height = self.DEFAULT_WINDOW_SIZE
        screen = QGuiApplication.primaryScreen()
        if screen is None:
            return default_width, default_height

        geometry = screen.availableGeometry()
        screen_width = geometry.width()
        screen_height = geometry.height()

        if screen_width >= default_width and screen_height >= default_height:
            return default_width, default_height

        scale = min(screen_width / default_width, screen_height / default_height)
        width = max(1, int(default_width * scale))
        height = max(1, int(default_height * scale))
        return width, height

    def closeEvent(self, event) -> None:  # noqa: N802
        for index in range(self.page_stack.count()):
            page = self.page_stack.widget(index)
            has_running_task = getattr(page, "has_running_task", None)
            if callable(has_running_task) and has_running_task():
                QMessageBox.warning(self, "提示", "当前仍有统计任务正在执行，请等待完成后再退出程序。")
                event.ignore()
                return
        super().closeEvent(event)
