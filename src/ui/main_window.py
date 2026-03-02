from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QStackedWidget,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from src.core.logging_bus import LoggingBus
from src.ui.charge_tab import ChargeTab
from src.ui.placeholders import AboutTab, UpdateLogTab, build_placeholder_tab


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.log_bus = LoggingBus()
        self._build_ui()
        self._apply_initial_size()

    def _build_ui(self) -> None:
        self.setWindowTitle("基带测试数据统计工具 V0.8")
        self.setMinimumSize(1100, 700)

        root_widget = QWidget()
        root_widget.setObjectName("appRoot")
        root_layout = QHBoxLayout(root_widget)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        self.page_stack = QStackedWidget()
        self.page_stack.setObjectName("pageStack")
        self.page_stack.addWidget(ChargeTab(self.log_bus))
        self.page_stack.addWidget(
            build_placeholder_tab("开发中，后续版本将实现续航测试功能。"),
        )
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
        screen = QGuiApplication.primaryScreen()
        if screen is None:
            self.resize(1280, 800)
            return
        geometry = screen.availableGeometry()
        width = int(geometry.width() * 0.78)
        height = int(geometry.height() * 0.78)
        self.resize(width, height)
