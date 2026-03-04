from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QRectF, QSettings, QSize, Qt
from PySide6.QtGui import QColor, QDragEnterEvent, QDropEvent, QPainter, QPen
from PySide6.QtWidgets import (
    QAbstractButton,
    QAbstractSpinBox,
    QApplication,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.core.charge_merge_service import process_charge_merge
from src.core.charge_statistics_service import process_charge_statistics
from src.core.file_collect import collect_files, collect_merge_groups
from src.core.logging_bus import LoggingBus
from src.core.models import BatchResult
from src.ui.components.file_upload import FileUploadWidget


@dataclass(slots=True)
class BatchPacingSettings:
    enabled: bool = False
    chunk_size: int = 4
    wait_seconds: int = 3


class SlideSwitch(QAbstractButton):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setCheckable(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedSize(46, 26)

    def sizeHint(self) -> QSize:  # noqa: D401
        return QSize(46, 26)

    def paintEvent(self, event) -> None:  # noqa: N802
        _ = event
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        base_rect = QRectF(self.rect()).adjusted(1, 1, -1, -1)
        radius = base_rect.height() / 2

        if self.isEnabled():
            on_track = QColor("#4b9fe3")
            off_track = QColor("#c8d6e4")
            border_color = QColor("#8fb2cf")
            knob_border = QColor("#9eb6cb")
        else:
            on_track = QColor("#b9cee2")
            off_track = QColor("#d8e1ea")
            border_color = QColor("#c2ced8")
            knob_border = QColor("#c2ced8")

        painter.setPen(QPen(border_color, 1))
        painter.setBrush(on_track if self.isChecked() else off_track)
        painter.drawRoundedRect(base_rect, radius, radius)

        margin = 3
        knob_size = base_rect.height() - margin * 2
        knob_x = (
            base_rect.right() - margin - knob_size
            if self.isChecked()
            else base_rect.left() + margin
        )
        knob_rect = QRectF(knob_x, base_rect.top() + margin, knob_size, knob_size)

        painter.setPen(QPen(knob_border, 1))
        painter.setBrush(QColor("#ffffff"))
        painter.drawEllipse(knob_rect)


class BatchPacingDialog(QDialog):
    def __init__(self, settings: BatchPacingSettings, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("batchPacingDialog")
        self.setWindowTitle("分批处理策略")
        self.setModal(True)
        self.resize(500, 330)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(12)

        header_frame = QFrame()
        header_frame.setObjectName("batchPacingHeader")
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(4)

        title_label = QLabel("分批处理策略")
        title_label.setObjectName("batchPacingTitle")
        subtitle_label = QLabel(
            "当单次处理文件较多时，启用分批等待可减少部分杀毒软件拦截导致的连续处理失败。"
        )
        subtitle_label.setObjectName("batchPacingSubtitle")
        subtitle_label.setWordWrap(True)
        header_layout.addWidget(title_label)
        header_layout.addWidget(subtitle_label)
        layout.addWidget(header_frame)

        options_frame = QFrame()
        options_frame.setObjectName("batchPacingOptions")
        options_layout = QVBoxLayout(options_frame)
        options_layout.setContentsMargins(14, 12, 14, 12)
        options_layout.setSpacing(10)

        self.toggle_card = QFrame()
        self.toggle_card.setObjectName("batchPacingToggleCard")
        toggle_layout = QHBoxLayout(self.toggle_card)
        toggle_layout.setContentsMargins(12, 10, 12, 10)
        toggle_layout.setSpacing(10)

        toggle_text_layout = QVBoxLayout()
        toggle_text_layout.setContentsMargins(0, 0, 0, 0)
        toggle_text_layout.setSpacing(2)
        toggle_title = QLabel("策略开关")
        toggle_title.setObjectName("batchPacingToggleTitle")
        toggle_hint = QLabel("开启后将按“单批上限 + 批间等待”分批执行任务。")
        toggle_hint.setObjectName("batchPacingToggleHint")
        toggle_hint.setWordWrap(True)
        toggle_text_layout.addWidget(toggle_title)
        toggle_text_layout.addWidget(toggle_hint)
        toggle_layout.addLayout(toggle_text_layout, stretch=1)

        toggle_control_layout = QVBoxLayout()
        toggle_control_layout.setContentsMargins(0, 0, 0, 0)
        toggle_control_layout.setSpacing(6)
        toggle_control_layout.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.status_badge = QLabel()
        self.status_badge.setObjectName("batchPacingStatusBadge")
        self.status_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.enable_switch = SlideSwitch()
        self.enable_switch.setObjectName("batchPacingSwitch")
        self.enable_switch.setChecked(settings.enabled)
        switch_row = QHBoxLayout()
        switch_row.setContentsMargins(0, 0, 0, 0)
        switch_row.setSpacing(8)
        switch_row.addWidget(self.enable_switch)
        switch_row.addWidget(self.status_badge)
        toggle_control_layout.addLayout(switch_row)
        toggle_layout.addLayout(toggle_control_layout)
        options_layout.addWidget(self.toggle_card)

        settings_row = QHBoxLayout()
        settings_row.setContentsMargins(0, 0, 0, 0)
        settings_row.setSpacing(10)

        chunk_label = QLabel("单批上限")
        chunk_label.setObjectName("batchPacingFieldLabel")
        settings_row.addWidget(chunk_label)

        self.chunk_size_spin = QSpinBox()
        self.chunk_size_spin.setRange(1, 200)
        self.chunk_size_spin.setValue(settings.chunk_size)
        self.chunk_size_spin.setFixedWidth(74)
        self.chunk_size_spin.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.chunk_size_spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        settings_row.addWidget(self.chunk_size_spin)
        self.chunk_size_unit_label = QLabel("个文件")
        self.chunk_size_unit_label.setObjectName("batchPacingUnit")
        settings_row.addWidget(self.chunk_size_unit_label)
        settings_row.addSpacing(12)

        wait_label = QLabel("批间等待")
        wait_label.setObjectName("batchPacingFieldLabel")
        settings_row.addWidget(wait_label)

        self.wait_seconds_spin = QSpinBox()
        self.wait_seconds_spin.setRange(1, 600)
        self.wait_seconds_spin.setValue(settings.wait_seconds)
        self.wait_seconds_spin.setFixedWidth(74)
        self.wait_seconds_spin.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.wait_seconds_spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        settings_row.addWidget(self.wait_seconds_spin)
        self.wait_seconds_unit_label = QLabel("秒")
        self.wait_seconds_unit_label.setObjectName("batchPacingUnit")
        settings_row.addWidget(self.wait_seconds_unit_label)
        settings_row.addStretch()
        options_layout.addLayout(settings_row)

        self.preview_label = QLabel()
        self.preview_label.setObjectName("batchPacingPreview")
        self.preview_label.setWordWrap(True)
        options_layout.addWidget(self.preview_label)
        layout.addWidget(options_frame)

        self.enable_switch.toggled.connect(self._sync_enabled_state)
        self.chunk_size_spin.valueChanged.connect(self._refresh_preview)
        self.wait_seconds_spin.valueChanged.connect(self._refresh_preview)
        self._sync_enabled_state(self.enable_switch.isChecked())

        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.setObjectName("batchPacingButtons")
        save_btn = button_box.button(QDialogButtonBox.StandardButton.Save)
        cancel_btn = button_box.button(QDialogButtonBox.StandardButton.Cancel)
        if save_btn is not None:
            save_btn.setText("保存")
            save_btn.setProperty("accent", "primary")
        if cancel_btn is not None:
            cancel_btn.setText("取消")
            cancel_btn.setProperty("accent", "subtle")
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        for button in (save_btn, cancel_btn):
            if button is None:
                continue
            style = button.style()
            style.unpolish(button)
            style.polish(button)
            button.update()

    def _sync_enabled_state(self, enabled: bool) -> None:
        self.chunk_size_spin.setEnabled(enabled)
        self.wait_seconds_spin.setEnabled(enabled)
        self.chunk_size_unit_label.setEnabled(enabled)
        self.wait_seconds_unit_label.setEnabled(enabled)
        self.enable_switch.setEnabled(True)
        state = "on" if enabled else "off"
        if self.toggle_card.property("state") != state:
            self.toggle_card.setProperty("state", state)
            toggle_style = self.toggle_card.style()
            toggle_style.unpolish(self.toggle_card)
            toggle_style.polish(self.toggle_card)
            self.toggle_card.update()
        if self.status_badge.property("state") != state:
            self.status_badge.setProperty("state", state)
            badge_style = self.status_badge.style()
            badge_style.unpolish(self.status_badge)
            badge_style.polish(self.status_badge)
            self.status_badge.update()
        self.status_badge.setText("已开启" if enabled else "已关闭")
        self._refresh_preview()

    def _refresh_preview(self) -> None:
        if not self.enable_switch.isChecked():
            state = "off"
            text = "当前状态：关闭。系统会连续处理所有文件，不做批间等待。"
        else:
            sample_total = 11
            chunk_size = max(1, self.chunk_size_spin.value())
            wait_seconds = self.wait_seconds_spin.value()
            full_batches = sample_total // chunk_size
            remainder = sample_total % chunk_size
            chunk_parts = [str(chunk_size)] * full_batches
            if remainder:
                chunk_parts.append(str(remainder))
            batch_count = len(chunk_parts)
            chunk_expr = " + ".join(chunk_parts)
            state = "on"
            text = (
                f"示例：上传 {sample_total} 个文件时，将分 {batch_count} 批处理（{chunk_expr}），"
                f"批间等待 {wait_seconds} 秒。"
            )
        if self.preview_label.property("state") != state:
            self.preview_label.setProperty("state", state)
            style = self.preview_label.style()
            style.unpolish(self.preview_label)
            style.polish(self.preview_label)
            self.preview_label.update()
        self.preview_label.setText(text)

    def current_settings(self) -> BatchPacingSettings:
        return BatchPacingSettings(
            enabled=self.enable_switch.isChecked(),
            chunk_size=self.chunk_size_spin.value(),
            wait_seconds=self.wait_seconds_spin.value(),
        )


class ProcessModeDialog(QDialog):
    MODE_SINGLE = "single"
    MODE_MERGE = "merge"

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.selected_mode: str | None = None
        self.setObjectName("modeSelectDialog")
        self.setWindowTitle("选择统计模式")
        self.setModal(True)
        self.resize(520, 360)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(12)

        header = QFrame()
        header.setObjectName("modeSelectHeader")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(4)

        title = QLabel("选择统计模式")
        title.setObjectName("modeSelectTitle")
        subtitle = QLabel("点击下方模式按钮后立即执行。两种模式对应原有两条处理流程。")
        subtitle.setObjectName("modeSelectSubtitle")
        subtitle.setWordWrap(True)
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        layout.addWidget(header)

        options = QFrame()
        options.setObjectName("modeSelectOptions")
        options_layout = QVBoxLayout(options)
        options_layout.setContentsMargins(14, 12, 14, 12)
        options_layout.setSpacing(10)

        single_card = self._build_mode_card(
            title="单文件模式",
            description="仅处理 Excel（.xlsx/.xls），用于统计“充电曲线”数据。",
            button_text="执行单文件模式",
            button_accent="success",
            card_kind="single",
            mode=self.MODE_SINGLE,
        )
        merge_card = self._build_mode_card(
            title="合并模式",
            description="按同名 Excel + .csv 配对处理，用于同时统计“充电曲线”和“充电温升”数据。",
            button_text="执行合并模式",
            button_accent="primary",
            card_kind="merge",
            mode=self.MODE_MERGE,
        )
        options_layout.addWidget(single_card)
        options_layout.addWidget(merge_card)
        layout.addWidget(options)

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 2, 0, 0)
        footer.addStretch()
        cancel_btn = QPushButton("取消")
        cancel_btn.setProperty("accent", "subtle")
        cancel_btn.clicked.connect(self.reject)
        footer.addWidget(cancel_btn)
        layout.addLayout(footer)

        style = cancel_btn.style()
        style.unpolish(cancel_btn)
        style.polish(cancel_btn)
        cancel_btn.update()

    def _build_mode_card(
        self,
        *,
        title: str,
        description: str,
        button_text: str,
        button_accent: str,
        card_kind: str,
        mode: str,
    ) -> QFrame:
        card = QFrame()
        card.setObjectName("modeOptionCard")
        card.setProperty("kind", card_kind)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(8)

        title_label = QLabel(title)
        title_label.setObjectName("modeOptionTitle")
        desc_label = QLabel(description)
        desc_label.setObjectName("modeOptionDesc")
        desc_label.setWordWrap(True)

        action_row = QHBoxLayout()
        action_row.setContentsMargins(0, 0, 0, 0)
        action_row.setSpacing(8)
        action_row.addStretch()
        action_btn = QPushButton(button_text)
        action_btn.setObjectName("modeOptionBtn")
        action_btn.setProperty("accent", button_accent)
        action_btn.clicked.connect(lambda _=False, selected=mode: self._select_mode(selected))
        action_row.addWidget(action_btn)

        card_layout.addWidget(title_label)
        card_layout.addWidget(desc_label)
        card_layout.addLayout(action_row)
        return card

    def _select_mode(self, mode: str) -> None:
        self.selected_mode = mode
        self.accept()


class SafeConfirmDialog(QDialog):
    def __init__(
        self,
        *,
        title: str,
        message: str,
        confirm_text: str = "仍继续执行",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("safeConfirmDialog")
        self.setWindowTitle("防误触提醒")
        self.setModal(True)
        self.resize(470, 280)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(12)

        header = QFrame()
        header.setObjectName("safeConfirmHeader")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(10)

        icon = QLabel("!")
        icon.setObjectName("safeConfirmIcon")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_layout.addWidget(icon)

        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(3)
        title_label = QLabel(title)
        title_label.setObjectName("safeConfirmTitle")
        subtitle_label = QLabel("请确认当前模式与上传文件类型匹配，避免误处理。")
        subtitle_label.setObjectName("safeConfirmSubtitle")
        subtitle_label.setWordWrap(True)
        text_layout.addWidget(title_label)
        text_layout.addWidget(subtitle_label)
        header_layout.addLayout(text_layout, stretch=1)
        layout.addWidget(header)

        body = QFrame()
        body.setObjectName("safeConfirmBody")
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(12, 10, 12, 10)
        body_layout.setSpacing(6)
        message_label = QLabel(message)
        message_label.setObjectName("safeConfirmMessage")
        message_label.setWordWrap(True)
        body_layout.addWidget(message_label)
        layout.addWidget(body)

        button_row = QHBoxLayout()
        button_row.setContentsMargins(0, 2, 0, 0)
        button_row.addStretch()
        cancel_btn = QPushButton("返回检查")
        cancel_btn.setProperty("accent", "subtle")
        confirm_btn = QPushButton(confirm_text)
        confirm_btn.setProperty("accent", "warn")
        cancel_btn.clicked.connect(self.reject)
        confirm_btn.clicked.connect(self.accept)
        button_row.addWidget(cancel_btn)
        button_row.addWidget(confirm_btn)
        layout.addLayout(button_row)

        for button in (cancel_btn, confirm_btn):
            style = button.style()
            style.unpolish(button)
            style.polish(button)
            button.update()

    @classmethod
    def ask(
        cls,
        parent: QWidget,
        *,
        title: str,
        message: str,
        confirm_text: str = "仍继续执行",
    ) -> bool:
        dialog = cls(
            title=title,
            message=message,
            confirm_text=confirm_text,
            parent=parent,
        )
        return dialog.exec() == QDialog.DialogCode.Accepted


class ChargeTab(QWidget):
    BATCH_PACING_ENABLED_KEY = "charge_tab/batch_pacing/enabled"
    BATCH_PACING_CHUNK_SIZE_KEY = "charge_tab/batch_pacing/chunk_size"
    BATCH_PACING_WAIT_SECONDS_KEY = "charge_tab/batch_pacing/wait_seconds"

    def __init__(self, log_bus: LoggingBus) -> None:
        super().__init__()
        self.log_bus = log_bus
        self.batch_pacing_settings = self._load_batch_pacing_settings()
        self._progress_total = 0
        self._progress_done = 0
        self.setAcceptDrops(True)
        self._build_ui()
        self._refresh_batch_pacing_hint()
        self.log_bus.subscribe(self._append_runtime_log)

    def _build_ui(self) -> None:
        self.setObjectName("chargeRoot")
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(20, 20, 20, 18)
        root_layout.setSpacing(12)

        page_title = QLabel("充电测试")
        page_title.setObjectName("pageTitle")
        page_subtitle = QLabel(
            "左侧配置文件上传、数据处理与输出设置；右侧查看运行日志。"
        )
        page_subtitle.setObjectName("pageSubtitle")
        page_header = QFrame()
        page_header.setObjectName("pageHero")
        page_header_layout = QVBoxLayout(page_header)
        page_header_layout.setContentsMargins(18, 14, 18, 14)
        page_header_layout.setSpacing(4)
        page_header_layout.addWidget(page_title)
        page_header_layout.addWidget(page_subtitle)
        root_layout.addWidget(page_header)

        horizontal_splitter = QSplitter()
        horizontal_splitter.setObjectName("chargeSplitter")
        horizontal_splitter.setOrientation(Qt.Orientation.Horizontal)
        horizontal_splitter.setChildrenCollapsible(False)

        left_panel = QFrame()
        left_panel.setObjectName("chargeLeftPanel")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(12)

        upload_group = QGroupBox("文件上传")
        upload_group.setObjectName("uploadGroup")
        upload_layout = QVBoxLayout(upload_group)
        upload_layout.setContentsMargins(14, 20, 14, 14)
        upload_layout.setSpacing(10)
        upload_hint = QLabel(
            "支持拖拽或选择文件/文件夹。统计数据处理 Excel（.xlsx/.xls）；合并后统计数据会自动匹配同名 Excel（.xlsx/.xls）+ .csv。"
        )
        upload_hint.setObjectName("groupHint")
        upload_hint.setWordWrap(True)
        upload_layout.addWidget(upload_hint)
        self.file_upload = FileUploadWidget()
        self.file_upload.paths_changed.connect(self._on_paths_changed)
        upload_layout.addWidget(self.file_upload)
        left_layout.addWidget(upload_group, stretch=1)

        action_group = QGroupBox("数据处理")
        action_group.setObjectName("actionGroup")
        action_layout = QVBoxLayout(action_group)
        action_layout.setContentsMargins(14, 20, 14, 14)
        action_layout.setSpacing(10)
        action_hint = QLabel("点击“统计数据”后可在弹窗中选择“单文件模式”或“合并模式”执行。")
        action_hint.setObjectName("groupHint")
        action_hint.setWordWrap(True)
        action_row = QHBoxLayout()
        action_row.setSpacing(10)
        self.statistics_btn = QPushButton("统计数据")
        self.statistics_btn.setProperty("accent", "primary")
        self.processing_progress = QProgressBar()
        self.processing_progress.setObjectName("processingProgressBar")
        self.processing_progress.setRange(0, 100)
        self.processing_progress.setValue(0)
        self.processing_progress.setFormat("0%")
        self.processing_progress.setTextVisible(True)
        self.processing_progress.setMinimumWidth(240)
        self.processing_progress.setProperty("state", "idle")
        action_row.addWidget(self.statistics_btn)
        action_row.addWidget(self.processing_progress, stretch=1)
        action_layout.addWidget(action_hint)
        action_layout.addLayout(action_row)
        left_layout.addWidget(action_group)

        output_group = QGroupBox("输出设置")
        output_group.setObjectName("outputGroup")
        output_layout = QVBoxLayout(output_group)
        output_layout.setContentsMargins(14, 20, 14, 14)
        output_layout.setSpacing(10)
        output_hint = QLabel(
            "默认输出到项目目录下 output 文件夹；如需降低批量处理失败风险，可在右侧“分批策略...”中设置分批等待。"
        )
        output_hint.setObjectName("groupHint")
        output_hint.setWordWrap(True)
        output_row = QHBoxLayout()
        output_row.setSpacing(10)
        output_label = QLabel("输出目录")
        output_label.setObjectName("fieldLabel")
        self.output_edit = QLineEdit(str((Path.cwd() / "output").resolve()))
        self.output_edit.setObjectName("pathEdit")
        self.browse_output_btn = QPushButton("浏览")
        self.batch_pacing_btn = QPushButton("分批策略...")
        self.batch_pacing_btn.setProperty("accent", "warn")
        output_row.addWidget(output_label)
        output_row.addWidget(self.output_edit, stretch=1)
        output_row.addWidget(self.browse_output_btn)
        output_row.addWidget(self.batch_pacing_btn)
        self.batch_pacing_hint = QLabel()
        self.batch_pacing_hint.setObjectName("groupHint")
        self.batch_pacing_hint.setWordWrap(True)
        output_layout.addWidget(output_hint)
        output_layout.addLayout(output_row)
        output_layout.addWidget(self.batch_pacing_hint)
        left_layout.addWidget(output_group)

        right_panel = QGroupBox("运行日志")
        right_panel.setObjectName("runtimeLogGroup")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(14, 20, 14, 14)
        right_layout.setSpacing(10)
        self.runtime_log_view = QTextEdit()
        self.runtime_log_view.setObjectName("runtimeLogView")
        self.runtime_log_view.setReadOnly(True)
        self.runtime_log_view.setPlaceholderText("处理过程、警告与错误会显示在这里。")
        right_layout.addWidget(self.runtime_log_view, stretch=1)
        right_foot = QHBoxLayout()
        right_foot.setContentsMargins(0, 6, 0, 0)
        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.setProperty("accent", "subtle")
        self.clear_log_btn.clicked.connect(self._on_clear_runtime_log)
        right_foot.addStretch()
        right_foot.addWidget(self.clear_log_btn)
        right_layout.addLayout(right_foot)

        horizontal_splitter.addWidget(left_panel)
        horizontal_splitter.addWidget(right_panel)
        horizontal_splitter.setSizes([760, 460])
        root_layout.addWidget(horizontal_splitter, stretch=1)

        self.browse_output_btn.clicked.connect(self._on_select_output)
        self.batch_pacing_btn.clicked.connect(self._open_batch_pacing_dialog)
        self.statistics_btn.clicked.connect(self._open_statistics_mode_dialog)

    def _log(self, level: str, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_bus.emit(level, f"[{timestamp}] {message}")

    def _append_runtime_log(self, level: str, message: str) -> None:
        level_color = {
            "INFO": "#3498db",
            "WARN": "#e67e22",
            "ERROR": "#e74c3c",
        }.get(level, "#3498db")
        level_bg = {
            "INFO": "#eaf4fc",
            "WARN": "#fef5e7",
            "ERROR": "#fdedec",
        }.get(level, "#eaf4fc")
        formatted_msg = self._format_log_message(message)
        self.runtime_log_view.append(
            f'<span style="background-color:{level_bg};color:{level_color};font-weight:700;'
            f'padding:2px 6px;border-radius:4px;">[{level}]</span> '
            f'{formatted_msg}'
        )

    def _format_log_message(self, message: str) -> str:
        import re
        result = message
        time_pattern = r'(\[\d{2}:\d{2}:\d{2}\])'
        result = re.sub(
            time_pattern,
            r'<span style="color:#2980b9;font-weight:600;">\1</span>',
            result
        )
        success_pattern = r'(\[成功\])'
        result = re.sub(
            success_pattern,
            r'<span style="color:#27ae60;font-weight:700;">\1</span>',
            result
        )
        fail_pattern = r'(\[失败\])'
        result = re.sub(
            fail_pattern,
            r'<span style="color:#e74c3c;font-weight:700;">\1</span>',
            result
        )
        arrow_pattern = r'(->)'
        result = re.sub(
            arrow_pattern,
            r'<span style="color:#9b59b6;font-weight:600;">→</span>',
            result
        )
        if '<span' not in result or result == message:
            result = f'<span style="color:#2c3e50;">{result}</span>'
        else:
            parts = re.split(r'(<span[^>]*>.*?</span>)', result)
            new_parts = []
            for part in parts:
                if part.startswith('<span'):
                    new_parts.append(part)
                elif part.strip():
                    new_parts.append(f'<span style="color:#2c3e50;">{part}</span>')
                else:
                    new_parts.append(part)
            result = ''.join(new_parts)
        return result

    def _on_clear_runtime_log(self) -> None:
        self.runtime_log_view.clear()

    def _on_paths_changed(self, paths: list[Path]) -> None:
        self._log("INFO", f"当前输入路径数量：{len(paths)}")

    def _on_select_output(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "选择输出目录", self.output_edit.text())
        if folder:
            self.output_edit.setText(folder)

    @staticmethod
    def _to_bool(value: object, default: bool) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            return value.strip().lower() in {"1", "true", "yes", "on"}
        if isinstance(value, (int, float)):
            return value != 0
        return default

    @staticmethod
    def _to_int(value: object, default: int) -> int:
        if isinstance(value, bool):
            return default
        if isinstance(value, (int, float)):
            return int(value)
        if isinstance(value, str):
            try:
                return int(value.strip())
            except ValueError:
                return default
        return default

    def _load_batch_pacing_settings(self) -> BatchPacingSettings:
        settings = QSettings()
        enabled = self._to_bool(
            settings.value(self.BATCH_PACING_ENABLED_KEY, False),
            default=False,
        )
        chunk_size = max(
            1,
            self._to_int(settings.value(self.BATCH_PACING_CHUNK_SIZE_KEY, 4), default=4),
        )
        wait_seconds = max(
            1,
            self._to_int(settings.value(self.BATCH_PACING_WAIT_SECONDS_KEY, 3), default=3),
        )
        return BatchPacingSettings(
            enabled=enabled,
            chunk_size=chunk_size,
            wait_seconds=wait_seconds,
        )

    def _save_batch_pacing_settings(self) -> None:
        settings = QSettings()
        settings.setValue(self.BATCH_PACING_ENABLED_KEY, self.batch_pacing_settings.enabled)
        settings.setValue(self.BATCH_PACING_CHUNK_SIZE_KEY, self.batch_pacing_settings.chunk_size)
        settings.setValue(self.BATCH_PACING_WAIT_SECONDS_KEY, self.batch_pacing_settings.wait_seconds)
        settings.sync()

    def _refresh_batch_pacing_hint(self) -> None:
        if self.batch_pacing_settings.enabled:
            self.batch_pacing_hint.setText(
                "说明：分批策略已启用。当前按每批 "
                f"{self.batch_pacing_settings.chunk_size} 个文件执行，批间等待 "
                f"{self.batch_pacing_settings.wait_seconds} 秒。"
            )
            return
        self.batch_pacing_hint.setText(
            "说明：分批策略默认关闭。开启后可按“每批文件数 + 等待秒数”分批执行。"
        )

    def _open_batch_pacing_dialog(self) -> None:
        dialog = BatchPacingDialog(self.batch_pacing_settings, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        updated = dialog.current_settings()
        if updated == self.batch_pacing_settings:
            return
        self.batch_pacing_settings = updated
        self._save_batch_pacing_settings()
        self._refresh_batch_pacing_hint()
        if updated.enabled:
            self._log(
                "INFO",
                f"分批策略已更新：每批 {updated.chunk_size} 个文件，批间等待 {updated.wait_seconds} 秒",
            )
            return
        self._log("INFO", "分批策略已关闭")

    def _effective_pacing(self) -> tuple[int | None, int]:
        if not self.batch_pacing_settings.enabled:
            return None, 0
        return self.batch_pacing_settings.chunk_size, self.batch_pacing_settings.wait_seconds

    def _current_inputs(self) -> list[Path]:
        return self.file_upload.get_paths()

    def _set_running(self, running: bool) -> None:
        for button in (
            self.browse_output_btn,
            self.batch_pacing_btn,
            self.statistics_btn,
            self.clear_log_btn,
        ):
            button.setDisabled(running)
        self.file_upload.set_controls_enabled(not running)

    def _validate_before_run(self) -> tuple[list[Path], Path] | None:
        inputs = self._current_inputs()
        if not inputs:
            QMessageBox.warning(self, "提示", "请先添加至少一个输入文件或文件夹。")
            return None
        output_text = self.output_edit.text().strip()
        if not output_text:
            QMessageBox.warning(self, "提示", "请先设置输出目录。")
            return None
        return inputs, Path(output_text)

    def _contains_csv_file(self, inputs: list[Path]) -> bool:
        files, _ = collect_files(inputs, {".csv"})
        return bool(files)

    def _count_excel_and_csv_files(self, inputs: list[Path]) -> tuple[int, int]:
        files, _ = collect_files(inputs, {".xlsx", ".xls", ".csv"})
        excel_count = sum(1 for file in files if file.suffix.lower() in {".xlsx", ".xls"})
        csv_count = sum(1 for file in files if file.suffix.lower() == ".csv")
        return excel_count, csv_count

    def _count_statistics_items(self, inputs: list[Path]) -> int:
        files, _ = collect_files(inputs, {".xlsx", ".xls"})
        return len(files)

    def _count_merge_items(self, inputs: list[Path]) -> int:
        groups, pair_errors, _ = collect_merge_groups(inputs)
        return len(groups) + len(pair_errors)

    def _set_progress_state(self, state: str) -> None:
        if self.processing_progress.property("state") == state:
            return
        self.processing_progress.setProperty("state", state)
        style = self.processing_progress.style()
        style.unpolish(self.processing_progress)
        style.polish(self.processing_progress)
        self.processing_progress.update()

    def _refresh_processing_progress(self) -> None:
        if self._progress_total <= 0:
            percent = 0
            format_text = "0%"
            tooltip_text = "处理进度：0/0"
        else:
            done = min(self._progress_done, self._progress_total)
            percent = int(done * 100 / self._progress_total)
            format_text = f"{percent}%  ({done}/{self._progress_total})"
            tooltip_text = f"处理进度：{done}/{self._progress_total}"
        self.processing_progress.setValue(percent)
        self.processing_progress.setFormat(format_text)
        self.processing_progress.setToolTip(tooltip_text)
        app = QApplication.instance()
        if app is not None:
            app.processEvents()

    def _start_processing_progress(self, total_items: int) -> None:
        self._progress_total = max(0, total_items)
        self._progress_done = 0
        self._set_progress_state("running" if self._progress_total > 0 else "idle")
        self._refresh_processing_progress()

    def _finish_processing_progress(self) -> None:
        if self._progress_total > 0:
            self._progress_done = self._progress_total
            self._set_progress_state("done")
        else:
            self._set_progress_state("idle")
        self._refresh_processing_progress()

    def _log_with_progress(self, level: str, message: str) -> None:
        if self._progress_total > 0:
            is_success = level == "INFO" and "[成功]" in message
            is_failed = level == "ERROR" and "[失败]" in message
            if is_success or is_failed:
                self._progress_done = min(self._progress_total, self._progress_done + 1)
                self._refresh_processing_progress()
        self._log(level, message)

    def _confirm_continue(self, title: str, message: str) -> bool:
        return SafeConfirmDialog.ask(
            self,
            title=title,
            message=message,
            confirm_text="仍继续执行",
        )

    def _open_statistics_mode_dialog(self) -> None:
        dialog = ProcessModeDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        if dialog.selected_mode == ProcessModeDialog.MODE_SINGLE:
            self._run_statistics()
            return
        if dialog.selected_mode == ProcessModeDialog.MODE_MERGE:
            self._run_merge()

    def _run_statistics(self) -> None:
        validated = self._validate_before_run()
        if validated is None:
            return
        inputs, output_dir = validated
        if self._contains_csv_file(inputs):
            should_continue = self._confirm_continue(
                "防误触提醒：单文件模式",
                "检测到当前上传内容包含 .csv 文件。\n"
                "“统计数据”通常仅处理 Excel 文件（.xlsx/.xls），可能存在误操作。\n"
                "是否仍继续执行“统计数据”？"
            )
            if not should_continue:
                self._log("WARN", "已取消执行：统计数据（检测到 .csv 文件）")
                return
        self._start_processing_progress(self._count_statistics_items(inputs))
        self._set_running(True)
        self._log("INFO", "开始执行：统计数据")
        chunk_size, wait_seconds = self._effective_pacing()
        if chunk_size is not None:
            self._log("INFO", f"分批策略：每批 {chunk_size} 个文件，批间等待 {wait_seconds} 秒")
        try:
            result = process_charge_statistics(
                inputs,
                output_dir,
                logger=self._log_with_progress,
                chunk_size=chunk_size,
                wait_seconds=wait_seconds,
            )
            self._log_batch_summary("统计数据", result)
        finally:
            self._finish_processing_progress()
            self._set_running(False)

    def _run_merge(self) -> None:
        validated = self._validate_before_run()
        if validated is None:
            return
        inputs, output_dir = validated
        excel_count, csv_count = self._count_excel_and_csv_files(inputs)
        if csv_count == 0:
            should_continue = self._confirm_continue(
                "防误触提醒：合并模式",
                "未检测到 .csv 文件。\n"
                "“合并后统计数据”需要使用 Excel（.xlsx/.xls）与 .csv 配对文件，可能存在误操作。\n"
                "是否仍继续执行“合并后统计数据”？"
            )
            if not should_continue:
                self._log("WARN", "已取消执行：合并后统计数据（未检测到 .csv 文件）")
                return
        elif excel_count != csv_count:
            should_continue = self._confirm_continue(
                "防误触提醒：合并模式",
                f"检测到 Excel 与 .csv 数量不一致（Excel={excel_count}，CSV={csv_count}）。\n"
                "这通常意味着混入了“单文件模式”输入，或存在未成对文件，可能导致误操作。\n"
                "是否仍继续执行“合并后统计数据”？"
            )
            if not should_continue:
                self._log(
                    "WARN",
                    f"已取消执行：合并后统计数据（Excel/CSV 数量不一致，Excel={excel_count}，CSV={csv_count}）",
                )
                return
        self._start_processing_progress(self._count_merge_items(inputs))
        self._set_running(True)
        self._log("INFO", "开始执行：合并后统计数据")
        chunk_size, wait_seconds = self._effective_pacing()
        if chunk_size is not None:
            self._log("INFO", f"分批策略：每批 {chunk_size} 个文件，批间等待 {wait_seconds} 秒")
        try:
            result = process_charge_merge(
                inputs,
                output_dir,
                logger=self._log_with_progress,
                chunk_size=chunk_size,
                wait_seconds=wait_seconds,
            )
            self._log_batch_summary("合并后统计数据", result)
        finally:
            self._finish_processing_progress()
            self._set_running(False)

    def _log_batch_summary(self, action: str, result: BatchResult) -> None:
        self._log(
            "INFO",
            f"{action}完成：总计 {result.total}，成功 {result.success}，失败 {result.failed}",
        )

    def closeEvent(self, event) -> None:  # noqa: N802
        self.log_bus.unsubscribe(self._append_runtime_log)
        super().closeEvent(event)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802
        if not event.mimeData().hasUrls():
            event.ignore()
            return
        dropped_paths: list[Path] = []
        for url in event.mimeData().urls():
            if url.isLocalFile():
                dropped_paths.append(Path(url.toLocalFile()))
        if dropped_paths:
            self.file_upload.add_paths(dropped_paths)
            self._log("INFO", f"拖拽导入 {len(dropped_paths)} 个路径")
        event.acceptProposedAction()
