from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QGroupBox,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QHBoxLayout,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.core.charge_merge_service import process_charge_merge
from src.core.charge_statistics_service import process_charge_statistics
from src.core.file_collect import collect_files
from src.core.logging_bus import LoggingBus
from src.core.models import BatchResult
from src.ui.components.file_upload import FileUploadWidget


class ChargeTab(QWidget):
    def __init__(self, log_bus: LoggingBus) -> None:
        super().__init__()
        self.log_bus = log_bus
        self.setAcceptDrops(True)
        self._build_ui()
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
        action_hint = QLabel("点击按钮后将按当前输入路径执行批处理，失败文件会在日志中标注。")
        action_hint.setObjectName("groupHint")
        action_hint.setWordWrap(True)
        action_row = QHBoxLayout()
        action_row.setSpacing(10)
        self.statistics_btn = QPushButton("统计数据")
        self.merge_btn = QPushButton("合并后统计数据")
        self.statistics_btn.setProperty("accent", "success")
        self.merge_btn.setProperty("accent", "primary")
        action_row.addWidget(self.statistics_btn)
        action_row.addWidget(self.merge_btn)
        action_row.addStretch()
        action_layout.addWidget(action_hint)
        action_layout.addLayout(action_row)
        left_layout.addWidget(action_group)

        output_group = QGroupBox("输出设置")
        output_group.setObjectName("outputGroup")
        output_layout = QVBoxLayout(output_group)
        output_layout.setContentsMargins(14, 20, 14, 14)
        output_layout.setSpacing(10)
        output_hint = QLabel("默认输出到项目目录下 output 文件夹，你也可以手动切换。")
        output_hint.setObjectName("groupHint")
        output_hint.setWordWrap(True)
        output_row = QHBoxLayout()
        output_row.setSpacing(10)
        output_label = QLabel("输出目录")
        output_label.setObjectName("fieldLabel")
        self.output_edit = QLineEdit(str((Path.cwd() / "output").resolve()))
        self.output_edit.setObjectName("pathEdit")
        self.browse_output_btn = QPushButton("浏览")
        output_row.addWidget(output_label)
        output_row.addWidget(self.output_edit, stretch=1)
        output_row.addWidget(self.browse_output_btn)
        output_layout.addWidget(output_hint)
        output_layout.addLayout(output_row)
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
        self.statistics_btn.clicked.connect(self._run_statistics)
        self.merge_btn.clicked.connect(self._run_merge)

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

    def _current_inputs(self) -> list[Path]:
        return self.file_upload.get_paths()

    def _set_running(self, running: bool) -> None:
        for button in (
            self.browse_output_btn,
            self.statistics_btn,
            self.merge_btn,
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

    def _confirm_continue(self, message: str) -> bool:
        reply = QMessageBox.question(
            self,
            "操作确认",
            message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return reply == QMessageBox.StandardButton.Yes

    def _run_statistics(self) -> None:
        validated = self._validate_before_run()
        if validated is None:
            return
        inputs, output_dir = validated
        if self._contains_csv_file(inputs):
            should_continue = self._confirm_continue(
                "检测到当前上传内容包含 .csv 文件。\n"
                "“统计数据”通常仅处理 Excel 文件（.xlsx/.xls），可能存在误操作。\n"
                "是否仍继续执行“统计数据”？"
            )
            if not should_continue:
                self._log("WARN", "已取消执行：统计数据（检测到 .csv 文件）")
                return
        self._set_running(True)
        self._log("INFO", "开始执行：统计数据")
        try:
            result = process_charge_statistics(inputs, output_dir, logger=self._log)
            self._log_batch_summary("统计数据", result)
        finally:
            self._set_running(False)

    def _run_merge(self) -> None:
        validated = self._validate_before_run()
        if validated is None:
            return
        inputs, output_dir = validated
        if not self._contains_csv_file(inputs):
            should_continue = self._confirm_continue(
                "未检测到 .csv 文件。\n"
                "“合并后统计数据”需要使用 Excel（.xlsx/.xls）与 .csv 配对文件，可能存在误操作。\n"
                "是否仍继续执行“合并后统计数据”？"
            )
            if not should_continue:
                self._log("WARN", "已取消执行：合并后统计数据（未检测到 .csv 文件）")
                return
        self._set_running(True)
        self._log("INFO", "开始执行：合并后统计数据")
        try:
            result = process_charge_merge(inputs, output_dir, logger=self._log)
            self._log_batch_summary("合并后统计数据", result)
        finally:
            self._set_running(False)

    def _log_batch_summary(self, action: str, result: BatchResult) -> None:
        self._log(
            "INFO",
            f"{action}完成：总计 {result.total}，成功 {result.success}，失败 {result.failed}",
        )
        for item in result.items:
            if item.status == "success":
                self._log("INFO", f"[成功] {item.name} -> {item.output_path}")
            else:
                self._log("ERROR", f"[失败] {item.name} -> {item.error}")

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
