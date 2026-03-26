from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QSettings, QThread, Qt
from PySide6.QtGui import QDoubleValidator, QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.core.file_collect import collect_files
from src.core.logging_bus import LoggingBus
from src.core.models import BatchResult
from src.core.voltage_sampling_service import process_voltage_sampling_statistics
from src.ui.background_task import BackgroundTaskWorker
from src.ui.charge_tab import BatchPacingDialog, BatchPacingSettings, SafeConfirmDialog
from src.ui.components.file_upload import FileUploadWidget


class VoltageSamplingModeDialog(QDialog):
    MODE_SINGLE = "single"

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.selected_mode: str | None = None
        self.setObjectName("modeSelectDialog")
        self.setWindowTitle("选择统计模式")
        self.setModal(True)
        self.resize(520, 300)

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
        subtitle = QLabel("点击下方模式按钮后立即执行。")
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
        options_layout.addWidget(
            self._build_mode_card(
                title="单文件模式",
                description="仅处理 Excel（.xlsx/.xls），用于统计分压采集测试数据并导出折线图。",
                button_text="执行单文件模式",
                button_accent="success",
                card_kind="single",
                mode=self.MODE_SINGLE,
            )
        )
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


class VoltageSamplingTab(QWidget):
    BATCH_PACING_ENABLED_KEY = "voltage_sampling_tab/batch_pacing/enabled"
    BATCH_PACING_CHUNK_SIZE_KEY = "voltage_sampling_tab/batch_pacing/chunk_size"
    BATCH_PACING_WAIT_SECONDS_KEY = "voltage_sampling_tab/batch_pacing/wait_seconds"

    def __init__(self, log_bus: LoggingBus) -> None:
        super().__init__()
        self.log_bus = log_bus
        self.batch_pacing_settings = self._load_batch_pacing_settings()
        self._progress_total = 0
        self._progress_done = 0
        self._worker_thread: QThread | None = None
        self._worker: BackgroundTaskWorker | None = None
        self._running_action: str = ""
        self.setAcceptDrops(True)
        self._build_ui()
        self._refresh_batch_pacing_hint()
        self.log_bus.subscribe(self._append_runtime_log)

    def _build_ui(self) -> None:
        self.setObjectName("chargeRoot")
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(20, 20, 20, 18)
        root_layout.setSpacing(12)

        page_title = QLabel("分压采集测试")
        page_title.setObjectName("pageTitle")
        page_subtitle = QLabel("左侧配置文件上传、采样电阻与输出设置；右侧查看运行日志。")
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
        upload_hint = QLabel("支持拖拽或选择文件/文件夹。分压采集测试仅处理 Excel（.xlsx/.xls）。")
        upload_hint.setObjectName("groupHint")
        upload_hint.setWordWrap(True)
        upload_layout.addWidget(upload_hint)
        self.file_upload = FileUploadWidget(
            supported_hint="支持 .xlsx / .xls 格式",
            file_filter="Excel Files (*.xlsx *.xls);;All Files (*.*)",
        )
        self.file_upload.paths_changed.connect(self._on_paths_changed)
        upload_layout.addWidget(self.file_upload)

        resistor_container = QWidget()
        resistor_layout = QHBoxLayout(resistor_container)
        resistor_layout.setContentsMargins(0, 0, 0, 0)
        resistor_layout.setSpacing(6)
        resistor_label = QLabel("采样电阻阻值：")
        resistor_label.setObjectName("fieldLabel")
        self.sampling_resistance_edit = QLineEdit("20")
        self.sampling_resistance_edit.setObjectName("pathEdit")
        self.sampling_resistance_edit.setFixedWidth(82)
        self.sampling_resistance_edit.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.sampling_resistance_edit.setValidator(QDoubleValidator(0.0, 1000000.0, 6, self))
        resistor_unit = QLabel("mΩ")
        resistor_unit.setObjectName("fieldLabel")
        resistor_layout.addWidget(resistor_label)
        resistor_layout.addWidget(self.sampling_resistance_edit)
        resistor_layout.addWidget(resistor_unit)
        self.file_upload.insert_button_row_widget(resistor_container)
        left_layout.addWidget(upload_group, stretch=1)

        action_group = QGroupBox("数据处理")
        action_group.setObjectName("actionGroup")
        action_layout = QVBoxLayout(action_group)
        action_layout.setContentsMargins(14, 20, 14, 14)
        action_layout.setSpacing(10)
        action_hint = QLabel("点击“统计数据”后可在弹窗中选择“执行单文件模式”。")
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
        output_hint = QLabel("默认输出到项目目录下 output 文件夹；如需降低批量处理失败风险，可在右侧“分批策略...”中设置分批等待。")
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
        level_color = {"INFO": "#3498db", "WARN": "#e67e22", "ERROR": "#e74c3c"}.get(level, "#3498db")
        level_bg = {"INFO": "#eaf4fc", "WARN": "#fef5e7", "ERROR": "#fdedec"}.get(level, "#eaf4fc")
        formatted_msg = self._format_log_message(message)
        self.runtime_log_view.append(
            f'<span style="background-color:{level_bg};color:{level_color};font-weight:700;'
            f'padding:2px 6px;border-radius:4px;">[{level}]</span> '
            f"{formatted_msg}"
        )

    def _format_log_message(self, message: str) -> str:
        import re

        result = message
        result = re.sub(r"(\[\d{2}:\d{2}:\d{2}\])", r'<span style="color:#2980b9;font-weight:600;">\1</span>', result)
        result = re.sub(r"(\[成功\])", r'<span style="color:#27ae60;font-weight:700;">\1</span>', result)
        result = re.sub(r"(\[失败\])", r'<span style="color:#e74c3c;font-weight:700;">\1</span>', result)
        result = re.sub(r"(->)", r'<span style="color:#9b59b6;font-weight:600;">→</span>', result)
        if "<span" not in result or result == message:
            return f'<span style="color:#2c3e50;">{result}</span>'

        parts = re.split(r"(<span[^>]*>.*?</span>)", result)
        formatted_parts: list[str] = []
        for part in parts:
            if part.startswith("<span"):
                formatted_parts.append(part)
            elif part.strip():
                formatted_parts.append(f'<span style="color:#2c3e50;">{part}</span>')
            else:
                formatted_parts.append(part)
        return "".join(formatted_parts)

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
        enabled = self._to_bool(settings.value(self.BATCH_PACING_ENABLED_KEY, False), default=False)
        chunk_size = max(1, self._to_int(settings.value(self.BATCH_PACING_CHUNK_SIZE_KEY, 4), default=4))
        wait_seconds = max(1, self._to_int(settings.value(self.BATCH_PACING_WAIT_SECONDS_KEY, 3), default=3))
        return BatchPacingSettings(enabled=enabled, chunk_size=chunk_size, wait_seconds=wait_seconds)

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
        self.batch_pacing_hint.setText("说明：分批策略默认关闭。开启后可按“每批文件数 + 等待秒数”分批执行。")

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
            self._log("INFO", f"分批策略已更新：每批 {updated.chunk_size} 个文件，批间等待 {updated.wait_seconds} 秒")
            return
        self._log("INFO", "分批策略已关闭")

    def _effective_pacing(self) -> tuple[int | None, int]:
        if not self.batch_pacing_settings.enabled:
            return None, 0
        return self.batch_pacing_settings.chunk_size, self.batch_pacing_settings.wait_seconds

    def _current_inputs(self) -> list[Path]:
        return self.file_upload.get_paths()

    def _set_running(self, running: bool) -> None:
        for button in (self.browse_output_btn, self.batch_pacing_btn, self.statistics_btn, self.clear_log_btn):
            button.setDisabled(running)
        self.file_upload.set_controls_enabled(not running)
        self.sampling_resistance_edit.setDisabled(running)

    def _parse_sampling_resistance_milliohm(self) -> float | None:
        text = self.sampling_resistance_edit.text().strip()
        if not text:
            QMessageBox.warning(self, "提示", "请先填写采样电阻阻值。")
            return None
        try:
            value = float(text)
        except ValueError:
            QMessageBox.warning(self, "提示", "采样电阻阻值格式无效，请输入数字。")
            return None
        if value <= 0:
            QMessageBox.warning(self, "提示", "采样电阻阻值必须大于 0 mΩ。")
            return None
        return value

    def _validate_before_run(self) -> tuple[list[Path], Path, float] | None:
        inputs = self._current_inputs()
        if not inputs:
            QMessageBox.warning(self, "提示", "请先添加至少一个输入文件或文件夹。")
            return None
        output_text = self.output_edit.text().strip()
        if not output_text:
            QMessageBox.warning(self, "提示", "请先设置输出目录。")
            return None
        sampling_resistance_milliohm = self._parse_sampling_resistance_milliohm()
        if sampling_resistance_milliohm is None:
            return None
        return inputs, Path(output_text), sampling_resistance_milliohm

    @staticmethod
    def _build_mode_output_dir(base_output_dir: Path, mode_label: str) -> Path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"{mode_label}_{timestamp}"
        candidate = base_output_dir / folder_name
        if not candidate.exists():
            candidate.mkdir(parents=True, exist_ok=False)
            return candidate

        suffix = 1
        while True:
            candidate = base_output_dir / f"{folder_name}({suffix})"
            if not candidate.exists():
                candidate.mkdir(parents=True, exist_ok=False)
                return candidate
            suffix += 1

    def _contains_non_excel_files(self, inputs: list[Path]) -> bool:
        files, _ = collect_files(inputs, {".csv", ".txt", ".log"})
        return bool(files)

    def _count_statistics_items(self, inputs: list[Path]) -> int:
        files, _ = collect_files(inputs, {".xlsx", ".xls"})
        return len(files)

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

    def _dispatch_statistics_task(
        self,
        *,
        action: str,
        output_dir: Path,
        total_items: int,
        task,
        task_args: tuple[object, ...],
        chunk_size: int | None,
        wait_seconds: int,
    ) -> None:
        if self._worker_thread is not None:
            self._log("WARN", "当前已有统计任务在执行，请稍后再试。")
            return

        self._running_action = action
        self._start_processing_progress(total_items)
        self._set_running(True)
        self._log("INFO", f"开始执行：{action}")
        self._log("INFO", f"本次输出目录：{output_dir}")
        if chunk_size is not None:
            self._log("INFO", f"分批策略：每批 {chunk_size} 个文件，批间等待 {wait_seconds} 秒")

        task_kwargs = {"chunk_size": chunk_size, "wait_seconds": wait_seconds}
        worker = BackgroundTaskWorker(task, args=task_args, kwargs=task_kwargs)
        thread = QThread(self)
        worker.moveToThread(thread)

        thread.started.connect(worker.run)
        worker.log.connect(self._log_with_progress)
        worker.result_ready.connect(self._on_task_result)
        worker.error.connect(self._on_task_error)
        worker.finished.connect(self._on_task_finished)
        worker.finished.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(self._reset_worker_state)

        self._worker = worker
        self._worker_thread = thread
        thread.start()

    def _on_task_result(self, result: object) -> None:
        if isinstance(result, BatchResult):
            self._log_batch_summary(self._running_action, result)
            return
        self._log("ERROR", "[失败] 后台任务返回了未知结果类型，无法输出汇总。")

    def _on_task_error(self, detail: str) -> None:
        action = self._running_action or "统计数据"
        self._log("ERROR", f"[失败] {action} 后台任务异常：{detail}")

    def _on_task_finished(self) -> None:
        self._finish_processing_progress()
        self._set_running(False)

    def _reset_worker_state(self) -> None:
        self._worker = None
        self._worker_thread = None
        self._running_action = ""

    def has_running_task(self) -> bool:
        return self._worker_thread is not None and self._worker_thread.isRunning()

    def _confirm_continue(self, title: str, message: str) -> bool:
        return SafeConfirmDialog.ask(self, title=title, message=message, confirm_text="仍继续执行")

    def _open_statistics_mode_dialog(self) -> None:
        dialog = VoltageSamplingModeDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        if dialog.selected_mode == VoltageSamplingModeDialog.MODE_SINGLE:
            self._run_statistics()

    def _run_statistics(self) -> None:
        validated = self._validate_before_run()
        if validated is None:
            return
        inputs, output_dir, sampling_resistance_milliohm = validated
        if self._contains_non_excel_files(inputs):
            should_continue = self._confirm_continue(
                "防误触提醒：单文件模式",
                "检测到当前上传内容包含非 Excel 文件（如 .csv/.txt/.log）。\n"
                "“执行单文件模式”通常仅处理 Excel 文件（.xlsx/.xls），可能存在误操作。\n"
                "是否仍继续执行“执行单文件模式”？",
            )
            if not should_continue:
                self._log("WARN", "已取消执行：执行单文件模式（检测到非 Excel 文件）")
                return

        try:
            mode_output_dir = self._build_mode_output_dir(output_dir, "单文件模式")
        except OSError as exc:
            self._log("ERROR", f"创建输出目录失败：{exc}")
            QMessageBox.critical(self, "错误", f"创建输出目录失败：{exc}")
            return

        self._log("INFO", f"采样电阻阻值：{sampling_resistance_milliohm:g} mΩ")
        chunk_size, wait_seconds = self._effective_pacing()
        self._dispatch_statistics_task(
            action="执行单文件模式",
            output_dir=mode_output_dir,
            total_items=self._count_statistics_items(inputs),
            task=process_voltage_sampling_statistics,
            task_args=(inputs, mode_output_dir, sampling_resistance_milliohm),
            chunk_size=chunk_size,
            wait_seconds=wait_seconds,
        )

    def _log_batch_summary(self, action: str, result: BatchResult) -> None:
        self._log("INFO", f"{action}完成：总计 {result.total}，成功 {result.success}，失败 {result.failed}")

    def closeEvent(self, event) -> None:  # noqa: N802
        if self.has_running_task():
            QMessageBox.warning(self, "提示", "当前仍有统计任务正在执行，请等待完成后再关闭窗口。")
            event.ignore()
            return
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
