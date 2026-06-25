"""
normalize_excel_gui.py — PyQt6 desktop GUI for the Persian Excel normaliser.

Layer model
-----------
  Presentation  ->  MainWindow   (widgets, layouts, dialogs, settings)
  Worker        ->  NormalizerWorker(QObject) + QThread
  Core          ->  normalize_workbook()  in normalize_excel.py

Threading model
---------------
NormalizerWorker subclasses QObject (not QThread). moveToThread() places it
on a dedicated QThread. All GUI mutations happen through Qt signals so the
main thread is the only thread touching widgets.
"""

from __future__ import annotations

import html
import logging
import sys
import traceback
from pathlib import Path
from typing import Optional

from PyQt6.QtCore import (
    QMimeData,
    QObject,
    QSettings,
    QThread,
    Qt,
    pyqtSignal,
    pyqtSlot,
)
from PyQt6.QtGui import (
    QAction,
    QCloseEvent,
    QDragEnterEvent,
    QDropEvent,
    QFont,
    QKeySequence,
)
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStatusBar,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

try:
    from normalize_excel import build_output_path, normalize_workbook
except ImportError as _err:
    print(f"ERROR: normalize_excel.py not found — {_err}", file=sys.stderr)
    sys.exit(1)

log: logging.Logger = logging.getLogger("excel_normalizer.gui")

APP_NAME = "Persian Excel Normalizer"
APP_VERSION = "2.3.0"
SETTINGS_ORG = "PersianDevTools"
SETTINGS_APP = "PersianExcelNormalizer"
MAX_LOG_LINES = 5_000
EXCEL_OPEN_FILTER = "Excel Files (*.xlsx *.xlsm);;All Files (*)"
EXCEL_SAVE_FILTER = (
    "Excel Workbook (*.xlsx);;"
    "Excel Macro-Enabled Workbook (*.xlsm);;"
    "All Files (*)"
)

LIGHT_THEME = """
QMainWindow, QDialog, QWidget { background-color: #f5f5f5; color: #1a1a1a; }
QLineEdit, QTextEdit { background-color: #ffffff; border: 1px solid #cccccc; border-radius: 3px; padding: 4px; color: #1a1a1a; }
QPushButton { background-color: #e6e6e6; border: 1px solid #bbbbbb; border-radius: 4px; padding: 5px 14px; color: #1a1a1a; }
QPushButton:hover { background-color: #d8d8d8; }
QPushButton:disabled { color: #aaaaaa; background-color: #f0f0f0; }
QPushButton#run_btn { background-color: #4a90d9; color: #ffffff; border: 1px solid #357abd; font-weight: bold; font-size: 11pt; }
QPushButton#cancel_btn { background-color: #d9534f; color: #ffffff; border: 1px solid #b52b27; }
QProgressBar { border: 1px solid #cccccc; border-radius: 3px; text-align: center; background-color: #e8e8e8; color: #1a1a1a; height: 16px; }
QProgressBar::chunk { background-color: #4a90d9; border-radius: 2px; }
QGroupBox { border: 1px solid #cccccc; border-radius: 5px; margin-top: 10px; padding-top: 6px; font-weight: bold; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; left: 8px; padding: 0 4px; }
QStatusBar { color: #555555; border-top: 1px solid #cccccc; }
"""

DARK_THEME = """
QMainWindow, QDialog, QWidget { background-color: #1e1e1e; color: #d4d4d4; }
QLineEdit, QTextEdit { background-color: #2d2d2d; border: 1px solid #555555; border-radius: 3px; padding: 4px; color: #d4d4d4; }
QPushButton { background-color: #3c3c3c; border: 1px solid #555555; border-radius: 4px; padding: 5px 14px; color: #d4d4d4; }
QPushButton:hover { background-color: #4a4a4a; }
QPushButton:disabled { color: #666666; background-color: #2d2d2d; border-color: #444444; }
QPushButton#run_btn { background-color: #0e639c; color: #ffffff; border: 1px solid #1177bb; font-weight: bold; font-size: 11pt; }
QPushButton#cancel_btn { background-color: #7a1a1a; color: #ffffff; border: 1px solid #9a2020; }
QProgressBar { border: 1px solid #555555; border-radius: 3px; text-align: center; background-color: #3c3c3c; color: #d4d4d4; height: 16px; }
QProgressBar::chunk { background-color: #0e639c; border-radius: 2px; }
QGroupBox { border: 1px solid #444444; border-radius: 5px; margin-top: 10px; padding-top: 6px; font-weight: bold; color: #d4d4d4; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; left: 8px; padding: 0 4px; }
QStatusBar { color: #888888; border-top: 1px solid #444444; }
"""


class _LogBridge(QObject):
    """Carries log records from any thread to the GUI thread."""

    new_record = pyqtSignal(str, str)  # (html_line, level_name)


class QtLogHandler(logging.Handler):
    """Thread-safe logging handler that forwards records via Qt signals."""

    _COLORS_LIGHT: dict[str, str] = {
        "DEBUG": "#888888",
        "INFO": "#1a1a1a",
        "WARNING": "#996600",
        "ERROR": "#cc2200",
        "CRITICAL": "#ee0000",
    }
    _COLORS_DARK: dict[str, str] = {
        "DEBUG": "#777777",
        "INFO": "#cccccc",
        "WARNING": "#f0a500",
        "ERROR": "#e05555",
        "CRITICAL": "#ff4444",
    }

    def __init__(self) -> None:
        super().__init__()
        self.bridge = _LogBridge()
        self._dark = False
        self.setFormatter(
            logging.Formatter(
                "%(asctime)s [%(levelname)-8s] %(message)s",
                datefmt="%H:%M:%S",
            )
        )

    def set_dark(self, dark: bool) -> None:
        self._dark = dark

    def emit(self, record: logging.LogRecord) -> None:
        try:
            line = self.format(record)
            palette = self._COLORS_DARK if self._dark else self._COLORS_LIGHT
            colour = palette.get(record.levelname, "#888888")
            escaped = html.escape(line, quote=False).replace(" ", "&nbsp;")
            self.bridge.new_record.emit(
                f'<span style="color:{colour};font-family:monospace;">{escaped}</span>',
                record.levelname,
            )
        except Exception:
            self.handleError(record)


class NormalizerWorker(QObject):
    """Executes normalisation jobs on a dedicated QThread."""

    progress = pyqtSignal(int, int)   # (processed_cells, total_cells)
    status = pyqtSignal(str)
    finished = pyqtSignal(list, bool)  # (completed_jobs, was_cancelled)
    error = pyqtSignal(str, str)       # (user_message, full_traceback)

    def __init__(
        self,
        jobs: list[tuple[Path, Path]],
        replace_zwnj: bool,
        parent: Optional[QObject] = None,
    ) -> None:
        super().__init__(parent)
        self._jobs = jobs
        self._replace_zwnj = replace_zwnj
        self._cancel_requested = False

    def cancel(self) -> None:
        """Request graceful cancellation. Safe to call from the GUI thread."""
        self._cancel_requested = True

    def _is_cancel_requested(self) -> bool:
        return self._cancel_requested

    @pyqtSlot()
    def run(self) -> None:
        completed: list[tuple[Path, Path]] = []
        total_files = len(self._jobs)

        for idx, (source, destination) in enumerate(self._jobs, start=1):
            if self._cancel_requested:
                log.info("Cancelled before file %d/%d.", idx, total_files)
                break

            self.status.emit(f"File {idx}/{total_files}: {source.name}")
            log.info("Processing %d/%d: %s", idx, total_files, source.name)

            try:
                normalize_workbook(
                    source,
                    destination,
                    replace_zwnj=self._replace_zwnj,
                    cancel_check=self._is_cancel_requested,
                    on_sheet_start=self._on_sheet_start,
                    on_progress=self.progress.emit,
                )
            except Exception as exc:  # noqa: BLE001
                detail = traceback.format_exc()
                log.error("Failed on '%s': %s", source.name, exc)
                self.error.emit(f"Error processing '{source.name}':\n{exc}", detail)
                continue

            if not self._cancel_requested:
                completed.append((source, destination))
                self.progress.emit(1, 1)
                log.info("Saved: %s", destination.name)

        if self._cancel_requested:
            self.status.emit("Cancelled by user.")
            log.info("Job cancelled by user request.")
        else:
            self.status.emit(
                "Completed successfully." if completed else "Finished — check log for errors."
            )

        self.finished.emit(completed, self._cancel_requested)

    def _on_sheet_start(self, idx: int, total: int, name: str) -> None:
        self.status.emit(f"Sheet {idx}/{total}: {name}")
        log.debug("  Sheet %d/%d: %s", idx, total, name)


class AboutDialog(QDialog):
    """Minimal application-information dialog."""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"About {APP_NAME}")
        self.setFixedSize(440, 300)
        layout = QVBoxLayout(self)
        title = QLabel(f"<h2>{APP_NAME}</h2>")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version = QLabel(f"Version {APP_VERSION}")
        version.setAlignment(Qt.AlignmentFlag.AlignCenter)
        description = QLabel(
            "Normalises Persian text in Excel files.\n\n"
            "• Arabic characters  →  Persian equivalents\n"
            "• ZWNJ (U+200C)  →  regular space  (optional)\n"
            "• Invisible control characters removed\n"
            "• NBSP / multiple spaces normalised\n"
            "• Formula cells are preserved untouched\n"
            "• .xlsm VBA archives are preserved by openpyxl keep_vba=True\n\n"
            "Core: normalize_excel.py  |  GUI: PyQt6"
        )
        description.setAlignment(Qt.AlignmentFlag.AlignCenter)
        description.setWordWrap(True)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(self.accept)
        layout.addWidget(title)
        layout.addWidget(version)
        layout.addWidget(description)
        layout.addStretch()
        layout.addWidget(buttons)


class MainWindow(QMainWindow):
    """Primary application window and worker/thread lifecycle owner."""

    def __init__(self) -> None:
        super().__init__()
        self._settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
        self._thread: Optional[QThread] = None
        self._worker: Optional[NormalizerWorker] = None
        self._input_paths: list[Path] = []
        self._current_theme = "light"

        self._log_handler = QtLogHandler()
        pkg_logger = logging.getLogger("excel_normalizer")
        pkg_logger.addHandler(self._log_handler)
        pkg_logger.setLevel(logging.DEBUG)
        pkg_logger.propagate = False

        self._build_ui()
        self._build_menu()
        self._wire_signals()
        self._restore_settings()
        self.setAcceptDrops(True)

    def _build_ui(self) -> None:
        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.setMinimumSize(780, 640)
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(8, 8, 8, 8)

        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.setChildrenCollapsible(False)
        top_widget = QWidget()
        top_layout = QVBoxLayout(top_widget)
        top_layout.addWidget(self._build_input_group())
        top_layout.addWidget(self._build_output_group())
        top_layout.addWidget(self._build_options_group())
        top_layout.addWidget(self._build_run_group())
        top_layout.addStretch()
        splitter.addWidget(top_widget)
        splitter.addWidget(self._build_log_group())
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 3)
        root_layout.addWidget(splitter)
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Ready")

    def _build_input_group(self) -> QGroupBox:
        group = QGroupBox("Input File(s)")
        row = QHBoxLayout(group)
        self._input_edit = QLineEdit()
        self._input_edit.setPlaceholderText(
            "Select or drag-and-drop one or more .xlsx / .xlsm files…"
        )
        self._input_edit.setReadOnly(True)
        browse_btn = QPushButton("Browse…")
        browse_btn.setFixedWidth(80)
        browse_btn.clicked.connect(self._browse_input)
        row.addWidget(self._input_edit)
        row.addWidget(browse_btn)
        return group

    def _build_output_group(self) -> QGroupBox:
        group = QGroupBox("Output File")
        row = QHBoxLayout(group)
        self._output_edit = QLineEdit()
        self._output_edit.setPlaceholderText(
            "Auto-generated: <input_stem>_normalized<same_suffix>"
        )
        self._output_browse_btn = QPushButton("Browse…")
        self._output_browse_btn.setFixedWidth(80)
        self._output_browse_btn.setToolTip(
            "Choose a custom .xlsx or .xlsm output path (single-file mode only)"
        )
        self._output_browse_btn.clicked.connect(self._browse_output)
        row.addWidget(self._output_edit)
        row.addWidget(self._output_browse_btn)
        self._sync_output_controls(running=False)
        return group

    def _build_options_group(self) -> QGroupBox:
        group = QGroupBox("Options")
        col = QVBoxLayout(group)
        self._zwnj_check = QCheckBox(
            "Preserve ZWNJ (U+200C)  — keep «نیم‌فاصله» intact"
        )
        self._zwnj_check.setChecked(False)
        col.addWidget(self._zwnj_check)
        return group

    def _build_run_group(self) -> QGroupBox:
        group = QGroupBox("Run")
        col = QVBoxLayout(group)
        btn_row = QHBoxLayout()
        self._run_btn = QPushButton("Normalize Excel")
        self._run_btn.setObjectName("run_btn")
        self._run_btn.setFixedHeight(40)
        self._run_btn.setEnabled(False)
        self._cancel_btn = QPushButton("Cancel")
        self._cancel_btn.setObjectName("cancel_btn")
        self._cancel_btn.setFixedSize(90, 40)
        self._cancel_btn.setEnabled(False)
        btn_row.addWidget(self._run_btn, stretch=1)
        btn_row.addWidget(self._cancel_btn)
        self._progress_bar = QProgressBar()
        self._progress_bar.setRange(0, 100)
        self._progress_bar.setValue(0)
        self._progress_bar.setTextVisible(True)
        self._status_label = QLabel("Ready")
        self._status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        col.addLayout(btn_row)
        col.addWidget(self._progress_bar)
        col.addWidget(self._status_label)
        return group

    def _build_log_group(self) -> QGroupBox:
        group = QGroupBox("Log")
        group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        col = QVBoxLayout(group)
        toolbar = QHBoxLayout()
        toolbar.addStretch()
        copy_btn = QPushButton("Copy Log")
        copy_btn.setFixedWidth(84)
        copy_btn.clicked.connect(self._copy_log)
        clear_btn = QPushButton("Clear Log")
        clear_btn.setFixedWidth(84)
        clear_btn.clicked.connect(self._clear_log)
        toolbar.addWidget(copy_btn)
        toolbar.addWidget(clear_btn)
        self._log_edit = QTextEdit()
        self._log_edit.setReadOnly(True)
        self._log_edit.setFont(QFont("Courier New", 9))
        self._log_edit.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self._log_edit.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self._log_edit.document().setMaximumBlockCount(MAX_LOG_LINES)
        col.addLayout(toolbar)
        col.addWidget(self._log_edit)
        return group

    def _build_menu(self) -> None:
        bar = self.menuBar()
        file_menu = bar.addMenu("File")
        open_act = QAction("Open File(s)…", self)
        open_act.setShortcut(QKeySequence("Ctrl+O"))
        open_act.triggered.connect(self._browse_input)
        quit_act = QAction("Quit", self)
        quit_act.setShortcut(QKeySequence("Ctrl+Q"))
        quit_act.triggered.connect(self.close)
        file_menu.addAction(open_act)
        file_menu.addSeparator()
        file_menu.addAction(quit_act)

        view_menu = bar.addMenu("View")
        light_act = QAction("Light Theme", self)
        light_act.triggered.connect(self._apply_light_theme)
        dark_act = QAction("Dark Theme", self)
        dark_act.triggered.connect(self._apply_dark_theme)
        clear_log_act = QAction("Clear Log", self)
        clear_log_act.setShortcut(QKeySequence("Ctrl+L"))
        clear_log_act.triggered.connect(self._clear_log)
        view_menu.addAction(light_act)
        view_menu.addAction(dark_act)
        view_menu.addSeparator()
        view_menu.addAction(clear_log_act)

        help_menu = bar.addMenu("Help")
        about_act = QAction(f"About {APP_NAME}", self)
        about_act.triggered.connect(self._show_about)
        help_menu.addAction(about_act)

    def _wire_signals(self) -> None:
        self._run_btn.clicked.connect(self._run_normalization)
        self._cancel_btn.clicked.connect(self._cancel_job)
        self._log_handler.bridge.new_record.connect(self._append_log_line)

    def _restore_settings(self) -> None:
        geometry = self._settings.value("window/geometry")
        if geometry:
            self.restoreGeometry(geometry)
        else:
            self.resize(860, 680)
            self._centre_on_screen()
        state = self._settings.value("window/state")
        if state:
            self.restoreState(state)
        self._apply_theme(self._settings.value("theme", "light", type=str))

    def _persist_settings(self) -> None:
        self._settings.setValue("window/geometry", self.saveGeometry())
        self._settings.setValue("window/state", self.saveState())
        self._settings.setValue("theme", self._current_theme)

    def _centre_on_screen(self) -> None:
        screen = QApplication.primaryScreen()
        if screen is None:
            return
        frame = self.frameGeometry()
        frame.moveCenter(screen.availableGeometry().center())
        self.move(frame.topLeft())

    def _apply_light_theme(self) -> None:
        self._apply_theme("light")

    def _apply_dark_theme(self) -> None:
        self._apply_theme("dark")

    def _apply_theme(self, name: str) -> None:
        self._current_theme = "dark" if name == "dark" else "light"
        dark = self._current_theme == "dark"
        self._log_handler.set_dark(dark)
        app = QApplication.instance()
        if app is not None:
            app.setStyleSheet(DARK_THEME if dark else LIGHT_THEME)

    def _browse_input(self) -> None:
        last_dir = self._settings.value("last_directory", str(Path.home()), type=str)
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Excel File(s)",
            last_dir,
            EXCEL_OPEN_FILTER,
        )
        if paths:
            self._settings.setValue("last_directory", str(Path(paths[0]).parent))
            self._load_paths(paths)

    def _browse_output(self) -> None:
        last_dir = self._settings.value("last_directory", str(Path.home()), type=str)
        default_suffix = (
            self._input_paths[0].suffix if len(self._input_paths) == 1 else ".xlsx"
        )
        suggested = (
            str(build_output_path(self._input_paths[0]))
            if len(self._input_paths) == 1
            else last_dir
        )
        selected_filter = (
            "Excel Macro-Enabled Workbook (*.xlsm)"
            if default_suffix.lower() == ".xlsm"
            else "Excel Workbook (*.xlsx)"
        )
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Normalised File As",
            suggested,
            EXCEL_SAVE_FILTER,
            selected_filter,
        )
        if path:
            output = Path(path)
            if not output.suffix:
                output = output.with_suffix(default_suffix)
            self._output_edit.setText(str(output))

    def _load_paths(self, raw_paths: list[str]) -> None:
        valid: list[Path] = []
        for raw in raw_paths:
            p = Path(raw)
            if p.suffix.lower() in {".xlsx", ".xlsm"} and p.exists():
                valid.append(p)
            else:
                log.warning("Skipping unsupported or missing file: %s", raw)

        if not valid:
            QMessageBox.warning(
                self,
                "No Valid Files",
                "None of the selected files are valid .xlsx / .xlsm files.",
            )
            return

        self._input_paths = valid
        if len(valid) == 1:
            self._input_edit.setText(str(valid[0]))
            self._output_edit.setText(str(build_output_path(valid[0])))
        else:
            self._input_edit.setText(f"{len(valid)} files selected")
            self._output_edit.clear()

        self._sync_output_controls(running=False)
        self._run_btn.setEnabled(True)
        self.statusBar().showMessage(f"{len(valid)} file(s) loaded")
        log.info("Loaded %d file(s).", len(valid))

    def _run_normalization(self) -> None:
        if self._is_running():
            return
        if not self._input_paths:
            QMessageBox.information(
                self, "No Input", "Please select at least one .xlsx / .xlsm file first."
            )
            return

        custom_output: Optional[Path] = None
        custom_text = self._output_edit.text().strip()
        if len(self._input_paths) == 1 and custom_text:
            custom_output = Path(custom_text)

        jobs: list[tuple[Path, Path]] = [
            (
                src,
                custom_output
                if custom_output is not None and len(self._input_paths) == 1
                else build_output_path(src),
            )
            for src in self._input_paths
        ]

        replace_zwnj = not self._zwnj_check.isChecked()
        self._thread = QThread(self)
        self._worker = NormalizerWorker(jobs, replace_zwnj)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.finished.connect(self._on_thread_done)
        self._worker.progress.connect(self._on_progress)
        self._worker.status.connect(self._on_status)
        self._worker.error.connect(self._on_worker_error)
        self._worker.finished.connect(self._on_worker_finished)

        self._progress_bar.setRange(0, 0)
        self._set_running_state(running=True)
        self._thread.start()
        log.info("Job started: %d file(s).", len(jobs))

    def _cancel_job(self) -> None:
        if self._worker is not None:
            self._worker.cancel()
            self._cancel_btn.setEnabled(False)
            self._status_label.setText("Cancelling…")
            self.statusBar().showMessage("Cancellation requested…")
            log.info("Cancellation requested by user.")

    @pyqtSlot(int, int)
    def _on_progress(self, done: int, total: int) -> None:
        if total <= 0:
            return
        if self._progress_bar.maximum() == 0:
            self._progress_bar.setRange(0, 100)
        percent = max(0, min(100, int(done / total * 100)))
        self._progress_bar.setValue(percent)

    @pyqtSlot(str)
    def _on_status(self, message: str) -> None:
        self._status_label.setText(message)
        self.statusBar().showMessage(message)

    @pyqtSlot(str, str)
    def _on_worker_error(self, user_message: str, detail: str) -> None:
        log.error("Worker error detail:\n%s", detail)
        QMessageBox.critical(self, "Processing Error", user_message)

    @pyqtSlot(list, bool)
    def _on_worker_finished(
        self, completed: list[tuple[Path, Path]], was_cancelled: bool
    ) -> None:
        if was_cancelled:
            log.info("Job finished: cancelled. %d file(s) saved.", len(completed))
        else:
            log.info("Job finished: %d file(s) completed successfully.", len(completed))
            if completed:
                self._progress_bar.setRange(0, 100)
                self._progress_bar.setValue(100)

    @pyqtSlot()
    def _on_thread_done(self) -> None:
        self._worker = None
        self._thread = None
        self._set_running_state(running=False)

    def _sync_output_controls(self, *, running: bool) -> None:
        single_file_mode = len(self._input_paths) == 1
        enabled = (not running) and single_file_mode
        self._output_edit.setEnabled(enabled)
        self._output_browse_btn.setEnabled(enabled)

    def _set_running_state(self, *, running: bool) -> None:
        self._run_btn.setEnabled((not running) and bool(self._input_paths))
        self._cancel_btn.setEnabled(running)
        self._input_edit.setEnabled(not running)
        self._sync_output_controls(running=running)
        self._zwnj_check.setEnabled(not running)
        self.setAcceptDrops(not running)

        if not running:
            self._progress_bar.setRange(0, 100)
            if not self._input_paths:
                self._progress_bar.setValue(0)

    def _is_running(self) -> bool:
        return self._thread is not None and self._thread.isRunning()

    @pyqtSlot(str, str)
    def _append_log_line(self, html_line: str, _level: str) -> None:
        cursor = self._log_edit.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        cursor.insertHtml(html_line + "<br>")
        self._log_edit.ensureCursorVisible()

    def _copy_log(self) -> None:
        clipboard = QApplication.clipboard()
        if clipboard:
            clipboard.setText(self._log_edit.toPlainText())
            self.statusBar().showMessage("Log copied to clipboard.", 3000)

    def _clear_log(self) -> None:
        self._log_edit.clear()
        self.statusBar().showMessage("Log cleared.", 2000)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802
        mime: QMimeData = event.mimeData()
        if mime.hasUrls() and any(
            u.toLocalFile().lower().endswith((".xlsx", ".xlsm"))
            for u in mime.urls()
        ):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802
        if self._is_running():
            event.ignore()
            return
        paths = [
            u.toLocalFile()
            for u in event.mimeData().urls()
            if u.toLocalFile().lower().endswith((".xlsx", ".xlsm"))
        ]
        if paths:
            self._settings.setValue("last_directory", str(Path(paths[0]).parent))
            self._load_paths(paths)
            event.acceptProposedAction()

    def _show_about(self) -> None:
        AboutDialog(self).exec()

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802
        """Persist settings and gracefully shut down any active worker."""
        if self._is_running() and self._worker is not None:
            reply = QMessageBox.question(
                self,
                "Job Running",
                "A normalisation job is still running.\n\n"
                "Cancel it and wait for the worker to stop before quitting?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                event.ignore()
                return

            self._worker.cancel()
            self._cancel_btn.setEnabled(False)
            self._status_label.setText("Cancelling before close…")
            self.statusBar().showMessage("Waiting for worker to stop…")
            if self._thread is not None and not self._thread.wait(5000):
                QMessageBox.warning(
                    self,
                    "Still Cancelling",
                    "The workbook is still being processed. Close was cancelled "
                    "to avoid interrupting a file write. Try again after the "
                    "worker finishes.",
                )
                log.warning("Close ignored: worker thread did not stop within 5 s.")
                event.ignore()
                return

        self._persist_settings()
        pkg_logger = logging.getLogger("excel_normalizer")
        if self._log_handler in pkg_logger.handlers:
            pkg_logger.removeHandler(self._log_handler)
        super().closeEvent(event)


def main() -> None:
    """Launch the Qt application."""
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(SETTINGS_ORG)
    app.setApplicationVersion(APP_VERSION)
    app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
