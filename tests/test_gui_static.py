from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GUI_SOURCE = (ROOT / "normalize_excel_gui.py").read_text(encoding="utf-8")


def test_gui_save_filter_allows_xlsm() -> None:
    assert "Excel Macro-Enabled Workbook (*.xlsm)" in GUI_SOURCE
    assert "*.xlsx *.xlsm" in GUI_SOURCE


def test_gui_close_event_is_typed_and_does_not_terminate_thread() -> None:
    assert "QCloseEvent" in GUI_SOURCE
    assert "def closeEvent(self, event: QCloseEvent)" in GUI_SOURCE
    assert ".terminate(" not in GUI_SOURCE


def test_gui_close_event_requests_thread_quit_before_wait() -> None:
    assert "self._thread.quit()" in GUI_SOURCE
    assert "self._thread.wait(5000)" in GUI_SOURCE


def test_gui_batch_mode_keeps_output_browse_disabled_after_completion() -> None:
    assert "def _sync_output_controls" in GUI_SOURCE
    assert "single_file_mode = len(self._input_paths) == 1" in GUI_SOURCE
    assert "self._output_browse_btn.setEnabled(enabled)" in GUI_SOURCE
    assert "self._sync_output_controls(running=running)" in GUI_SOURCE


def test_gui_worker_uses_safe_cancellation_and_batch_progress() -> None:
    assert "NormalizationCancelled" in GUI_SOURCE
    assert "cancel_check=self._is_cancel_requested" in GUI_SOURCE
    assert "def _emit_batch_progress" in GUI_SOURCE
    assert "_PROGRESS_UNITS_PER_FILE" in GUI_SOURCE
    assert "self.progress.emit(1, 1)" not in GUI_SOURCE
    assert "lambda d, t" not in GUI_SOURCE


def test_gui_output_dialog_rejects_extension_mismatch() -> None:
    assert "Output Extension Mismatch" in GUI_SOURCE
    assert "output.suffix.lower() != default_suffix.lower()" in GUI_SOURCE
