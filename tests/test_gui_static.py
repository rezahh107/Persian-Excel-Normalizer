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


def test_gui_batch_mode_keeps_output_browse_disabled_after_completion() -> None:
    assert "def _sync_output_controls" in GUI_SOURCE
    assert "single_file_mode = len(self._input_paths) == 1" in GUI_SOURCE
    assert "self._output_browse_btn.setEnabled(enabled)" in GUI_SOURCE
    assert "self._sync_output_controls(running=running)" in GUI_SOURCE


def test_gui_worker_uses_method_callbacks_instead_of_trivial_lambdas() -> None:
    assert "cancel_check=self._is_cancel_requested" in GUI_SOURCE
    assert "on_progress=self.progress.emit" in GUI_SOURCE
    assert "lambda d, t" not in GUI_SOURCE
