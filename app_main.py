# -*- coding: utf-8 -*-
r"""專案架構與環境:
  完整使用說明請見 README.md
  所有輸入/設定預設放在 user_data/
"""

from __future__ import annotations

import sys
import traceback

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QMessageBox

from app_mainwindow import MainWindow


def main() -> int:

    smoke_test = "--smoke-test" in sys.argv[1:]

    qt_args = [sys.argv[0]] + [arg for arg in sys.argv[1:] if arg != "--smoke-test"]

    app = QApplication(qt_args)

    try:
        w = MainWindow()
        w.show()

        if smoke_test:
            QTimer.singleShot(400, app.quit)

        return app.exec()

    except Exception:
        error_message = f"發生未預期的嚴重錯誤，應用程式即將關閉。\n\n錯誤資訊：\n{traceback.format_exc()}"
        QMessageBox.critical(None, "應用程式錯誤", error_message)
        return 1

if __name__ == "__main__":

    raise SystemExit(main())
