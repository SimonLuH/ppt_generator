import sys
import traceback
import logging
from PyQt5.QtWidgets import QMessageBox, QApplication

def show_error(exc_type, exc_value, exc_traceback) -> None:
    """
    全局异常处理：消息框显示并记录日志。
    """
    error_message = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    QMessageBox.critical(None, "程序出错", error_message)
    sys.exit(1)
