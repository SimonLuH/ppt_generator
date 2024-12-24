import sys
from PyQt5.QtWidgets import QApplication
from client_gui.utils.logger import configure_logging
from client_gui.utils.exception_handler import show_error
from client_gui.gui.main_window import PPTClientGUI

def main():
    """
    GUI入口。
    """
    configure_logging()
    sys.excepthook = show_error
    app = QApplication(sys.argv)
    window = PPTClientGUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
