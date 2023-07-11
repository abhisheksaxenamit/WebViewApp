from pathlib import Path

from PySide6.QtWidgets import QMainWindow, QApplication


class WebViewApp(QMainWindow):

    file_name = str(Path.home())+"/Advertising.xlsx"

    def __init__(self):
        super().__init__()
        self.current_url = ""
        self.setWindowTitle("WebView App")


if __name__ == '__main__':
    app = QApplication([])
    window = WebViewApp()
    window.show()
    app.exec()