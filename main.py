from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import QMainWindow, QApplication, QSizePolicy, QSplitter


class WebViewApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.file_name = str(Path.home()) + "/Advertising.xlsx"
        self.webview = QWebEngineView()
        self.current_url = ""
        self.setWindowTitle("WebView App")
        self.set_webview()
        self.create_app()

    def set_webview(self):
        self.webview.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.webview.page().profile().setHttpUserAgent(
            "Mozilla/5.0 AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3")
        self.webview.page().profile().clearHttpCache()
        self.webview.page().setHtml(
            "<html><head><style>body { background-image: url("
            "'https://cdn.dribbble.com/users/642843/screenshots/2196696/attachments/406228"
            "/romain_trystram_wallpaper_wetransfer_dribbble.jpg'); background-repeat: no-repeat; background-position: "
            "center; background-size: cover; }</style></head><body></body></html>")
        self.webview.setContextMenuPolicy(Qt.CustomContextMenu)
        # self.webview.customContextMenuRequested.connect(self.show_context_menu)

    def create_app(self):
        splitter = QSplitter()
        splitter.addWidget(self.webview)
        self.setCentralWidget(splitter)
        self.showMaximized()


if __name__ == '__main__':
    app = QApplication([])
    window = WebViewApp()
    window.show()
    app.exec()
