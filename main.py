import sys
from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt, QUrl
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import QMainWindow, QApplication, QSizePolicy, QSplitter, QWidget, QVBoxLayout, QComboBox, \
    QPushButton, QScrollArea


class WebViewApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.file_name = str(Path.home()) + "/Advertising.xlsx"
        self.webview = QWebEngineView()
        self.current_url = ""
        self.setWindowTitle("WebView App")

        # Get the different sheets of the excel #
        self.sheets = self.get_sheets_from_excel()
        print(f'{self.sheets}')

        # Temp Sheet 1 #
        self.sheet_val = "BANKS"

        # Set the Web View Engine sheet #
        self.set_webview()

        # Setup the button_widget #
        self.buttons_widget = QWidget()
        self.buttons_widget.setMaximumSize(280, 16777215)
        self.buttons_layout = QVBoxLayout(self.buttons_widget)
        self.buttons_layout.setContentsMargins(10, 10, 10, 10)
        self.buttons_widget.setStyleSheet("background-color: black;")

        # Set Dropdown Menu for different Sheets #
        self.sheet_dropdown = QComboBox(self)
        self.sheet_dropdown.setMaximumSize(260, 240)
        self.create_dropdown_menu()

        # Adding Buttons to the widget #
        self.set_buttons_widget()

        # Make the final APP #
        self.set_action_bar()
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

    def set_buttons_widget(self):
        # Get the excel sheet data
        excel_data = self.read_excel_file(self.sheet_val)
        excel_data = excel_data.sort_values(self.sheet_val)
        button_list = excel_data[self.sheet_val].tolist()
        link_list = excel_data["LINK"].tolist()
        print(f'{button_list}')

        for idx, button_name in enumerate(button_list):
            for link in link_list[idx].splitlines():
                button = QPushButton(button_name)
                button.setStyleSheet("background-color: white; color: black; font: Bold")
                button.setFixedSize(240, 50)
                button.clicked.connect(lambda *args, url=link: self.open_web_gui(url))
                self.buttons_layout.addWidget(button)

    def create_app(self):
        splitter = QSplitter()
        splitter.addWidget(self.webview)
        splitter.addWidget(self.action_bar_widget)
        self.setCentralWidget(splitter)
        self.showMaximized()

    def get_sheets_from_excel(self):
        try:
            return pd.read_excel(self.file_name, sheet_name=None).keys()
        except KeyError as e:
            print("Expected column headers not found")
            sys.exit(1)
        except TypeError as e:
            print("Type Error")
            sys.exit(1)
        except FileNotFoundError as e:
            print("Excel file not found " + str(e))
            sys.exit(1)

    def open_web_gui(self, url):
        self.current_url = url
        # url = "https://wetransfer.com/"
        print(f"open_web_gui : {self.current_url}")
        try:
            self.webview.load(QUrl(url))
            self.webview.show()
        except Exception as e:
            print(e)

    def set_action_bar(self):
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.buttons_widget)
        self.scroll_area.setFixedWidth(280)

        self.action_bar = QVBoxLayout()
        self.action_bar.addWidget(self.sheet_dropdown)
        self.action_bar.addWidget(self.scroll_area)

        self.action_bar_widget = QWidget()
        self.action_bar_widget.setFixedWidth(290)
        self.action_bar_widget.setLayout(self.action_bar)

    def read_excel_file(self, sheet):
        print(f'Sheet: {self.file_name}')
        dataframe = pd.read_excel(self.file_name, sheet_name=sheet)
        return dataframe

    def create_dropdown_menu(self):
        self.sheet_dropdown.addItems(self.sheets)
        self.sheet_dropdown.setStyleSheet("background-color: black; color: white;")
        self.sheet_val = self.sheet_dropdown.currentText()
        self.sheet_dropdown.currentText()
        self.sheet_dropdown.currentTextChanged.connect(self.handle_dropdown_selection)

    def handle_dropdown_selection(self, selected_value):
        print(f'Calling handle_dropdown_selection with value {selected_value}')
        self.sheet_val = selected_value
        # Clear the buttons layout
        self.buttons_layout = self.buttons_widget.layout()
        print(f"Buttons in the widget: {self.buttons_layout.count()}")
        while self.buttons_layout.count():
            item = self.buttons_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.set_buttons_widget()
        self.buttons_widget.update()


if __name__ == '__main__':
    app = QApplication([])
    window = WebViewApp()
    window.show()
    app.exec()
