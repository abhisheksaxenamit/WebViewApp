import sys
import webbrowser
from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import QMainWindow, QApplication, QSizePolicy, QSplitter, QWidget, QVBoxLayout, QComboBox, \
    QPushButton, QScrollArea, QLineEdit, QMenu


class WebViewApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.sheet_val = None
        self.copy_link = None
        self.searching = False
        self.action_bar_widget = None
        self.cheat_sheet = {}
        self.file_name = str(Path.home()) + "/Advertising.xlsx"
        self.webview = QWebEngineView()
        self.current_url = ""
        self.setWindowTitle("WebView App")

        # Get the different sheets of the Excel #
        self.sheets = self.get_sheets_from_excel()
        print(f'{self.sheets}')

        # Set the Web View Engine sheet #
        self.set_webview()

        # Set up the button_widget #
        self.buttons_widget = QWidget()
        self.buttons_widget.setMaximumSize(280, 16777215)
        self.buttons_layout = QVBoxLayout(self.buttons_widget)
        self.buttons_layout.setContentsMargins(10, 10, 10, 10)
        self.buttons_widget.setStyleSheet("background-color: black;")

        # Gather Data from the Excel to the dict
        self.excel_data_to_dict()
        # Set Search tab #
        self.search_bar = QLineEdit()
        self.search_bar.setClearButtonEnabled(True)
        self.search_bar.returnPressed.connect(self.search)

        search_icon = QIcon.fromTheme("edit-find")
        search_action = self.search_bar.addAction(search_icon, QLineEdit.LeadingPosition)
        search_action.triggered.connect(self.search)
        self.search_bar_layout = QVBoxLayout()
        self.search_bar_layout.addWidget(self.search_bar)

        # Set Dropdown Menu for different Sheets #
        self.sheet_dropdown = QComboBox(self)
        self.sheet_dropdown.setMaximumSize(260, 240)
        self.create_dropdown_menu()

        # Adding Buttons to the widget #
        print(f"searching {self.searching} Calling main -> set_buttons_widget")
        self.set_buttons_widget()

        # Make the final APP #
        self.set_action_bar()
        self.create_app()

    def search(self):
        search_text = self.search_bar.text()
        print(f"searching {self.searching} : {search_text.strip()} search()")
        if search_text.strip() != "":
            # print(f"searching {self.searching} search()")
            self.searching = True
        else:
            # print(f"searching {self.searching} search()")
            self.searching = False
        matches = []
        for key in self.cheat_sheet.keys():
            if search_text.lower() in key.lower():
                # print(f"{key}")
                matches.append((key, self.cheat_sheet[key]))
        self.create_buttons_from_search(matches)

        # Clear the search bar after the search
        self.search_bar.clear()

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
        self.webview.customContextMenuRequested.connect(self.show_context_menu)

    def create_buttons_from_search(self, matches):
        self.clear_button_widget()
        for match in matches:
            for link in match[1].splitlines():
                button = QPushButton(match[0])
                button.setStyleSheet("background-color: white; color: black; font: Bold")
                button.setFixedSize(240, 50)
                button.clicked.connect(lambda *args, url=link: self.open_web_gui(url))
                self.buttons_layout.addWidget(button)
        self.set_buttons_widget()

    def clear_button_widget(self):
        self.buttons_layout = self.buttons_widget.layout()
        print(f"Buttons in the widget: {self.buttons_layout.count()}")
        while self.buttons_layout.count():
            item = self.buttons_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def handle_dropdown_selection(self, selected_value):
        print(f'Calling handle_dropdown_selection with value {selected_value}')
        self.sheet_val = selected_value
        # Clear the buttons layout
        self.clear_button_widget()
        self.set_buttons_widget()
        self.buttons_widget.update()

    def set_buttons_widget(self):
        # Get the Excel sheet data
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

    def get_sheets_from_excel(self):
        try:
            return pd.read_excel(self.file_name, sheet_name=None).keys()
        except KeyError:
            print("Expected column headers not found")
            sys.exit(1)
        except TypeError as e:
            print(f"Type Error {e}")
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

    def read_excel_file(self, sheet):
        # print(f'Sheet: {self.file_name}')
        dataframe = pd.read_excel(self.file_name, sheet_name=sheet)
        return dataframe

    def excel_data_to_dict(self):
        for page in self.sheets:
            excel_data = pd.read_excel(self.file_name, sheet_name=page)
            for idx, row in excel_data.iterrows():
                key = row[page]
                if key in self.cheat_sheet:
                    # Handle duplicate keys
                    count = 1
                    new_key = f"{key}_{count}"
                    while new_key in self.cheat_sheet:
                        count += 1
                        new_key = f"{key}_{count}"
                    key = new_key
                self.cheat_sheet[key] = row['LINK']
        # print(f'{self.cheat_sheet}')
        print(f'Total links: {len(self.cheat_sheet)}')

    def create_dropdown_menu(self):
        self.sheet_dropdown.addItems(self.sheets)
        self.sheet_dropdown.setStyleSheet("background-color: black; color: white;")
        self.sheet_val = self.sheet_dropdown.currentText()
        self.sheet_dropdown.currentText()
        print(f"searching {self.searching} create_dropdown_menu -> handle_dropdown_selection")
        self.sheet_dropdown.currentTextChanged.connect(self.handle_dropdown_selection)

    def show_context_menu(self, point):
        context_menu = QMenu(self)
        # Copy Link #
        copy_link_action = QAction("Copy Link", self)
        copy_link_action.triggered.connect(self.copy_link)
        context_menu.addAction(copy_link_action)

        # Open Web Browser #
        open_in_browser = QAction("Open in Web Browser", self)
        open_in_browser.triggered.connect(self.open_in_web_browser)
        context_menu.addAction(open_in_browser)

        context_menu.exec_(self.mapToGlobal(point))

    def open_in_web_browser(self):
        print(f"open_in_web_browser {self.current_url}")
        if self.current_url is not None:
            webbrowser.open(self.current_url)

    def set_action_bar(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(self.buttons_widget)
        scroll_area.setFixedWidth(280)

        action_bar = QVBoxLayout()
        action_bar.addWidget(self.search_bar)
        action_bar.addWidget(self.sheet_dropdown)
        action_bar.addWidget(scroll_area)

        self.action_bar_widget = QWidget()
        self.action_bar_widget.setFixedWidth(290)
        self.action_bar_widget.setLayout(action_bar)

    def create_app(self):
        splitter = QSplitter()
        splitter.addWidget(self.webview)
        splitter.addWidget(self.action_bar_widget)
        self.setCentralWidget(splitter)
        self.showMaximized()


if __name__ == '__main__':
    app = QApplication([])
    window = WebViewApp()
    window.show()
    app.exec()
