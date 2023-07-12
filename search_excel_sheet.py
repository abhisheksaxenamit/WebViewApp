from pathlib import Path

import pandas as pd
from PySide6.QtWidgets import QApplication, QMainWindow, QLineEdit, QVBoxLayout, QWidget


class SearchView(QMainWindow):
    def __init__(self):
        super().__init__()
        self.file_name = str(Path.home()) + "/Advertising.xlsx"

        sheets = pd.read_excel(self.file_name, sheet_name=None).keys()
        self.cheat_sheet = {}
        for page in sheets:
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

        self.search_bar = QLineEdit()
        self.search_bar.returnPressed.connect(self.search)
        self.search_bar_layout = QVBoxLayout()

        # self.search_bar_layout.addWidget(self.search_bar)

        self.search_bar_layout.addWidget(self.search_bar)
        central_widget = QWidget()
        central_widget.setLayout(self.search_bar_layout)

        self.setCentralWidget(central_widget)
        self.showMaximized()

    def search(self):
        search_text = self.search_bar.text()
        matches = []

        for key in self.cheat_sheet.keys():
            if search_text.lower() in key.lower():
                print(f"{key}")
                matches.append((key, self.cheat_sheet[key]))

        if matches:
            for match in matches:
                print(f"Match: {match[0]}, Link: {match[1]}")
        else:
            print("No matches found.")

        # Clear the search bar after the search
        self.search_bar.clear()



if __name__ == '__main__':
    app = QApplication([])
    window = SearchView()
    window.show()
    app.exec()