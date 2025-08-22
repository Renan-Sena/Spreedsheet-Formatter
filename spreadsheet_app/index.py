import sys
from PySide6.QtWidgets import ( QApplication )

from core.spreadsheet import SpreadSheetApp 

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # 🌙 Dark Mode global
    app.setStyleSheet("""
QMainWindow { background-color: #2b2b2b; }
QPushButton {
    background-color: #007acc;
    color: white;
    border-radius: 6px;
    padding: 6px 12px;
    font-weight: bold;
}
QPushButton:hover { background-color: #3399ff; }
QTableView {
    background-color: #1e1e1e;
    color: #eeeeee;
    gridline-color: #444444;
    selection-background-color: #3399ff;
    selection-color: white;
    alternate-background-color: #2e2e2e;
}
QHeaderView::section {
    background-color: #3c3c3c;
    color: white;
    padding: 4px;
    border: 1px solid #444;
}
QLineEdit {
    background-color: #3c3c3c;
    color: white;
    border: 1px solid #555;
    border-radius: 4px;
    padding: 4px;
}
QLabel { color: #eeeeee; }
QStatusBar { background-color: #333333; color: #bbbbbb; }
QDialog { background-color: #2b2b2b; }
""")

    win = SpreadSheetApp()
    win.show()
    sys.exit(app.exec())
