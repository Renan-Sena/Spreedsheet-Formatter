from PySide6.QtWidgets import (
    QVBoxLayout,
    QDialog, 
    QLabel, 
    QLineEdit, 
    QDialogButtonBox
)

class RemoveRowsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Remove Row")
        self.idx = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Enter the index of the row to remove:"))
        self.le = QLineEdit()
        layout.addWidget(self.le)
        self.err = QLabel("")
        self.err.setStyleSheet("color:#ff5555; font-weight:bold;")
        layout.addWidget(self.err)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._ok)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _ok(self):
        txt = self.le.text().strip()
        if not txt.isdigit():
            self.err.setText("Invalid index.")
            return
        self.idx = int(txt)
        self.accept()