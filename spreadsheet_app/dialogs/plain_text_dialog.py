from PySide6.QtWidgets import (
    QVBoxLayout,
    QDialog, 
    QLabel, 
    QLineEdit, 
    QDialogButtonBox
)

class PlainTextDialog(QDialog):
    def __init__(self, title, label, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.value = None
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(label))
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
        v = self.le.text().strip()
        if v == "":
            self.err.setText("Value cannot be empty.")
            return
        self.value = v
        self.accept()