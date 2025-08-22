from PySide6.QtWidgets import (
    QWidget, 
    QVBoxLayout, 
    QHBoxLayout,
    QDialog, 
    QLabel, 
    QLineEdit, 
    QDialogButtonBox,
    QRadioButton, 
    QGridLayout
)

class InsertRowsDialog(QDialog):
    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add Row")
        self.setMinimumSize(420, 520)
        self.values = None
        self.pos = None
        self.before = False

        layout = QVBoxLayout(self)

        gridw = QWidget()
        grid = QGridLayout(gridw)
        self.inputs = {}
        for i, col in enumerate(columns):
            grid.addWidget(QLabel(col), i, 0)
            le = QLineEdit()
            self.inputs[col] = le
            grid.addWidget(le, i, 1)

        layout.addWidget(gridw)

        pos_row = QHBoxLayout()
        pos_row.addWidget(QLabel("Insert row before/after which index?"))
        self.pos_edit = QLineEdit()
        self.pos_edit.setPlaceholderText("E.g.: 0, 1, 2 ...")
        pos_row.addWidget(self.pos_edit)
        layout.addLayout(pos_row)

        rb_row = QHBoxLayout()
        self.rb_before = QRadioButton("Before")
        self.rb_after = QRadioButton("After")
        self.rb_after.setChecked(True)
        rb_row.addWidget(self.rb_before)
        rb_row.addWidget(self.rb_after)
        layout.addLayout(rb_row)

        self.err = QLabel("")
        self.err.setStyleSheet("color:#ff5555; font-weight:bold;")
        layout.addWidget(self.err)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._ok)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _ok(self):
        pos_txt = self.pos_edit.text().strip()
        if not pos_txt.isdigit():
            self.err.setText("Index must be a valid integer.")
            return
        self.pos = int(pos_txt)
        self.before = self.rb_before.isChecked()

        vals = {col: le.text() for col, le in self.inputs.items()}
        self.values = vals
        self.accept()