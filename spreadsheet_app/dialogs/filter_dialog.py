from PySide6.QtWidgets import (
    QWidget, 
    QVBoxLayout, 
    QDialog, 
    QLabel, 
    QLineEdit, 
    QFormLayout, 
    QDialogButtonBox, 
    QScrollArea
)

class FilterDialog(QDialog):
    def __init__(self, columns, existing_filters=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Filters")
        self.setMinimumSize(420, 480)

        layout = QVBoxLayout(self)

        info = QLabel("Set filters for the columns (empty = no filter).")
        info.setWordWrap(True)
        layout.addWidget(info)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        form = QFormLayout(inner)
        self.inputs = {}

        for col in columns:
            le = QLineEdit()
            if existing_filters and col in existing_filters:
                le.setText(str(existing_filters[col]))
            self.inputs[col] = le
            form.addRow(QLabel(col + ":"), le)

        scroll.setWidget(inner)
        layout.addWidget(scroll)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_filters(self):
        """
        Opens the dialog and returns a dict {column: text} with only non-empty values.
        Returns None if the user cancels.
        """
        if self.exec() != QDialog.Accepted:
            return None
        result = {}
        for col, le in self.inputs.items():
            txt = le.text().strip()
            if txt:
                result[col] = txt
        return result