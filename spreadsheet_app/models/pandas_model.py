import pandas as pd
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex

class PandasModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame, parent=None):
        super().__init__(parent)
        self._df = df

    @property
    def df(self):
        return self._df

    def update_df(self, df: pd.DataFrame):
        self.beginResetModel()
        self._df = df
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        r, c = index.row(), index.column()
        value = self._df.iat[r, c]

        if role in (Qt.DisplayRole, Qt.EditRole):
            if pd.api.types.is_datetime64_any_dtype(self._df.dtypes[c]):
                if pd.isna(value):
                    return ""
                try:
                    return pd.to_datetime(value).strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    return str(value) if value is not None else ""
            return "" if pd.isna(value) else str(value)
        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable

    def setData(self, index, value, role=Qt.EditRole):
        if role != Qt.EditRole or not index.isValid():
            return False
        r, c = index.row(), index.column()
        col = self._df.columns[c]
        dtype = self._df[col].dtype

        try:
            if pd.api.types.is_integer_dtype(dtype):
                self._df.at[r, col] = int(value) if value != "" else pd.NA
            elif pd.api.types.is_float_dtype(dtype):
                self._df.at[r, col] = float(value) if value != "" else pd.NA
            elif pd.api.types.is_datetime64_any_dtype(dtype):
                self._df.at[r, col] = pd.to_datetime(value) if value != "" else pd.NaT
            else:
                self._df.at[r, col] = value
        except Exception:
            return False

        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        return True

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section]) if section < len(self._df.columns) else ""
        else:
            return str(section)