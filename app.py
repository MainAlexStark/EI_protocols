from email import header
import sys

from PyQt6.QtCore import QSize, Qt
from PyQt6.QtCore import QAbstractTableModel, Qt, QVariant
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QPushButton,
    QVBoxLayout, QLineEdit, QLabel, QFileDialog, QWidget, QCheckBox, QDialog, QHBoxLayout
)
from PyQt6.QtWidgets import QTableView
import pandas as pd


from EI_protocols_utils.utils.settings import settings
from utils.user_info import save_paths, load_paths
data = load_paths(filename=settings.user_info_path)

class SettingsDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Настройки")
        self.setFixedSize(QSize(300, 200))
        layout = QVBoxLayout(self)
        
        self.app_settings = load_paths(filename=settings.app_settings_path)
        for key, value in self.app_settings.items():
            checkbox = QCheckBox(f"{key}")
            checkbox.setChecked(value)
            checkbox.stateChanged.connect(self.update_setting(key))
            layout.addWidget(checkbox)
            
    def update_setting(self, key):
        def handler(state):
            self.app_settings[key] = bool(state)
            save_paths(self.app_settings, filename=settings.app_settings_path)
        return handler
    
    
    
class PandasModel(QAbstractTableModel):
    """Модель, адаптирующая pandas.DataFrame под Qt TableView."""
    
    def __init__(self, df):
        super().__init__()
        self._df = df

    def rowCount(self, parent=None):
        return len(self._df.index)

    def columnCount(self, parent=None):
        return len(self._df.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return QVariant()

        if role == Qt.ItemDataRole.DisplayRole:
            value = self._df.iat[index.row(), index.column()]
            return str(value)

        return QVariant()

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return QVariant()

        if orientation == Qt.Orientation.Horizontal:
            return str(self._df.columns[section])
        else:
            return str(self._df.index[section])

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("ЕИ Протоколы")

        # Создаём центральный виджетcd 
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # Создаём layout для центрального виджета
        main_layout = QHBoxLayout()
        central_widget.setLayout(main_layout)

        # Секция создания протокола
        create_protocol_layout = QVBoxLayout()

        self.journal_path_button = QPushButton("Путь к журналу")
        self.journal_path_button.setFixedSize(200, 40)
        create_protocol_layout.addWidget(self.journal_path_button)
        self.journal_path_button.clicked.connect(self.select_journal_path)
        
        self.journal_path_label = QLabel(data.get("journal_path", ""))
        create_protocol_layout.addWidget(self.journal_path_label)
        
        self.protocols_path_button = QPushButton("Путь к протоколам")
        self.protocols_path_button.setFixedSize(200, 40)
        create_protocol_layout.addWidget(self.protocols_path_button)
        self.protocols_path_button.clicked.connect(self.select_protocols_path)

        self.protocols_path_label = QLabel(data.get("protocols_path", ""))
        create_protocol_layout.addWidget(self.protocols_path_label)
        
        self.from_row_input = QLineEdit()
        self.from_row_input.setPlaceholderText("С какой строки")
        self.from_row_input.setFixedSize(200, 40)
        create_protocol_layout.addWidget(self.from_row_input)
        
        self.to_row_input = QLineEdit()
        self.to_row_input.setPlaceholderText("По какую строку")
        self.to_row_input.setFixedSize(200, 40)
        create_protocol_layout.addWidget(self.to_row_input)

        self.create_protocol_button = QPushButton("Создать протокол")
        self.create_protocol_button.setFixedSize(200, 40)
        create_protocol_layout.addWidget(self.create_protocol_button)

        self.settings_button = QPushButton("Настройки")
        self.settings_button.setFixedSize(200, 40)
        self.settings_button.clicked.connect(self.open_settings)
        create_protocol_layout.addWidget(self.settings_button)

        main_layout.addLayout(create_protocol_layout)
        
        
        ### Layout с таблицей
        self.table_layout = QVBoxLayout()
        
        self.table = QTableView()
        if data.get("journal_path", ""): self.load_excel_to_table(data["journal_path"])
        self.table_layout.addWidget(self.table)
        
        main_layout.addLayout(self.table_layout)


    ##################### Методы работы с таблицей
    def load_excel_to_table(self, path):
        try:
            df = pd.read_excel(path)

            # создаём модель
            self.model = PandasModel(df)
            self.table.setModel(self.model)
            self.table.selectionModel().selectionChanged.connect(self.on_selection_changed)

            # подгоняем колонки
            self.table.horizontalHeader().setStretchLastSection(True)
            self.table.resizeColumnsToContents()

            # Восстанавливаем сохранённые ширины
            settings_data = load_paths(filename=settings.app_settings_path)
            widths = settings_data.get("column_widths", {})

            for col_index in range(len(df.columns)):
                col_name = df.columns[col_index]
                if col_name in widths:
                    self.table.setColumnWidth(col_index, widths[col_name])

            # Слушаем изменения ширины пользователя
            header = self.table.horizontalHeader()
            header.sectionResized.connect(self.save_column_width)

        except Exception as e:
            print("Ошибка загрузки Excel:", e)
            
    def save_column_width(self, index, old_size, new_size):
        # Создаём словарь если нет
        settings_data = load_paths(filename=settings.app_settings_path)
        if "column_widths" not in settings_data:
            settings_data["column_widths"] = {}

        # Берём имя столбца по индексу
        column_name = self.model._df.columns[index]
        settings_data["column_widths"][column_name] = new_size

        save_paths(settings_data, filename=settings.app_settings_path)
        
    def get_selected_rows(self):
        indexes = self.table.selectionModel().selectedIndexes()
        rows = sorted(set(index.row() for index in indexes))
        return rows
    
    def on_selection_changed(self, selected, deselected):
        rows = self.get_selected_rows()
        self.from_row_input.setText(str(min(rows)+2) if rows else "")
        self.to_row_input.setText(str(max(rows)+2) if rows else "")
    ################


    def select_journal_path(self):
        file, _ = QFileDialog.getOpenFileName(self, caption="Выберите файл", filter="Excel файлы (*.xlsx)")
        if file:
            self.journal_path_label.setText(file)
            data["journal_path"] = file
            save_paths(data, filename=settings.user_info_path)
            self.load_excel_to_table(file)
            
    def select_protocols_path(self):
        directory = QFileDialog.getExistingDirectory(self, caption="Выберите папку для протоколов")
        if directory:
            self.protocols_path_label.setText(directory)
            data["protocols_path"] = directory
            save_paths(data, filename=settings.user_info_path)
            
    def open_settings(self):
        self.settings_dialog = SettingsDialog()
        self.settings_dialog.show()

app = QApplication(sys.argv)
window = MainWindow()
window.show()
app.exec()
