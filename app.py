import sys

from PyQt6.QtCore import QSize, Qt
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QPushButton,
    QVBoxLayout, QLineEdit, QLabel, QFileDialog, QWidget
)

from EI_protocols_utils.utils.settings import settings
from utils.user_info import save_paths, load_paths
data = load_paths(filename=settings.user_info_path)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("ЕИ Протоколы")

        # Создаём центральный виджет
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # Создаём layout для центрального виджета
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)

        # Секция создания протокола
        create_protocol_layout = QVBoxLayout()

        self.journal_path_button = QPushButton("Путь к журналу")
        create_protocol_layout.addWidget(self.journal_path_button)
        self.journal_path_button.clicked.connect(self.select_journal_path)
        
        self.journal_path_label = QLabel("")
        create_protocol_layout.addWidget(self.journal_path_label)
        
        self.protocols_path_button = QPushButton("Путь к протоколам")
        create_protocol_layout.addWidget(self.protocols_path_button)
        self.protocols_path_button.clicked.connect(self.select_protocols_path)

        self.protocols_path_label = QLabel()
        create_protocol_layout.addWidget(self.protocols_path_label)

        self.create_protocol_button = QPushButton("Создать протокол")
        create_protocol_layout.addWidget(self.create_protocol_button)

        self.settings_button = QPushButton("Настройки")
        create_protocol_layout.addWidget(self.settings_button)

        main_layout.addLayout(create_protocol_layout)

    def select_journal_path(self):
        file, _ = QFileDialog.getOpenFileName(self, caption="Выберите файл", filter="Excel файлы (*.xlsx)")
        if file:
            self.journal_path_label.setText(file)
            
    def select_protocols_path(self):
        directory = QFileDialog.getExistingDirectory(self, caption="Выберите папку для протоколов")
        if directory:
            self.protocols_path_label.setText(directory)


app = QApplication(sys.argv)
window = MainWindow()
window.show()
app.exec()
