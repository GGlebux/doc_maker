import os

from PyQt6.QtWidgets import (QTextEdit, QVBoxLayout,
                             QDialog, QScrollArea)


class LogViewer(QDialog):
    def __init__(self, log_file="app.log"):
        super().__init__()
        self.setWindowTitle("Логи приложения")
        self.setMinimumSize(600, 400)

        # Основной виджет с прокруткой
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)

        scroll.setWidget(self.text_edit)

        layout = QVBoxLayout()
        layout.addWidget(scroll)
        self.setLayout(layout)

        self.log_file = log_file
        self.refresh_logs()

    def refresh_logs(self):
        """Обновляет содержимое логов из файла"""
        if os.path.exists(self.log_file):
            with open(self.log_file, "r", encoding="utf-8") as f:
                self.text_edit.setPlainText(f.read())
                # Прокрутка вниз
                self.text_edit.verticalScrollBar().setValue(
                    self.text_edit.verticalScrollBar().maximum()
                )