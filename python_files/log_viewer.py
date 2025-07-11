import os

from PyQt6.QtCore import QRegularExpression, QTimer
from PyQt6.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor, QFont
from PyQt6.QtWidgets import (QTextEdit, QVBoxLayout,
                             QDialog, QFileDialog, QHBoxLayout, QPushButton)


class LogHighlighter(QSyntaxHighlighter):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.highlighting_rules = []

        # Ошибки (красный)
        error_format = QTextCharFormat()
        error_format.setForeground(QColor(255, 0, 0))
        error_format.setFontWeight(QFont.Weight.Bold)
        self.highlighting_rules.append(
            (QRegularExpression(r"ERROR.*|Traceback.*"), error_format)
        )

        # Предупреждения (жёлтый)
        warning_format = QTextCharFormat()
        warning_format.setForeground(QColor(200, 100, 0))
        self.highlighting_rules.append(
            (QRegularExpression(r"WARNING.*"), warning_format)
        )

        # Инфо (зеленый)
        error_format = QTextCharFormat()
        error_format.setForeground(QColor(102, 204, 0))
        error_format.setFontWeight(QFont.Weight.Bold)
        self.highlighting_rules.append(
            (QRegularExpression(r"INFO.*"), error_format)
        )

        # Даты (синий)
        date_format = QTextCharFormat()
        date_format.setForeground(QColor(153, 102, 204))
        self.highlighting_rules.append(
            (QRegularExpression(r"\d{2}:\d{2}:\d{2} \d{2}-\d{2}-\d{4}"), date_format)
        )

        # Форматы
        traceback_format = QTextCharFormat()
        traceback_format.setForeground(QColor(200, 50, 50))  # Красноватый
        traceback_format.setFontWeight(QFont.Weight.Bold)

        file_path_format = QTextCharFormat()
        file_path_format.setForeground(QColor(42, 130, 218))  # Синий как в VS Code
        file_path_format.setFontItalic(True)

        line_num_format = QTextCharFormat()
        line_num_format.setForeground(QColor(106, 135, 89))  # Зеленоватый

        error_type_format = QTextCharFormat()
        error_type_format.setForeground(QColor(255, 100, 100))
        error_type_format.setFontWeight(QFont.Weight.Bold)

        code_block_format = QTextCharFormat()
        code_block_format.setBackground(QColor(40, 44, 52))  # Темный фон как в VS Code
        code_block_format.setForeground(QColor(171, 178, 191))  # Светло-серый текст
        code_block_format.setFontFamily("Consolas")
        code_block_format.setFontPointSize(10)

        # Правила для Traceback
        self.highlighting_rules.extend([
            (QRegularExpression(r"Traceback \(most recent call last\):"), traceback_format),
            (QRegularExpression(r'File "(.+?)"'), file_path_format),  # Путь к файлу
            (QRegularExpression(r', line \d+'), line_num_format),  # Номер строки
            (QRegularExpression(r'^\w+Error:.*$'), error_type_format),  # Тип ошибки
            (QRegularExpression(r'^    .*$'), code_block_format)  # Код с отступом
        ])



    def highlightBlock(self, text):
        for pattern, fmt in self.highlighting_rules:
            match_iterator = pattern.globalMatch(text)
            while match_iterator.hasNext():
                match = match_iterator.next()
                self.setFormat(match.capturedStart(), match.capturedLength(), fmt)


class LogViewer(QDialog):
    def __init__(self, log_file="app.log"):
        super().__init__()
        self.setWindowTitle("Логи приложения")
        self.setMinimumSize(800, 600)

        # Основные элементы
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.text_edit.setFontFamily("Courier New")
        self.text_edit.setFontPointSize(10)

        # Подсветка синтаксиса
        self.highlighter = LogHighlighter(self.text_edit.document())

        # Кнопки управления
        self.btn_refresh = QPushButton("Обновить")
        self.btn_clear = QPushButton("Очистить логи")
        self.btn_save = QPushButton("Сохранить как...")

        # Настройка layout
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.btn_refresh)
        button_layout.addWidget(self.btn_clear)
        button_layout.addWidget(self.btn_save)

        main_layout = QVBoxLayout()
        main_layout.addLayout(button_layout)
        main_layout.addWidget(self.text_edit)
        self.setLayout(main_layout)

        # Подключение сигналов
        self.btn_refresh.clicked.connect(self.refresh_logs)
        self.btn_clear.clicked.connect(self.clear_logs)
        self.btn_save.clicked.connect(self.save_logs)

        # Автообновление
        self.timer = QTimer()
        self.timer.timeout.connect(self.refresh_logs)
        self.timer.start(5000)  # Обновление каждые 5 сек

        self.log_file = log_file
        self.refresh_logs()

    def refresh_logs(self):
        """Загружает логи из файла"""
        if os.path.exists(self.log_file):
            with open(self.log_file, "r", encoding="utf-8") as f:
                self.text_edit.setPlainText(f.read())
                self.scroll_to_bottom()

    def scroll_to_bottom(self):
        """Прокрутка вниз"""
        scrollbar = self.text_edit.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def clear_logs(self):
        """Очищает файл логов"""
        with open(self.log_file, "w", encoding="utf-8") as f:
            f.write("")
        self.refresh_logs()

    def save_logs(self):
        """Сохраняет логи в выбранный файл"""
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Сохранить логи", "", "Лог файлы (*.log);;Все файлы (*)"
        )
        if file_name:
            with open(file_name, "w", encoding="utf-8") as f:
                f.write(self.text_edit.toPlainText())
