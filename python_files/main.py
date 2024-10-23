import sys

from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow, QButtonGroup
)

from design.design import Ui_MainWindow
from python_files.clear import Cleaner
from python_files.data import Data
from python_files.excel import Excel
from python_files.word import Word


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("приЁмка")
        self.setupUi(self)
        self.updateUi()

        #ToDo: Создать диалоговые окна и сообщения statusBar

        # Классы для заполнения Word, Excel, Cleaner (очистка формы), Data (все данные)
        self.data = Data(self)
        self.excel = Excel(self, self.data)
        self.word = Word(self, self.data)
        self.cleaner = Cleaner(self)

        self.rating_button.clicked.connect(self.first_path)
        self.common_button.clicked.connect(self.second_path)
        self.aic_button.clicked.connect(self.third_path)
        self.stream_button.clicked.connect(self.fourth_path)

        self.fill_word_button.clicked.connect(self.word.start_up)
        self.fill_excel_button.clicked.connect(self.excel.start_up)
        self.clear_button.clicked.connect(self.cleaner.clear_form)
        self.exit_button.clicked.connect(sys.exit)

    def updateUi(self):
        """Добавляем недостающую логику на интерфейс"""
        self.base_education = QButtonGroup(self)
        self.base_education.addButton(self.nine)
        self.base_education.addButton(self.eleven)

        self.form_education = QButtonGroup(self)
        self.form_education.addButton(self.full_time)
        self.form_education.addButton(self.correspondence)

        self.age = QButtonGroup(self)
        self.age.addButton(self.minor)
        self.age.addButton(self.adult)

        self.more_group = QButtonGroup(self)
        self.more_group.setExclusive(False)
        self.more_group.addButton(self.svo)
        self.more_group.addButton(self.target_direction)

        self.finance = QButtonGroup(self)
        self.finance.addButton(self.budget)
        self.finance.addButton(self.commerce)

        self.stream = QButtonGroup(self)
        self.stream.addButton(self.first_stream)
        self.stream.addButton(self.second_stream)

        self.statement = QButtonGroup(self)
        self.statement.addButton(self.prof)
        self.statement.addButton(self.spec)
        self.statement.addButton(self.spec_with_exam)

        self.spec_var_first.currentTextChanged.connect(self.stream_toggle)

    def stream_toggle(self, text):
        """Тублер виджетов, связанных с выбором потока для определенных специальностей"""
        if text in ['54.02.01 Дизайн (по отраслям)',
                    '07.02.01 Архитектура',
                    '55.02.02 Анимация и анимационное кино (по видам)']:
            self.label_34.setEnabled(True)
            self.label_35.setEnabled(True)
            self.groupBox_6.setEnabled(True)
            self.stream_button.setEnabled(True)
            self.excel.four_flag = True
        else:
            self.label_34.setEnabled(False)
            self.label_35.setEnabled(False)
            self.groupBox_6.setEnabled(False)
            self.stream_button.setEnabled(False)
            self.excel.four_flag = False

    def first_path(self):
        self.excel.first_excel = self.excel.select_excel_file("Выберите РЕЙТИНГ Excel")

    def second_path(self):
        self.excel.second_excel= self.excel.select_excel_file("Выберите ОБЩИЙ Excel")

    def third_path(self):
        self.excel.third_excel = self.excel.select_excel_file("Выберите АИС Excel")

    def fourth_path(self):
        self.excel.fourth_excel = self.excel.select_excel_file("Выберите ПОТОК Excel")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
