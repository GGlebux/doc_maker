import sys

from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow, QButtonGroup
)

from design.design import Ui_MainWindow
from python_files.clear import Cleaner
from python_files.excel import Excel
from python_files.word import Word


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("приЁмка")
        self.setupUi(self)
        self.updateUi()

        self.excel = Excel(self)
        self.word = Word(self)
        self.cleaner = Cleaner(self)

        self.fill_word_button.clicked.connect(self.word.start_up)
        self.fill_excel_button.clicked.connect(self.excel.start_up)
        self.clear_button.clicked.connect(self.cleaner.clear_form)
        self.exit_button.clicked.connect(sys.exit)

    def updateUi(self):
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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
