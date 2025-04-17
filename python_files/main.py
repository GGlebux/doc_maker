import sys

from PyQt6 import QtCore
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow, QButtonGroup, QMessageBox
)

from design.design import Ui_MainWindow
from python_files.clear import Cleaner
from python_files.data import Data
from python_files.excel import Excel
from python_files.word import Word


def simple_toggle(check_box, label, path, button, set_flag_func):
    """Статический переключатель для активации виджетов"""
    if check_box.isChecked():
        label.setEnabled(True)
        button.setEnabled(True)
        set_flag_func(True)
    else:
        label.setEnabled(False)
        button.setEnabled(False)
        path.clear()
        set_flag_func(False)


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("приЁмка")

        self.updateUi()

        # Классы для заполнения Word, Excel, Cleaner (очистка формы), Data (все данные)
        self.data = Data(self)
        self.excel = Excel(self, self.data)
        self.word = Word(self, self.data)
        self.cleaner = Cleaner(self)

        self.rating_button.clicked.connect(lambda: self.select_excel_path('rating_excel',
                                                                          'Выберите РЕЙТИНГ Excel',
                                                                          self.path_rating))
        self.total_button.clicked.connect(lambda: self.select_excel_path('total_excel',
                                                                         'Выберите ОБЩИЙ Excel',
                                                                         self.path_total))
        self.aic_button.clicked.connect(lambda: self.select_excel_path('aic_excel',
                                                                       'Выберите АИС Excel',
                                                                       self.path_aic))
        self.stream_button.clicked.connect(lambda: self.select_excel_path('stream_excel',
                                                                          'Выберите ПОТОК Excel',
                                                                          self.path_stream))
        self.svo_button.clicked.connect(lambda: self.select_excel_path('svo_excel',
                                                                       'Выберите СВО Excel',
                                                                       self.path_svo))
        self.dormitory_button.clicked.connect(lambda: self.select_excel_path('dormitory_excel',
                                                                             'Выберите ОБЩЕЖИТИЕ Excel',
                                                                             self.path_dormitory))
        self.orphan_button.clicked.connect(lambda: self.select_excel_path('orphan_excel',
                                                                          'Выберите СИРОТЫ Excel',
                                                                          self.path_orphan))

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
        self.more_group.addButton(self.orphan)

        self.finance = QButtonGroup(self)
        self.finance.setExclusive(False)
        self.finance.addButton(self.budget)
        self.finance.addButton(self.commerce)

        self.stream = QButtonGroup(self)
        self.stream.addButton(self.first_stream)
        self.stream.addButton(self.second_stream)

        self.statement = QButtonGroup(self)
        self.statement.addButton(self.prof)
        self.statement.addButton(self.spec)
        self.statement.addButton(self.spec_with_exam)

        self.gender = QButtonGroup(self)
        self.gender.addButton(self.female)
        self.gender.addButton(self.male)

        # Если выбрали одну из необходимых специальностей, то разблокируем выбор потоков
        self.spec_var_first.currentTextChanged.connect(self.stream_toggle)

        # Если выбрали СВО, то позволяем заполнить соотвествующий документ
        self.svo.checkStateChanged.connect(lambda: simple_toggle(self.svo,
                                                                 self.svo_label,
                                                                 self.path_svo,
                                                                 self.svo_button,
                                                                 lambda boolean: setattr(self.excel,
                                                                                         'svo_flag',
                                                                                         boolean)))

        # Если выбрали СИРОТА, то позволяем заполнить соотвествующий документ
        self.orphan.checkStateChanged.connect(lambda: simple_toggle(self.orphan,
                                                                    self.orphan_label,
                                                                    self.path_orphan,
                                                                    self.orphan_button,
                                                                    lambda boolean: setattr(self.excel,
                                                                                            'orphan_flag',
                                                                                            boolean)))

        # Если выбрали ОБЩЕЖИТИЕ, то позволяем заполнить соотвествующий документ
        self.dormitory.checkStateChanged.connect(self.dormitory_toggle)

    def event(self, event):
        if event.type() == QtCore.QEvent.Type.KeyPress:
            if event.key() in (QtCore.Qt.Key.Key_Return, QtCore.Qt.Key.Key_Enter):
                self.focusNextPrevChild(True)
        return super().event(event)

    def stream_toggle(self, text):
        """Тублер активации виджетов, связанных с выбором потока для определенных специальностей"""
        if text in ['54.02.01 Дизайн (по отраслям)',
                    '07.02.01 Архитектура',
                    '55.02.02 Анимация и анимационное кино (по видам)']:
            self.stream_label1.setEnabled(True)
            self.stream_label2.setEnabled(True)
            self.stream_group_box.setEnabled(True)
            self.stream_button.setEnabled(True)
            self.excel.stream_flag = True
        else:
            self.stream_label1.setEnabled(False)
            self.stream_label2.setEnabled(False)
            self.stream_group_box.setEnabled(False)
            self.stream_button.setEnabled(False)
            self.first_stream.setChecked(True)
            self.path_stream.clear()
            self.excel.stream_flag = False

    def dormitory_toggle(self):
        """Тублер для переключения полей связанных с общежитием"""
        # Регулируем группу кнопок, отвечающую за гендер
        if self.dormitory.isChecked():
            self.gender_box.setEnabled(True)
            self.male.setEnabled(True)
            self.female.setEnabled(True)
        else:
            self.gender_box.setEnabled(False)
            self.male.setChecked(True)
            self.male.setEnabled(False)
            self.female.setEnabled(False)
        # Регулируем виджеты для сохранения файла ОБЩЕЖИТИЕ excel
        simple_toggle(self.dormitory,
                      self.dormitory_label,
                      self.path_dormitory,
                      self.dormitory_button,
                      lambda boolean: setattr(self.excel,
                                              'dormitory_flag',
                                              boolean))

    def select_excel_path(self, excel_name, title, path):
        setattr(self.excel, excel_name, self.excel.select_excel_file(title))
        path.setText(getattr(self.excel, excel_name))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
