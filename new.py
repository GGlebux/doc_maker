import os
import sys

from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QLabel,
    QLineEdit,
    QComboBox,
    QCheckBox,
    QPushButton,
    QDateEdit,
    QVBoxLayout,
    QHBoxLayout,
    QFileDialog,
    QMessageBox, QScrollArea, QFormLayout, QGroupBox, QRadioButton
)
from cx_Freeze import setup, Executable
from docxtpl import DocxTemplate
from openpyxl import load_workbook

# Константы
SAVE_DIRECTORY = ''
EXCEL_FILE_FIRST = ''
EXCEL_FILE_SECOND = ''
EXCEL_FILE_THIRD = ''
UNIVERSE = 'СКЛЯРОВ ЛОХ'

# Глобальные переменные
svo = None
target_direction = None


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Мой бухгалтер")
        self.setGeometry(300, 0, 750, 1070)

        # Создаем центральный виджет и макет
        centralWidget = QWidget()
        self.setCentralWidget(centralWidget)
        mainLayout = QVBoxLayout(centralWidget)

        # Создаем область прокрутки
        scrollArea = QScrollArea(centralWidget)
        scrollArea.setWidgetResizable(True)
        scrollAreaWidgetContents = QWidget()
        scrollAreaLayout = QVBoxLayout(scrollAreaWidgetContents)
        scrollArea.setWidget(scrollAreaWidgetContents)
        mainLayout.addWidget(scrollArea)

        # Создаем QFormLayout для ввода данных
        formLayout = QFormLayout()

        # Добавляем поля ввода в QFormLayout
        self.reg_number_label = QLabel("Регистрационный номер:")
        self.reg_number = QLineEdit()
        formLayout.addRow(self.reg_number_label, self.reg_number)
        self.surname_label = QLabel("Фамилия:")
        self.surname = QLineEdit()
        formLayout.addRow(self.surname_label, self.surname)
        self.name_label = QLabel("Имя:")
        self.name = QLineEdit()
        formLayout.addRow(self.name_label, self.name)
        self.patronymic_label = QLabel("Отчество:")
        self.patronymic = QLineEdit()
        formLayout.addRow(self.patronymic_label, self.patronymic)
        self.date_birthday_label = QLabel("Дата рождения:")
        self.date_birthday = QDateEdit()
        self.date_birthday.setDisplayFormat("dd.MM.yyyy")
        formLayout.addRow(self.date_birthday_label, self.date_birthday)
        self.snils_label = QLabel("СНИЛС:")
        self.snils = QLineEdit()
        formLayout.addRow(self.snils_label, self.snils)
        self.inn_label = QLabel("ИНН:")
        self.inn = QLineEdit()
        formLayout.addRow(self.inn_label, self.inn)
        self.citizenship_label = QLabel("Гражданство:")
        self.citizenship = QLineEdit()
        formLayout.addRow(self.citizenship_label, self.citizenship)
        self.id_doc_label = QLabel("Документ, удостоверяющий личность:")
        self.id_doc = QLineEdit()
        formLayout.addRow(self.id_doc_label, self.id_doc)
        self.series_label = QLabel("Серия:")
        self.series = QLineEdit()
        formLayout.addRow(self.series_label, self.series)
        self.number_label = QLabel("Номер:")
        self.number = QLineEdit()
        formLayout.addRow(self.number_label, self.number)
        self.date_id_doc_label = QLabel("Дата выдачи:")
        self.date_id_doc = QDateEdit()
        self.date_id_doc.setDisplayFormat("dd.MM.yyyy")
        formLayout.addRow(self.date_id_doc_label, self.date_id_doc)
        self.office_doc_label = QLabel("Кем выдан:")
        self.office_doc = QLineEdit()
        formLayout.addRow(self.office_doc_label, self.office_doc)
        self.address_label = QLabel("Адрес:")
        self.address = QLineEdit()
        formLayout.addRow(self.address_label, self.address)
        self.tel_number_label = QLabel("Номер телефона:")
        self.tel_number = QLineEdit()
        formLayout.addRow(self.tel_number_label, self.tel_number)

        # Добавляем поля ввода в QFormLayout
        self.spec_var_first_label = QLabel("Специальность:")
        self.spec_var_first = QComboBox()
        self.spec_var_first.addItems(["54.02.01 Дизайн (по отраслям)", "07.02.01 Архитектура",
                                      "55.02.02 Анимация и анимационное кино (по видам)"])
        formLayout.addRow(self.spec_var_first_label, self.spec_var_first)
        self.spec_var_second_label = QLabel("Форма обучения:")
        self.spec_var_second = QComboBox()
        self.spec_var_second.addItems(["Очная", "Заочная"])
        formLayout.addRow(self.spec_var_second_label, self.spec_var_second)
        self.spec_var_third_label = QLabel("Основание получения образования:")
        self.spec_var_third = QComboBox()
        self.spec_var_third.addItems(["Среднее общее образование", "Среднее профессиональное образование"])
        formLayout.addRow(self.spec_var_third_label, self.spec_var_third)
        self.parent_work_label = QLabel("Место работы родителей:")
        self.parent_work = QLineEdit()
        formLayout.addRow(self.parent_work_label, self.parent_work)
        self.certificate_score_label = QLabel("Балл аттестата:")
        self.certificate_score = QLineEdit()
        formLayout.addRow(self.certificate_score_label, self.certificate_score)
        self.form_education_label = QLabel("Форма обучения:")
        self.form_education = QComboBox()
        self.form_education.addItems(["Очная", "Заочная"])
        formLayout.addRow(self.form_education_label, self.form_education)

        # Добавляем QFormLayout в layout для прокрутки
        scrollAreaLayout.addLayout(formLayout)

        # Создаем группу радиокнопок
        self.ageGroup = QGroupBox("Возраст:")
        self.ageLayout = QVBoxLayout(self.ageGroup)
        self.adultRadio = QRadioButton("Совершеннолетний")
        self.minorRadio = QRadioButton("Несовершеннолетний")
        self.ageLayout.addWidget(self.adultRadio)
        self.ageLayout.addWidget(self.minorRadio)

        # Добавляем группу радиокнопок в layout для прокрутки
        scrollAreaLayout.addWidget(self.ageGroup)

        # Создаем QGroupBox для радиокнопок
        groupBox = QGroupBox("Дополнительные сведения:")
        groupBoxLayout = QVBoxLayout(groupBox)

        # Добавляем радиокнопки в QGroupBox
        self.svo_checkbox = QCheckBox("СВО")
        groupBoxLayout.addWidget(self.svo_checkbox)
        self.target_direction_checkbox = QCheckBox("Целевое направление")
        groupBoxLayout.addWidget(self.target_direction_checkbox)

        # Добавляем QGroupBox в layout для прокрутки
        scrollAreaLayout.addWidget(groupBox)

        # Создаем QFormLayout для родителя
        parentFormLayout = QFormLayout()

        # Добавляем поля ввода для родителя в QFormLayout
        self.parent_fio_label = QLabel("ФИО родителя:")
        self.parent_fio = QLineEdit()
        parentFormLayout.addRow(self.parent_fio_label, self.parent_fio)
        self.parent_ser_num_pass_label = QLabel("Серия и номер паспорта родителя:")
        self.parent_ser_num_pass = QLineEdit()
        parentFormLayout.addRow(self.parent_ser_num_pass_label, self.parent_ser_num_pass)
        self.parent_pass_info_label = QLabel("Кем и когда выдан паспорт:")
        self.parent_pass_info = QLineEdit()
        parentFormLayout.addRow(self.parent_pass_info_label, self.parent_pass_info)
        self.parent_address_label = QLabel("Адрес родителя:")
        self.parent_address = QLineEdit()
        parentFormLayout.addRow(self.parent_address_label, self.parent_address)
        self.parent_number_label = QLabel("Номер телефона родителя:")
        self.parent_number = QLineEdit()
        parentFormLayout.addRow(self.parent_number_label, self.parent_number)

        # Добавляем QFormLayout для родителя в layout для прокрутки
        scrollAreaLayout.addLayout(parentFormLayout)

        # Создаем кнопки
        buttonLayout = QHBoxLayout()
        self.fill_word_button = QPushButton("Заполнить Word")
        self.fill_word_button.clicked.connect(self.fill_word)
        buttonLayout.addWidget(self.fill_word_button)
        self.fill_excel_button = QPushButton("Заполнить Excel")
        self.fill_excel_button.clicked.connect(self.fill_excel)
        buttonLayout.addWidget(self.fill_excel_button)
        self.clear_button = QPushButton("Очистить")
        self.clear_button.clicked.connect(self.clear_form)
        buttonLayout.addWidget(self.clear_button)
        self.exit_button = QPushButton("Выход")
        self.exit_button.clicked.connect(self.close)
        buttonLayout.addWidget(self.exit_button)

        # Добавляем кнопки в layout для прокрутки
        scrollAreaLayout.addLayout(buttonLayout)

    def clear_form(self):
        """Очистить все поля ввода Entry"""
        global svo, target_direction
        svo = None
        target_direction = None
        self.reg_number.clear()
        self.surname.clear()
        self.name.clear()
        self.patronymic.clear()
        self.date_birthday.setDate(QDate.currentDate())
        self.snils.clear()
        self.inn.clear()
        self.citizenship.clear()
        self.id_doc.clear()
        self.series.clear()
        self.number.clear()
        self.date_id_doc.setDate(QDate.currentDate())
        self.office_doc.clear()
        self.address.clear()
        self.tel_number.clear()
        self.spec_var_first.setCurrentIndex(0)
        self.spec_var_second.setCurrentIndex(0)
        self.spec_var_third.setCurrentIndex(0)
        self.parent_work.clear()
        self.certificate_score.clear()
        self.form_education.setCurrentIndex(0)
        self.parent_fio.clear()
        self.parent_ser_num_pass.clear()
        self.parent_pass_info.clear()
        self.parent_address.clear()
        self.parent_number.clear()
        self.svo_checkbox.setChecked(False)
        self.target_direction_checkbox.setChecked(False)

    def get_input_data(self):
        """Возвращает словарь со всеми переменными введенными в интерфейсе и их значения"""
        return {
            'reg_number': self.reg_number.text(),
            'surname': self.surname.text(),
            'name': self.name.text(),
            'patronymic': self.patronymic.text(),
            'date_birthday': self.date_birthday.date().toString("dd.MM.yyyy"),
            'snils': self.snils.text(),
            'inn': self.inn.text(),
            'citizenship': self.citizenship.text(),
            'id_doc': self.id_doc.text(),
            'series': self.series.text(),
            'number': self.number.text(),
            'date_id_doc': self.date_id_doc.date().toString("dd.MM.yyyy"),
            'office_doc': self.office_doc.text(),
            'address': self.address.text(),
            'tel_number': self.tel_number.text(),
            'spec_var_first': self.spec_var_first.currentText(),
            'spec_var_second': self.spec_var_second.currentText(),
            'spec_var_third': self.spec_var_third.currentText(),
            'parent_work': self.parent_work.text(),
            'certificate_score': self.certificate_score.text(),
            'form_education': self.form_education.currentText(),
            'svo': "Да" if self.svo_checkbox.isChecked() else "Нет",
            'target_direction': "Да" if self.target_direction_checkbox.isChecked() else "Нет",
            'parent_fio': self.parent_fio.text(),
            'parent_ser_num': self.parent_ser_num_pass.text(),
            'parent_pass_info': self.parent_pass_info.text(),
            'parent_address': self.parent_address.text(),
            'parent_number': self.parent_number.text(),
            'base_education': self.spec_var_third.currentText(),
            'finance': ""}

    def fill_word(self):
        """Заполнить документ Word информацией введенной из интерфейса"""
        doc = None
        doc2 = None

        # Заполнение заявления
        # Проверка вида заполнения документа
        if self.spec_var_first.currentText() == "54.02.01 Дизайн (по отраслям)":
            doc = doc_prof
        elif self.spec_var_first.currentText() == "07.02.01 Архитектура" or self.spec_var_first.currentText() == "55.02.02 Анимация и анимационное кино (по видам)":
            doc = doc_special
        elif self.spec_var_first.currentText() == "54.02.01 Дизайн (по отраслям)" or self.spec_var_first.currentText() == "07.02.01 Архитектура" or self.spec_var_first.currentText() == "55.02.02 Анимация и анимационное кино (по видам)":
            doc = doc_sp_with_ex
        else:
            QMessageBox.warning(self, "Ошибка",
                                "При выборе специальности с экзаменом необходимо выбрать одну из следующих специальностей:" +
                                "\n07.02.01 Архитектура,\n54.02.01 Дизайн (по отраслям)," +
                                "\n55.02.02 Анимация и анимационное кино (по видам)")
            return

        data = self.get_input_data()
        if doc:
            doc.render(data)
            filename, _ = QFileDialog.getSaveFileName(self, "Сохранить заявление", "Word Files (*.docx)")
            if filename:
                doc.save(filename)

        # Заполнение согласия на обработку персональных данных
        if self.adultRadio.isChecked():
            doc2 = doc_adult
        elif self.minorRadio.isChecked():
            doc2 = doc_minor
        if doc2:
            doc2.render(data)
            filename, _ = QFileDialog.getSaveFileName(self, "Сохранить согласие на обработку данных",
                                                      "Word Files (*.docx)")
            if filename:
                doc2.save(filename)
                QMessageBox.information(self, "Успешно", "Файлы Word успешно созданы!")

    def fill_excel(self):
        """Заполнить документ Excel информацией введенной из интерфейса"""
        global EXCEL_FILE_FIRST, EXCEL_FILE_SECOND, EXCEL_FILE_THIRD

        # Проверка выбраны ли существующие файлы для заполнения
        while not EXCEL_FILE_FIRST:
            EXCEL_FILE_FIRST = self.select_excel_file("Выберите первый файл Excel")

        while not EXCEL_FILE_SECOND:
            EXCEL_FILE_SECOND = self.select_excel_file("Выберите второй файл Excel")

        while not EXCEL_FILE_THIRD:
            EXCEL_FILE_THIRD = self.select_excel_file("Выберите третий файл Excel")

        '''Заполняем первый excel'''
        # Открываем существующий файл первый Excel
        work_book_first = load_workbook(EXCEL_FILE_FIRST)
        e = work_book_first.active

        #  Заполняем шапку excel
        head_excel_first = ['Номер заявления',
                            'ФИО абитуриента',
                            'Оригинал/копия аттестата',
                            'Средний балл',
                            'Специальность (1)',
                            'Специальность (2)',
                            'Специальность (3)',
                            'Форма обучения']

        alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
                    'U',
                    'V', 'W', 'X', 'Y', 'Z']

        if e['A1'].value is None:
            for i in range(8):
                e[f'{alphabet[i]}1'] = head_excel_first[i]

        #  Находим первую пустую строку в столбце "B"
        empty_row = 1
        while e[f'B{empty_row}'].value is not None:
            empty_row += 1

        # Записываем данные в пустую строку
        data = self.get_input_data()
        e[f'B{empty_row}'] = data['surname'] + ' ' + data['name'] + ' ' + data['patronymic']
        e[f'D{empty_row}'] = data['certificate_score']
        e[f'E{empty_row}'] = data['spec_var_first']
        e[f'F{empty_row}'] = data['spec_var_second']
        e[f'G{empty_row}'] = data['spec_var_third']
        e[f'H{empty_row}'] = data['form_education']

        work_book_first.save(EXCEL_FILE_FIRST)
        work_book_first.close()

        '''Заполняем второй excel'''
        # Открываем существующий файл второй Excel
        work_book_second = load_workbook(EXCEL_FILE_SECOND)
        e2 = work_book_second.active

        #  Заполняем шапку excel
        head_excel_second = ['Регистрационный номер',
                             'Фамилия',
                             'Имя',
                             'Отчество',
                             'Дата рождения',
                             'СНИЛС',
                             'ИНН',
                             'Гражданство',
                             'Документ, удостоверяющий личность',
                             'Серия',
                             'Номер',
                             'Кем выдан',
                             'Когда выдан',
                             'Проживающий по адресу',
                             'Телефон',
                             'Специальность 1',
                             'Сведения о родителях',
                             'Средний балл аттестата',
                             'Форма обучения',
                             'Участник СВО',
                             'Целевое направление']

        if e2['A1'].value is None:
            for i in range(21):
                e2[f'{alphabet[i]}1'] = head_excel_second[i]

            # Находим первую пустую строку в столбце "B"
            empty_row_sec = 1
            while e2[f'B{empty_row_sec}'].value is not None:
                empty_row_sec += 1

            # Проверяем, не существует ли уже такой же ФИО в столбце B
            fio = self.get_input_data()['surname'] + ' ' + self.get_input_data()['name'] + ' ' + self.get_input_data()[
                'patronymic']
            for row in range(2, empty_row_sec):
                if e2[f'B{row}'].value == fio:
                    QMessageBox.warning(self, "Ошибка", "Абитуриент с такими ФИО уже есть в таблице!")
                    return

            # Записываем данные в пустую строку (если ФИО уникально)
            # Загружаем все данные из интерфейса как словарь
            data2 = self.get_input_data()
            # Удаляем вторую и третью специальность
            del data2['spec_var_second']
            del data2['spec_var_third']
            # Переводим словарь в список
            data2 = [val for val in data2.values()]
            # Объединяем все сведения о родителях
            data2[16] = ' '.join(data2[21:]) + ' ' + data2[16]
            # Заполняем пустую строку данными из списка
            for i in range(21):
                e2[f'{alphabet[i]}{empty_row_sec}'] = data2[i]

            work_book_second.save(EXCEL_FILE_SECOND)
            work_book_second.close()

    def select_excel_file(self, title):
        """Открывает диалоговое окно для выбора файла Excel"""
        global SAVE_DIRECTORY
        filename, _ = QFileDialog.getOpenFileName(self, title, SAVE_DIRECTORY, "Excel Files (*.xlsx)")
        if filename:
            SAVE_DIRECTORY = filename
            return filename
        return None

    def save_file(self, title):
        """Открывает диалоговое окно для сохранения файла"""
        global SAVE_DIRECTORY
        filename, _ = QFileDialog.getSaveFileName(self, title, SAVE_DIRECTORY, "Word Files (*.docx)")
        if filename:
            SAVE_DIRECTORY = filename
            return True
        return False


def load_template(template_name):
    template_path = os.path.join("patterns", template_name)
    doc = DocxTemplate(template_path)
    return doc


setup(
    name="Admission Office",
    version="2.0",
    description="My Awesome Application",
    executables=[Executable("new.py", base="Win32GUI")],
    # Включаем папку шаблонов
    include_files=[("patterns", "patterns")], )

if __name__ == "__main__":
    doc_prof = load_template('prof.docx')
    doc_special = load_template('special.docx')
    doc_sp_with_ex = load_template('sp_with_ex.docx')
    doc_adult = load_template('adult.docx')
    doc_minor = load_template('minor.docx')

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
