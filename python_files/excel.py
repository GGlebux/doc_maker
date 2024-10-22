import os
from multiprocessing.process import parent_process

from PyQt6.QtCore import QDir
from PyQt6.QtWidgets import QMessageBox, QFileDialog
from docxtpl import DocxTemplate
from openpyxl.reader.excel import load_workbook

from python_files.data import Data


class Excel:
    def __init__(self, parent):
        self.parent = parent
        self.data = Data(parent)
        self.alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I',
                         'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R',
                         'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
        self.first_excel = None
        self.second_excel = None
        self.third_excel = None

    def start_up(self):
        """Запуск заполнения"""
        self.fill_first_excel()
        self.fill_second_excel()
        self.fill_third_excel()

    def fill_first_excel(self):
        """Заполняет первый excel"""
        # Проверка выбраны ли существующие файлы для заполнения
        while not self.first_excel:
            self.first_excel = self.select_excel_file("Выберите РЕЙТИНГ Excel")

        # Открываем существующий файл первый Excel
        work_book_first = load_workbook(self.first_excel)
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

        if e['A1'].value is None:
            for i in range(8):
                e[f'{self.alphabet[i]}1'] = head_excel_first[i]

        #  Находим первую пустую строку в столбце "B"
        empty_row = 1
        while e[f'B{empty_row}'].value is not None:
            empty_row += 1

        # Записываем данные в пустую строку
        data = self.data.get_input_data()
        e[f'B{empty_row}'] = data['surname'] + ' ' + data['name'] + ' ' + data['patronymic']
        e[f'D{empty_row}'] = data['certificate_score']
        e[f'E{empty_row}'] = data['spec_var_first']
        e[f'F{empty_row}'] = data['spec_var_second']
        e[f'G{empty_row}'] = data['spec_var_third']
        e[f'H{empty_row}'] = data['form_education']

        work_book_first.save(self.first_excel)
        work_book_first.close()

    def fill_second_excel(self):
        """Заполняет второй excel"""
        # Проверка выбраны ли существующие файлы для заполнения
        while not self.second_excel:
            self.second_excel = self.select_excel_file("Выберите ОБЩИЙ Excel")

        # Открываем существующий файл второй Excel
        work_book_second = load_workbook(self.second_excel)
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
                e2[f'{self.alphabet[i]}1'] = head_excel_second[i]

            # Находим первую пустую строку в столбце "B"
            empty_row_sec = 1
            while e2[f'B{empty_row_sec}'].value is not None:
                empty_row_sec += 1

            # ToDo: Проверка, последней и текущей инфы на дубликаты
            # fio = self.get_input_data()['surname'] + ' ' + self.get_input_data()['name'] + ' ' + \
            #       self.get_input_data()[
            #           'patronymic']
            # for row in range(2, empty_row_sec):
            #     if e2[f'B{row}'].value == fio:
            #         QMessageBox.warning(self.parent, "Ошибка", "Абитуриент с такими ФИО уже есть в таблице!")
            #         return

            # Загружаем все данные из интерфейса как словарь
            data = self.data.get_input_data()
            # Удаляем вторую и третью специальность
            del data['spec_var_second']
            del data['spec_var_third']
            # Переводим словарь в список
            data = [val for val in data.values()]
            # Объединяем все сведения о родителях
            data[16] = ' '.join(data[21:]) + ' ' + data[16]
            # Заполняем пустую строку данными из списка
            for i in range(21):
                e2[f'{self.alphabet[i]}{empty_row_sec}'] = data[i]

            work_book_second.save(self.second_excel)
            work_book_second.close()

    def fill_third_excel(self):
        """Заполняет третий excel"""
        # Проверка выбраны ли существующие файлы для заполнения
        while not self.third_excel:
            self.third_excel = self.select_excel_file("Выберите АИС Excel")

        # Открываем существующий файл третий Excel
        work_book_third = load_workbook(self.third_excel)
        e3 = work_book_third.active

        #  Заполняем шапку excel
        head_excel_third = ['Номер заявления',
                            '№ ЕПГУ',
                            'Фамилия абитуриента',
                            'Имя абитуриента',
                            'Имя абитуриента',
                            'Дата рождения',
                            'Серия удостоверяющего документа',
                            'Номер удостоверяющего документа',
                            'СНИЛС',
                            'Дата подачи заявления',
                            'Источник подачи заявления',
                            'Специальность (1)',
                            'Специальность (2)',
                            'Специальность (3)',
                            'Средний балл аттестата',
                            'Тип финансирования',
                            'Форма обучения',
                            'Базовое образование',
                            'Статус заявления',
                            'Статус специальности']

        if e3['A1'].value is None:
            for i in range(20):
                e3[f'{self.alphabet[i]}1'] = head_excel_third[i]

        #  Находим первую пустую строку в столбце "С"
        empty_row = 1
        while e3[f'C{empty_row}'].value is not None:
            empty_row += 1

        work_book_third.save(self.third_excel)
        work_book_third.close()

    def select_excel_file(self, title):
        """Открывает диалоговое окно для выбора файла Excel"""
        filename, _ = QFileDialog.getOpenFileName(self.parent, title, QDir().homePath(), "Excel Files (*.xlsx)")
        if filename:
            return filename
        return None
