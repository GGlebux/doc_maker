from PyQt6.QtCore import QDir
from PyQt6.QtWidgets import QMessageBox, QFileDialog
from openpyxl.reader.excel import load_workbook


class Excel:
    def __init__(self, parent, data):
        self.parent = parent
        self.data = data
        self.alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I',
                         'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R',
                         'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
        self.four_flag = False
        self.first_excel = None
        self.second_excel = None
        self.third_excel = None
        self.fourth_excel = None

        self.one = False
        self.two = False
        self.three = False
        self.four = False

    def start_up(self):
        """Запуск заполнения"""
        try:
            # Проверка выбраны ли существующие файлы для заполнения
            while not self.first_excel:
                self.parent.first_path()
            while not self.second_excel:
                self.parent.second_path()
            while not self.third_excel:
                self.parent.third_path()
            if self.four_flag:
                while not self.fourth_excel:
                    self.parent.fourth_path()
                self.fill_fourth_excel()
                self.four_flag = False

            # Заполнение
            self.fill_first_excel()
            self.fill_second_excel()
            self.fill_third_excel()
            if self.one and self.two and self.three and (self.four == self.four_flag):
                QMessageBox.information(self.parent, "Успешно", "Файлы Excel успешно созданы!")
                self.one = False
                self.two = False
                self.three = False
                self.four = False
        except PermissionError:
            QMessageBox.warning(self.parent, "Ошибка",
                                "Закройте все окна Excel и повторите попытку\n(иначе данные не сохранятся)")

    def fill_first_excel(self):
        """Заполняет первый excel"""
        # Открываем первый существующий файл Excel
        wb = load_workbook(self.first_excel)
        e = wb.active

        #  Заполняем шапку excel
        head_excel = ['Номер заявления',
                      'ФИО абитуриента',
                      'Оригинал/копия аттестата',
                      'Средний балл',
                      'Специальность (1)',
                      'Специальность (2)',
                      'Специальность (3)',
                      'Форма обучения']

        if e['A1'].value is None:
            for i in range(8):
                e[f'{self.alphabet[i]}1'] = head_excel[i]

        #  Находим первую пустую строку в столбце "B"
        empty_row = 1
        while e[f'B{empty_row}'].value is not None:
            empty_row += 1

        data = self.data.get_input_data()

        # Проверка на уникальность
        try:
            old_fio = e[f'B{empty_row - 1}'].value
            new_fio = data['surname'] + ' ' + data['name'] + ' ' + data['patronymic']
            if not self.check_unique(old_fio, new_fio, '<Рейтинг>'):
                return
        except TypeError:
            QMessageBox.warning(self.parent, 'Предупреждение',
                                'Недостаточно данных для проверки уникальности данных\n(возможны дубликаты)')

        # Записываем данные в пустую строку
        e[f'A{empty_row}'] = data['reg_number']
        e[f'B{empty_row}'] = data['surname'] + ' ' + data['name'] + ' ' + data['patronymic']
        e[f'D{empty_row}'] = data['certificate_score']
        e[f'E{empty_row}'] = data['spec_var_first']
        e[f'F{empty_row}'] = data['spec_var_second']
        e[f'G{empty_row}'] = data['spec_var_third']
        e[f'H{empty_row}'] = data['form_education']

        wb.save(self.first_excel)
        wb.close()
        self.one = True

    def fill_second_excel(self):
        """Заполняет второй excel"""
        # Открываем второй существующий файл Excel
        wb = load_workbook(self.second_excel)
        e = wb.active

        #  Заполняем шапку excel
        head_excel = ['Регистрационный номер',
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

        if e['A1'].value is None:
            for i in range(21):
                e[f'{self.alphabet[i]}1'] = head_excel[i]

        # Находим первую пустую строку в столбце "B"
        empty_row = 1
        while e[f'B{empty_row}'].value is not None:
            empty_row += 1

        # Загружаем все данные из интерфейса как словарь
        data = self.data.get_input_data()

        # Проверка значений на уникальность
        try:
            old_pasport = e[f'J{empty_row - 1}'].value + e[f'K{empty_row - 1}'].value
            new_passport = data['series'].strip() + data['number'].strip()
            if not self.check_unique(old_pasport, new_passport, '<Общий>'):
                return
        except TypeError:
            QMessageBox.warning(self.parent, 'Предупреждение',
                                'Недостаточно данных для проверки уникальности данных\n(возможны дубликаты)')
            self.alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I',
                             'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R',
                             'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

        e[f'A{empty_row}'] = data['reg_number']
        e[f'B{empty_row}'] = data['surname']
        e[f'C{empty_row}'] = data['name']
        e[f'D{empty_row}'] = data['patronymic']
        e[f'E{empty_row}'] = data['date_birthday']
        e[f'F{empty_row}'] = data['snils']
        e[f'G{empty_row}'] = data['inn']
        e[f'H{empty_row}'] = data['citizenship']
        e[f'I{empty_row}'] = data['id_doc']
        e[f'J{empty_row}'] = data['series']
        e[f'K{empty_row}'] = data['number']
        e[f'L{empty_row}'] = data['office_doc']
        e[f'M{empty_row}'] = data['date_id_doc']
        e[f'N{empty_row}'] = data['address']
        e[f'O{empty_row}'] = data['tel_number']
        e[f'P{empty_row}'] = data['spec_var_first']
        e[f'Q{empty_row}'] = (data['parent_fio'] + '; ' + data['parent_ser_num'] + '; ' +
                              data['parent_pass_info'] + '; ' + data['parent_address'] + '; ' +
                              data['parent_work'] + '; ' + data['parent_number'])
        e[f'R{empty_row}'] = data['certificate_score']
        e[f'S{empty_row}'] = data['form_education']
        e[f'T{empty_row}'] = data['svo']
        e[f'U{empty_row}'] = data['target_direction']

        wb.save(self.second_excel)
        wb.close()
        self.two = True

    def fill_third_excel(self):
        """Заполняет третий excel"""
        # Открываем третий существующий файл Excel
        wb = load_workbook(self.third_excel)
        e = wb.active

        #  Заполняем шапку excel
        head_excel = ['Номер заявления',
                      '№ ЕПГУ',
                      'Фамилия абитуриента',
                      'Имя абитуриента',
                      'Отчество абитуриента',
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

        if e['A1'].value is None:
            for i in range(20):
                e[f'{self.alphabet[i]}1'] = head_excel[i]

        #  Находим первую пустую строку в столбце "С"
        empty_row = 1
        while e[f'C{empty_row}'].value is not None:
            empty_row += 1

        data = self.data.get_input_data()

        # Проверка значений на уникальность
        try:
            old_pasport = e[f'G{empty_row - 1}'].value + e[f'H{empty_row - 1}'].value
            new_passport = data['series'].strip() + data['number'].strip()
            if not self.check_unique(old_pasport, new_passport, '<АИС>'):
                return
        except TypeError:
            QMessageBox.warning(self.parent, 'Предупреждение',
                                'Недостаточно данных для проверки уникальности данных\n(возможны дубликаты)')

        # Записываем данные в пустую строку
        e[f'A{empty_row}'] = data['reg_number']
        e[f'C{empty_row}'] = data['surname']
        e[f'D{empty_row}'] = data['name']
        e[f'E{empty_row}'] = data['patronymic']
        e[f'F{empty_row}'] = data['date_birthday']
        e[f'G{empty_row}'] = data['series']
        e[f'H{empty_row}'] = data['number']
        e[f'I{empty_row}'] = data['snils']
        e[f'L{empty_row}'] = data['spec_var_first']
        e[f'M{empty_row}'] = data['spec_var_second']
        e[f'N{empty_row}'] = data['spec_var_third']
        e[f'O{empty_row}'] = data['certificate_score']
        e[f'P{empty_row}'] = data['finance']
        e[f'Q{empty_row}'] = data['form_education']
        e[f'R{empty_row}'] = data['base_education']

        wb.save(self.third_excel)
        wb.close()
        self.three = True

    def fill_fourth_excel(self):
        """Заполняет четвертый excel"""
        # Открываем первый существующий файл Excel
        wb = load_workbook(self.fourth_excel)
        e = wb.active

        #  Заполняем шапку excel
        head_excel = ['ФИО абитуриента',
                      'Специальность 1'
                      'Поток']

        if e['A1'].value is None:
            for i in range(2):
                e[f'{self.alphabet[i]}1'] = head_excel[i]

        #  Находим первую пустую строку в столбце "A"
        empty_row = 1
        while e[f'A{empty_row}'].value is not None:
            empty_row += 1

        data = self.data.get_input_data()

        # Проверка значений на уникальность
        try:
            old_fio = e[f'A{empty_row - 1}'].value
            new_fio = data['surname'].strip() + ' ' + data['name'].strip() + ' ' + data['patronymic'].strip()
            if not self.check_unique(old_fio, new_fio, '<Поток>'):
                return
        except TypeError:
            QMessageBox.warning(self.parent, 'Предупреждение',
                                'Недостаточно данных для проверки уникальности данных\n(возможны дубликаты)')

        # Записываем данные в пустую строку
        e[f'A{empty_row}'] = data['surname'] + ' ' + data['name'] + ' ' + data['patronymic']
        e[f'B{empty_row}'] = data['spec_var_first']
        e[f'C{empty_row}'] = data['stream']

        wb.save(self.fourth_excel)
        wb.close()
        self.four = True

    def select_excel_file(self, title):
        """Открывает диалоговое окно для выбора файла Excel"""
        filename, _ = QFileDialog.getOpenFileName(self.parent, title, QDir().homePath(), "Excel Files (*.xlsx)")
        if filename:
            return filename
        return None

    def check_unique(self, old_value, new_value, table):
        """Проверяет есть ли в таблице данные, подобные новым"""
        if old_value.strip() == new_value.strip():
            QMessageBox.warning(self.parent, "Ошибка", f"Такой абитуриент уже есть в таблице: {table}!")
            return False
        return True
