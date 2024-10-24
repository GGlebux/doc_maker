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

    def start_up(self):
        """Запуск заполнения"""
        # Проверка выбраны ли существующие файлы для заполнения
        while not self.first_excel:
            self.first_excel = self.select_excel_file("Выберите РЕЙТИНГ Excel")
        while not self.second_excel:
            self.second_excel = self.select_excel_file("Выберите ОБЩИЙ Excel")
        while not self.third_excel:
            self.third_excel = self.select_excel_file("Выберите АИС Excel")
        if self.four_flag:
            while not self.fourth_excel:
                self.fourth_excel = self.select_excel_file("Выберите ПОТОК Excel")
            self.fill_fourth_excel()
        # Заполнение
        print('Заупскаю 1')
        self.fill_first_excel()
        print('Заупскаю 2')
        self.fill_second_excel()
        print('Заупскаю 3')
        self.fill_third_excel()
        QMessageBox.information(self.parent, "Успешно", "Файлы Excel успешно созданы!")

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

        # Проверка на уникальность
        data = self.data.get_input_data()
        old_fio = e[f'B{empty_row - 1}'].value
        new_fio = data['surname'] + ' ' + data['name'] + ' ' + data['patronymic']
        if empty_row != 1 and self.check_unique(old_fio, new_fio, '<Рейтинг>'):
            return

        # Записываем данные в пустую строку
        e[f'B{empty_row}'] = data['surname'] + ' ' + data['name'] + ' ' + data['patronymic']
        e[f'D{empty_row}'] = data['certificate_score']
        e[f'E{empty_row}'] = data['spec_var_first']
        e[f'F{empty_row}'] = data['spec_var_second']
        e[f'G{empty_row}'] = data['spec_var_third']
        e[f'H{empty_row}'] = data['form_education']

        wb.save(self.first_excel)
        wb.close()
        print('Первый заполнен')

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
        old_pasport = e[f'J{empty_row - 1}'].value + e[f'K{empty_row - 1}'].value
        new_passport = data['series'].strip() + data['number'].strip()
        if empty_row != 1 and self.check_unique(old_pasport, new_passport, '<Общий>'):
            return

        # Удаляем вторую и третью специальность
        del data['spec_var_second']
        del data['spec_var_third']
        # Переводим словарь в список
        data = [val for val in data.values()]
        # Объединяем все сведения о родителях
        data[16] = ' '.join(data[21:]) + ' ' + data[16]
        # Заполняем пустую строку данными из списка
        for i in range(21):
            e[f'{self.alphabet[i]}{empty_row}'] = data[i]

        wb.save(self.second_excel)
        wb.close()
        print('Второй заполнен')

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

        if e['A1'].value is None:
            for i in range(20):
                e[f'{self.alphabet[i]}1'] = head_excel[i]

        #  Находим первую пустую строку в столбце "С"
        empty_row = 1
        while e[f'C{empty_row}'].value is not None:
            empty_row += 1

        data = self.data.get_input_data()

        # Проверка значений на уникальность
        old_pasport = e[f'G{empty_row - 1}'].value + e[f'H{empty_row - 1}'].value
        new_passport = data['series'].strip() + data['number'].strip()
        if empty_row != 1 and self.check_unique(old_pasport, new_passport, '<АИС>'):
            return

        # Записываем данные в пустую строку
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
        print('Третий заполнен')

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
            for i in range(3):
                e[f'{self.alphabet[i]}1'] = head_excel[i]

        #  Находим первую пустую строку в столбце "A"
        empty_row = 1
        while e[f'A{empty_row}'].value is not None:
            empty_row += 1

        data = self.data.get_input_data()

        # Проверка значений на уникальность
        old_fio = e[f'A{empty_row - 1}'].value
        new_fio = data['surname'].strip() + ' ' + data['name'].strip() + ' ' + data['patronymic'].strip()
        if empty_row != 1 and self.check_unique(old_fio, new_fio, '<Поток>'):
            print('Блокирую')
            return

        # Записываем данные в пустую строку
        e[f'A{empty_row}'] = data['surname'] + ' ' + data['name'] + ' ' + data['patronymic']
        e[f'B{empty_row}'] = data['spec_var_first']
        e[f'С{empty_row}'] = data['stream']

        wb.save(self.fourth_excel)
        wb.close()
        print('Четвертый заполнен')

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