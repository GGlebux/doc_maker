import os

from PyQt6.QtCore import QDir
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from docxtpl import DocxTemplate


class Word:
    def __init__(self, parent, data):
        self.parent = parent
        self.data = data
        self.pro = self.parent.spec_var_first.currentText()
        self.adult = self.parent.adult.isChecked()
        self.minor = self.parent.minor.isChecked()
        self.one = False
        self.two = False

    def start_up(self):
        """Запуск заполнения"""
        data = self.data.get_input_data()
        self.fill_word_application(data)
        self.fill_word_data_processing(data)
        if self.one and self.two:
            QMessageBox.information(self.parent, "Успешно", "Файлы Word успешно созданы!")
            self.one = False
            self.two = False

    def fill_word_application(self, data):
        """Заполнение заявления"""
        doc = None
        application = data['statement']

        # Проверка вида заполнения документа
        if application == 'Профессия':
            doc = load_template('prof.docx')
        elif application == 'Специальность':
            doc = load_template('special.docx')
        elif application == 'Спец. с экзаменом':
            if (self.pro == '54.02.01 Дизайн (по отраслям)'
                    or self.pro == '07.02.01 Архитектура'
                    or self.pro == '55.02.02 Анимация и анимационное кино (по видам)'):
                doc = load_template('sp_with_ex.docx')
            else:
                QMessageBox.warning(self.parent, "Ошибка",
                                    "При выборе специальности с экзаменом необходимо " +
                                    "выбрать одну из следующих специальностей:" +
                                    "\n07.02.01 Архитектура,\n54.02.01 Дизайн (по отраслям)," +
                                    "\n55.02.02 Анимация и анимационное кино (по видам)")
                return

        if doc:
            filename = self.save_word_file("Сохранить заявление")
            if filename:
                doc.render(self.data.get_input_data())
                doc.save(filename)
                self.one = True

    def fill_word_data_processing(self, data):
        """Заполнение согласия на обработку персональных данных"""
        doc = None
        if self.adult:
            doc = load_template('adult.docx')
        elif self.minor:
            doc = load_template('minor.docx')
        if doc:
            filename = self.save_word_file("Сохранить согласие на обработку данных")
            if filename:
                doc.render(data)
                doc.save(filename)
                self.two = True

    def save_word_file(self, title):
        """Открывает диалоговое окно для сохранения файла"""
        filename, _ = QFileDialog.getSaveFileName(self.parent, title, QDir().homePath(), "Word Files (*.docx)")
        if filename:
            return filename
        return


def load_template(template_name):
    """Загружает шаблоны Word"""
    # Используем абсолютный путь к файлу
    # ToDo: Для работы в коде - '../patterns', текущий вариант для правильной компиляции
    doc = DocxTemplate(os.path.abspath(f'patterns/{template_name}'))
    return doc
