import os

from PyQt6.QtCore import QDir
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from docxtpl import DocxTemplate


class Word:
    def __init__(self, parent, data):
        self.parent = parent
        self.data = data
        self.pro = self.parent.spec_var_first.currentText()
        self.one = False


    def start_up(self):
        """Запуск заполнения"""
        # Валидация формы
        if not self.parent.validator.validate():
            return
        data = self.data.get_input_data()
        self.fill_word_application(data)
        if self.one:
            QMessageBox.information(self.parent, "Успешно", "Файлы Word успешно созданы!")
            self.one = False

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
            doc = load_template('sp_with_ex.docx')

        if doc:
            filename = self.save_word_file("Сохранить заявление")
            if filename:
                doc.render(self.data.get_input_data())
                doc.save(filename)
                self.one = True

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
