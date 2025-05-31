import os
import sys

from PyQt6.QtCore import QDir
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from docxtpl import DocxTemplate


def get_base_dir():
    """Возвращает правильную корневую папку (для разработки и для сборки)."""
    if hasattr(sys, '_MEIPASS'):
        # Режим собранного exe (Nuitka)
        return sys._MEIPASS
    else:
        # Режим разработки: поднимаемся на уровень выше из папки python_files
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))



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
        filename = None
        application = data['statement']

        # Проверка вида заполнения документа
        match application:
            case 'Профессия': doc = self.load_template('prof.docx')
            case 'Специальность': doc = self.load_template('special.docx')
            case 'Спец. с экзаменом': doc = self.load_template('sp_with_ex.docx')
            case _: self.parent.logger.warning('Не выбран Word')

        if doc:
            while not filename:
                filename = self.save_word_file("Сохранить заявление")
            doc.render(self.data.get_input_data())
            doc.save(filename)
            self.one = True
            self.parent.logger.info(f'Файл word заполнен по шаблону <{application}> в <{filename}>')

    def save_word_file(self, title):
        """Открывает диалоговое окно для сохранения файла"""
        filename, _ = QFileDialog.getSaveFileName(self.parent, title, QDir().homePath(), "Word Files (*.docx)")
        if filename:
            return filename
        return

    def load_template(self, template_name):
        """Загружает шаблоны Word с учётом сборки и разработки."""
        base_dir = get_base_dir()
        template_path = os.path.join(base_dir, 'patterns', template_name)
        if not os.path.exists(template_path):
            self.parent.logger.error(f'Шаблон не найден: {template_path}')
            return Exception
        return DocxTemplate(template_path)