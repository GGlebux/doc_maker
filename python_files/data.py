class Data:
    def __init__(self, parent):
        self.parent = parent

    def get_input_data(self):
        """Возвращает словарь со всеми переменными введенными в интерфейсе и их значения"""
        # в идеале вот это чудо заменить на liteSQL
        return {
            'reg_number': self.parent.reg_number.text().strip(),
            'surname': self.parent.surname.text().strip(),
            'name': self.parent.name.text().strip(),
            'patronymic': self.parent.patronymic.text().strip(),
            'date_birthday': self.parent.date_birthday.date().toString('dd.MM.yyyy'),
            'snils': self.parent.snils.text().strip(),
            'inn': self.parent.inn.text().strip(),
            'citizenship': self.parent.citizenship.text().strip(),
            'id_doc': self.parent.id_doc.text().strip(),
            'series': self.parent.series.text().strip(),
            'number': self.parent.number.text().strip(),
            'date_id_doc': self.parent.date_id_doc.date().toString('dd.MM.yyyy'),
            'office_doc': self.parent.office_doc.text().strip(),
            'address': self.parent.address.text().strip(),
            'tel_number': self.parent.tel_number.text().strip(),
            'spec_var_first': self.parent.spec_var_first.currentText(),
            'spec_var_second': self.parent.spec_var_second.currentText(),
            'spec_var_third': self.parent.spec_var_third.currentText(),
            'certificate_score': self.parent.certificate_score.text(),
            'form_education': self.parent.form_education.checkedButton().text(),
            'svo': 'Да' if self.parent.svo.isChecked() else 'Нет',
            'target_direction': 'Да' if self.parent.target_direction.isChecked() else 'Нет',
            'parent_fio': self.parent.parent_fio.text().strip(),
            'parent_ser_num': self.parent.parent_ser_num.text().strip(),
            'parent_pass_info': self.parent.parent_pass_info.text().strip(),
            'parent_address': self.parent.parent_address.text().strip(),
            'parent_work': self.parent.parent_work.text().strip(),
            'parent_number': self.parent.parent_number.text().strip(),
            'base_education': self.parent.base_education.checkedButton().text(),
            'statement': self.parent.statement.checkedButton().text(),
            'finance': {'budget': '+' if self.parent.budget.isChecked() else '-',
                        'commerce': '+' if self.parent.commerce.isChecked() else '-'},
            'stream': self.parent.stream.checkedButton().text(),
            'gender': self.parent.gender.checkedButton().text(),
            'orphan': 'Да' if self.parent.orphan.isChecked() else 'Нет',
            'certificate': '✓' if self.parent.certificate.checkedButton().text() == 'Оригинал' else '×'}

