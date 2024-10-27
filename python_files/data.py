class Data:
    def __init__(self, parent):
        self.parent = parent

    def get_input_data(self):
        """Возвращает словарь со всеми переменными введенными в интерфейсе и их значения"""
        # в идеале вот это чудо заменить на liteSQL
        return {
            'reg_number': self.parent.reg_number.text(),
            'surname': self.parent.surname.text(),
            'name': self.parent.name.text(),
            'patronymic': self.parent.patronymic.text(),
            'date_birthday': self.parent.date_birthday.date().toString("dd.MM.yyyy"),
            'snils': self.parent.snils.text(),
            'inn': self.parent.inn.text(),
            'citizenship': self.parent.citizenship.text(),
            'id_doc': self.parent.id_doc.text(),
            'series': self.parent.series.text(),
            'number': self.parent.number.text(),
            'date_id_doc': self.parent.date_id_doc.date().toString("dd.MM.yyyy"),
            'office_doc': self.parent.office_doc.text(),
            'address': self.parent.address.text(),
            'tel_number': self.parent.tel_number.text(),
            'spec_var_first': self.parent.spec_var_first.currentText(),
            'spec_var_second': self.parent.spec_var_second.currentText(),
            'spec_var_third': self.parent.spec_var_third.currentText(),
            'parent_work': self.parent.parent_work.text(),
            'certificate_score': self.parent.certificate_score.text(),
            'form_education': self.parent.form_education.checkedButton().text(),
            'svo': "Да" if self.parent.svo.isChecked() else "Нет",
            'target_direction': "Да" if self.parent.target_direction.isChecked() else "Нет",
            'parent_fio': self.parent.parent_fio.text(),
            'parent_ser_num': self.parent.parent_ser_num_pass.text(),
            'parent_pass_info': self.parent.parent_pass_info.text(),
            'parent_address': self.parent.parent_address.text(),
            'parent_number': self.parent.parent_number.text(),
            'base_education': self.parent.base_education.checkedButton().text(),
            'statement': self.parent.statement.checkedButton().text(),
            'finance': self.parent.finance.checkedButton().text(),
            'stream': self.parent.stream.checkedButton().text()}
