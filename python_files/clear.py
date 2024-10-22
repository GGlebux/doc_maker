from math import trunc

from PyQt6.QtCore import QDate


class Cleaner:
    def __init__(self, parent):
        self.parent = parent

    def clear_form(self):
        """Очищает форму (поля, кнопки, чекбоксы)"""
        self.parent.reg_number.clear()
        self.parent.surname.clear()
        self.parent.name.clear()
        self.parent.patronymic.clear()
        self.parent.date_birthday.setDate(QDate.currentDate())
        self.parent.snils.clear()
        self.parent.inn.clear()
        self.parent.citizenship.clear()
        self.parent.id_doc.clear()
        self.parent.series.clear()
        self.parent.number.clear()
        self.parent.date_id_doc.setDate(QDate.currentDate())
        self.parent.office_doc.clear()
        self.parent.address.clear()
        self.parent.tel_number.clear()
        self.parent.spec_var_first.setCurrentIndex(0)
        self.parent.spec_var_second.setCurrentIndex(0)
        self.parent.spec_var_third.setCurrentIndex(0)
        self.parent.nine.setChecked(True)
        self.parent.parent_work.clear()
        self.parent.certificate_score.setValue(5.00)
        self.parent.minor.setChecked(True)
        self.parent.full_time.setChecked(True)
        self.parent.parent_fio.clear()
        self.parent.parent_ser_num_pass.clear()
        self.parent.parent_pass_info.clear()
        self.parent.parent_address.clear()
        self.parent.parent_number.clear()
        self.parent.svo.setChecked(False)
        self.parent.target_direction.setChecked(False)
