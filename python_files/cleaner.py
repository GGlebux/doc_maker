from PyQt6.QtCore import QDate

from python_files.static import log_exception


class Cleaner:
    def __init__(self, parent):
        self.parent = parent

    def clear_form(self):
        """Очищает форму (поля, кнопки, чекбоксы)"""
        try:
            self.parent.reg_number.clear()
            self.parent.surname.clear()
            self.parent.name.clear()
            self.parent.patronymic.clear()
            self.parent.date_birthday.setDate(QDate.currentDate())
            self.parent.snils.clear()
            self.parent.inn.clear()
            self.parent.citizenship.setText('РФ')
            self.parent.id_doc.setText('Паспорт')
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
            self.parent.spec_with_exam.setChecked(True)
            self.parent.full_time.setChecked(True)
            self.parent.parent_work.clear()
            self.parent.certificate_score.setValue(5.00)
            self.parent.minor.setChecked(True)
            self.parent.full_time.setChecked(True)
            self.parent.parent_fio.clear()
            self.parent.parent_ser_num.clear()
            self.parent.parent_pass_info.clear()
            self.parent.parent_address.clear()
            self.parent.parent_number.clear()
            self.parent.svo.setChecked(False)
            self.parent.target_direction.setChecked(False)
            self.parent.budget.setChecked(True)
            self.parent.first_stream.setChecked(True)
            self.parent.budget.setChecked(True)
            self.parent.commerce.setChecked(False)
            self.parent.orphan.setChecked(False)
            self.parent.dormitory.setChecked(False)
            self.parent.copy.setChecked(True)
            for bnt in self.parent.form_education.buttons():
                bnt.setChecked(False)
                bnt.setEnabled(False)
            self.parent.spec_with_exam.setChecked(True)
            self.parent.spec_with_exam.setEnabled(True)
            self.parent.statusbar.showMessage('Форма очищена', 2000)
        except Exception as e:
            log_exception(self.parent, e, 'очистке формы')