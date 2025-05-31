from PyQt6.QtCore import QRegularExpression
from PyQt6.QtGui import QRegularExpressionValidator
from PyQt6.QtWidgets import QMessageBox

spec = 'Специальность'
spec_with_exam = 'Спец. с экзаменом'
prof = 'Профессия'
all_spec = ['08.02.01',
            '08.02.04',
            '08.02.13',
            '08.02.08',
            '08.02.14',
            '08.02.15',
            '09.02.07',
            '29.02.11',
            '38.02.08',
            '42.02.02']
all_spec_with_exam = ['07.02.01',
                      '54.02.01',
                      '55.02.02']
all_prof = ['54.01.01',
            '54.01.20',
            '08.01.28']

all_full = ['07.02.01',
                '08.01.28',
                '08.02.14',
                '08.02.15',
                '09.02.07',
                '29.02.11',
                '38.02.08',
                '42.02.02',
                '54.02.01',
                '55.02.02',
                '54.01.01',
                '54.01.20']

all_full_easy = ['08.02.01', '08.02.08']

def is_present_sp_form(sp, form):
    """Проверяет пару направление-форма_обучения"""
    full = 'Очная'
    easy = 'Заочная'
    full_easy = 'Очно-заочная'



    first = any(sp.startswith(number) for number in all_full) and form == full
    second = any(sp.startswith(number) for number in all_full_easy) and form in [full, easy]
    third = sp.startswith('08.02.13') and form in [full, full_easy]
    fourth = sp.startswith('08.02.04') and form == full_easy

    if any([first, second, third, fourth]):
        return True
    return False


def is_present_sp_type(sp, type_state):
    """Проверяет пару направление-тип_поступления"""
    first = any(sp.startswith(number) for number in all_spec) and type_state == spec
    second = any(sp.startswith(number) for number in all_spec_with_exam) and type_state == spec_with_exam
    third = any(sp.startswith(number) for number in all_prof) and type_state == prof

    if any([first, second, third]):
        return True
    return False


def check_sp_form_type(data):
    """Проверяет специальность/форму_обучения/тип_поступления"""
    log = 'Несопоставимые поля:\n'
    sp = data['spec_var_first']
    form = data['form_education']
    type_state = data['statement']

    sp_form = is_present_sp_form(sp, form)
    sp_type = is_present_sp_type(sp, type_state)

    log += f'<{sp}> и <{form}>\n' if not sp_form else ''
    log += f'<{sp}> и <{type_state}>\n' if not sp_type else ''

    return sp_form and sp_type, log


class Validator:
    def __init__(self, parent, data):
        self.parent = parent
        self.data = data

        self.reg_num_validator = QRegularExpressionValidator(QRegularExpression(r'^\d+[а-яё]{1,4}$'))
        self.parent.reg_number.setValidator(self.reg_num_validator)

    def validate(self):
        """Валидирует форму и оповещает пользователя"""
        error_counter = 0
        data = self.data.get_input_data()
        couple_cp_form, log = check_sp_form_type(data)
        if not self.__check_finance_choice():
            QMessageBox.warning(self.parent,
                                'Ошибка',
                                'Необходимо выбрать хотя бы один вид финансирования!')
            error_counter += 1

        if not couple_cp_form:
            QMessageBox.warning(self.parent,
                                'Ошибка',
                                log)
            error_counter += 1

        result = error_counter == 0
        if not result:
            self.parent.logger.warning("Валидация не пройдена")
        return result

    def __check_finance_choice(self):
        """Должен быть выбран тип финансирования"""
        if sum(btn.isChecked() for btn in self.parent.finance.buttons()) == 0:
            return False
        return True

    def get_correct_type_btn(self, value):
        value = value[:8]
        if value in all_spec:
            return self.parent.spec
        elif value in all_spec_with_exam:
            return self.parent.spec_with_exam
        elif value in all_prof:
            return self.parent.prof

    def get_correct_form_btn(self, value):
        value = value[:8]
        if value == '08.02.04':
            return [self.parent.combined]
        elif value == '08.02.13':
            return [self.parent.full_time, self.parent.combined]
        elif value in all_full_easy:
            return [self.parent.full_time, self.parent.correspondence]
        elif value in all_full:
            return [self.parent.full_time]

