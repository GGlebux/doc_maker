from PyQt6.QtCore import QRegularExpression
from PyQt6.QtGui import QRegularExpressionValidator
from PyQt6.QtWidgets import QMessageBox


def is_present_sp_form(sp, form):
    """Проверяет пару направление-форма обучения"""
    full = 'Очная'
    easy = 'Заочная'
    full_easy = 'Очно-заочная'
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

    first = any(sp.startswith(number) for number in all_full) and form == full
    second = any(sp.startswith(number) for number in all_full_easy) and form in [full, easy]
    third = sp.startswith('08.02.13') and form in [full, full_easy]
    fourth = sp.startswith('08.02.04') and form == full_easy

    print('Очка=' + str(first))
    print('Очка/заочка=' + str(second))
    print('3=' + str(third))
    print('4=' + str(fourth))

    if any([first, second, third, fourth]):
        return True
    return False


def is_present_sp_type(sp, type_state):
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

    first = any(sp.startswith(number) for number in all_spec) and type_state == spec
    second = any(sp.startswith(number) for number in all_spec_with_exam) and type_state == spec_with_exam
    third = any(sp.startswith(number) for number in all_prof) and type_state == prof

    if any([first, second , third]):
        return True
    return False


def check_sp_form_type(data):
    log = 'Несопоставимые поля:\n'
    sp = data['spec_var_first']
    form = data['form_education']
    type_state = data['statement']

    sp_form = is_present_sp_form(sp, form)
    sp_type = is_present_sp_type(sp, type_state)

    print('Итоги (форма и тип)= ' + str(sp_form) +', ' + str(sp_type))

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
        data = self.data.get_input_data()
        error_counter = 0
        if not self.__check_finance_choice():
            QMessageBox.warning(self.parent,
                                'Ошибка',
                                'Необходимо выбрать хотя бы один вид финансирования!')
            error_counter += 1

        couple_cp_form, log = check_sp_form_type(data)
        print(couple_cp_form, log)
        if not couple_cp_form:
            QMessageBox.warning(self.parent,
                                'Ошибка',
                                log)
            error_counter += 1
        return error_counter == 0

    def __check_finance_choice(self):
        """Должен быть выбран тип финансирования"""
        if sum(btn.isChecked() for btn in self.parent.finance.buttons()) == 0:
            return False
        return True
