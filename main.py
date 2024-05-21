# подключаем библиотеки
from tkinter import *
from tkinter import messagebox, filedialog, ttk, Tk
from docxtpl import DocxTemplate
from openpyxl import load_workbook, Workbook

# подключаем графическую библиотеку
window = Tk()
# заголовок окна
window.title("Мой бухгалтер")

# Константы
SAVE_DIRECTORY = ''
EXCEL_FILE = ''


def on_closing():
    '''Обрабатывает закрытие окна'''
    if messagebox.askokcancel('Внимание', "Закрыть программу?"):
        window.destroy()


def error(info):
    '''Выводит диалоговое окно с описание ошибки'''
    messagebox.askokcancel('Внимание', info)


def context():
    return {'reg_number': reg_number.get(),
            'surname': surname.get(),
            'name': name.get(),
            'patronymic': patronymic.get(),
            'date_birthday': date_birthday.get(),
            'snils': snils.get(),
            'inn': inn.get(),
            'citizenship': citizenship.get(),
            'id_doc': id_doc.get(),
            'series': series.get(),
            'number': number.get(),
            'date_id_doc': date_id_doc.get(),
            'office_doc': office_doc.get(),
            'address': address.get(),
            'tel_number': tel_number.get(),
            'spec_var_first': spec_var_first.get(),
            'spec_var_second': spec_var_second.get(),
            'spec_var_third': spec_var_third.get(),
            'parents_info': parents_info.get()
            }


def fill_word():
    '''Заполнить документ Word информацией введенной из интерфейса'''

    # Проверка вида заполнения документа
    if choice.get() == profession:
        doc = DocxTemplate('patterns/prof.docx')
    elif choice.get() == specialty:
        doc = DocxTemplate('patterns/special.docx')
    elif choice.get() == specialty_with_exam:
        if (spec_var_first.get() == '54.02.01 Дизайн (по отраслям)'
                or spec_var_first.get() == '07.02.01 Архитектура'
                or spec_var_first.get() == '55.02.02 Анимация и анимационное кино (по видам)'):
            doc = DocxTemplate('patterns/sp_with_ex.docx')
        else:
            return error(
                'При выборе специальности с экзаменом необходимо выбрать одну из следующих специальностей:\n07.02.01 Архитектура,\n54.02.01 Дизайн (по отраслям),\n55.02.02 Анимация и анимационное кино (по видам)')

    doc.render(context())

    all_way = save_file()
    if all_way:
        doc.save(all_way)


def save_file(type_file='word'):
    '''Возвращает путь, имя файла и расширение, с которым его необходимо сохранить'''
    if type_file == 'excel':
        file_path = filedialog.asksaveasfilename(defaultextension='xlsx', filetypes=[("Excel files", "*.xlsx")])
    else:
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
    return file_path


def select_excel_file():
    '''Возвращает уже созданный файл excel для дальнейшего редактирования'''

    global EXCEL_FILE
    global l22

    EXCEL_FILE = filedialog.askopenfilename(title='Выбор excel файла для заполнения', defaultextension='xlsx',
                                            filetypes=[("Excel files", "*.xlsx")])

    if EXCEL_FILE:
        # обновление строки состояния выбранного файла
        l22.config(text=EXCEL_FILE)
        return EXCEL_FILE


def create_excel_file():
    '''Создает новый excel файл и возвращает его для дальнейшего редактирования'''

    global EXCEL_FILE
    global l22

    work_book = Workbook()
    EXCEL_FILE = save_file(type_file='excel')

    if EXCEL_FILE:
        # обновление строки состояния выбранного файла
        l22.config(text=EXCEL_FILE)

        work_book.save(EXCEL_FILE)
        return EXCEL_FILE


def fill_excel():
    '''Заполнить документ excel информацией введенной из интерфейса'''
    global EXCEL_FILE

    # Проверка выбран ли существующий файл для заполнения
    while not EXCEL_FILE:
        EXCEL_FILE = select_excel_file()

    # Открываем существующий файл Excel
    work_book = load_workbook(EXCEL_FILE)
    work_page = work_book.active

    #   ToDo: Заполняем шапку excel
    '''Нужны конкретные заголовки столбцов и в каком порядке они будут идти'''


    #   ToDo: Находим первую пустую строку в столбце "A"
    empty_row = 1
    while work_page[f'A{empty_row}'].value is not None:
        empty_row += 1

    #   ToDo: Записываем данные в пустую строку


    work_book.save(EXCEL_FILE)


# сообщаем системе о том, что делать, когда окно закрывается
window.protocol("WM_DELETE_WINDOW", on_closing)

'''Cоздание надписи для полей ввода и размещение их по сетке'''

l1 = Label(window, text='Регистрационный номер')
l1.grid(row=0, column=0, sticky=E)

l2 = Label(window, text='Фамилия')
l2.grid(row=1, column=0, sticky=E)

l3 = Label(window, text='Имя')
l3.grid(row=2, column=0, sticky=E)

l4 = Label(window, text='Отчество')
l4.grid(row=3, column=0, sticky=E)

l5 = Label(window, text='Дата рождения')
l5.grid(row=4, column=0, sticky=E)

l6 = Label(window, text='СНИЛС')
l6.grid(row=5, column=0, sticky=E)

l7 = Label(window, text='ИНН')
l7.grid(row=6, column=0, sticky=E)

l8 = Label(window, text='Гражданство')
l8.grid(row=7, column=0, sticky=E)

l9 = Label(window, text='Документ, удостоверяющий личность')
l9.grid(row=8, column=0, sticky=E)

l10 = Label(window, text='Серия')
l10.grid(row=9, column=0, sticky=E)

l12 = Label(window, text='Номер')
l12.grid(row=11, column=0, sticky=E)

l13 = Label(window, text='Когда и кем выдан')
l13.grid(row=12, column=0, sticky=E)

l14 = Label(window, text='Проживающего(ей) по адресу')
l14.grid(row=13, column=0, sticky=E)

l15 = Label(window, text='Телефон')
l15.grid(row=14, column=0, sticky=E)

l16 = Label(window, text='Специальность 1')
l16.grid(row=15, column=0, sticky=E)

l17 = Label(window, text='Специальность 2')
l17.grid(row=16, column=0, sticky=E)

l18 = Label(window, text='Специальность 3')
l18.grid(row=17, column=0, sticky=E)

l19 = Label(window, text='Сведения о родителях')
l19.grid(row=18, column=0, sticky=E)

l21 = Label(window, text='Заполнение файла excel')
l21.grid(row=20, column=0)

l22 = Label(window, text='')
l22.grid(row=20, column=1)

'''Создание полей для ввода данных'''
reg_number = StringVar()
e1 = Entry(window, textvariable=reg_number, width=30)
e1.grid(row=0, column=1, sticky=W)

surname = StringVar()
e2 = Entry(window, textvariable=surname, width=30)
e2.grid(row=1, column=1, sticky=W)

name = StringVar()
e3 = Entry(window, textvariable=name, width=30)
e3.grid(row=2, column=1, sticky=W)

patronymic = StringVar()
e3 = Entry(window, textvariable=patronymic, width=30)
e3.grid(row=3, column=1, sticky=W)

date_birthday = StringVar()
e4 = Entry(window, textvariable=date_birthday, width=30)
e4.grid(row=4, column=1, sticky=W)

snils = StringVar()
e5 = Entry(window, textvariable=snils, width=30)
e5.grid(row=5, column=1, sticky=W)

inn = StringVar()
e6 = Entry(window, textvariable=inn, width=30)
e6.grid(row=6, column=1, sticky=W)

citizenship = StringVar()
e7 = Entry(window, textvariable=citizenship, width=30)
e7.grid(row=7, column=1, sticky=W)

id_doc = StringVar()
e8 = Entry(window, textvariable=id_doc, width=30)
e8.grid(row=8, column=1, sticky=W)

series = StringVar()
e9 = Entry(window, textvariable=series, width=30)
e9.grid(row=9, column=1, sticky=W)

number = StringVar()
e10 = Entry(window, textvariable=number, width=30)
e10.grid(row=11, column=1, sticky=W)

date_id_doc = StringVar()
e11 = Entry(window, textvariable=date_id_doc, width=25)
e11.grid(row=12, column=1, sticky=W)

office_doc = StringVar()
e12 = Entry(window, textvariable=office_doc, width=25)
e12.grid(row=12, column=1)

address = StringVar()
e13 = Entry(window, textvariable=address, width=60)
e13.grid(row=13, column=1, sticky=W)

tel_number = StringVar()
e14 = Entry(window, textvariable=tel_number, width=30)
e14.grid(row=14, column=1, sticky=W)

parents_info = StringVar()
e15 = Entry(window, textvariable=parents_info, width=60)
e15.grid(row=18, column=1, sticky=W)

profession = 'Профессия'
specialty = 'Специальность'
specialty_with_exam = 'Специальность c экзаменом'

'''Выбор шаблона заполнения word файла'''
choice = StringVar(value=profession)

rb = Radiobutton(window, text=profession, value=profession, variable=choice)
rb.grid(row=3, column=3)

rb1 = Radiobutton(window, text=specialty, value=specialty, variable=choice)
rb1.grid(row=4, column=3)

rb2 = Radiobutton(window, text=specialty_with_exam, value=specialty_with_exam, variable=choice)
rb2.grid(row=5, column=3)

'''Выбор трех специальностей'''
specializations = ['38.02.08 Торговое дело',
                   '08.02.13 Монтаж и эксплуатация внутренних сантехнических устройств, кондиционирования воздуха и вентиляции',
                   '08.02.01 Строительство и эксплуатация зданий и сооружений',
                   '29.02.11 Полиграфическое производство',
                   '08.02.14 Эксплуатация и обслуживание многоквартирного дома',
                   '08.01.28 Мастер отделочных строительных и декоративных работ',
                   '54.01.20 Графический дизайнер',
                   '09.02.07 Информационные системы и программирование',
                   '42.02.02 Издательское дело',
                   '54.01.01 Исполнитель художественно - оформительских работ',
                   '54.02.01 Дизайн (по отраслям)',
                   '08.02.08 Монтаж и эксплуатация оборудования и систем газоснабжения',
                   '07.02.01 Архитектура',
                   '55.02.02 Анимация и анимационное кино (по видам)']

spec_var_first = StringVar(value=specializations[0])
spec_var_second = StringVar(value=specializations[0])
spec_var_third = StringVar(value=specializations[0])

combobox1 = ttk.Combobox(textvariable=spec_var_first, values=specializations, width=60)
combobox1.grid(row=15, column=1)

combobox1 = ttk.Combobox(textvariable=spec_var_second, values=specializations, width=60)
combobox1.grid(row=16, column=1)

combobox1 = ttk.Combobox(textvariable=spec_var_third, values=specializations, width=60)
combobox1.grid(row=17, column=1)

# Кнопки взаимодействия

btn1 = Button(window, text="Заполнить Word", width=12, command=fill_word)
btn1.grid(row=13, column=3)

btn2 = Button(window, text="Заполнить Excel", width=12, command=fill_excel)
btn2.grid(row=14, column=3)

btn3 = Button(window, text='Выбрать файл', width=12, command=select_excel_file)
btn3.grid(row=20, column=2)

btn4 = Button(window, text='Создать новый', width=12, command=create_excel_file)
btn4.grid(row=20, column=3)

# пусть окно работает всё время до закрытия
window.mainloop()
