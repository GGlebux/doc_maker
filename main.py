# подключаем графическую библиотеку для создания интерфейсов
from tkinter import *
from docxtpl import DocxTemplate

# подключаем графическую библиотеку
window = Tk()
# заголовок окна
window.title("Мой бухгалтер")


# обрабатываем закрытие окна
def on_closing():
    '''# показываем диалоговое окно с кнопкой
    if messagebox.askokcancel("", "Закрыть программу?"):
        # удаляем окно и освобождаем память'''
    window.destroy()


# сообщаем системе о том, что делать, когда окно закрывается
window.protocol("WM_DELETE_WINDOW", on_closing)

'''Cоздание надписи для полей ввода и размещение их по сетке'''

l1 = Label(window, text="Регистрационный номер")
l1.grid(row=0, column=0)

l2 = Label(window, text='Фамилия')
l2.grid(row=1, column=0)

l3 = Label(window, text='Имя')
l3.grid(row=2, column=0)

l4 = Label(window, text='Отчество')
l4.grid(row=3, column=0)

l5 = Label(window, text='Дата рождения')
l5.grid(row=4, column=0)

l6 = Label(window, text='СНИЛС')
l6.grid(row=5, column=0)

l7 = Label(window, text='ИНН')
l7.grid(row=6, column=0)

l8 = Label(window, text='Гражданство')
l8.grid(row=7, column=0)

l9 = Label(window, text='Документ, удостоверяющий личность')
l9.grid(row=8, column=0)

l10 = Label(window, text='Серия')
l10.grid(row=9, column=0)

l12 = Label(window, text='Номер')
l12.grid(row=11, column=0)

l13 = Label(window, text='Когда и кем выдан')
l13.grid(row=12, column=0)

l14 = Label(window, text='Проживающего(ей) по адресу')
l14.grid(row=13, column=0)

l15 = Label(window, text='Телефон')
l15.grid(row=14, column=0)

'''Создание полей для ввода данных'''

reg_number = StringVar()
e1 = Entry(window, textvariable=reg_number, width=30)
e1.grid(row=0, column=1)

surname = StringVar()
e2 = Entry(window, textvariable=surname)
e2.grid(row=1, column=1)

name = StringVar()
e3 = Entry(window, textvariable=name)
e3.grid(row=2, column=1)

patronymic = StringVar()
e3 = Entry(window, textvariable=patronymic)
e3.grid(row=3, column=1)

date_birthday = StringVar()
e3 = Entry(window, textvariable=date_birthday)
e3.grid(row=4, column=1)

snils = StringVar()
e3 = Entry(window, textvariable=snils)
e3.grid(row=5, column=1)

inn = StringVar()
e3 = Entry(window, textvariable=inn)
e3.grid(row=6, column=1)

citizenship = StringVar()
e3 = Entry(window, textvariable=citizenship)
e3.grid(row=7, column=1)

id_doc = StringVar()
e3 = Entry(window, textvariable=id_doc)
e3.grid(row=8, column=1)

series = StringVar()
e3 = Entry(window, textvariable=series)
e3.grid(row=9, column=1)

number = StringVar()
e3 = Entry(window, textvariable=number)
e3.grid(row=11, column=1)

date_id_doc = StringVar()
e3 = Entry(window, textvariable=date_id_doc)
e3.grid(row=12, column=1)

office_doc = StringVar()
e3 = Entry(window, textvariable=office_doc, width=45)
e3.grid(row=12, column=2)

address = StringVar()
e3 = Entry(window, textvariable=address, width=45)
e3.grid(row=13, column=1)

tel_number = StringVar()
e3 = Entry(window, textvariable=tel_number, width=45)
e3.grid(row=14, column=1)


def rad_but_show():
    print(f'Выбраны профессии {NONE}')


'''rb_status = 0
rb = Radiobutton(window, text='Профессии', variable=rb_status, command=rad_but_show(), state=NORMAL)
rb.grid(row=3, column=3)

rb1 = Radiobutton(window, text='Специальности', variable=rb_status, command=rad_but_show(), state=NORMAL)
rb1.grid(row=4, column=3)

rb2 = Radiobutton(window, text='Специальности с экз', variable=rb_status, command=rad_but_show(), state=NORMAL)
rb2.grid(row=5, column=3)
'''


def show_message():
    print(date_id_doc.get())


# создаём кнопки действий и привязываем их к своим функциям
# кнопки размещаем тоже по сетке
'''b1 = Button(window, text="Посмотреть все", width=12)
b1.grid(row=2, column=3)  # size of the button

b2 = Button(window, text="Поиск", width=12)
b2.grid(row=3, column=3)

b3 = Button(window, text="Добавить", width=12)
b3.grid(row=4, column=3)

b4 = Button(window, text="Обновить", width=12)
b4.grid(row=5, column=3)

b5 = Button(window, text="Удалить", width=12)
b5.grid(row=6, column=3)

b6 = Button(window, text="Закрыть", width=12, command=on_closing)
b6.grid(row=7, column=3)'''


def fill():
    '''Заполнить документ Word информацией введенной из интерфейса'''
    doc = DocxTemplate('prof.docx')
    context = {'reg_number': reg_number.get(),
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
               'day_doc': date_id_doc.get().split('.')[0],
               'month_doc': date_id_doc.get().split('.')[1],
               'year_doc': date_id_doc.get().split('.')[1][-2:],
               'office_doc': office_doc.get(),
               'address': address.get(),
               'tel_number': tel_number.get()
               }
    doc.render(context)
    doc.save("res.docx")


btn = Button(window, text="Заполнить", width=12, command=fill)
btn.grid(row=7, column=7)

# пусть окно работает всё время до закрытия
window.mainloop()
