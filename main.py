# подключаем графическую библиотеку для создания интерфейсов
from tkinter import *
from docxtpl import DocxTemplate
from tkinter import messagebox

# подключаем графическую библиотеку
window = Tk()
# заголовок окна
window.title("Мой бухгалтер")


# обрабатываем закрытие окна
def on_closing():
    # показываем диалоговое окно с кнопкой
    if messagebox.askokcancel("", "Закрыть программу?"):
        # удаляем окно и освобождаем память
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
e4 = Entry(window, textvariable=date_birthday)
e4.grid(row=4, column=1)

snils = StringVar()
e5 = Entry(window, textvariable=snils)
e5.grid(row=5, column=1)

inn = StringVar()
e6 = Entry(window, textvariable=inn)
e6.grid(row=6, column=1)

citizenship = StringVar()
e7 = Entry(window, textvariable=citizenship)
e7.grid(row=7, column=1)

id_doc = StringVar()
e8 = Entry(window, textvariable=id_doc)
e8.grid(row=8, column=1)

series = StringVar()
e9 = Entry(window, textvariable=series)
e9.grid(row=9, column=1)

number = StringVar()
e10 = Entry(window, textvariable=number)
e10.grid(row=11, column=1)

date_id_doc = StringVar()
e11 = Entry(window, textvariable=date_id_doc)
e11.grid(row=12, column=1)

office_doc = StringVar()
e12 = Entry(window, textvariable=office_doc, width=45)
e12.grid(row=12, column=2)

address = StringVar()
e13 = Entry(window, textvariable=address, width=45)
e13.grid(row=13, column=1)

tel_number = StringVar()
e14 = Entry(window, textvariable=tel_number, width=45)
e14.grid(row=14, column=1)

profession = 'Профессия'
specialty = 'Специальность'
specialty_with_exam = 'Специальность c экзаменом'

choice = StringVar(value=profession)

rb = Radiobutton(window, text=profession, value=profession, variable=choice)
rb.grid(row=3, column=3)

rb1 = Radiobutton(window, text=specialty, value=specialty, variable=choice)
rb1.grid(row=4, column=3)

rb2 = Radiobutton(window, text=specialty_with_exam, value=specialty_with_exam, variable=choice)
rb2.grid(row=5, column=3)





def fill():
    '''Заполнить документ Word информацией введенной из интерфейса'''

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
               'day_doc': date_id_doc.get(),
               'month_doc': date_id_doc.get(),
               'year_doc': date_id_doc.get(),
               'office_doc': office_doc.get(),
               'address': address.get(),
               'tel_number': tel_number.get()
               }
    if choice == profession:
        doc = DocxTemplate('patterns/prof.docx')
    elif choice == specialty:
        doc = DocxTemplate('patterns/special.docx')
    else:
        doc = DocxTemplate('patterns/sp_with_ex.docx')
    doc.render(context)
    doc.save(f'{surname.get()} {name.get()} {patronymic.get()}.docx')
    '''    '''


btn = Button(window, text="Заполнить", width=12, command=fill)
btn.grid(row=7, column=7)

# пусть окно работает всё время до закрытия
window.mainloop()
