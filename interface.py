from tkinter import *
from tkinter import filedialog, messagebox
from main import generate_document
import openpyxl
from my_functions import get_data, _onKeyRelease, check_entry

root = Tk()

def check_contractTypeRB():
    if contractTypeRB.get() == '0':
        contractTypeRB1.config(bg="#FFA2A2")
        contractTypeRB2.config(bg="#FFA2A2")
        contractTypeRB3.config(bg="#FFA2A2")
        contractTypeRB4.config(bg="#FFA2A2")
    else:
        contractTypeRB1.config(bg="white")
        contractTypeRB2.config(bg="white")
        contractTypeRB3.config(bg="white")
        contractTypeRB4.config(bg="white")

def check_ourLegalrb():
    if ourLegalrb.get() == '0':
        ourLegalrb1.config(bg="#FFA2A2")
        ourLegalrb2.config(bg="#FFA2A2")
    else:
        ourLegalrb1.config(bg="white")
        ourLegalrb2.config(bg="white")

def open_file_dialog(): # Функция открывает файл и забирает данные из него
    # Выбираем файл
    file_path = filedialog.askopenfilename()
    # Открываем файл
    workbook = openpyxl.load_workbook(file_path)
    # Выбираем активный лист
    sheet = workbook.active
    # Копируем значение ячейки C5
    website = sheet['C5'].value
    inn = sheet['C17'].value
    bik = sheet['C22'].value
    checkingAccount = sheet['C19'].value
    managerName = sheet['C23'].value
    managerEmail = sheet['C24'].value
    contactPerson = sheet['C25'].value
    contactPersonPhone = sheet['C26'].value
    contactPersonEmail = sheet['C27'].value
    ContragentNumber = sheet['C38'].value
    ourLegal = sheet['C41'].value
    set_text(website, inn, bik, checkingAccount, managerName, managerEmail, contactPerson, contactPersonPhone, contactPersonEmail, ContragentNumber, ourLegal)
    check_all()


# функция заполняет поля значениями
def set_text(website, inn, bik, checkingAccount, managerName, managerEmail, contactPerson, contactPersonPhone, contactPersonEmail, ContragentNumber, ourLegal):
    websiteField.delete(0, END)
    websiteField.insert(0, website)
    innField.delete(0, END)
    innField.insert(0, inn)
    bikField.delete(0, END)
    bikField.insert(0, bik)
    checkingAccountField.delete(0, END)
    checkingAccountField.insert(0, checkingAccount)
    managerNameField.delete(0, END)
    managerNameField.insert(0, managerName)
    managerEmailField.delete(0, END)
    managerEmailField.insert(0, managerEmail)
    contactPersonField.delete(0, END)
    contactPersonField.insert(0, contactPerson)
    contactPersonPhoneField.delete(0, END)
    contactPersonPhoneField.insert(0, contactPersonPhone)
    contactPersonEmailField.delete(0, END)
    contactPersonEmailField.insert(0, contactPersonEmail)
    contractDateField.delete(0, END)
    contractDateField.insert(0, get_data())
    ContragentNumberField.delete(0, END)
    ContragentNumberField.insert(0, ContragentNumber)
    print(ourLegal)
    if ourLegal == 'ЦИТ':
        ourLegalrb1.select()
    else:
        ourLegalrb2.select()
    costField.delete(0, END)
    # costField.insert(0, '60000')
    check_all()
def generate_docx():
    # Получаем данные от пользователя
    send_inn = innField.get()
    website = websiteField.get()
    send_bik = bikField.get()
    checking_account = checkingAccountField.get()
    manager_name = managerNameField.get()
    manager_email = managerEmailField.get()
    contact_person = contactPersonField.get()
    contact_person_phone = contactPersonPhoneField.get()
    contact_person_email = contactPersonEmailField.get()
    contract_type = contractTypeRB.get()
    cost = costField.get()
    contract_date = contractDateField.get()
    contract_number = contractNumberField.get()
    contragent_number = ContragentNumberField.get()
    conditions = conditionsField.get()
    our_legal = ourLegalrb.get()

    #     messagebox.showinfo("Ошибка", "Тип договора не выбран")
    # Тут блок с проверкой валидности ввода данных
    if len(str(send_bik)) != 9:
        messagebox.showinfo("Ошибка", "БИК неверный. БИК должен содержать 9 цифр.")
        return
    check_all()
    # тут можно написать функцию проверки корректности заполнения полей
    # Генерируем документ
    generate_document(send_inn, website, contact_person, send_bik, checking_account, manager_name, manager_email, \
                      contact_person_phone, contact_person_email, contract_type, cost, contract_date, contract_number, contragent_number, conditions, our_legal)

root.title('Generator docx file')
root.geometry('400x772+2000+150')
root.bind_all("<Key>", _onKeyRelease, "+")

# Создаем фрейм (область для размещения других объектов)
frame_top = Frame(root, bg='#313830', bd=5)
frame_top.place(relwidth=1, relheight=1)

# Заголовок
HeaderLabel = Label(frame_top, text='Генератор документов', bg='#ffb700', font=40)
HeaderLabel.place(x=-3, y=-3, width=395, height=28)

# Сайт текст
websiteLabel = Label(frame_top, text='Сайт', bg='#ffb700', font=40)
websiteLabel.place(x=-3, y=30, width=148, height=35)

# Сайт поле
websiteField = Entry(frame_top, bg='white', font=30)
websiteField.bind("<FocusOut>", lambda event: check_entry(websiteField))
websiteField.place(x=147, y=30, width=245, height=35)

# ИНН текст
innLabel = Label(frame_top, text='ИНН', bg='#ffb700', font=40)
innLabel.place(x=-3, y=70, width=148, height=35)

# ИНН поле
innField = Entry(frame_top, bg='white', font=30)
innField.bind("<FocusOut>", lambda event: check_entry(innField))
innField.place(x=147, y=70, width=245, height=35)

# БИК текст
bikLabel = Label(frame_top, text='БИК', bg='#ffb700', font=40)
bikLabel.place(x=-3, y=110, width=148, height=35)

# БИК поле
bikField = Entry(frame_top, bg='white', font=30)
bikField.bind("<FocusOut>", lambda event: check_entry(bikField))
bikField.place(x=147, y=110, width=245, height=35)

# Расчетный счет текст
checkingAccountLabel = Label(frame_top, text='Расчетный счет', bg='#ffb700', font=40)
checkingAccountLabel.place(x=-3, y=150, width=148, height=35)

# Расчетный счет поле
checkingAccountField = Entry(frame_top, bg='white', font=30)
checkingAccountField.bind("<FocusOut>", lambda event: check_entry(checkingAccountField))
checkingAccountField.place(x=147, y=150, width=245, height=35)

# Менеджер ФИО текст
managerNameLabel = Label(frame_top, text='Менеджер ФИО', bg='#ffb700', font=40)
managerNameLabel.place(x=-3, y=190, width=148, height=35)

# Менеджер ФИО поле
managerNameField = Entry(frame_top, bg='white', font=30)
managerNameField.bind("<FocusOut>", lambda event: check_entry(managerNameField))
managerNameField.place(x=147, y=190, width=245, height=35)

# Менеджер емайл текст
managerEmailLabel = Label(frame_top, text='Менеджер e-mail', bg='#ffb700', font=40)
managerEmailLabel.place(x=-3, y=230, width=148, height=35)

# Менеджер емайл поле
managerEmailField = Entry(frame_top, bg='white', font=30)
managerEmailField.bind("<FocusOut>", lambda event: check_entry(managerEmailField))
managerEmailField.place(x=147, y=230, width=245, height=35)

# Контактное лицо текст
contactPersonLabel = Label(frame_top, text='Контактное лицо', bg='#ffb700', font=40)
contactPersonLabel.place(x=-3, y=270, width=148, height=35)

# Контактное лицо поле
contactPersonField = Entry(frame_top, bg='white', font=30)
contactPersonField.bind("<FocusOut>", lambda event: check_entry(contactPersonField))
contactPersonField.place(x=147, y=270, width=245, height=35)

# Контактное лицо телефон текст
contactPersonPhoneLabel = Label(frame_top, text='Контактное лицо\nтелефон', bg='#ffb700', font=40)
contactPersonPhoneLabel.place(x=-3, y=310, width=148, height=35)

# Контактное лицо телефон поле
contactPersonPhoneField = Entry(frame_top, bg='white', font=30)
contactPersonPhoneField.bind("<FocusOut>", lambda event: check_entry(contactPersonPhoneField))
contactPersonPhoneField.place(x=147, y=310, width=245, height=35)

# Контактное лицо емайл текст
contactPersonEmailLabel = Label(frame_top, text='Контактное лицо\ne-mail', bg='#ffb700', font=40)
contactPersonEmailLabel.place(x=-3, y=350, width=148, height=35)

# Контактное лицо емайл поле
contactPersonEmailField = Entry(frame_top, bg='white', font=30)
contactPersonEmailField.bind("<FocusOut>", lambda event: check_entry(contactPersonEmailField))
contactPersonEmailField.place(x=147, y=350, width=245, height=35)

# Тип договора текст
contractTypeLabel = Label(frame_top, text='Тип договора', bg='#ffb700', font=40)
contractTypeLabel.place(x=-3, y=390, width=148, height=85)

# Тип договора поле
contractTypeRB = StringVar(value=0)
contractTypeRB1 = Radiobutton(frame_top, text='SEO. Без этапов гарантий и бонуса', variable=contractTypeRB, value='SEO. Без этапов гарантий и бонуса', anchor='w', command=check_contractTypeRB)
contractTypeRB2 = Radiobutton(frame_top, text='Контекст. Яндекс Директ', variable=contractTypeRB, value='Контекст. Яндекс Директ', anchor='w', command=check_contractTypeRB)
contractTypeRB3 = Radiobutton(frame_top, text='Лицензия Битрикс', variable=contractTypeRB, value='Лицензия Битрикс', anchor='w', command=check_contractTypeRB)
contractTypeRB4 = Radiobutton(frame_top, text='Разработка. Создание сайта', variable=contractTypeRB, value='Разработка. Создание сайта', anchor='w', command=check_contractTypeRB)
contractTypeRB1.place(x=147, y=390, width=245)
contractTypeRB2.place(x=147, y=410, width=245)
contractTypeRB3.place(x=147, y=430, width=245)
contractTypeRB4.place(x=147, y=450, width=245)

# Стоимость текст
costLabel = Label(frame_top, text='Стоимость', bg='#ffb700', font=40)
costLabel.place(x=-3, y=480, width=148, height=35)

# Стоимость поле
costField = Entry(frame_top, bg='white', font=30)
costField.bind("<FocusOut>", lambda event: check_entry(costField))
costField.place(x=147, y=480, width=245, height=35)

# Дата договора текст
contractDateLabel = Label(frame_top, text='Дата договора', bg='#ffb700', font=40)
contractDateLabel.place(x=-3, y=520, width=148, height=35)

# Дата договора поле
contractDateField = Entry(frame_top, bg='white', font=30)
contractDateField.bind("<FocusOut>", lambda event: check_entry(contractDateField))
contractDateField.place(x=147, y=520, width=245, height=35)
#
# Номер заказчика текст
ContragentNumberLabel = Label(frame_top, text='Номер заказчика', bg='#ffb700', font=40)
ContragentNumberLabel.place(x=-3, y=560, width=148, height=35)

# Номер заказчика поле
ContragentNumberField = Entry(frame_top, bg='white', font=30)
ContragentNumberField.bind("<FocusOut>", lambda event: check_entry(ContragentNumberField))
ContragentNumberField.place(x=147, y=560, width=245, height=35)

# Порядковый номер договора текст
contractNumberLabel = Label(frame_top, text='Порядковый номер\nдоговора', bg='#ffb700', font=40)
contractNumberLabel.place(x=-3, y=600, width=148, height=35)

# Порядковый номер договора поле
contractNumberField = Entry(frame_top, bg='white', font=30)
contractNumberField.bind("<FocusOut>", lambda event: check_entry(contractNumberField))
contractNumberField.place(x=147, y=600, width=245, height=35)

# Особые условия договора текст
conditionsLabel = Label(frame_top, text='Особые условия\nдоговора', bg='#ffb700', font=40)
conditionsLabel.place(x=-3, y=640, width=148, height=35)

# Особые условия договора поле
conditionsField = Entry(frame_top, bg='white', font=30)
# contractNumberField.bind("<FocusOut>", lambda event: check_entry(contractNumberField))
conditionsField.place(x=147, y=640, width=245, height=35)

# Наше юр лицо текст
ourLegalEntityLabel = Label(frame_top, text='Наше юр лицо', bg='#ffb700', font=40)
ourLegalEntityLabel.place(x=-3, y=680, width=148, height=45)

# Наше юр лицо поле
ourLegalrb = StringVar(value=0)
# Listbox(frame_top, bg='white', font=30)
ourLegalrb1 = Radiobutton(frame_top, text='ЦИТ "Информ-С"', variable=ourLegalrb, value='ЦИТ "Информ-С"', anchor='w', command=check_ourLegalrb)
ourLegalrb2 = Radiobutton(frame_top, text='ПИАЦ "Информ-С"', variable=ourLegalrb, value='ПИАЦ "Информ-С"', anchor='w', command=check_ourLegalrb)
ourLegalrb1.place(x=147, y=680, width=245)
ourLegalrb2.place(x=147, y=700, width=245)

# Создаем кнопку и при нажатии будет срабатывать метод "set_text"
btn = Button(frame_top, text='Заполнить поля\nиз карточки', command=open_file_dialog)
btn.place(x=-3, y=730, width=148, height=35)

# Создаем кнопку и при нажатии будет срабатывать метод "generate_docx"
btn = Button(frame_top, text='Создать\nдоговор', command=generate_docx)
btn.place(x=147, y=730, width=245, height=35)

# # Создаем кнопку и при нажатии будет срабатывать метод "generate_xlsx"
# btn = Button(frame_top, text='Сохранить\nкарточку', command=generate_xlsx)
# btn.place(x=-3, y=770, width=148, height=35)

def check_all():
    check_entry(websiteField)
    check_entry(innField)
    check_entry(bikField)
    check_entry(checkingAccountField)
    check_entry(managerNameField)
    check_entry(managerEmailField)
    check_entry(contactPersonField)
    check_entry(contactPersonPhoneField)
    check_entry(contactPersonEmailField)
    check_entry(costField)
    check_entry(contractDateField)
    check_entry(contractNumberField)
    check_entry(ContragentNumberField)
    check_contractTypeRB()
    check_ourLegalrb()


root.mainloop()