from docxtpl import DocxTemplate
from dadata import Dadata # это подгрузка данных юр лиц и банка https://github.com/hflabs/dadata-py
from interface import *
from num2words import num2words
import pymorphy2
import openpyxl

# 6316102218 ПИАЦ бик 044525411
# 7715598102 ЦИТ
def generate_document(send_inn, website, contact_person, send_bik, checking_account, manager_name, manager_email, \
                      contact_person_phone, contact_person_email, contract_type, cost, contract_date, contract_number, \
                      contragent_number, conditions, our_legal):
    # Собираем недостающие данные
    # Делаем запрос данных об организации (код функции в файле get_rq) и сохряняем их в переменную data
    dadata = Dadata('0fc7d60da65943f6aa3ba2f4a289b50bc024d18f') # токен сервиса дадата
    result_bank = dadata.find_by_id("bank", send_bik) # делаем запрос в дадату
    result_party = dadata.find_by_id("party", send_inn) # делаем запрос в дадату
    # Создаем номер договора
    if contract_type == 'SEO. Без этапов гарантий и бонуса':
        number = contragent_number + '/ПДВ-1/' + contract_number
    elif contract_type == 'Контекст. Яндекс Директ':
        number = contragent_number + '/АКНТ-1/' + contract_number
    elif contract_type == 'Разработка. Создание сайта':
        number = contragent_number + '/СЗД/' + contract_number
    else:
        number = contragent_number + '/РАЗ/' + contract_number

    # Собираем словарь, из которого будем рендерить документ DOCX
    # Данные заказчика
    context = {}
    context['Сайт'] = website
    context['Наименование_полное'] = result_party[0]['data']['name']['full_with_opf'] # Полное название
    context['Наименование_краткое'] = result_party[0]['data']['name']['short_with_opf'] # Краткое название
    context['Должность_руководителя'] = result_party[0]['data']['management']['post'].lower().capitalize() # Должность
    context['Должность_руководителя_РП'] = director(result_party[0]['data']['management']['post'])
    context['ФИО_руководителя'] = result_party[0]['data']['management']['name'].lower().capitalize() # ФИО
    context['ФИО_руководителя_РП'] = padej(result_party[0]['data']['management']['name'])
    context['ФИО_руководителя_кратко'] = shorten_fio(result_party[0]['data']['management']['name'])
    context['пол'] = gender(result_party[0]['data']['management']['name'])
    context['Юр_адрес'] = result_party[0]['data']['address']['unrestricted_value'] # Юр. адрес
    context['ИНН'] = 'ИНН ' + send_inn
    context['ОГРН'] = 'ОГРН ' + result_party[0]['data']['ogrn'] # ОГРН
    context['КПП'] = 'КПП ' + result_party[0]['data']['kpp'] # КПП
    context['Расчетный_счет'] = 'Р/с ' + checking_account
    context['Наименование_банка'] = result_bank[0]['data']['name']['payment'] # Банк
    context['Кор_счет'] = 'К/с ' + result_bank[0]['data']['correspondent_account'] # Кор.счет
    context['БИК'] = ', БИК ' + send_bik
    context['Менеджер_проекта'] = manager_name + ", +7(846)300-27-99, " + manager_email
    context['Контактное_лицо'] = contact_person + ", " + contact_person_phone + ", " + contact_person_email
    context['Стоимость'] = str(cost) + ' (' + num2words(int(cost), lang='ru') + ') ' + rubl(int(cost))
    context['Дата_договора'] = contract_date
    context['Номер_заказчика'] = contragent_number
    context['Номер_договора'] = contract_number
    context['Особое_условие'] = conditions
    # Подставляем наши реквизиты
    context['OUR_ФИО_руководителя'] = "Мухитова Юлия Маратовна"  # ФИО
    context['OUR_ФИО_руководителя_РП'] = "Мухитовой Юлии Маратовны"
    context['OUR_ФИО_руководителя_кратко'] = "Мухитова Ю. М."
    if our_legal == 'ЦИТ "Информ-С"':
        context['OUR_Наименование_полное'] = 'Общество с ограниченной ответственностью Центр интернет-технологий «Информ-С»'  # Полное название
        context['OUR_Наименование_краткое'] = 'ООО Центр интернет-технологий «Информ-С»'  # Краткое название
        context['OUR_Должность_руководителя'] = 'Генеральный директор'  # Должность
        context['OUR_Должность_руководителя_РП'] = 'Генерального директора'
        context['OUR_Юр_адрес'] = "443068, Самарская область, г. Самара, Ново-Садовая ул, д. 106, офис 33"  # Юр. адрес
        context['OUR_ИНН'] = 'ИНН 7715598102'
        context['OUR_ОГРН'] = 'ОГРН 1067746480042'  # ОГРН
        context['OUR_КПП'] = 'КПП 631601001'  # КПП
        context['OUR_Расчетный_счет'] = 'Р/с 40702810501300002309'
        context['OUR_Наименование_банка'] = "АО «АЛЬФА-БАНК», г. Москва"  # Банк
        context['OUR_Кор_счет'] = 'К/с 30101810200000000593'  # Кор.счет
        context['OUR_БИК'] = 'БИК 044525593'
    elif our_legal == 'ПИАЦ "Информ-С"':
        context['OUR_Наименование_полное'] = 'Общество с ограниченной ответственностью ПИАЦ «Информ-С»'
        context['OUR_Наименование_краткое'] = 'ООО ПИАЦ «Информ-С»'  # Краткое название
        context['OUR_Должность_руководителя'] = "Директор"  # Должность
        context['OUR_Должность_руководителя_РП'] = "Директора"
        context['OUR_Юр_адрес'] = "443068, Самарская область, г. Самара, Ново-Садовая ул, д. 106, офис 28"  # Юр. адрес
        context['OUR_ИНН'] = 'ИНН 6316102218'
        context['OUR_ОГРН'] = 'ОГРН 1056316047028'  # ОГРН
        context['OUR_КПП'] = 'КПП 631601001'   # КПП
        context['OUR_Расчетный_счет'] = 'Р/с 40702810913180002984'
        context['OUR_Наименование_банка'] = 'ФИЛИАЛ "ЦЕНТРАЛЬНЫЙ" БАНКА ВТБ (ПАО)'  # Банк
        context['OUR_Кор_счет'] = 'К/с 30101810145250000411'  # Кор.счет
        context['OUR_БИК'] = 'БИК 044525411'


    doc = DocxTemplate("templates/" + contract_type + ".docx")
    # подставляем контекст в шаблон
    doc.render(context)
    # сохраняем и смотрим, что получилось
    contract_date_yhar = contract_date[-7:]
    doc.save(contract_date_yhar[:4] + "_" + number.replace('/','_') + '_' + website.replace('/','_') + ".docx")
    print(result_party[0]['value'])
    print(our_legal)
    print(contract_type)
    print(number)

    # # ============ Эта часть кода генерирует xlsx =======================
    # workbook = openpyxl.load_workbook('template.xlsx') # Открываем существующий файл xlsx
    # sheet = workbook.active  # Выбираем активный лист
    # # Записываем значение из переменной в ячейки
    # value = 'Hello'
    # sheet['C4'] = value
    # sheet['C5'] = website
    # sheet['C6'] = value
    # sheet['C7'] = value
    # sheet['C8'] = value
    # sheet['C9'] = value
    # sheet['C10'] = value
    # sheet['C11'] = value
    # sheet['C12'] = 'Устава'

    # # Сохраняем файл
    # workbook.save('example.xlsx')
    # ============= конец генерации xlsx ============

# Эта функция склоняет слово рубль в зависимости от числа
def rubl(var):
    if var % 10 == 1 and var % 100 != 11:
        rubles = 'рубль'
    elif var % 10 in [2, 3, 4] and var % 100 not in [12, 13, 14]:
        rubles = 'рубля'
    else:
        rubles = 'рублей'
    return rubles

# Эта функция определяет пол подписанта и ставит слово действующего/ей
def gender(fio):
    parts = fio.split(" ")
    if len(parts) == 3:
        surname, name, patronymic = parts
        if patronymic.endswith("ич"):
            gender = "действующего"
        elif patronymic.endswith("на"):
            gender = "действующей"
        else:
            gender = "действующего/ей"
    else:
        gender = "действующего/ей"
    return gender


# Эта функция слоняет ФИО в родительный падеж
morph = pymorphy2.MorphAnalyzer()
def padej(fio):
    parts = fio.split()
    remaining_words = ' '.join(parts[3:])  # если ФИО содержит больше трех слов, тогда хвостик мы сохраняем в отдельную переменную
    if len(remaining_words) != 0:
        remaining_words = " " + remaining_words
    else:
        remaining_words = ""
    fio = ' '.join(parts[:3])
    parts = fio.split(" ")
    if len(parts) == 3:
        surname, name, patronymic = parts
        surname = morph.parse(surname)[0].inflect({"gent"}).word.title()
        name = morph.parse(name)[0].inflect({"gent"}).word.title()
        patronymic = morph.parse(patronymic)[0].inflect({"gent"}).word.title()
        fio_gen = f"{surname} {name} {patronymic}" + remaining_words
    else:
        fio_gen = "неизвестное ФИО"
    return fio_gen

# Эта функция склоняет должность в родительный падеж
def director(doljnost):
    doljnost = doljnost.lower().capitalize()
    if "Генеральный директор" in doljnost:
        new_doljnost = 'Генерального директора'
    elif "Директор" in doljnost:
        new_doljnost = "Директора"
    else:
        new_doljnost = doljnost
    return new_doljnost

def shorten_fio(full_name):
    name_parts = full_name.split()
    shortened_name = name_parts[0] + " " + name_parts[1][0] + ". " + name_parts[2][0] + "."
    return shortened_name
