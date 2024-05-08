import datetime
import interface
# Функция формирования даты
def get_data():
    today = datetime.date.today()
    den = today.day
    mesjac = today.month
    god = today.year

    if den <= 9:
        den = '0' + str(den)

    mesjac_text = {
        1: "января",
        2: "февраля",
        3: "марта",
        4: "апреля",
        5: "мая",
        6: "июня",
        7: "июля",
        8: "августа",
        9: "сентября",
        10: "октября",
        11: "ноября",
        12: "декабря"
    }

    if mesjac in mesjac_text:
        mesjac = mesjac_text[mesjac]
    else:
        mesjac = "неизвестный месяц"

    return('«' + str(den) + '» ' + mesjac + ' ' + str(god) + ' г.')

def check_entry(entry):
    if entry.get() == "":
        entry.config(bg="#FFA2A2")
    else:
        entry.config(bg="white")



# Функция вставки значения из буфера обмена при любой раскладке клавиатуры
def _onKeyRelease(event):
    ctrl  = (event.state & 0x4) != 0
    if event.keycode==88 and  ctrl and event.keysym.lower() != "x":
        event.widget.event_generate("<<Cut>>")

    if event.keycode==86 and  ctrl and event.keysym.lower() != "v":
        event.widget.event_generate("<<Paste>>")

    if event.keycode==67 and  ctrl and event.keysym.lower() != "c":
        event.widget.event_generate("<<Copy>>")

