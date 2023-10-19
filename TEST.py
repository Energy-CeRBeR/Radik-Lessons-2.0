import telebot
from telebot import types
import openpyxl
from openpyxl.reader.excel import load_workbook
import time
import string


#bot = telebot.TeleBot("6011141241:AAEp8y5Yyul5jGpLLZc7XxDYq9lhlgTNxoU")
bot = telebot.TeleBot("5754400160:AAFvJI6-SIXzUnT5nXOQhWils_QwKAIyTj4")


@bot.message_handler(commands=["start"])
def start(message):
    #print("start")
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn1 = types.InlineKeyboardButton("Узнать расписание", callback_data="fac")
    markup.add(btn1)
    bot.send_message(message.chat.id, text="Добро пожаловать в РГРТУ!", reply_markup=markup)

@bot.callback_query_handler(func=lambda callback: callback.data == "fac")
def selectFac(callback):
    #print("selectFac", callback.data)
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn1 = types.InlineKeyboardButton("ФРТ", callback_data="frt")
    btn2 = types.InlineKeyboardButton("ФЭ", callback_data="fe")
    btn3 = types.InlineKeyboardButton("ФАИТУ", callback_data="faitu")
    btn4 = types.InlineKeyboardButton("ФВТ", callback_data="fvt")
    btn5 = types.InlineKeyboardButton("ИЭФ", callback_data="ief")
    home = types.InlineKeyboardButton("В главное меню", callback_data="home")
    markup.row(btn1, btn2)
    markup.row(btn3, btn4, btn5)
    markup.add(home)
    bot.send_message(callback.message.chat.id, text="Выберите факультет", reply_markup=markup)

fac = ""
@bot.callback_query_handler(func=lambda callback: callback.data in ["frt", "fe", "faitu", "fvt", "ief", "home"])
def selectCourse(callback):
    #print("selectCourse", callback.data)
    if callback.data == "home":
        markup = types.InlineKeyboardMarkup(row_width=1)
        btn1 = types.InlineKeyboardButton("Узнать расписание", callback_data="fac")
        markup.add(btn1)
        bot.send_message(callback.message.chat.id, text="Добро пожаловать в РГРТУ!", reply_markup=markup)

    else:
        global fac
        fac = callback.data
        markup = types.InlineKeyboardMarkup(row_width=1)
        btn1 = types.InlineKeyboardButton("1 курс", callback_data="course_1")
        btn2 = types.InlineKeyboardButton("2 курс", callback_data="course_2")
        btn3 = types.InlineKeyboardButton("3 курс", callback_data="course_3")
        btn4 = types.InlineKeyboardButton("4 курс", callback_data="course_4")
        back = types.InlineKeyboardButton("Вернуться назад", callback_data="fac")
        home = types.InlineKeyboardButton("В главное меню", callback_data="home")
        markup.row(btn1, btn2)
        if fac in ["frt", "fvt", "ief"]:
            btn5 = types.InlineKeyboardButton("5 курс", callback_data="course_5")
            markup.row(btn3, btn4, btn5)
        else:
            markup.row(btn3, btn4)
        markup.add(back)
        markup.add(home)
        bot.send_message(callback.message.chat.id, "Выберите курс", reply_markup=markup)


wb = ""
sheet = ""
lines_days = {"понедельник": [4, 11], "вторник": [12, 19], "среда": [20, 27], "четверг": [28, 37],
              "пятница": [38, 45], "суббота": [46, 53]}

letters = string.ascii_uppercase[3:]

column_groups = dict()
course = ""
@bot.callback_query_handler(func=lambda callback: callback.data in ["course_1", "course_2", "course_3", "course_4",
                                                                    "course_5", "fac", "home"])
def selectGroup(callback):
    #print("selectGroup", callback.data)
    if callback.data == "fac":
        markup = types.InlineKeyboardMarkup(row_width=1)
        btn1 = types.InlineKeyboardButton("ФРТ", callback_data="frt")
        btn2 = types.InlineKeyboardButton("ФЭ", callback_data="fe")
        btn3 = types.InlineKeyboardButton("ФАИТУ", callback_data="faitu")
        btn4 = types.InlineKeyboardButton("ФВТ", callback_data="fvt")
        btn5 = types.InlineKeyboardButton("ИЭФ", callback_data="ief")
        back = types.InlineKeyboardButton("Вернуться назад", callback_data="fac")
        home = types.InlineKeyboardButton("В главное меню", callback_data="home")
        markup.row(btn1, btn2)
        markup.row(btn3, btn4, btn5)
        markup.add(back, home)
        bot.send_message(callback.message.chat.id, text="Выберите факультет", reply_markup=markup)

    elif callback.data == "home":
        markup = types.InlineKeyboardMarkup(row_width=1)
        btn1 = types.InlineKeyboardButton("Узнать расписание", callback_data="fac")
        markup.add(btn1)
        bot.send_message(callback.message.chat.id, text="Добро пожаловать в РГРТУ!", reply_markup=markup)

    else:
        global course, fac, column_groups, wb, sheet
        column_groups = dict()
        course = callback.data[-1]
        files = {"frt": ["frt_1", "frt_2", "frt_3", "frt_4", "frt_5"],
                 "fe": ["fe_1", "fe_2", "fe_3", "fe_4", "fe_5"],
                 "faitu": ["faitu_1", "faitu_2", "faitu_3", "faitu_4", "faitu_5"],
                 "fvt": ["fvt_1", "fvt_2", "fvt_3", "fvt_4", "fvt_5"],
                 "ief": ["ief_1", "ief_2", "ief_3", "ief_4", "ief_5"]}

        wb = openpyxl.reader.excel.load_workbook(filename=f"{files[fac][int(course) - 1]}.xlsx")
        wb.active = 0
        sheet = wb.active

        for l in letters:
            cell = f"{l}3"
            if sheet[cell].value is not None:
                column_groups[sheet[cell].value] = cell[0]
            else:
                break

        markup = types.InlineKeyboardMarkup(row_width=1)
        buttons = []
        for gr in column_groups.keys():
            btn = types.InlineKeyboardButton(gr, callback_data=gr)
            buttons.append(btn)

        back = types.InlineKeyboardButton("Вернуться назад", callback_data="course")
        home = types.InlineKeyboardButton("В главное меню", callback_data="home")
        markup.row(*buttons[:len(buttons) // 2])
        markup.row(*buttons[len(buttons) // 2:])
        markup.add(home)
        markup.add(back)
        bot.send_message(callback.message.chat.id, "Выберите номер вашей группы:", reply_markup=markup)


number = ""
@bot.callback_query_handler(func=lambda callback: callback.data in column_groups.keys()
                                               or callback.data in ["course", "home"])
def selectType(callback):
    #print("selectType", callback.data)
    if callback.data == "course":
        markup = types.InlineKeyboardMarkup(row_width=1)
        btn1 = types.InlineKeyboardButton("1 курс", callback_data="course_1")
        btn2 = types.InlineKeyboardButton("2 курс", callback_data="course_2")
        btn3 = types.InlineKeyboardButton("3 курс", callback_data="course_3")
        btn4 = types.InlineKeyboardButton("4 курс", callback_data="course_4")
        back = types.InlineKeyboardButton("Вернуться назад", callback_data="fac")
        home = types.InlineKeyboardButton("В главное меню", callback_data="home")
        markup.row(btn1, btn2)
        if fac in ["frt", "fvt", "ief"]:
            btn5 = types.InlineKeyboardButton("5 курс", callback_data="course_5")
            markup.row(btn3, btn4, btn5)
        else:
            markup.row(btn3, btn4)
        markup.add(back)
        markup.add(home)
        bot.send_message(callback.message.chat.id, "Выберите курс", reply_markup=markup)

    elif callback.data == "home":
        markup = types.InlineKeyboardMarkup(row_width=1)
        btn1 = types.InlineKeyboardButton("Узнать расписание", callback_data="group")
        markup.add(btn1)
        bot.send_message(callback.message.chat.id, text="Добро пожаловать в РГРТУ!", reply_markup=markup)

    else:
        global number
        number = callback.data
        markup = types.InlineKeyboardMarkup(row_width=1)
        type1 = types.InlineKeyboardButton("Числитель", callback_data="числитель")
        type2 = types.InlineKeyboardButton("Знаменатель", callback_data="знаменатель")
        back = types.InlineKeyboardButton("Вернуться назад", callback_data="back_to_groups")
        home = types.InlineKeyboardButton("В главное меню", callback_data="home")
        markup.row(type1, type2)
        markup.add(back, home)
        bot.send_message(callback.message.chat.id, "Выберите формат недели:", reply_markup=markup)


type_week = ""
@bot.callback_query_handler(func=lambda callback: callback.data in ["числитель", "знаменатель", "back_to_groups", "home"])
def selectDay(callback):
    #print("selectDay", callback.data)
    if callback.data == "back_to_groups":
        markup = types.InlineKeyboardMarkup(row_width=1)
        buttons = []
        for gr in column_groups.keys():
            btn = types.InlineKeyboardButton(gr, callback_data=gr)
            buttons.append(btn)

        back = types.InlineKeyboardButton("Вернуться назад", callback_data="course")
        home = types.InlineKeyboardButton("В главное меню", callback_data="home")
        markup.row(*buttons[:len(buttons) // 2])
        markup.row(*buttons[len(buttons) // 2:])
        markup.add(back, home)
        bot.send_message(callback.message.chat.id, "Выберите номер вашей группы:", reply_markup=markup)

    elif callback.data == "home":
        markup = types.InlineKeyboardMarkup(row_width=1)
        btn1 = types.InlineKeyboardButton("Узнать расписание", callback_data="group")
        markup.add(btn1)
        bot.send_message(callback.message.chat.id, text="Добро пожаловать в РГРТУ!", reply_markup=markup)
    else:
        global type_week
        type_week = callback.data
        markup = types.InlineKeyboardMarkup(row_width=1)
        day1 = types.InlineKeyboardButton("Понедельник", callback_data="понедельник")
        day2 = types.InlineKeyboardButton("Вторник", callback_data="вторник")
        day3 = types.InlineKeyboardButton("Среда", callback_data="среда")
        day4 = types.InlineKeyboardButton("Четверг", callback_data="четверг")
        day5 = types.InlineKeyboardButton("Пятница", callback_data="пятница")
        day6 = types.InlineKeyboardButton("Суббота", callback_data="суббота")
        back = types.InlineKeyboardButton("Вернуться назад", callback_data="back_to_types")
        home = types.InlineKeyboardButton("В главное меню", callback_data="home")
        markup.row(day1, day2, day3)
        markup.row(day4, day5, day6)
        markup.add(back, home)
        bot.send_message(callback.message.chat.id, "Выберите день недели:", reply_markup=markup)


'''
Столбец D - 3434, Столбец E - 343
Столбец B - время
Столбец C - тип недели
строки 4 - 53 столбцов от D - информация о предмете
Понедельник: 4 - 11
Вторник: 12 - 19
Среда: 20 - 27
Четверг: 28 - 37
Пятница: 38 - 45
Суббота: 46 - 53
'''


frt1 = {"понедельник": [4, 13], "вторник": [14, 23], "среда": [24, 33], "четверг": [34, 43],
              "пятница": [44, 53], "суббота": [54, 59]}
fe1 = {"понедельник": [4, 11], "вторник": [12, 19], "среда": [20, 27], "четверг": [28, 35],
              "пятница": [36, 43], "суббота": [44, 51]}
faitu1 = {"понедельник": [4, 11], "вторник": [12, 19], "среда": [20, 27], "четверг": [28, 35],
              "пятница": [36, 43], "суббота": [44, 44]}
fvt1 = {"понедельник": [4, 11], "вторник": [12, 19], "среда": [20, 27], "четверг": [28, 37],
              "пятница": [38, 45], "суббота": [46, 53]}
ief1 = {"понедельник": [4, 11], "вторник": [12, 19], "среда": [20, 27], "четверг": [28, 35],
              "пятница": [36, 43], "суббота": [44, 44]}

frt2 = {"понедельник": [4, 13], "вторник": [14, 23], "среда": [24, 33], "четверг": [34, 43],
              "пятница": [44, 53], "суббота": [54, 61]}
fe2 = {"понедельник": [4, 11], "вторник": [12, 19], "среда": [20, 27], "четверг": [28, 35],
              "пятница": [36, 43], "суббота": [44, 51]}
faitu2 = {"понедельник": [4, 13], "вторник": [14, 21], "среда": [22, 29], "четверг": [30, 37],
              "пятница": [38, 45], "суббота": [46, 55]}
fvt2 = {"понедельник": [4, 13], "вторник": [14, 23], "среда": [24, 33], "четверг": [34, 43],
              "пятница": [44, 53], "суббота": [54, 61]}
ief2 = {"понедельник": [4, 12], "вторник": [13, 20], "среда": [21, 28], "четверг": [29, 36],
              "пятница": [37, 45], "суббота": [46, 49]}

frt3 = {"понедельник": [4, 15], "вторник": [16, 17], "среда": [18, 29], "четверг": [30, 41],
              "пятница": [42, 53], "суббота": [54, 65]}
fe3 = {"понедельник": [4, 11], "вторник": [12, 19], "среда": [20, 27], "четверг": [28, 37],
              "пятница": [38, 47], "суббота": [48, 55]}
faitu3 = {"понедельник": [4, 13], "вторник": [14, 21], "среда": [22, 31], "четверг": [32, 41],
              "пятница": [42, 43], "суббота": [44, 49]}
fvt3 = {"понедельник": [4, 15], "вторник": [16, 25], "среда": [26, 37], "четверг": [38, 47],
              "пятница": [48, 57], "суббота": [58, 67]}
ief3 = {"понедельник": [4, 11], "вторник": [12, 21], "среда": [22, 31], "четверг": [32, 41],
              "пятница": [42, 43], "суббота": [44, 53]}

frt4 = {"понедельник": [4, 13], "вторник": [14, 25], "среда": [26, 37], "четверг": [38, 39],
              "пятница": [40, 51], "суббота": [52, 61]}
fe4 = {"понедельник": [4, 13], "вторник": [14, 25], "среда": [26, 35], "четверг": [36, 45],
              "пятница": [46, 53], "суббота": [54, 61]}
faitu4 = {"понедельник": [4, 13], "вторник": [14, 23], "среда": [24, 25], "четверг": [26, 35],
              "пятница": [36, 45], "суббота": [46, 53]}
fvt4 = {"понедельник": [4, 13], "вторник": [14, 21], "среда": [22, 31], "четверг": [32, 41],
              "пятница": [42, 51], "суббота": [52, 59]}
ief4 = {"понедельник": [4, 11], "вторник": [12, 19], "среда": [20, 21], "четверг": [22, 31],
              "пятница": [32, 39], "суббота": [40, 47]}

frt5 = {"понедельник": [4, 15], "вторник": [16, 27], "среда": [28, 29], "четверг": [30, 41],
              "пятница": [42, 53], "суббота": [54, 63]}
fvt5 = {"понедельник": [4, 9], "вторник": [10, 17], "среда": [18, 27], "четверг": [28, 35],
              "пятница": [36, 39], "суббота": [40, 47]}
ief5 = {"понедельник": [4, 9], "вторник": [10, 13], "среда": [14, 17], "четверг": [18, 23],
              "пятница": [24, 29], "суббота": [30, 37]}

patterns = {"frt": [frt1, frt2, frt3, frt4, frt5],
                 "fe": [fe1, fe2, fe3, fe4],
                 "faitu": [faitu1, faitu2, faitu3, faitu4],
                 "fvt": [fvt1, fvt2, fvt3, fvt4, fvt5],
                 "ief": [ief1, ief2, ief3, ief4, ief5]}

num_pair = {"08:10": "1 пара", "09:55": "2 пара", "11:40": "3 пара", "13:35": "4 пара", "15:20": "5 пара"}
num_pair2 = {"11:40": "1 пара", "13:35": "2 пара", "15:20": "3 пара", "17:05": "4 пара", "18:50": "5 пара", "20:25": "6 пара"}
tr_fac = {"frt": "ФРТ", "fe": "ФЭ", "faitu": "ФАИТУ", "fvt": "ФВТ", "ief": "ИЭФ"}

def merged(cell):
    for mergedCell in sheet.merged_cells.ranges:
        #print(mergedCell)
        if cell in mergedCell:
            return True
    return False


day = ""
@bot.callback_query_handler(func=lambda callback: callback.data in lines_days.keys()
                                                  or callback.data in ["back_to_types", "home"])
def showTimesheet(callback):
   # print("showTimesheet", callback.data)
    if callback.data == "back_to_types":
        markup = types.InlineKeyboardMarkup(row_width=1)
        type1 = types.InlineKeyboardButton("Числитель", callback_data="числитель")
        type2 = types.InlineKeyboardButton("Знаменатель", callback_data="знаменатель")
        back = types.InlineKeyboardButton("Вернуться назад", callback_data="back_to_groups")
        home = types.InlineKeyboardButton("В главное меню", callback_data="home")
        markup.row(type1, type2)
        markup.add(back, home)
        bot.send_message(callback.message.chat.id, "Выберите формат недели:", reply_markup=markup)

    elif callback.data == "home":
        markup = types.InlineKeyboardMarkup(row_width=1)
        btn1 = types.InlineKeyboardButton("Узнать расписание", callback_data="group")
        markup.add(btn1)
        bot.send_message(callback.message.chat.id, text="Добро пожаловать в РГРТУ!", reply_markup=markup)

    else:
        global day, type_week
        day = callback.data
        a = patterns[fac][int(course) - 1][day][0]
        b = patterns[fac][int(course) - 1][day][1]

        #a = lines_days[day][0]
        #b = lines_days[day][1]
        output = ""
        for s in range(a, b + 1):
            try:
                if sheet[f"C{s}"].value.lower().strip() == f"{type_week[:4]}.":
                    cell = f"{column_groups[number]}{s}"
                    if (f"{type_week[:4]}." == "числ.") or (not(merged(cell))):
                        info = sheet[cell].value.replace("\n", "   ").split("   ")
                    else:
                        cell = f"{column_groups[number]}{s - 1}"
                        info = sheet[cell].value.replace("\n", "   ").split("   ")

                    if info[0] != "":
                        temp = []
                        if type_week[:4] == "числ":
                            tm = sheet[f"B{s}"].value.replace(".", ":").split("-")
                        else:
                            tm = sheet[f"B{s - 1}"].value.replace(".", ":").split("-")

                        start_lesson = tm[0]
                        finish_lesson = tm[1]
                        x, y = tuple(int(q) for q in start_lesson.split(":"))
                        t1 = (x*60 + y + 45) // 60
                        t2 = x*60 + y + 45 - t1 * 60
                        if t1 < 10: t1 = f"0{t1}"
                        if t2 < 10: t2 = f"0{t2}"
                        timeout = f"{t1}:{t2}"

                        if course == "1" or course == "2":
                            temp.append(f"{num_pair[start_lesson]}\n")
                        else:
                            temp.append(f"{num_pair2[start_lesson]}\n")

                        temp.append(f"Предмет: {info[0]}")

                        if len(info) >= 2:
                            temp.append(f"Преподаватель: {info[1]}")

                        if len(info) >= 3:
                            temp.append(f"Аудитория: {info[2]}")
                        temp.append(f"Начало: {start_lesson}")
                        temp.append(f"Конец: {finish_lesson}")
                        temp.append(f"Перерыв: {timeout} (5 минут)")
                        for item in temp:
                            output += f"{item}\n"
                        output += "\n" * 2
            except:
                 ...

        if output != "":
            if course == "1" or course == "2":
                output = f"{course} курс {tr_fac[fac]}, группа {number}, 1 смена\n\n{day}:\n\n" + output
            else:
                output = f"{course} курс {tr_fac[fac]}, группа {number}, 2 смена\n{day}:\n\n" + output

            bot.send_message(callback.message.chat.id, output)

        else:
            bot.send_message(callback.message.chat.id, "На данный день занятий не запланировано!")

        time.sleep(1)

        markup = types.InlineKeyboardMarkup(row_width=1)
        select_day = types.InlineKeyboardButton("Расписание на другой день", callback_data=type_week)
        new_params = types.InlineKeyboardButton("Изменить параметры", callback_data="fac")
        home = types.InlineKeyboardButton("В главное меню", callback_data="home")
        markup.add(new_params, select_day, home)
        bot.send_message(callback.message.chat.id, "Выберите следующее действие:", reply_markup=markup)




bot.polling(none_stop=True)