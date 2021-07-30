import telebot
import openpyxl
import re
import time
import logging
from telebot import types

group_list = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']
group_box_list = []
global JoinedUsers
bot = telebot.TeleBot("1888736433:AAGH9jNUJ6CjUQ1pdZi8HQIVVE22WeNpWtw")
logging.basicConfig(format = u'%(levelname)-8s [%(asctime)s] %(message)s', level = logging.INFO, filename = u'adminpanel.log')

def excel():
    global wb
    wb = openpyxl.reader.excel.load_workbook (filename ="groups.xlsx")
    global sheetlen
    sheetlen = len(wb.sheetnames)


excel()

def userscheck():
    with open ('users.txt', 'r') as f:
        text = f.read()
        JoinedUsers = text.split('\n')
        f.close()
userscheck()

@bot.message_handler(commands=['start'])
def start(message):
    search_text = str(message.chat.id)
    print(search_text)
    with open('users.txt', 'r+') as file:
        lines = [line.rstrip('\n') for line in open('users.txt')]
        lines.append(search_text)
        lines2 = list(set(lines))
        for item in lines2:
            if item != '':
                file.write("%s\n" % item)
        file.close()

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    btn0 = types.KeyboardButton('Список групп 👥')
    btn1 = types.KeyboardButton('Поиск 🔎')
    btn4 = types.KeyboardButton('Расписание звонков 🔔')
    btn5 = types.KeyboardButton('Справка 💡️')
    btn6 = types.KeyboardButton('Социальные сети ℹ️')
    markup.add(btn0, btn1, btn4, btn5, btn6)
    send_message = f"<b>Привет, {message.from_user.first_name}</b>! \nДля управления используй кнопки внизу\nИспользуй /help для просмотра всех доступных команд"

    userscheck()

    with open('admins.txt', 'r') as adminfile:
        lines_admin = [line.rstrip('\n') for line in open('admins.txt')]
        lines2_admin = list(set(lines_admin))
        lines2_admin_str = ','.join(lines2_admin)
        if search_text in lines2_admin_str:
                btn7 = types.KeyboardButton('Администрация️')
                markup.add(btn7)
    bot.send_message(message.chat.id, send_message, parse_mode='html', reply_markup=markup)

@bot.message_handler(commands=['groups'])
def group_btn(message):

    box_chose_week = types.InlineKeyboardMarkup(row_width=2)
    previous_week = types.InlineKeyboardButton(text='Предыдущая', callback_data='previous_week')
    currenta_week = types.InlineKeyboardButton(text ='Текущая️', callback_data='currenta_week')
    next_week = types.InlineKeyboardButton(text = 'Следующая️', callback_data='next_week')
    box_chose_week.add(currenta_week)
    box_chose_week.add(previous_week,next_week)

    bot.send_message(message.chat.id, 'Выбор недели', reply_markup=box_chose_week)

@bot.message_handler(commands=['search'])
def search_btn(message):
    box_chose_week_search = types.InlineKeyboardMarkup(row_width=2)
    previous_week_search = types.InlineKeyboardButton(text='Предыдущая', callback_data='previous_week_search')
    currenta_week_search = types.InlineKeyboardButton(text ='Текущая', callback_data='currenta_week_search')
    next_week_search = types.InlineKeyboardButton(text = 'Следующая', callback_data='next_week_search')
    box_chose_week_search.add(currenta_week_search)
    box_chose_week_search.add(previous_week_search, next_week_search)
    bot.send_message(message.chat.id, 'Выбор недели', reply_markup=box_chose_week_search)

@bot.message_handler(commands=['time'])
def time_btn(message):
    wb.active = sheetlen - 2
    sheet = wb.active
    time_box = types.InlineKeyboardMarkup(row_width=1)
    monday    = types.InlineKeyboardButton(text = "Понедельник", callback_data= "monday")
    tuesday   = types.InlineKeyboardButton(text = "Вторник", callback_data="tuesday")
    wednesday = types.InlineKeyboardButton(text = "Среда", callback_data="wednesday")
    thursday  = types.InlineKeyboardButton(text = "Четверг", callback_data="thursday")
    friday    = types.InlineKeyboardButton(text = "Пятница", callback_data="friday")
    saturday  = types.InlineKeyboardButton(text = "Суббота", callback_data="saturday")
    time_box.add(monday,tuesday,wednesday,thursday,friday,saturday)
    bot.send_message(message.chat.id, "Выбор дня надели:", reply_markup=time_box)

@bot.message_handler(commands=['social'])
def social_btn(message):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn_site = types.InlineKeyboardButton(text='Сайт', url='http://www.gcbip.by')
    btn_insta = types.InlineKeyboardButton(text='Инстаграм', url='https://www.instagram.com/college_bip/?hl=ru')
    btn_telegram = types.InlineKeyboardButton(text='Телеграм', url='https://t.me/gcbip')
    btn_youtube = types.InlineKeyboardButton(text='Ютуб', url='https://www.youtube.com/channel/UCcy4LxxcMsrJUlxgDV3Tkdg')
    markup.add(btn_site, btn_insta, btn_telegram, btn_youtube)
    bot.send_message(message.chat.id, 'Колледж в социальных сетях:', reply_markup=markup)

@bot.message_handler(commands=['help'])
def help_btn(message):
    send_message = "Доступные команды:\n/start - Запуск бота\n/groups - Список групп\n/search - Поиск\n/time - Расписание звонков\n/social - Социальные сети"
    bot.send_message(message.chat.id, send_message, parse_mode='html')


@bot.message_handler(content_types=['adminpanel'])
def admin_panel(message):
    logging.info(f'Пользователь: {message.from_user.first_name} {message.from_user.last_name} id: {message.chat.id} Открыл панель администратора')
    message_type_box = types.InlineKeyboardMarkup(row_width=1)
    textonly = types.InlineKeyboardButton(text="Отправить текстовое сообщение", callback_data="textonly")
    photoandtext = types.InlineKeyboardButton(text="Отправить графическое сообщение", callback_data="photoandtext")
    newdocument = types.InlineKeyboardButton(text="Загрузить Excel-файл", callback_data="newdocument")
    message_type_box.add(textonly, photoandtext, newdocument)
    bot.send_message(message.chat.id, "Выбор действия:", reply_markup=message_type_box)

@bot.message_handler(content_types=['text'])
def mess(message):
    if message.text == 'Справка 💡️': return help_btn(message)
    elif message.text == 'Расписание звонков 🔔':return time_btn(message)
    elif message.text == 'Социальные сети ℹ️':return social_btn(message)
    elif message.text == 'Список групп 👥':return group_btn(message)
    elif message.text == 'Администрация️': return admin_panel(message)
    elif message.text == 'Поиск 🔎':return search_btn(message)



@bot.callback_query_handler (func = lambda call: True)
def callback_inline(call):
    markup_back_to_time = types.InlineKeyboardMarkup(row_width=1)
    backtotime = types.InlineKeyboardButton(text="Вернуться назад", callback_data="backtotime")
    markup_back_to_time.add(backtotime)

    markup_back_to_group = types.InlineKeyboardMarkup(row_width=1)

    markup_back_to_choseweek = types.InlineKeyboardMarkup(row_width=1)
    backtochoseweek = types.InlineKeyboardButton(text="Вернуться к выбору недели", callback_data="backtochoseweek")
    markup_back_to_choseweek.add(backtochoseweek)

    message_type_box = types.InlineKeyboardMarkup(row_width=1)
    textonly = types.InlineKeyboardButton(text="Отправить текстовое сообщение", callback_data="textonly")
    photoandtext = types.InlineKeyboardButton(text="Отправить графическое сообщение", callback_data="photoandtext")
    excelfile = types.InlineKeyboardButton(text="Загрузить Excel-файл", callback_data="excelfile")
    message_type_box.add(textonly, photoandtext, excelfile)



    result = ''
    time_list = []
    score_list = []
    subject_list = []


    if call.data == "backtochoseweek":
        bot.delete_message(call.message.chat.id, call.message.message_id)
        return group_btn(call.message)

    if call.data =="excelfile":
        logging.info(f'id: {call.message.chat.id} Открыл меню загрузки Excel-файла')
        bot.delete_message(call.message.chat.id, call.message.message_id)
        send = bot.send_message(call.message.chat.id, 'Отправьте файл:')
        bot.register_next_step_handler(send,handle_docs_photo)
    if call.data =="newdocument":
        logging.info(
            f'id: {call.message.chat.id} Открыл меню загрузки Excel-файла')
        bot.delete_message(call.message.chat.id, call.message.message_id)
        send = bot.send_message(call.message.chat.id, 'Отправьте файл:')
        bot.register_next_step_handler(send,handle_docs_photo)

    if call.data =="return_search":
        bot.delete_message(call.message.chat.id, call.message.message_id)
        return search_btn(call.message)



    if call.data == "previous_week":
        global back_to_group_day
        back_to_group_day = 0
        wb.active = sheetlen - 3
        global sheet
        sheet = wb.active
        group_box = types.InlineKeyboardMarkup(row_width=3)
        for i in range(0, len(group_list)):
            if (sheet[group_list[i] + str(2)].value) != ' ':
                group_box_list.append(f'id{i}')
                group_box_list[i] = types.InlineKeyboardButton(text=sheet[group_list[i] + str(2)].value,
                                                               callback_data=f"group_id{i}")
        group_box.add(group_box_list[0], group_box_list[1], group_box_list[2], group_box_list[3], group_box_list[4],
                      group_box_list[5], group_box_list[6], group_box_list[7], group_box_list[8], group_box_list[9],
                      group_box_list[10], group_box_list[11], group_box_list[12], group_box_list[13],
                      group_box_list[14], backtochoseweek)

        bot.edit_message_text('Выбрана неделя: Предыдущая\nВыбор группы: ', call.message.chat.id, call.message.message_id, reply_markup=group_box)
    if call.data == "currenta_week":
        back_to_group_day = 1
        wb.active = sheetlen - 2
        sheet = wb.active
        group_box = types.InlineKeyboardMarkup(row_width=3)
        for i in range(0, len(group_list)):
            if (sheet[group_list[i] + str(2)].value) != ' ':
                group_box_list.append(f'id{i}')
                group_box_list[i] = types.InlineKeyboardButton(text=sheet[group_list[i] + str(2)].value,
                                                               callback_data=f"group_id{i}")
        group_box.add(group_box_list[0], group_box_list[1], group_box_list[2], group_box_list[3], group_box_list[4],
                      group_box_list[5], group_box_list[6], group_box_list[7], group_box_list[8], group_box_list[9],
                      group_box_list[10], group_box_list[11], group_box_list[12], group_box_list[13],
                      group_box_list[14], backtochoseweek)

        bot.edit_message_text(f'Выбрана неделя: Текущая\nВыбор группы: ', call.message.chat.id, call.message.message_id, reply_markup=group_box)
    if call.data == "next_week":
        back_to_group_day = 2
        wb.active = sheetlen - 1
        sheet = wb.active
        group_box = types.InlineKeyboardMarkup(row_width=3)
        for i in range(0, len(group_list)):
            if (sheet[group_list[i] + str(2)].value) != ' ':
                group_box_list.append(f'id{i}')
                group_box_list[i] = types.InlineKeyboardButton(text=sheet[group_list[i] + str(2)].value,
                                                               callback_data=f"group_id{i}")
        group_box.add(group_box_list[0], group_box_list[1], group_box_list[2], group_box_list[3], group_box_list[4],
                      group_box_list[5], group_box_list[6], group_box_list[7], group_box_list[8], group_box_list[9],
                      group_box_list[10], group_box_list[11], group_box_list[12], group_box_list[13],
                      group_box_list[14], backtochoseweek)

        bot.edit_message_text('Выбрана неделя: Следующая\nВыбор группы: ', call.message.chat.id, call.message.message_id, reply_markup=group_box)


    if call.data =="previous_week_search":
        bot.delete_message(call.message.chat.id, call.message.message_id)
        back_to_group_day = 0
        wb.active = sheetlen - 3
        sheet = wb.active
        send = bot.send_message(call.message.chat.id, 'Выбрана неделя: Предыдущая \nВвод данных для поиска (Преподаватель, Группа, Предмет, Кабинет)')
        bot.register_next_step_handler(send, search_teacher)
    if call.data == "currenta_week_search":
        bot.delete_message(call.message.chat.id, call.message.message_id)
        back_to_group_day = 1
        wb.active = sheetlen - 2
        sheet = wb.active
        send = bot.send_message(call.message.chat.id, 'Выбрана неделя: Текущая \nВвод данных для поиска (Преподаватель, Группа, Предмет, Кабинет)')
        bot.register_next_step_handler(send, search_teacher)
    if call.data == "next_week_search":
        bot.delete_message(call.message.chat.id, call.message.message_id)
        back_to_group_day = 2
        wb.active = sheetlen - 1
        sheet = wb.active
        send = bot.send_message(call.message.chat.id, 'Выбрана неделя: Следующая \nВвод данных для поиска (Преподаватель, Группа, Предмет, Кабинет)')
        bot.register_next_step_handler(send, search_teacher)








    if call.data =="send_refresh":
        bot.edit_message_text('Отправка отменена. Выберите действие', call.message.chat.id, call.message.message_id,
                              reply_markup=message_type_box)

    if call.data =="send_to_all_onlytext": return step_onlytext(call.message)

    if call.data =="send_to_all_photo_with_text": return step_text(call.message)

    if call.data =="textonly":
        logging.info(
            f'id: {call.message.chat.id} Открыл меню отправки текста всем пользователям')
        send = bot.edit_message_text('Отправьте текстовое сообщение', call.message.chat.id, call.message.message_id,)
        bot.register_next_step_handler(send, step_onlytext_preview)
    if call.data =="photoandtext":
        logging.info(
            f'id: {call.message.chat.id} Открыл меню отправки Текста с изображением')
        send = bot.edit_message_text('Отправьте ссылку на изображение:', call.message.chat.id, call.message.message_id,)
        bot.register_next_step_handler(send, step_photo)

    if call.data =="monday":
        sheet = wb.active
        day = sheet['A3'].value
        for i in range(3,9):
            time_list.append(sheet['C' + str(i)].value)
            score_list.append(sheet['B' + str(i)].value)
            result = result + f'{score_list[i-3]} — {time_list[i-3]}\n'
        bot.edit_message_text(f'Расписание на {day} \n{result}', call.message.chat.id, call.message.message_id, reply_markup=markup_back_to_time)
    if call.data == "tuesday":
        sheet = wb.active
        day = sheet['A9'].value
        for i in range(9, 15):
            time_list.append(sheet['C' + str(i)].value)
            score_list.append(sheet['B' + str(i)].value)
            result = result + f'{score_list[i - 9]} — {time_list[i - 9]}\n'
        bot.edit_message_text(f'Расписание на {day} \n{result}', call.message.chat.id, call.message.message_id, reply_markup=markup_back_to_time)
    if call.data == "wednesday":
        sheet = wb.active
        day = sheet['A15'].value
        for i in range(15, 21):
            time_list.append(sheet['C' + str(i)].value)
            score_list.append(sheet['B' + str(i)].value)
            result = result + f'{score_list[i - 15]} — {time_list[i - 15]}\n'
        bot.edit_message_text(f'Расписание на {day} \n{result}', call.message.chat.id, call.message.message_id, reply_markup=markup_back_to_time)
    if call.data =="thursday":
        sheet = wb.active
        day = sheet['A21'].value
        for i in range(21, 27):
            time_list.append(sheet['C' + str(i)].value)
            score_list.append(sheet['B' + str(i)].value)
            result = result + f'{score_list[i - 21]} — {time_list[i - 21]}\n'
        bot.edit_message_text(f'Расписание на {day} \n{result}', call.message.chat.id, call.message.message_id, reply_markup=markup_back_to_time)
    if call.data =="friday":
        sheet = wb.active
        day = sheet['A27'].value
        for i in range(27, 33):
            time_list.append(sheet['C' + str(i)].value)
            score_list.append(sheet['B' + str(i)].value)
            result = result + f'{score_list[i - 27]} — {time_list[i - 27]}\n'
        bot.edit_message_text(f'Расписание на {day} \n{result}', call.message.chat.id, call.message.message_id, reply_markup=markup_back_to_time)
    if call.data =="saturday":
        sheet = wb.active
        day = sheet['A33'].value
        for i in range(33, 39):
            time_list.append(sheet['C' + str(i)].value)
            score_list.append(sheet['B' + str(i)].value)
            result = result + f'{score_list[i - 33]} — {time_list[i - 33]}\n'
        bot.edit_message_text(f'Расписание на {day} \n{result}', call.message.chat.id, call.message.message_id, reply_markup=markup_back_to_time)

    if call.data =="backtotime":
        bot.delete_message(call.message.chat.id, call.message.message_id)
        return time_btn(call.message)

    markup_group_day_box = types.InlineKeyboardMarkup(row_width=2)

    for cc in range(0, 20):
        if call.data == f"group_id{cc}":
            if back_to_group_day == 0:
                text = 'Предыдущая'
            elif back_to_group_day == 1:
                text = 'Текущая'
            elif back_to_group_day == 2:
                text = 'Следующая'
            sheet = wb.active
            monday = types.InlineKeyboardButton(text=sheet['A' + str(3)].value, callback_data=f"monday_group{cc}")
            tuesday = types.InlineKeyboardButton(text=sheet['A' + str(9)].value, callback_data=f"tuesday_group{cc}")
            wednesday = types.InlineKeyboardButton(text=sheet['A' + str(15)].value, callback_data=f"wednesday_group{cc}")
            thursday = types.InlineKeyboardButton(text=sheet['A' + str(21)].value, callback_data=f"thursday_group{cc}")
            friday = types.InlineKeyboardButton(text=sheet['A' + str(27)].value, callback_data=f"friday_group{cc}")
            saturday = types.InlineKeyboardButton(text=sheet['A' + str(33)].value, callback_data=f"saturday_group{cc}")
            week = types.InlineKeyboardButton(text='Вся неделя', callback_data=f"week_group{cc}")
            if back_to_group_day == 0:
                backtogroup = types.InlineKeyboardButton(text="Вернуться назад", callback_data="previous_week")
            if back_to_group_day == 1:
                backtogroup = types.InlineKeyboardButton(text="Вернуться назад", callback_data="currenta_week")
            if back_to_group_day == 2:
                backtogroup = types.InlineKeyboardButton(text="Вернуться назад", callback_data="next_week")
            markup_group_day_box.add(monday,tuesday,wednesday,thursday,friday,saturday)
            markup_group_day_box.add(week)
            markup_group_day_box.add(backtogroup)
            markup_group_day_box.add(backtochoseweek)

            bot.edit_message_text(f'Выбрана неделя: {text}\nВыбрана группа: {sheet[group_list[cc] + str(2)].value}', call.message.chat.id, call.message.message_id, reply_markup=markup_group_day_box)
    for cc in range(0, 20):
        if call.data == f"monday_group{cc}":
            if back_to_group_day == 0:
                text = 'Предыдущая'
            elif back_to_group_day == 1:
                text = 'Текущая'
            elif back_to_group_day == 2:
                text = 'Следующая'
            sheet = wb.active
            day = sheet['A3'].value
            for i in range(3, 9):
                time_list.append(sheet['C' + str(i)].value)
                subject_list.append(sheet[f'{group_list[cc]}' + str(i)].value)
                if subject_list[i-3] == None: subject_list[i-3] = ' '
                result = result + f'{time_list[i - 3]}  —  {subject_list[i - 3]}\n'
            backtogroup = types.InlineKeyboardButton(text="Вернуться назад", callback_data=f"group_id{cc}")
            if back_to_group_day == 0:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы", callback_data="previous_week")
            if back_to_group_day == 1:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы", callback_data="currenta_week")
            if back_to_group_day == 2:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы", callback_data="next_week")
            markup_back_to_group.add(backtogroup, backtogroups, backtochoseweek)
            bot.edit_message_text(
                f'Выбрана неделя: {text}\nРезультат {sheet[group_list[cc] + str(2)].value} {day}: \n \n{result}',
                call.message.chat.id,
                call.message.message_id, reply_markup=markup_back_to_group)
    for cc in range(0, 20):
        if call.data == f"tuesday_group{cc}":
            if back_to_group_day == 0:
                text = 'Предыдущая'
            elif back_to_group_day == 1:
                text = 'Текущая'
            elif back_to_group_day == 2:
                text = 'Следующая'
            day = sheet['A9'].value
            for i in range(9, 15):
                time_list.append(sheet['C' + str(i)].value)
                subject_list.append(sheet[f'{group_list[cc]}' + str(i)].value)
                if subject_list[i - 9] == None: subject_list[i - 9] = ' '
                result = result + f'{time_list[i - 9]}  —  {subject_list[i - 9]}\n'
            backtogroup = types.InlineKeyboardButton(text="Вернуться назад", callback_data=f"group_id{cc}")
            if back_to_group_day == 0:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы",
                                                          callback_data="previous_week")
            if back_to_group_day == 1:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы",
                                                          callback_data="currenta_week")
            if back_to_group_day == 2:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы", callback_data="next_week")
            markup_back_to_group.add(backtogroup, backtogroups, backtochoseweek)
            bot.edit_message_text(
                f'Выбрана неделя: {text}\nРезультат {sheet[group_list[cc] + str(2)].value} {day}: \n \n{result}',
                call.message.chat.id,
                call.message.message_id, reply_markup=markup_back_to_group)
    for cc in range(0, 20):
        if call.data == f"wednesday_group{cc}":
            if back_to_group_day == 0:
                text = 'Предыдущая'
            elif back_to_group_day == 1:
                text = 'Текущая'
            elif back_to_group_day == 2:
                text = 'Следующая'
            day = sheet['A15'].value
            for i in range(15, 21):
                time_list.append(sheet['C' + str(i)].value)
                subject_list.append(sheet[f'{group_list[cc]}' + str(i)].value)
                if subject_list[i - 15] == None: subject_list[i - 15] = ' '
                result = result + f'{time_list[i - 15]}  —  {subject_list[i - 15]}\n'
            backtogroup = types.InlineKeyboardButton(text="Вернуться назад", callback_data=f"group_id{cc}")
            if back_to_group_day == 0:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы",
                                                          callback_data="previous_week")
            if back_to_group_day == 1:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы",
                                                          callback_data="currenta_week")
            if back_to_group_day == 2:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы", callback_data="next_week")
            markup_back_to_group.add(backtogroup, backtogroups, backtochoseweek)
            bot.edit_message_text(
                f'Выбрана неделя: {text}\nРезультат {sheet[group_list[cc] + str(2)].value} {day}: \n \n{result}',
                call.message.chat.id,
                call.message.message_id, reply_markup=markup_back_to_group)
    for cc in range(0, 20):
        if call.data == f"thursday_group{cc}":
            if back_to_group_day == 0:
                text = 'Предыдущая'
            elif back_to_group_day == 1:
                text = 'Текущая'
            elif back_to_group_day == 2:
                text = 'Следующая'
            day = sheet['A21'].value
            for i in range(21, 27):
                time_list.append(sheet['C' + str(i)].value)
                subject_list.append(sheet[f'{group_list[cc]}' + str(i)].value)
                if subject_list[i - 21] == None: subject_list[i - 21] = ' '
                result = result + f'{time_list[i - 21]}  —  {subject_list[i - 21]}\n'
            backtogroup = types.InlineKeyboardButton(text="Вернуться назад", callback_data=f"group_id{cc}")
            if back_to_group_day == 0:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы",
                                                          callback_data="previous_week")
            if back_to_group_day == 1:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы",
                                                          callback_data="currenta_week")
            if back_to_group_day == 2:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы", callback_data="next_week")
            markup_back_to_group.add(backtogroup, backtogroups, backtochoseweek)
            bot.edit_message_text(
                f'Выбрана неделя: {text}\nРезультат {sheet[group_list[cc] + str(2)].value} {day}: \n \n{result}',
                call.message.chat.id,
                call.message.message_id, reply_markup=markup_back_to_group)
    for cc in range(0, 20):
        if call.data == f"friday_group{cc}":
            if back_to_group_day == 0:
                text = 'Предыдущая'
            elif back_to_group_day == 1:
                text = 'Текущая'
            elif back_to_group_day == 2:
                text = 'Следующая'
            day = sheet['A27'].value
            for i in range(27, 33):
                time_list.append(sheet['C' + str(i)].value)
                subject_list.append(sheet[f'{group_list[cc]}' + str(i)].value)
                if subject_list[i - 27] == None: subject_list[i - 27] = ' '
                result = result + f'{time_list[i - 27]}  —  {subject_list[i - 27]}\n'
            backtogroup = types.InlineKeyboardButton(text="Вернуться назад", callback_data=f"group_id{cc}")
            if back_to_group_day == 0:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы",
                                                          callback_data="previous_week")
            if back_to_group_day == 1:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы",
                                                          callback_data="currenta_week")
            if back_to_group_day == 2:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы", callback_data="next_week")
            markup_back_to_group.add(backtogroup, backtogroups, backtochoseweek)
            bot.edit_message_text(
                f'Выбрана неделя: {text}\nРезультат {sheet[group_list[cc] + str(2)].value} {day}: \n \n{result}',
                call.message.chat.id,
                call.message.message_id, reply_markup=markup_back_to_group)
    for cc in range(0, 20):
        if call.data == f"saturday_group{cc}":
            if back_to_group_day == 0:
                text = 'Предыдущая'
            elif back_to_group_day == 1:
                text = 'Текущая'
            elif back_to_group_day == 2:
                text = 'Следующая'
            day = sheet['A33'].value
            for i in range(33, 39):
                time_list.append(sheet['C' + str(i)].value)
                subject_list.append(sheet[f'{group_list[cc]}' + str(i)].value)
                if subject_list[i - 33] == None: subject_list[i - 33] = ' '
                result = result + f'{time_list[i - 33]}  —  {subject_list[i - 33]}\n'
            backtogroup = types.InlineKeyboardButton(text="Вернуться назад", callback_data=f"group_id{cc}")
            if back_to_group_day == 0:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы",
                                                          callback_data="previous_week")
            if back_to_group_day == 1:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы",
                                                          callback_data="currenta_week")
            if back_to_group_day == 2:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы", callback_data="next_week")
            markup_back_to_group.add(backtogroup, backtogroups, backtochoseweek)
            bot.edit_message_text(f'Выбрана неделя: {text}\nРезультат {sheet[group_list[cc] + str(2)].value} {day}: \n \n{result}', call.message.chat.id,
                                  call.message.message_id, reply_markup=markup_back_to_group)




    for cc in range(0, 20):
        if call.data == f"week_group{cc}":
            result_1= sheet['A3'].value
            result_2= sheet['A9'].value
            result_3 = sheet['A15'].value
            result_4 = sheet['A21'].value
            result_5 = sheet['A27'].value
            result_6 = sheet['A33'].value
            result_week_group = ''
            text = ''
            if back_to_group_day == 0:
                text = 'Предыдущая'
            elif back_to_group_day == 1:
                text = 'Текущая'
            elif back_to_group_day == 2:
                text = 'Следующая'

            for number in range(0, 6):
                result_time = sheet['C' + str(number + 3)].value
                result_group_name = sheet[group_list[cc] + str(2)].value
                result_cell = sheet[group_list[cc] + str(number + 3)].value
                if result_cell != None:
                    result_week_group = f'{result_week_group}\n{result_time} — {result_cell}'
                elif result_cell == None:
                    result_cell = ' '
                    result_week_group = f'{result_week_group}\n{result_time} — {result_cell}'
            result_week_group = f'{result_1}\n{result_week_group} \n\n {result_2}\n'
            for number in range(5, 11):
                result_time = sheet['C' + str(number + 4)].value
                result_group_name = sheet[group_list[cc] + str(2)].value
                result_cell = sheet[group_list[cc] + str(number + 4)].value
                if result_cell != None:
                    result_week_group = f'{result_week_group}\n{result_time} — {result_cell}'
                elif result_cell == None:
                    result_cell = ' '
                    result_week_group = f'{result_week_group}\n{result_time} — {result_cell}'
            result_week_group = f'{result_week_group} \n\n {result_3}\n'
            for number in range(11, 17):
                result_time = sheet['C' + str(number + 4)].value
                result_group_name = sheet[group_list[cc] + str(2)].value
                result_cell = sheet[group_list[cc] + str(number + 4)].value
                if result_cell != None:
                    result_week_group = f'{result_week_group}\n{result_time} — {result_cell}'
                elif result_cell == None:
                    result_cell = ' '
                    result_week_group = f'{result_week_group}\n{result_time} — {result_cell}'
            result_week_group = f'{result_week_group} \n\n {result_4}\n'
            for number in range(17, 23):
                result_time = sheet['C' + str(number + 4)].value
                result_group_name = sheet[group_list[cc] + str(2)].value
                result_cell = sheet[group_list[cc] + str(number + 4)].value
                if result_cell != None:
                    result_week_group = f'{result_week_group}\n{result_time} — {result_cell}'
                elif result_cell == None:
                    result_cell = ' '
                    result_week_group = f'{result_week_group}\n{result_time} — {result_cell}'
            result_week_group = f'{result_week_group} \n\n {result_5}\n'
            for number in range(23, 29):
                result_time = sheet['C' + str(number + 4)].value
                result_group_name = sheet[group_list[cc] + str(2)].value
                result_cell = sheet[group_list[cc] + str(number + 4)].value
                if result_cell != None:
                    result_week_group = f'{result_week_group}\n{result_time} — {result_cell}'
                elif result_cell == None:
                    result_cell = ' '
                    result_week_group = f'{result_week_group}\n{result_time} — {result_cell}'
            result_week_group = f'{result_week_group} \n\n {result_6}\n'
            for number in range(29, 35):
                result_time = sheet['C' + str(number + 4)].value
                result_group_name = sheet[group_list[cc] + str(2)].value
                result_cell = sheet[group_list[cc] + str(number + 4)].value
                if result_cell != None:
                    result_week_group = f'{result_week_group}\n{result_time} — {result_cell}'
                elif result_cell == None:
                    result_cell = ' '
                    result_week_group = f'{result_week_group}\n{result_time} — {result_cell}'
            result_week_group = f'{result_week_group}'

            backtogroup = types.InlineKeyboardButton(text="Вернуться назад", callback_data=f"group_id{cc}")
            if back_to_group_day == 0:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы",
                                                          callback_data="previous_week")
            if back_to_group_day == 1:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы",
                                                          callback_data="currenta_week")
            if back_to_group_day == 2:
                backtogroups = types.InlineKeyboardButton(text="Вернуться к выбору группы", callback_data="next_week")
            markup_back_to_group.add(backtogroup, backtogroups, backtochoseweek)
            bot.edit_message_text(f'Выбрана неделя: {text}\nВыбрана группа: {result_group_name}:\n\n{result_week_group}', call.message.chat.id,
                                  call.message.message_id, reply_markup=markup_back_to_group)





def step_photo(message):
    global photo
    photo = message.text
    try:
        l = re.search("(?P<url>https?://[^\s]+)", photo).group()
        send = bot.send_message(message.chat.id, 'Отправьте текстовое сообщение:')
        bot.register_next_step_handler(send, step_text_preview)
        logging.info(
            f'Пользователь: {message.from_user.first_name} {message.from_user.last_name} id: {message.chat.id} ввел ссылку на изображение')
    except Exception:
        bot.send_message(message.chat.id, 'Неверная ссылка, попробуйте заново')
        logging.info(
            f'Пользователь: {message.from_user.first_name} {message.from_user.last_name} id: {message.chat.id} ввел неверную ссылку на изображение')
        return admin_panel(message)

def step_text_preview(message):
    global text_to_photo
    text_to_photo = message.text
    try:
        message_type_box = types.InlineKeyboardMarkup(row_width=2)
        send_to_all_photo_with_text = types.InlineKeyboardButton(text="Отправить",
                                                          callback_data="send_to_all_photo_with_text")
        send_refresh = types.InlineKeyboardButton(text="Повторить", callback_data="send_refresh")
        message_type_box.add(send_refresh, send_to_all_photo_with_text)
        bot.send_message(message.chat.id, "Предпросмотр сообщения:")
        bot.send_photo(message.chat.id, photo, text_to_photo)
        bot.send_message(message.chat.id, "Отправлять?", reply_markup=message_type_box)
        logging.info(
            f'id: {message.chat.id} ввел текст (с изображением)')
    except Exception:
        bot.send_message(message.chat.id,f'Ошибка на этапе отправки изображения')



def step_text(message):
    with open ('users.txt', 'r') as f:
        text = f.read()
        JoinedUsers = text.split('\n')
        f.close()
    try:
        i = 0
        for user in JoinedUsers:
            if user != '':
                bot.send_photo(user, photo, text_to_photo)
                i = i + 1
        bot.edit_message_text(f'Успешно. Сообщение доставлено {i} раз(а)',message.chat.id,message.message_id)
        logging.info(
            f'Пользователь: {message.from_user.first_name} {message.from_user.last_name} id: {message.chat.id} Отправил избражение с текстом всем участникам')
    except Exception:
        logging.info(
            f'id: {message.chat.id} Ошибка на этапе изображение с текстом')
        bot.edit_message_text(f'Ошибка на этапе отправки изображения с текстом',message.chat.id,message.message_id)


def step_onlytext_preview(message):
    global text_to_send
    text_to_send = message.text
    try:
        message_type_box = types.InlineKeyboardMarkup(row_width=2)
        send_to_all_onlytext = types.InlineKeyboardButton(text="Отправить",callback_data="send_to_all_onlytext")
        send_refresh = types.InlineKeyboardButton(text="Повторить", callback_data="send_refresh")
        message_type_box.add(send_refresh, send_to_all_onlytext)
        bot.send_message(message.chat.id, "Предпросмотр сообщения:")
        bot.send_message(message.chat.id, text_to_send)
        bot.send_message(message.chat.id,"Отправлять?", reply_markup=message_type_box)
        logging.info(
            f'id: {message.chat.id} Ввел данные для отправки текста')
    except Exception:
        logging.info(
            f'id: {message.chat.id} Ошибка на отправке текста')
        print('Ошибка на этапе отправки рассылки {Только текст}')
def step_onlytext(message):
    with open ('users.txt', 'r') as f:
        text = f.read()
        JoinedUsers = text.split('\n')
        f.close()
    try:
        i = 0
        for user in JoinedUsers:

            if user != '':
                bot.send_message(user, text_to_send)
                i = i + 1
        bot.edit_message_text(f'Успешно. Сообщение доставлено {i} раз(а)',message.chat.id,message.message_id)
        logging.info(
        f'id: {message.chat.id} Отправил текст всем участникам')
    except Exception:
        logging.info(
            f'id: {message.chat.id} Ошибка на отправке текста')

def handle_docs_photo(message):
    try:
        chat_id = message.chat.id

        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        check_format = message.document.file_name.split(".")
        sravnenie = check_format[1]
        if sravnenie == "xlsx":
            src = '' + message.document.file_name;
            with open(src, 'wb') as new_file:
                new_file.write(downloaded_file)
            bot.reply_to(message, f'Файл {message.document.file_name} сохранен')
            time.sleep(6)
            bot.send_message(message.chat.id, 'Данные в Telegram обновлены')
            logging.info(
                f'id: {message.chat.id} обновил Excel Файл')
            excel()

        elif sravnenie == "xlsm":
            src = '' + message.document.file_name;
            with open(src, 'wb') as new_file:
                new_file.write(downloaded_file)
            bot.reply_to(message, f'Файл {message.document.file_name} сохранен')
            time.sleep(6)
            bot.send_message(message.chat.id, 'Данные в Telegram обновлены')
            logging.info(
                f'id: {message.chat.id} обновил Excel Файл')
            excel()

        elif sravnenie == "xlsb":
            src = '' + message.document.file_name;
            with open(src, 'wb') as new_file:
                new_file.write(downloaded_file)
            bot.reply_to(message, f'Файл {message.document.file_name} сохранен')
            time.sleep(6)
            bot.send_message(message.chat.id, 'Данные в Telegram обновлены')
            logging.info(
                f'id: {message.chat.id} обновил Excel Файл')
            excel()
        else:
            logging.info(
                f'id: {message.chat.id} Неудалось загрузить файл (Неверный формат / Не Excel файл)')
            bot.send_message(message.chat.id, "Неверный формат / Попытка загрузить не Excel файл")
            return admin_panel(message)
    except Exception as e:
        bot.reply_to(message, e)
    return admin_panel(message)


def search_teacher(message):
    return_search_box = types.InlineKeyboardMarkup(row_width=1)
    return_search = types.InlineKeyboardButton(text="Повторить", callback_data="return_search")
    return_search_box.add(return_search)
    global text_to_search_teacher
    text_to_search_teacher = message.text

    result_all = ''
    result_day = ''
    result_time = ''
    result_group_name = ''
    result_all = ''
    result_group_id = ''

    result_1 = sheet['A3'].value
    result_2 = sheet['A9'].value
    result_3 = sheet['A15'].value
    result_4 = sheet['A21'].value
    result_5 = sheet['A27'].value
    result_6 = sheet['A33'].value

    for group_id in range(0,14):
        for number in range(0,6):
           if sheet[group_list[group_id] + str(number + 3)].value != None:
               index_value = sheet[group_list[group_id] + str(number + 3)].value
               result_group_id = sheet[group_list[group_id] + str(2)].value
               if text_to_search_teacher.lower() in index_value.lower() or text_to_search_teacher.lower() in result_group_id.lower():
                    result_time = sheet['C' + str(number + 3)].value
                    result_group_name = f'Группа: {sheet[group_list[group_id] + str(2)].value}'
                    result_group_object = sheet[group_list[group_id] + str(number + 3)].value
                    result_all = f'{result_all}\n {result_group_name}:\n {result_time} - {result_group_object}  \n ---------------------'
    result_all = f'{result_1}\n{result_all}\n\n{result_2}\n'
    for group_id in range(0, 14):
        for number in range(6, 12):
            if sheet[group_list[group_id] + str(number + 3)].value != None:
                index_value = sheet[group_list[group_id] + str(number + 3)].value
                result_group_id = sheet[group_list[group_id] + str(2)].value
                if text_to_search_teacher.lower() in index_value.lower() or text_to_search_teacher.lower() in result_group_id.lower():
                    result_time = sheet['C' + str(number + 3)].value
                    result_group_name = f'Группа: {sheet[group_list[group_id] + str(2)].value}'
                    result_group_object = sheet[group_list[group_id] + str(number + 3)].value
                    result_all = f'{result_all}\n {result_group_name}:\n {result_time} - {result_group_object}  \n ---------------------'
    result_all = f'\n{result_all}\n\n{result_3}\n'
    for group_id in range(0, 14):
        for number in range(12, 18):
            if sheet[group_list[group_id] + str(number + 3)].value != None:
                index_value = sheet[group_list[group_id] + str(number + 3)].value
                result_group_id = sheet[group_list[group_id] + str(2)].value
                if text_to_search_teacher.lower() in index_value.lower() or text_to_search_teacher.lower() in result_group_id.lower():
                    result_time = sheet['C' + str(number + 3)].value
                    result_group_name = f'Группа: {sheet[group_list[group_id] + str(2)].value}'
                    result_group_object = sheet[group_list[group_id] + str(number + 3)].value
                    result_all = f'{result_all}\n {result_group_name}:\n {result_time} - {result_group_object}  \n ---------------------'

    result_all = f'\n{result_all}\n\n{result_4}\n'
    for group_id in range(0, 14):
        for number in range(18, 24):
            if sheet[group_list[group_id] + str(number + 3)].value != None:
                index_value = sheet[group_list[group_id] + str(number + 3)].value
                result_group_id = sheet[group_list[group_id] + str(2)].value
                if text_to_search_teacher.lower() in index_value.lower() or text_to_search_teacher.lower() in result_group_id.lower():
                    result_time = sheet['C' + str(number + 3)].value
                    result_group_name = f'Группа: {sheet[group_list[group_id] + str(2)].value}'
                    result_group_object = sheet[group_list[group_id] + str(number + 3)].value
                    result_all = f'{result_all}\n {result_group_name}:\n {result_time} - {result_group_object}  \n ---------------------'
    result_all = f'\n{result_all}\n\n{result_5}\n'
    for group_id in range(0, 14):
        for number in range(24, 30):
            if sheet[group_list[group_id] + str(number + 3)].value != None:
                index_value = sheet[group_list[group_id] + str(number + 3)].value
                result_group_id = sheet[group_list[group_id] + str(2)].value
                if text_to_search_teacher.lower() in index_value.lower() or text_to_search_teacher.lower() in result_group_id.lower():
                    result_time = sheet['C' + str(number + 3)].value
                    result_group_name = f'Группа: {sheet[group_list[group_id] + str(2)].value}'
                    result_group_object = sheet[group_list[group_id] + str(number + 3)].value
                    result_all = f'{result_all}\n {result_group_name}:\n {result_time} - {result_group_object}  \n ---------------------'
    result_all = f'\n{result_all}\n\n{result_6}\n'
    for group_id in range(0, 14):
        for number in range(30, 36):
            if sheet[group_list[group_id] + str(number + 3)].value != None:
                index_value = sheet[group_list[group_id] + str(number + 3)].value
                result_group_id = sheet[group_list[group_id] + str(2)].value
                if text_to_search_teacher.lower() in index_value.lower() or text_to_search_teacher.lower() in result_group_id.lower():
                    result_time = sheet['C' + str(number + 3)].value
                    result_group_name = f'Группа: {sheet[group_list[group_id] + str(2)].value}'
                    result_group_object = sheet[group_list[group_id] + str(number + 3)].value
                    result_all = f'{result_all}\n {result_group_name}:\n {result_time} - {result_group_object}  \n ---------------------'
    result_all = f'\n{result_all}'
    bot.send_message(message.chat.id,'Результат поиска:')
    return_search_box = types.InlineKeyboardMarkup(row_width=1)
    return_search = types.InlineKeyboardButton(text="Повторить", callback_data="return_search")
    return_search_box.add(return_search)
    bot.send_message(message.chat.id, f'{result_all}\n', reply_markup=return_search_box)

bot.polling (none_stop = True)
