import requests
import telebot
import json
import os
import sys
import sqlite3
import apiai
import xlrd
import xlwt
import time
from datetime import datetime

import db_helper
import start
import stats

import sign
import ask
import Asks
import Analysis


def load_config():
    if not os.path.isfile('config.json'):
        sys.exit("Не найден файл конфигураций.")
    else:
        with open("config.json", encoding='utf-8') as config_file:
            return json.load(config_file)


def load_chats(config):
    chats = {}
    if os.path.isfile('chats.json'):
        with open('chats.json', encoding='utf-8') as chats_file:
            try:
                chats = json.load(chats_file)
            except Exception as ex:
                chats = {
                    "private": {
                        "guest": {},
                        "student": {},
                        "admin": {}
                    },
                    "group": {},
                    "channel": {},
                    "supergroup": {}}
                print(ex)

    if config['admin'] != '':
        chats['private']['admin'] = {
            config['admin']: 'default'
        }
        if config['admin'] in chats['private']['student']:
            del chats['private']['student'][config['admin']]
    else:
        chats['private']['admin'] = {}

    if config['channel'] != '':
        chats['channel'][config['channel']] = "default"

    with open('chats.json', 'w', encoding='utf-8') as chats_file:
        json.dump(chats, chats_file)
    return chats


def save_chats():
    global chats
    with open('chats.json', 'w', encoding='utf-8') as chats_file:
        json.dump(chats, chats_file)


def lessons_bot_AI(small_talk_token, message, bot):
    if small_talk_token != '':
        request = apiai.ApiAI(small_talk_token).text_request()
        request.lang = 'ru'
        request.session_id = 'LessonsBotAI'
        request.query = message.text.strip()
        responseJson = json.loads(request.getresponse().read().decode('utf-8'))
        response = responseJson['result']['fulfillment']['speech']
        if response:
            bot.send_message(message.chat.id, text=response)
        else:
            bot.send_message(message.chat.id, text='Я Вас не совсем понял.')


def command_help(message, bot, chats, commands_dict):
    chat = message.chat
    user_id = str(chat.id)

    if chat.type == 'private':
        commands = ''
        for a in commands_dict['all'].items():
            commands += '{} - {}\n'.format(a[0], a[1])
        if user_id in chats['private']['admin']:
            for a in commands_dict['admin'].items():
                commands += '{} - {}\n'.format(a[0], a[1])
        elif user_id in chats['private']['student']:
            for a in commands_dict['student'].items():
                commands += '{} - {}\n'.format(a[0], a[1])
        else:
            for a in commands_dict['guest'].items():
                commands += '{} - {}\n'.format(a[0], a[1])
        bot.send_message(user_id, commands)
    if chat.type == 'group':
        pass


def command_start(message, bot, chats, config):
    chat = message.chat
    user_id = str(str(chat.id))
    if chat.type == 'private':
        if len(chats['private']['admin']) == 0 or user_id in chats['private']['admin']:
            chats['private']['admin'] = {
                user_id: 'default'
            }
            if user_id in chats['private']['student']:
                del chats['private']['student'][user_id]
            bot.send_message(user_id,
                             'Вы являетесь админом этого бота.\nКомандой /help можете получить доступные команды.')
            save_chats()
        else:
            answ = start.start(config)
            bot.send_message(user_id, answ['about'])
            bot.send_message(user_id, answ['info'])
    if chat.type == 'group':
        pass


def download_doc(bot, token, document, f_name, formats):
    res = {'is_ok': True}
    file_format = document.file_name.split('.')[-1]
    if file_format in formats:
        file_info = bot.get_file(document.file_id)
        file = requests.get(
            'https://api.telegram.org/file/bot{0}/{1}'.format(token, file_info.file_path))
        if int(file.status_code) == 200:
            with open(f_name + '.' + file_format, 'wb') as output:
                output.write(file.content)
        else:
            res['is_ok'] = False
            res['answ'] = 'Не удалось загрузить файл.'
    else:
        res['is_ok'] = False
        res['answ'] = 'Отправьте файл в формате ' + ', '.join(formats) + '.'
    return res


def save_table_as_xls(db_name, table_name):
    if table_name in ['ratings', 'absents']:
        conn = sqlite3.connect(db_name)
        cur = conn.cursor()
        studs = cur.execute('SELECT * FROM students')
        stud_names = [description[0] for description in cur.description]
        studs = studs.fetchall()
        rows = cur.execute('SELECT * FROM {}'.format(table_name))
        names = [description[0] for description in cur.description]
        rows = rows.fetchall()
        conn.close()
        if len(rows) > 0:
            wb = xlwt.Workbook()
            sheet = wb.add_sheet(table_name)
            columns = list(list(names[:1]) + list(stud_names[1:]) + list(names[1:]))
            for j, h in enumerate(columns):
                sheet.write(0, j, h)
            for i in range(len(rows)):
                rows[i] = list(rows[i][0:1]) + list(studs[i][1:]) + list(rows[i][1:])
                for j in range(len(rows[i])):
                    sheet.write(i + 1, j, rows[i][j])
            wb.save(table_name + '.xls')
            return True

    else:
        rows = db_helper.execute_select(db_name, 'SELECT * FROM {}'.format(table_name))
        if len(rows) > 0:
            wb = xlwt.Workbook()
            sheet = wb.add_sheet(table_name)
            for i in range(len(rows)):
                for j in range(len(rows[i])):
                    sheet.write(i, j, rows[i][j])
            wb.save(table_name + '.xls')
            return True
    return False


commands_dict = {
    "all": {
        "/start": "начать диалог с ботом",
        "/help": "получить доступные команды"
    },
    "guest": {
        "/signup": 'подписаться(через пробел введите \'номер зачетной книжки, фамилия, имя, отчество\')'
    },
    "admin": {
        '/analysis': 'анализ',
        '/loadstudents': u"загрузить список студентов(вам придется заново заполнить рейтинги и пропуски студентов!⚡⚡⚡)",
        "/loadhws": u"загрузить новый список домашних заданий(вам придется заново заполнить рейтинги студентов!⚡⚡⚡)",
        "/loadratings": "загрузить новый список выполненных заданий студентов",
        "/loadabsents": "загрузить список с пропусками студентов",
        "/getstudents": "выгрузить список студентов",
        "/gethws": "выгрузить список домашних заданий",
        "/getratings": "выгрузить список выполненных домашних заданий",
        "/getabsents": "выгрузить список с пропусками студентов",
        "/showstudents": "показать список студентов",
        "/showhws": "показать список домащних заданий",
        "/showratings": "показать успеваемость студентов",
        "/showabsents": "показать пропуски студентов",
        "/addhw": "добавить домашнее задание",
        "/delhw": "удалить домашнее задание",
        "/addstudent": "добавить студента(через пробел введите \'номер зачетной книжки, фамилия,имя,отчество\')",
        "/delstudent": "удалить студента(через пробел укажите номер зачетной книжки студента)",
        "/channel": 'ответить в канал через бота(через пробел текст)',
        '/getask': 'получить первый вопрос в стеке вопросов',
        '/answer': 'ответить на вопрос в канал(вопрос будет удален)',
        '/delask': 'удалить первый вопрос в стеке вопросов',
        "/default": "режим по умолчанию(можно поговорить с ботом)",
        "/getchats": "выгрузить файл chats",
        "/getconfig": "выгрузить файл конфигураций",
        "/loadconfig": "загрузить файл конфигураций(⚡⚡⚡)"
    },
    "student": {
        "/ask": "отправить вопрос преподавателю(через пробел напишите текст вопроса)",
        "/absents": "узнать информацию о своих пропусках(через пробел укажите номер зачетной книжки)",
        "/allhws": "получить полный список домашних заданий",
        "/hws": "получить список текущих домашних заданий(через пробел укажите номер зачетной книжки)",
        "/signout": "отписаться(через пробел укажите номер зачетной книжки)",
        "/default": "режим по умолчанию(можно поговорить с ботом)"
    }
}

db = 'db/lessons_db.db'

config = load_config()
chats = load_chats(config)

db_helper.prepare_db(db)

bot = telebot.TeleBot(config['token'], threaded=False)
@bot.message_handler(content_types=['text'],
                     commands=['start', 'help', 'signup', 'signout', 'default', 'loadratings', 'loadabsents',
                               'loadstudents', 'loadhws', 'getstudents', 'gethws', 'getratings', 'getabsents', 'addhw',
                               'delhw', "addstudent", "delstudent", 'hws', 'absents', "allhws", "showstudents",
                               "showhws", "showratings", "showabsents", 'getchats', 'getconfig', 'loadconfig', 'ask',
                               'channel', 'getask', 'answer', 'delask', 'analysis'])
def command_handler(message):
    global chats, config, commands_dict
    chat = message.chat
    user_id = str(chat.id)
    command = (message.text[1:]).split(' ')[0]
    text = message.text[2 + len(command):].strip()
    try:
        if chat.type == 'private':
            if command == 'start':
                command_start(message, bot, chats, config)
            elif command == 'help':
                command_help(message, bot, chats, commands_dict)
            else:
                if user_id in chats['private']['admin']:
                    if command == 'analysis':
                        file_name = Analysis.analysis(db)
                        doc = open(file_name, 'rb')
                        bot.send_document(user_id, doc)
                    if command == 'getask':
                        ask.get_ask(bot, db, user_id)
                    if command == 'answer':
                        if config['channel'] != '':
                            if text != '':
                                Asks.send_answer(bot, db, user_id, config['channel'], text)
                            else:
                                bot.send_message(user_id, 'Введите сообщение')
                        else:
                            bot.send_message(user_id, "Не задан канал.")
                    if command == 'delask':
                        Asks.get_delete_ask(bot, db, user_id)
                    if command == 'channel':
                        if config['channel'] != '':
                            if text != '':
                                bot.send_message(config['channel'], text)
                            else:
                                bot.send_message(user_id, "Введите текст сообщения")
                        else:
                            bot.send_message(user_id, "Не задан канал.")
                    if command == 'addhw':
                        chats['private']['admin'][user_id] = 'addhw'
                        bot.send_message(user_id,
                                         "Чтобы добавить домашнее задание, отправьте Excel-файл либо введите текст (Тема| Описание).\n")
                    if command == 'delhw':
                        chats['private']['admin'][user_id] = 'delhw'
                        hws = db_helper.execute_select(db, 'SELECT id,name FROM homeworks')

                        if len(hws) > 0:
                            hws = '\n'.join(list(map(lambda x: str(x[0]) + '. ' + x[1], hws)))
                            bot.send_message(user_id,
                                             "Чтобы удалить домашнее задание, выберите соответствующий номер:\n" + hws)
                        else:
                            bot.send_message(user_id, "В базе данных нет дз.")
                    if command == 'loadstudents':
                        chats['private']['admin'][user_id] = 'loadstudents'
                        bot.send_message(user_id, "Чтобы загрузить новый список студентов, отправьте Excel-файл.\n")
                    if command == 'loadhws':
                        chats['private']['admin'][user_id] = 'loadhws'
                        bot.send_message(user_id, "Чтобы загрузить новый список дз, отправьте Excel-файл.\n")
                    if command == 'loadratings':
                        chats['private']['admin'][user_id] = 'loadratings'
                        bot.send_message(user_id, "Чтобы загрузить новый рейтинг студентов, отправьте Excel-файл.\n")
                    if command == 'loadabsents':
                        chats['private']['admin'][user_id] = 'loadabsents'
                        bot.send_message(user_id, "Чтобы загрузить новый список пропусков, отправьте Excel-файл.\n")
                    if command == 'getstudents':
                        tablename = 'students'
                        if save_table_as_xls(db, tablename):
                            doc = open(tablename + '.xls', 'rb')
                            bot.send_document(user_id, doc)
                        else:
                            bot.send_message(user_id, "Список пуст.")
                    if command == 'gethws':
                        tablename = 'homeworks'
                        if save_table_as_xls(db, tablename):
                            doc = open(tablename + '.xls', 'rb')
                            bot.send_document(user_id, doc)
                        else:
                            bot.send_message(user_id, "Список пуст.")
                    if command == 'getratings':
                        tablename = 'ratings'
                        if save_table_as_xls(db, tablename):
                            doc = open(tablename + '.xls', 'rb')
                            bot.send_document(user_id, doc)
                        else:
                            bot.send_message(user_id, "Список пуст.")
                    if command == 'getabsents':
                        tablename = 'absents'
                        if save_table_as_xls(db, tablename):
                            doc = open(tablename + '.xls', 'rb')
                            bot.send_document(user_id, doc)
                        else:
                            bot.send_message(user_id, "Список пуст.")
                    if command == 'getchats':
                        with open('chats_to_send.json', 'w', encoding='utf-8') as chats_file:
                            json.dump(chats, chats_file)
                        doc = open('chats_to_send.json', 'rb')
                        bot.send_document(user_id, doc)
                    if command == 'getconfig':
                        with open('config_to_send.json', 'w', encoding='utf-8') as config_file:
                            json.dump(config, config_file)
                        doc = open('config_to_send.json', 'rb')
                        bot.send_document(user_id, doc)
                    if command == 'loadconfig':
                        chats['private']['admin'][user_id] = 'loadconfig'
                        bot.send_message(user_id, "Чтобы загрузить новый config, отправьте JSON-файл.\n")
                    if command == 'showstudents':
                        rows = db_helper.execute_select(db, 'SELECT * FROM students')
                        if len(rows) > 0:
                            answ = ''
                            for i in range(len(rows)):
                                answ += str(i + 1) + ". {} - {} {}. {}.\n".format(rows[i][0], rows[i][1], rows[i][2][0],
                                                                                  rows[i][3][0])
                            bot.send_message(user_id, answ)
                        else:
                            bot.send_message(user_id, "Список пуст.")
                    if command == 'showhws':
                        rows = db_helper.execute_select(db, 'SELECT * FROM homeworks')
                        if len(rows) > 0:
                            answ = ''
                            for i in range(len(rows)):
                                answ += "{}. {}\n{}\n".format(rows[i][0], rows[i][1], rows[i][2])
                            bot.send_message(user_id, answ)
                        else:
                            bot.send_message(user_id, "Список пуст.")
                    if command == 'showratings':
                        rows = db_helper.execute_select(db, 'SELECT * FROM students')
                        if len(rows) > 0:
                            answ = ''
                            for i in range(len(rows)):
                                rat = stats.get_hws(db, rows[i][0])['done']
                                answ += str(i + 1) + ". {} - {} {}. {}.:  {}\n".format(rows[i][0], rows[i][1],
                                                                                       rows[i][2][0], rows[i][3][0],
                                                                                       rat)
                            bot.send_message(user_id, answ)
                        else:
                            bot.send_message(user_id, "Список пуст.")
                    if command == 'showabsents':
                        rows = db_helper.execute_select(db, 'SELECT * FROM students')
                        if len(rows) > 0:
                            answ = ''
                            for i in range(len(rows)):
                                rat = stats.get_abs(db, rows[i][0])['absents']
                                answ += str(i + 1) + ". {} - {} {}. {}.:  {}\n".format(rows[i][0], rows[i][1],
                                                                                       rows[i][2][0], rows[i][3][0],
                                                                                       rat)
                            bot.send_message(user_id, answ)
                        else:
                            bot.send_message(user_id, "Список пуст.")
                    if command == "addstudent":
                        args = text.split(',')
                        args = list(map(lambda x: x.strip(), args))
                        args = list(filter(lambda x: len(x) > 0, args))
                        if len(args) == 4:
                            user = db_helper.execute_select(db, 'SELECT * FROM students WHERE number={};'.format(
                                args[0]))

                            if len(user) == 0 or (len(user) == 1 and user[0][1] == args[1] and user[0][2] == args[2] and
                                                  user[0][3] == args[3]):
                                db_helper.add_student(db, args)
                                bot.send_message(user_id, "Успешно.")
                            else:
                                bot.send_message(user_id, "Неверно указаны данные.")
                        else:
                            bot.send_message(user_id, "Неверный ввод.")
                    if command == "delstudent":
                        arg = text
                        if len(arg) > 0:
                            user = db_helper.execute_select(db, 'SELECT * FROM students WHERE number={}'.format(
                                '\'' + arg + '\''))
                            if len(user) > 0:
                                db_helper.delete_student(db, arg)
                                bot.send_message(user_id, "Успешно.")
                            else:
                                bot.send_message(user_id, "Неверно указан номер зачетной книжки.")
                        else:
                            bot.send_message(user_id, "Неверный ввод.")
                    if command == 'signup':
                        bot.send_message(user_id, 'Вы уже являетесь админом.')
                    if command == 'signout':
                        bot.send_message(user_id, 'Вы уже являетесь админом.')
                    if command == 'default':
                        chats['private']['admin'][user_id] = 'default'

                elif user_id in chats['private']['student']:
                    if command == 'ask':
                        # print(message)
                        if text != '':
                            ask.save_ask(bot, db, message, user_id, text)
                        else:
                            bot.send_message(user_id, 'Введите сообщение')
                    if command == 'hws':
                        arg = text
                        answ = stats.get_hws(db, arg)
                        if answ['isok']:
                            result = "Выполнено заданий: {}.\n".format(answ['done'])
                            if len(answ['hws']) > 0:
                                hws = list(map(lambda x: '{}. {}\n{}\n'.format(x[0], x[1], x[2]), answ['hws']))
                                result += 'Не выполнено:\n' + ''.join(hws)
                            bot.send_message(user_id, result)
                        else:
                            bot.send_message(user_id, "Неверный ввод.")
                    if command == 'absents':
                        arg = text
                        answ = stats.get_abs(db, arg)
                        if answ['isok']:
                            bot.send_message(user_id, "Пропущено занятий: {}".format(answ['absents']))
                        else:
                            bot.send_message(user_id, "Неверный ввод.")
                    if command == "allhws":
                        hws = db_helper.execute_select(db, 'SELECT * FROM homeworks;')
                        if len(hws) > 0:
                            answ = ''
                            for i in range(len(hws)):
                                answ += '{}. {}\n{}\n\n'.format(hws[i][0], hws[i][1], hws[i][2])
                            bot.send_message(user_id, answ)
                        else:
                            bot.send_message(user_id, "Нет домашних заданий.")
                    if command == 'signup':
                        bot.send_message(user_id, "Вы уже являетесь слушателем курса.")
                    if command == 'signout':
                        finish_date = datetime.strptime(config['finish_date'], '%Y-%m-%d')
                        days_to_finish = (finish_date - datetime.now()).days

                        if days_to_finish < 7:
                            arg = text
                            if len(arg) > 0:
                                user = db_helper.execute_select(db, 'SELECT * FROM students WHERE number={}'.format(
                                    '\'' + arg + '\''))

                                if len(user) > 0:
                                    db_helper.delete_student(db, arg)
                                    if user_id in chats['private']['student']:
                                        del chats['private']['student'][user_id]
                                    bot.send_message(user_id,
                                                     "Вы отписались.\nКомандой /help можете получить доступные команды.")
                                    save_chats()
                                else:
                                    bot.send_message(user_id, "Неверно указан номер зачетной книжки.")
                            else:
                                bot.send_message(user_id, "Неверный ввод.")
                        else:
                            bot.send_message(user_id, 'Курс еще не окончен.')
                    if command == 'default':
                        chats['private']['student'][user_id] = 'default'

                else:
                    if command == 'signup':
                        start_date = datetime.strptime(config['start_date'], '%Y-%m-%d')
                        days_from_start = (datetime.now() - start_date).days

                        if days_from_start <= int(config['days_after']):
                            args = text.split(',')
                            args = list(map(lambda x: x.strip(), args))
                            args = list(filter(lambda x: len(x) > 0, args))
                            if len(args) == 4:
                                user = db_helper.execute_select(db, 'SELECT * FROM students WHERE number={};'.format(
                                    args[0]))

                                if len(user) == 0 or (
                                        len(user) == 1 and user[0][1] == args[1] and user[0][2] == args[
                                    2] and
                                        user[0][3] == args[3]):
                                    db_helper.add_student(db, args)

                                    chats['private']['student'][user_id] = 'default'
                                    bot.send_message(user_id,
                                                     "Вы подписались.\nКомандой /help можете получить доступные команды.")
                                    save_chats()
                                else:
                                    bot.send_message(user_id, "Неверно указаны данные.")
                            else:
                                bot.send_message(user_id, "Неверный ввод.")
                        else:
                            flag = False
                            args = text.split(',')
                            args = list(map(lambda x: x.strip(), args))
                            args = list(filter(lambda x: len(x) > 0, args))
                            if len(args) == 4:
                                user = db_helper.execute_select(db, 'SELECT * FROM students WHERE number={};'.format(
                                    args[0]))

                                if (len(user) == 1 and user[0][1] == args[1] and user[0][2] == args[2] and user[0][3] == args[3]):
                                    flag = True
                                    db_helper.add_student(db, args)

                                    chats['private']['student'][user_id] = 'default'
                                    bot.send_message(user_id,
                                                     "Вы подписались.\nКомандой /help можете получить доступные команды.")
                                    save_chats()
                                else:
                                    bot.send_message(user_id, "Неверно указаны данные.")
                            else:
                                bot.send_message(user_id, "Неверный ввод.")
                            if not flag:
                                bot.send_message(user_id, 'К сожалению, запись на курс закрыта.')
        elif chat.type == 'group':
            pass
        elif chat.type == 'channel':
            pass
    except Exception as ex:
        bot.send_message(user_id, 'Произошла ошибка: ' + str(ex))
        print(ex)


@bot.message_handler(content_types=['text', 'document'])
def get_message(message):
    global commands_dict, chats, config
    chat = message.chat
    user_id = str(chat.id)
    try:
        if chat.type == 'private':
            if user_id in chats['private']['admin']:
                if chats['private']['admin'][user_id] == 'loadconfig':
                    if message.content_type == 'document':
                        res = download_doc(bot, config['token'], message.document, 'new_config', ['json'])
                        if res['is_ok']:
                            with open('new_config.json', 'r', encoding='utf-8') as config_file:
                                config = json.load(config_file)
                        else:
                            bot.send_message(user_id, res['answ'])
                if chats['private']['admin'][user_id] == 'addhw':
                    if message.content_type == 'document':
                        res = download_doc(bot, config['token'], message.document, 'hw', ['xls', 'xlsx'])
                        if res['is_ok']:
                            file_format = message.document.file_name.split('.')[-1]
                            rb = xlrd.open_workbook('hw.' + file_format)
                            sheet = rb.sheet_by_index(0)
                            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]

                            if len(vals) > 0:
                                val = vals[0]
                                if len(val) == 2:
                                    db_helper.add_hw(db, val)
                                    bot.send_message(user_id, "Успешно.")
                                    chats['private']['admin'][user_id] = 'default'
                                else:
                                    bot.send_message(user_id, "Неверный ввод.")
                            else:
                                bot.send_message(user_id, "Файл пуст.")
                        else:
                            bot.send_message(user_id, res['answ'])

                    elif message.content_type == 'text':
                        text = message.text.strip()
                        args = text.split('|')
                        args = list(map(lambda x: x.strip(), args))
                        args = list(filter(lambda x: len(x) > 0, args))

                        if len(args) == 2:
                            db_helper.add_hw(db, args)
                            bot.send_message(user_id, "Успешно.")
                            chats['private']['admin'][user_id] = 'default'
                        else:
                            bot.send_message(user_id, "Неверный ввод.")
                elif chats['private']['admin'][user_id] == 'delhw':
                    text = message.text.strip()
                    if text.isdigit():
                        arg = int(text)
                        hw = db_helper.execute_select(db, 'SELECT * FROM homeworks WHERE id={}'.format(arg))
                        if len(hw) > 0:
                            db_helper.delete_hw(db, arg)
                            bot.send_message(user_id, "Успешно.")
                            chats['private']['admin'][user_id] = 'default'
                        else:
                            bot.send_message(user_id, "Не удалось найти дз с таким номером.")
                    else:
                        bot.send_message(user_id, "Неверный ввод.")
                elif chats['private']['admin'][user_id] == 'loadhws':
                    if message.content_type == 'document':
                        res = download_doc(bot, config['token'], message.document, 'hws', ['xls', 'xlsx'])
                        if res['is_ok']:
                            file_format = message.document.file_name.split('.')[-1]
                            rb = xlrd.open_workbook('hws.' + file_format)
                            sheet = rb.sheet_by_index(0)
                            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]

                            if len(vals) > 0:
                                db_helper.delete_hws(db)
                                not_added = 0
                                for val in vals:
                                    if len(val) == 2:
                                        db_helper.add_hw(db, val)
                                    else:
                                        not_added += 1
                                if not_added == 0:
                                    bot.send_message(user_id, "Успешно.")
                                    chats['private']['admin'][str(user_id)] = 'default'
                                else:
                                    bot.send_message(user_id, "Не добавлено дз: {}.".format(str(not_added)))
                            else:
                                bot.send_message(user_id, "Файл пуст.")
                        else:
                            bot.send_message(user_id, res['answ'])
                elif chats['private']['admin'][user_id] == 'loadabsents':
                    if message.content_type == 'document':
                        res = download_doc(bot, config['token'], message.document, 'absents', ['xls', 'xlsx'])
                        if res['is_ok']:
                            file_format = message.document.file_name.split('.')[-1]
                            rb = xlrd.open_workbook('absents.' + file_format)
                            sheet = rb.sheet_by_index(0)
                            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
                            if len(vals) > 0:
                                del vals[0]

                            if len(vals) > 0:
                                db_helper.recreate_absents(db)
                                for i in range(len(vals[0][4:])):
                                    db_helper.add_absents_column(db)

                                not_updated = 0
                                for val in vals:
                                    stud = db_helper.execute_select(db,
                                                                    'SELECT * FROM students WHERE number=\'{}\''.format(
                                                                        val[0]))
                                    if len(stud) > 0:
                                        val = val[0:1] + val[4:]
                                        for i in range(len(val)):
                                            val[i] = '\'{}\''.format(val[i])
                                        db_helper.add_absent(db, val)
                                    else:
                                        not_updated += 1
                                if not_updated == 0:
                                    bot.send_message(user_id, "Успешно.")
                                    chats['private']['admin'][user_id] = 'default'
                                else:
                                    bot.send_message(user_id, "Не удалось обновить: {}.".format(str(not_updated)))
                            else:
                                bot.send_message(user_id, "Файл пуст.")
                        else:
                            bot.send_message(user_id, res['answ'])
                elif chats['private']['admin'][user_id] == 'loadratings':
                    if message.content_type == 'document':
                        res = download_doc(bot, config['token'], message.document, 'ratings', ['xls', 'xlsx'])
                        if res['is_ok']:
                            file_format = message.document.file_name.split('.')[-1]
                            rb = xlrd.open_workbook('ratings.' + file_format)
                            sheet = rb.sheet_by_index(0)
                            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
                            if len(vals) > 0:
                                del vals[0]

                            if len(vals) > 0:
                                db_helper.recreate_ratings(db)

                                not_updated = 0
                                for val in vals:
                                    stud = db_helper.execute_select(db,
                                                                    'SELECT * FROM students WHERE number=\'{}\''.format(
                                                                        val[0]))
                                    if len(stud) > 0:
                                        val = val[0:1] + val[4:]
                                        for i in range(len(val)):
                                            val[i] = '\'{}\''.format(val[i])
                                        db_helper.update_rating(db, val)
                                    else:
                                        not_updated += 1
                                if not_updated == 0:
                                    bot.send_message(user_id, "Успешно.")
                                    chats['private']['admin'][user_id] = 'default'
                                else:
                                    bot.send_message(user_id, "Не удалось обновить: {}.".format(str(not_updated)))
                            else:
                                bot.send_message(user_id, "Файл пуст.")
                        else:
                            bot.send_message(user_id, res['answ'])
                elif chats['private']['admin'][user_id] == 'loadstudents':
                    if message.content_type == 'document':
                        res = download_doc(bot, config['token'], message.document, 'students', ['xls', 'xlsx'])
                        if res['is_ok']:
                            file_format = message.document.file_name.split('.')[-1]
                            rb = xlrd.open_workbook('students.' + file_format)
                            sheet = rb.sheet_by_index(0)
                            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]

                            if len(vals) > 0:
                                db_helper.delete_students(db)
                                chats['private']['student'] = {}
                                not_updated = 0
                                for val in vals:
                                    if len(val) == 4:
                                        db_helper.add_student(db, val)
                                    else:
                                        not_updated += 1
                                if not_updated == 0:
                                    bot.send_message(user_id, "Успешно.")
                                    chats['private']['admin'][user_id] = 'default'
                                else:
                                    bot.send_message(user_id, "Не удалось добавить: {}.".format(str(not_updated)))
                            else:
                                bot.send_message(user_id, "Файл пуст.")
                        else:
                            bot.send_message(user_id, res['answ'])
                elif chats['private']['admin'][user_id] == 'default':
                    if message.content_type == 'text':
                        lessons_bot_AI(config['small_talk'], message, bot)
            elif user_id in chats['private']['student']:
                if chats['private']['student'][str(user_id)] == 'default':
                    lessons_bot_AI(config['small_talk'], message, bot)
            else:
                lessons_bot_AI(config['small_talk'], message, bot)
        if chat.type == 'group':
            pass
        if chat.type == 'channel':
            pass
    except Exception as ex:
        bot.send_message(message.chat.id, 'Произошла ошибка: ' + str(ex))
        print(ex)


if __name__ == '__main__':
    while True:
        try:
            bot.polling(none_stop=True)

            while True:
                time.sleep(50)
                pass
        except KeyboardInterrupt:
            exit()
        except Exception as ex:
            print(ex)
