import logging
import os
import sys
from urllib.parse import quote
from itertools import zip_longest
from emoji import emojize
from pyexcel_xlsx import get_data
from model import save, Schedule, Tags, DoesNotExist, Users, fn
from telegram import ReplyKeyboardMarkup, ParseMode, ReplyKeyboardRemove, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import Updater, CommandHandler, RegexHandler, MessageHandler, Filters, ConversationHandler, CallbackQueryHandler
from config import ALLTESTS, ADMIN_ID, PROD, MySQL, ADMINS, log, OKSANA, html_template
import pdfkit
from PyPDF2 import PdfFileMerger
import shutil

pdfs_path = './schedules'

dbs = {'courses': Schedule,
       'tags': Tags}

days_of_week = {'Суббота': 'Субботам',
                'Воскресенье': 'Воскресеньям'}

SECOND, THIRD, FORTH, FIFTH, SIX = range(5)
start_keyboard = [['Записаться'], ['Расписание'], [KeyboardButton('Перезвонить мне', request_contact=True)]]
user_data = {}


def compile_msg(d):
    s = ''
    _id, place, course_name, metro, address, comments, age_from, age_to, \
    period, price, lecturer, day_of_week, time, db_dt = d
    day_of_week = days_of_week[day_of_week.capitalize()]
    address_url = quote(address)
    s += f'<b>=={course_name}==</b>\n'
    s += f'<b>Адрес:</b> <a href="http://maps.google.com/?q={address_url}">{place}, {address}</a>\n'
    s += f'<b>Метро:</b> {metro}\n'
    s += f'<b>Для кого:</b> {comments.capitalize()}\n'
    s += f'<b>Возраст:</b> {age_from}-{age_to} лет\n'
    s += f'<b>По</b> {day_of_week.lower()} c {time}\n'
    s += f'<b>Когда:</b> {period}\n'
    s += f'<b>Сколько стоит:</b> {price}\n'
    s += f'<b>Преподаватель:</b> {lecturer}\n'
    return s


@log
def start(bot, update):
    username = update.message.from_user.username
    name = update.message.from_user.first_name
    uid = update.message.from_user.id
    if user_data.get(uid):
        del user_data[uid]
    try:
        Users.get(Users.telegram_id == uid)
    except DoesNotExist:
        Users.create(telegram_id=uid, username=username, name=name)
    bot.sendMessage(uid, 'Привет! Выбирайте действие ' + emojize(':winking_face:'),
                    disable_web_page_preview=True,
                    reply_markup=ReplyKeyboardMarkup(start_keyboard))
    return ConversationHandler.END


@log
def start_enroll(bot, update):
    uid = update.message.from_user.id
    bot.sendMessage(uid, 'Введите возраст ребенка', reply_markup=ReplyKeyboardRemove())
    return SECOND


@log
def age_preferences(bot, update):
    uid = update.message.from_user.id
    message = update.message.text.strip()
    try:
        age = int(message)
        user_data[uid] = {'age': age}
        tags = MySQL(f'''select distinct
                            tag2
                        from tags
                        where true
                            and age_from <= {age}
                            and age_to >= {age}
                        order by tag2''').fetchall()
        tags = [[t[0]] for t in tags] + [['Назад']]
        user_data[uid]['tags'] = tags
        bot.sendMessage(uid, 'Выберите одно из увлечений ребенка',
                        reply_markup=ReplyKeyboardMarkup(tags))
        return THIRD
    except ValueError:
        bot.sendMessage(uid, 'Введите возраст цифрой')


@log
def prefs2_prefs3(bot, update):
    uid = update.message.from_user.id
    message = update.message.text
    if [message] not in user_data[uid]['tags']:
        return
    if message == 'Назад':
        bot.sendMessage(uid, 'Вот предыдущее меню',
                        disable_web_page_preview=True,
                        reply_markup=ReplyKeyboardMarkup(start_keyboard))
        return ConversationHandler.END
    age = user_data[uid]['age']
    user_data[uid]['interests'] = [message]
    tags = MySQL(f'''select distinct
                        tag3
                    from tags
                    where true
                        and age_from <= {age}
                        and age_to >= {age}
                        and tag2 = "{message}"
                    order by tag3''').fetchall()
    tags = [[t[0]] for t in tags] + [['Назад']]
    user_data[uid]['tags'] = tags
    bot.sendMessage(uid, 'Выберите одно из увлечений ребенка',
                        reply_markup=ReplyKeyboardMarkup(tags, one_time_keyboard=True))
    try:
        Tags.get(Tags.tag4.is_null(is_null=False))
    except DoesNotExist:
        return FIFTH
    return FORTH


@log
def prefs3_prefs4(bot, update):
    uid = update.message.from_user.id
    message = update.message.text
    if [message] not in user_data[uid]['tags']:
        return
    if message == 'Назад':
        bot.sendMessage(uid, 'Вот предыдущее меню',
                        disable_web_page_preview=True,
                        reply_markup=ReplyKeyboardMarkup(start_keyboard))
        return ConversationHandler.END
    user_data[uid]['interests'].append(message)
    age = user_data[uid]['age']
    tags = MySQL(f'''select distinct
                        tag4
                    from tags
                    where true
                        and age_from <= {age}
                        and age_to >= {age}
                        and tag3 = "{message}"
                    order by tag4''').fetchall()
    tags = [[t[0]] for t in tags] + [['Назад']]
    bot.sendMessage(uid, 'Выберите одно из увлечений ребенка',
                        reply_markup=ReplyKeyboardMarkup(tags, one_time_keyboard=True))
    return FIFTH


@log
def preferences_send_courses(bot, update):
    uid = update.message.from_user.id
    message = update.message.text
    if [message] not in user_data[uid]['tags']:
        return
    if message == 'Назад':
        bot.sendMessage(uid, 'Вот предыдущее меню',
                        disable_web_page_preview=True,
                        reply_markup=ReplyKeyboardMarkup(start_keyboard))
        return ConversationHandler.END
    user_data[uid]['interests'].append(message)

    age = user_data[uid]['age']
    user_data[uid]['interests'].append(None)
    tag2, tag3, tag4 = user_data[uid]['interests']
    msgs = []
    course_form_tags = Tags.get(Tags.age_from <= age,
                                Tags.age_to >= age,
                                Tags.tag2 == tag2,
                                Tags.tag3 == tag3,
                                Tags.tag4 == tag4).course_name
    data = Schedule.\
        select().\
        where(
              Schedule.course_name == course_form_tags).\
        order_by(Schedule.course_name)
    data = [c for c in data.tuples()]
    for d in data:
        s = compile_msg(d)
        msgs.append(s)

    user_data[uid]['courses'] = msgs
    user_data[uid]['approved'] = False
    len_courses = len(msgs)
    if len_courses == 0:
        bot.sendMessage(uid, 'Для этого возраста у нас пока нет подходящих курсов :(', reply_markup=ReplyKeyboardMarkup(start_keyboard))
        return ConversationHandler.END

    buttons = [[
        InlineKeyboardButton(emojize(':black_rightwards_arrow:'), callback_data='right_0')
    ],
        [InlineKeyboardButton('Записаться на этот курс', callback_data='approve_0')],
        [InlineKeyboardButton('Назад', callback_data='back_1')]
    ]

    msg = f'1/{len_courses}\n' + msgs[0]
    reply_markupi = InlineKeyboardMarkup(buttons)
    bot.sendMessage(uid,
                    msg,
                    reply_markup=reply_markupi,
                    parse_mode=ParseMode.HTML,
                    disable_web_page_preview=True)
    return SIX


@log
def confirm_value(bot, update):
    query = update.callback_query
    uid = query.from_user.id
    message = query.data
    user_courses = user_data[uid]['courses']
    action, i = message.split('_')
    i = int(i)
    if action == 'right':
        i += 1
    elif action == 'left':
        i -= 1
    elif action == 'back':
        bot.editMessageText(chat_id=uid,
                            message_id=query.message.message_id,
                            text='Предыдущее меню')
        bot.sendMessage(uid, 'Выбирай действие ' + emojize(':winking_face:'),
                        disable_web_page_preview=True,
                        disable_notification=True,
                        reply_markup=ReplyKeyboardMarkup(start_keyboard))
        return ConversationHandler.END
    elif action == 'approve':
        user_data[uid]['approved'] = i
        keyboard = [[KeyboardButton('Перезвонить мне', request_contact=True)], ['Назад']]
        bot.editMessageText(chat_id=uid,
                            message_id=query.message.message_id,
                            text='Отличный выбор!')
        bot.sendMessage(uid,
                        'Отправьте свой номер телефона, чтобы наш оператор связался с вами ' + emojize(':winking_face:'),
                        reply_markup=ReplyKeyboardMarkup(keyboard))
        return ConversationHandler.END

    buttons = [[

    ],
        [InlineKeyboardButton('Записаться на этот курс', callback_data=f'approve_{i}')],
        [InlineKeyboardButton('Назад', callback_data='back_1')]
    ]
    if i > 0:
        if i == len(user_courses) - 1:
            buttons[0].append(InlineKeyboardButton(emojize(':leftwards_black_arrow:'), callback_data=f'left_{i}'))
        else:
            buttons[0].extend([InlineKeyboardButton(emojize(':leftwards_black_arrow:'), callback_data=f'left_{i}'),
                               InlineKeyboardButton(emojize(':black_rightwards_arrow:'), callback_data=f'right_{i}')
                               ])
    if i == 0:
        buttons[0].append(InlineKeyboardButton(emojize(':black_rightwards_arrow:'), callback_data='right_0'))

    reply_markupi = InlineKeyboardMarkup(buttons)
    len_courses = len(user_courses)
    msg = f'{i+1}/{len_courses}\n' + user_courses[i]

    bot.editMessageText(chat_id=uid,
                        message_id=query.message.message_id,
                        text=msg, reply_markup=reply_markupi,
                        parse_mode=ParseMode.HTML,
                        disable_web_page_preview=True)


@log
def start_schedule(bot, update):
    uid = update.message.from_user.id
    # keyboard = [['Метро'], ['Курс'], ['Возраст'], ['Полное расписание'], ['Назад']]
    keyboard = [['Метро'], ['Курс'], ['Возраст'], ['Назад']]
    bot.sendMessage(uid, 'По чему будем искать?', reply_markup=ReplyKeyboardMarkup(keyboard))
    return SECOND


@log
def select_category(bot, update):
    uid = update.message.from_user.id
    message = update.message.text
    if not user_data.get(uid):
        user_data[uid] = {}
    if message == 'Метро':
        user_data[uid]['category'] = 'metro'
        data = [s.metro for s in Schedule.select(Schedule.metro).distinct().order_by(Schedule.metro)]
    elif message == 'Возраст':
        user_data[uid]['category'] = 'age_from'
        data = [a.age_from for a in Schedule.select(Schedule.age_from).distinct().order_by(Schedule.age_from)]
    elif message == 'Курс':
        user_data[uid]['category'] = 'course_name'
        data = [c.course_name for c in Schedule.select(Schedule.course_name).distinct().order_by(Schedule.course_name)]
    elif message == 'Назад':
        bot.sendMessage(uid, 'Вот предыдущее меню',
                        disable_web_page_preview=True,
                        reply_markup=ReplyKeyboardMarkup(start_keyboard))
        return ConversationHandler.END
    elif message == 'Полное расписание':
        user_data[uid]['category'] = 'all'
        bot.sendMessage(uid, 'Тут должно быть раписание на весь месяц', reply_markup=ReplyKeyboardMarkup(start_keyboard))
        return ConversationHandler.END
    else:
        bot.sendMessage(uid, 'Выбирайте один из вариантов на клавиатуре')
        return
    keyboard = [[f'{k}+'] if isinstance(k, int) else [k] for k in data] + [['Назад']]
    bot.sendMessage(uid, 'Выбирай', reply_markup=ReplyKeyboardMarkup(keyboard))
    return THIRD


@log
def get_course_data(bot, update):
    uid = update.message.from_user.id
    message = update.message.text.strip('+')
    if message == 'Назад':
        bot.sendMessage(uid, 'Вот предыдущее меню',
                        disable_web_page_preview=True,
                        reply_markup=ReplyKeyboardMarkup(start_keyboard))
        return ConversationHandler.END
    pdf_name_final = os.path.join(pdfs_path, 'Раписание_{}.pdf'.format(message))
    if os.path.exists(pdf_name_final):
        bot.sendDocument(uid, document=open(pdf_name_final, 'rb'))
        bot.sendMessage(uid, 'Выбирайте действие ' + emojize(':winking_face:'),
                        reply_markup=ReplyKeyboardMarkup(start_keyboard))
        return ConversationHandler.END

    query = '''
    select
      place,
      course_name,
      metro,
      address,
      comments,
      age_from,
      age_to,
      period,
      price,
      lecturer,
      day_of_week,
      time
    from schedule
    where TRUE
      and {}
    ORDER by course_name
    '''

    column = user_data[uid]['category']
    config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')
    s = f'{column} >= {message}' if column == 'age' else f'{column} = "{message}"'
    data = MySQL(query.format(s)).fetchall()
    pdfs = []
    part = 6
    c = 0
    data = [data[n:n+part] for n in range(0, len(data), part)]
    for element in data:
        s = ''
        for r in element:
            c += 1
            place, course_name, metro, address, comments, age_from, age_to, \
            period, price, lecturer, days_of_week, _time = r
            s += f'<tr><td>Площадка</td><td>{place}</td></tr>' \
                f'<tr><td>Курс</td><td>{course_name}</td></tr>' \
                f'<tr><td>Метро</td><td>{metro}</td></tr>' \
                f'<tr><td>Адрес</td><td>{address}</td></tr>' \
                f'<tr><td>Комментарии</td><td>{comments}</td></tr>' \
                f'<tr><td>Возраст</td><td>{age_from}-{age_to} лет</td></tr>' \
                f'<tr><td>Когда</td><td>{period}</td></tr>' \
                f'<tr><td>Цена</td><td>{price}</td></tr>' \
                f'<tr><td>Лектор</td><td>{lecturer}</td></tr>' \
                f'<tr><td>День недели</td><td>{days_of_week.capitalize()}</td></tr>' \
                f'<tr><td></td><td></td></tr>' \
                f'<tr><td></td><td></td></tr>'

        string = html_template.format(s)
        pdf_name = str(c) + '.pdf'
        pdfkit.from_string(string, pdf_name, configuration=config)
        pdfs.append(pdf_name)

    if not os.path.exists(pdfs_path):
        os.makedirs(pdfs_path)
    if len(pdfs) == 1:
        shutil.move(pdfs[0], pdf_name_final)
    else:
        merger = PdfFileMerger()
        for pdf in pdfs:
            merger.append(pdf)
            os.remove(pdf)
        merger.write(pdf_name_final)
    bot.sendDocument(uid, document=open(pdf_name_final, 'rb'))
    bot.sendMessage(uid, 'Выбирайте действие ' + emojize(':winking_face:'),
                    reply_markup=ReplyKeyboardMarkup(start_keyboard))
    return ConversationHandler.END


@log
def process_file(bot, update):
    uid = update.message.from_user.id
    if uid in ADMINS:
        file_id = update.message.document.file_id
        fname = update.message.document.file_name
        new_file = bot.getFile(file_id)
        new_file.download(fname)
        sheets = get_data(fname)
        tables = [Schedule, Tags]
        for t in tables:
            if t.table_exists():
                t.drop_table()
            t.create_table()
        msg_report = ''
        for sheet in sheets:
            columns = ('place', 'course_name', 'metro', 'address', 'comments',
                       'age_from', 'age_to', 'period', 'price', 'lecturer', 'day_of_week', 'time')
            _data = []
            for row in sheets[sheet][1:]:
                if not row:
                    continue
                if sheet.lower() == 'tags':
                    columns = ['course_name', 'age_from', 'age_to', 'tag2', 'tag3', 'tag4']
                    _data.append(dict(zip_longest(columns,
                                                  [r.strip('"\'!?[]{},. \n')
                                                   if not isinstance(r, (float, int, type(None))) else r
                                                   for r in row],
                                                  fillvalue=None)))
                else:
                    _data.append(dict(zip_longest(columns,
                                                  [r.strip('"\'!?[]{},. \n')
                                                   if not isinstance(r, (float, int, (type(None)))) else r
                                                   for r in row],
                                                  fillvalue=None)))
            if save(_data, dbs[sheet]):
                msg_report += 'Данные на странице <b>{}</b> сохранил\n'.format(sheet)

            else:
                msg_report += 'Что-то не так с данными на странице <b>{}</b>'.format(sheet),
        bot.sendMessage(uid, msg_report, parse_mode=ParseMode.HTML)

        if not os.path.exists(pdfs_path):
            os.makedirs(pdfs_path)

        for f in os.listdir(pdfs_path):
            if f.endswith('pdf'):
                rmvpath = os.path.join(pdfs_path, f)
                print(rmvpath)
                os.remove(rmvpath)
        os.remove(fname)


@log
def process_contact(bot, update):
    uid = update.message.from_user.id
    bot.sendMessage(uid, 'Оператор уже читает ваше сообщение ' + emojize(':smiling_face_with_sunglasses:'),
                    reply_markup=ReplyKeyboardMarkup(start_keyboard))
    contact = update.message.contact
    first_name = contact.first_name
    last_name = contact.last_name
    phone_number = contact.phone_number
    user_course = user_data.get(uid)
    msg = ''
    if user_course:
        approved = user_course['approved']
        chosen_course = user_course['courses'][approved]
        interests = user_course['interests']

        msg += f'{first_name} {last_name}\n'
        msg += f'+{phone_number}\n'
        msg += f'Интересы ребенка: {interests}\n\n'
        msg += f'{chosen_course}\n'
        del user_data[uid]
    else:
        msg += f'{first_name} {last_name}\n'
        msg += f'+{phone_number}\n\n'
        msg += f'Помочь в выбором курса'

    bot.sendMessage(OKSANA,
                    msg,
                    disable_web_page_preview=True,
                    parse_mode=ParseMode.HTML)


@log
def cancel(bot, update):
    uid = update.message.from_user.id
    if user_data.get(uid):
        del user_data[uid]
    bot.sendMessage(uid, 'Вот предыдущее меню',
                    reply_markup=ReplyKeyboardMarkup(start_keyboard))
    return ConversationHandler.END


if __name__ == '__main__':
    updater = None
    token = None
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    if len(sys.argv) > 1:
        token = sys.argv[-1]
        if token.lower() == 'prod':
            updater = Updater(PROD)
            logging.basicConfig(filename=BASE_DIR + '/out.log',
                                filemode='a',
                                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                                level=logging.INFO)
    else:
        updater = Updater(ALLTESTS)
        logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                            level=logging.DEBUG)
    dp = updater.dispatcher
    enroll = ConversationHandler(
        entry_points=[RegexHandler('^Записаться$', start_enroll)],
        states={SECOND: [MessageHandler(Filters.text, age_preferences)],
                THIRD: [MessageHandler(Filters.text, prefs2_prefs3)],
                FORTH: [MessageHandler(Filters.text, prefs2_prefs3)],
                FIFTH: [MessageHandler(Filters.text, preferences_send_courses)],
                SIX: [CallbackQueryHandler(confirm_value)]},
        fallbacks=[RegexHandler('^Назад$', cancel),
                   CommandHandler('start', start)]
    )

    schedule = ConversationHandler(
        entry_points=[RegexHandler('^Расписание$', start_schedule)],
        states={
            SECOND: [MessageHandler(Filters.text, select_category)],
            THIRD: [MessageHandler(Filters.text, get_course_data)]
                },
        fallbacks=[RegexHandler('^Назад$', cancel),
                   CommandHandler('start', start)]
    )
    dp.add_handler(enroll)
    dp.add_handler(schedule)
    dp.add_handler(CommandHandler('start', start))
    dp.add_handler(RegexHandler('^Назад$', cancel))
    dp.add_handler(MessageHandler(Filters.document, process_file))
    dp.add_handler(MessageHandler(Filters.contact, process_contact))
    updater.start_polling()
    updater.idle()
