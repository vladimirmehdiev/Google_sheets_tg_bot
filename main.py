import telebot
from telebot import types
import gspread
from oauth2client.service_account import ServiceAccountCredentials

bot = telebot.TeleBot('#####')  # Подключение через API к боту


def get_data_from_google_sheet():
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            "tg-exel-d3ec78530c3d.json",
            [
                "#####",
                "####"
            ]
        )

        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url("#####")
        sheet = spreadsheet.sheet1
        data = sheet.get_all_records()

        return data
    except Exception as e:
        print(f"Произошла ошибка при чтении Google Таблицы: {e}")
        return None

#-----------------------------------------------------------------

# Глобальные переменные для хранения выбора пользователя
selected_program = None
selected_year = None
selected_indicator = None
selected_region = None
select_value = None

# Обработка выбора программы
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn_cis = types.KeyboardButton('ЦИС')
    btn_grant = types.KeyboardButton('Гранты на патентование')
    markup.row(btn_cis, btn_grant)
    bot.send_message(message.chat.id,
                     f'Здравия желаю, {message.from_user.first_name} {message.from_user.last_name}. Выберите вашу программу:',
                     reply_markup=markup)

# Обработка выбора показателя
@bot.message_handler(func=lambda message: message.text in ['ЦИС'])
def handle_program_selection(message):
    global selected_program
    selected_program = message.text
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn1 = types.KeyboardButton('Кол-во поданных заявок на изобретения, полезные модели и промышленные образцы')
    btn2 = types.KeyboardButton('Кол-во поданных заявок на изобретения')
    markup.row(btn1, btn2)
    btn3 = types.KeyboardButton('Кол-во поданных заявок на полезные модели')
    btn4 = types.KeyboardButton('Кол-во поданных заявок на промышленные образцы')
    markup.row(btn3, btn4)
    bot.send_message(message.chat.id, f'Вы выбрали программу: {selected_program}. Теперь выберите показатель:',
                     reply_markup=markup)

# Обработка выбора показателя
@bot.message_handler(func=lambda message: message.text in ['Гранты на патентование'])
def handle_program_selection1(message):
    global selected_program
    selected_program = message.text
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn1 = types.KeyboardButton('Кол-во поданных заявок')
    btn2 = types.KeyboardButton('Запрашиваемая сумма')
    markup.row(btn1, btn2)
    btn3 = types.KeyboardButton('Поданное число патентов')
    btn4 = types.KeyboardButton('Количество отклоненных и отозванных заявок')
    markup.row(btn3, btn4)
    btn5 = types.KeyboardButton('Сумма по отклоненным и отозванным заявкам')
    btn6 = types.KeyboardButton('Число патентов по отклоненным и отозванным заявкам')
    markup.row(btn5, btn6)
    btn7 = types.KeyboardButton('Кол-во одобренных заявок')
    btn8 = types.KeyboardButton('Одобренная сумма')
    markup.row(btn7, btn8)
    btn9 = types.KeyboardButton('Одобренное число патентов')
    btn10 = types.KeyboardButton('Кол-во выплаченных заявок')
    markup.row(btn9, btn10)
    btn11 = types.KeyboardButton('Выплаченная сумма')
    btn12 = types.KeyboardButton('Выплаченное число патентов')
    markup.row(btn11, btn12)
    btn13 = types.KeyboardButton('Кол-во компаний по выплаченным заявкам')
    markup.row(btn13)
    bot.send_message(message.chat.id, f'Вы выбрали программу: {selected_program}. Теперь выберите показатель:',
                     reply_markup=markup)

@bot.message_handler(func=lambda message: message.text in ['Кол-во поданных заявок на изобретения, полезные модели и промышленные образцы',
                                                           'Кол-во поданных заявок на изобретения', 'Кол-во поданных заявок на полезные модели',
                                                           'Кол-во поданных заявок на промышленные образцы', 'Кол-во поданных заявок',
                                                           'Запрашиваемая сумма', 'Поданное число патентов', 'Количество отклоненных и отозванных заявок',
                                                           'Сумма по отклоненным и отозванным заявкам', 'Число патентов по отклоненным и отозванным заявкам',
                                                           'Кол-во одобренных заявок', 'Одобренная сумма', 'Одобренное число патентов', 'Кол-во выплаченных заявок',
                                                           'Выплаченная сумма', 'Выплаченное число патентов', 'Кол-во компаний по выплаченным заявкам'])
def handle_indicator_selection(message):
    global selected_indicator
    selected_indicator  = message.text
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn1 = types.KeyboardButton('2018')
    btn2 = types.KeyboardButton('2019')
    markup.row(btn1, btn2)
    btn3 = types.KeyboardButton('2020')
    btn4 = types.KeyboardButton('2021')
    markup.row(btn3, btn4)
    btn5 = types.KeyboardButton('2022')
    btn6 = types.KeyboardButton('2023')
    markup.row(btn5, btn6)
    btn7 = types.KeyboardButton('2024')
    btn8 = types.KeyboardButton('2018-2024')
    markup.row(btn7, btn8)
    bot.send_message(message.chat.id, f'Вы выбрали показатель: {selected_indicator}. Теперь выберите год:',
                     reply_markup=markup)

@bot.message_handler(func=lambda message: message.text in ['2018', '2019', '2020', '2021', '2022', '2023', '2024', '2018-2024'])
def handle_year_selection(message):
    global selected_year
    selected_year = message.text

    # Если выбрана программа "ЦИС", то продолжаем с выбором региона
    if selected_program == 'ЦИС':
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn_moscow = types.KeyboardButton('Москва')
        btn_russia_moscow = types.KeyboardButton('Россия (вкл. Москву)')
        markup.row(btn_moscow)
        markup.row(btn_russia_moscow)
        bot.send_message(message.chat.id, f'Вы выбрали год: {selected_year}. Теперь выберите регион:',
                         reply_markup=markup)
    else:
        # Для "Гранты на патентование" сразу переходим к расчету
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn_calculate = types.KeyboardButton('Рассчитать сколько было подано заявок')
        btn_back = types.KeyboardButton('Назад')
        markup.row(btn_calculate)
        markup.row(btn_back)
        bot.send_message(message.chat.id, f'Вы выбрали год: {selected_year}. Теперь выберите действие:',
                         reply_markup=markup)


@bot.message_handler(func=lambda message: message.text in ['Москва', 'Россия (вкл. Москву)'])
def handle_region_selection(message):
    global selected_region
    selected_region = message.text
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn_calculate = types.KeyboardButton('Рассчитать сколько было подано заявок')
    btn_back = types.KeyboardButton('Назад')
    markup.row(btn_calculate)
    markup.row(btn_back)
    bot.send_message(message.chat.id, f'Вы выбрали регион: {selected_region}. Теперь выберите действие', reply_markup=markup)

# Расчет по выбранной программе и году
@bot.message_handler(func=lambda message: message.text == 'Рассчитать сколько было подано заявок')
def calculate_applications(message):
    global selected_indicator, selected_year, select_value
    data = get_data_from_google_sheet()

    if data is None or not data:
        bot.send_message(message.chat.id, "Данные не найдены или Google Таблица пуста.")
        return

    total = 0
    response = f"Данные из Google Таблицы по показателю {selected_indicator} в году {selected_year}:\n"

    years = selected_year
    if selected_year == '2018-2024' and selected_program != 'Гранты на патентование':
        years = ['2018', '2019', '2020', '2021', '2022', '2023', '2024']
    else:
        years = [selected_year]

    for row in data:
        # Фильтруем строки по выбранной программе и году
        if row.get('Наименование показателя') == selected_indicator and selected_year in row:
            if row.get('Регион') == selected_region or row.get('Программа (2)') == 'Гранты на патентование':
                for year in years:
                    if year in row and isinstance(row[year], (int, float)):  # Проверка, что значение числовое
                        total += row[year]
                select_value = row.get('Ед. Изм.')


    response += f"\nСумма чисел в столбце {selected_year}: {total} {select_value}"
    bot.send_message(message.chat.id, response)

#-----------------------------------------------------------------
@bot.message_handler(func=lambda message: message.text == 'Назад')
def go_back(message):
    global selected_year, selected_program
    selected_year = None  # Сбрасываем выбранный год
    selected_program = None  # Сбрасываем выбранный регион

    # Возвращаемся в главное меню
    start(message)

bot.polling(non_stop=True)