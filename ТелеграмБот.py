import telebot
import openpyxl
from flask import Flask, request
import os
import openpyxl

EXCEL_FILE = 'препараты.xlsx'

def create_excel_if_not_exists():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Препараты"
        ws.append(['Название препарата', 'Описание', 'Дозировка'])
        # Пример строки (можно удалить или заменить)
        ws.append(['Парацетамол', 'Обезболивающее и жаропонижающее средство', '500 мг 3 раза в день'])
        wb.save(EXCEL_FILE)

create_excel_if_not_exists()

# Укажите ваш токен бота и webhook url
TOKEN = '8410859830:AAEqRSwrAXfUQaf4POcBRy9f_w9CS5_TNxE'
bot = telebot.TeleBot(TOKEN)
WEBHOOK_URL = 'https://ВАШ-РЕНДЕР-АДРЕС.onrender.com/'  # замените на ваш адрес Render

# Путь к файлу Excel
EXCEL_FILE = 'препараты.xlsx'

# Функция для чтения данных из Excel
def read_drugs_from_excel():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb.active
        drugs = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            name, description, dosage = row
            drugs.append({
                'name': name,
                'description': description,
                'dosage': dosage
            })
        return drugs
    except Exception as e:
        return []

# Команда /start
@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.send_message(message.chat.id, "Здравствуйте! Я бот-справочник по препаратам. Введите название препарата или используйте команду /list для просмотра всех препаратов.")

# Команда /list для вывода всех препаратов
@bot.message_handler(commands=['list'])
def list_drugs(message):
    drugs = read_drugs_from_excel()
    if not drugs:
        bot.send_message(message.chat.id, "Не удалось загрузить список препаратов.")
        return
    reply = "Список препаратов:\n"
    for drug in drugs:
        reply += f"- {drug['name']}\n"
    bot.send_message(message.chat.id, reply)

# Поиск препарата по названию
@bot.message_handler(func=lambda message: True)
def search_drug(message):
    query = message.text.strip().lower()
    drugs = read_drugs_from_excel()
    found = None
    for drug in drugs:
        if drug['name'] and query in drug['name'].lower():
            found = drug
            break
    if found:
        reply = f"Название: {found['name']}\nОписание: {found['description']}\nДозировка: {found['dosage']}"
    else:
        reply = "Препарат не найден. Попробуйте другое название или используйте /list."
    bot.send_message(message.chat.id, reply)

# Flask-приложение для Render
app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def webhook():
    if request.method == 'POST':
        json_str = request.get_data().decode('UTF-8')
        update = telebot.types.Update.de_json(json_str)
        bot.process_new_updates([update])
        return '', 200
    return 'Бот работает!', 200

# Установка webhook при запуске (без before_first_request)
import os

def set_webhook():
    bot.remove_webhook()
    bot.set_webhook(url=WEBHOOK_URL)

set_webhook()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
