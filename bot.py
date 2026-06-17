# ============================================================
#  БОТ ДЛЯ СБОРА ЗАЯВОК + EMAIL + КНОПКИ
# ============================================================

import telebot
import openpyxl
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from telebot.types import ReplyKeyboardMarkup, KeyboardButton

# ===================== НАСТРОЙКИ =====================
# Токен от BotFather
TOKEN = 'ВАШ_ТОКЕН'

# Настройки почты (Яндекс)
EMAIL_SENDER = 'orlovm94@yandex.ru'
EMAIL_PASSWORD = 'ВАШ_ПАРОЛЬ_ПРИЛОЖЕНИЯ'

EMAIL_SERVER = 'smtp.yandex.ru'
EMAIL_PORT = 465

# Куда отправлять заявки
EMAIL_RECIPIENT = 'orlovm94@yandex.ru'

# Excel создаётся В ПАПКЕ С БОТОМ
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, 'Заявки.xlsx')
# ============================================================

bot = telebot.TeleBot(TOKEN)
user_states = {}

# ============ КНОПКИ ============
def get_main_keyboard():
    """Главная клавиатура с кнопками."""
    markup = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=False)
    markup.add(KeyboardButton('📝 Оставить заявку'))
    markup.add(KeyboardButton('❌ Отменить'))
    return markup

# ============ ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ============
def create_excel_if_not_exists():
    """Создаёт Excel-файл с заголовками, если его нет."""
    if not os.path.exists(EXCEL_FILE):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = 'Заявки'
        headers = ['Дата', 'Имя', 'Телефон', 'Услуга', 'Telegram ID']
        sheet.append(headers)
        workbook.save(EXCEL_FILE)
        print(f'Файл {EXCEL_FILE} создан.')

def save_to_excel(data):
    """Сохраняет заявку в Excel."""
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active
    row = [
        data['date'],
        data['name'],
        data['phone'],
        data['service'],
        data['user_id']
    ]
    sheet.append(row)
    workbook.save(EXCEL_FILE)
    print(f'Заявка от {data["name"]} сохранена в Excel.')

def send_email_notification(data):
    """Отправляет письмо с заявкой на email."""
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECIPIENT
        msg['Subject'] = 'Новая заявка с Telegram-бота!'

        body = f"""
Здравствуйте!

Поступила новая заявка через Telegram-бота:

📅 Дата: {data['date']}
👤 Имя: {data['name']}
📞 Телефон: {data['phone']}
🔧 Услуга: {data['service']}
📱 Telegram ID: {data['user_id']}

Свяжитесь с клиентом как можно скорее!
        """

        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        # Яндекс — SSL, порт 465
        server = smtplib.SMTP_SSL(EMAIL_SERVER, EMAIL_PORT)
        # Логин без @yandex.ru
        server.login(EMAIL_SENDER.split('@')[0], EMAIL_PASSWORD)
        server.sendmail(EMAIL_SENDER, EMAIL_RECIPIENT, msg.as_string())
        server.quit()

        print(f'Письмо отправлено на {EMAIL_RECIPIENT}')
        return True

    except Exception as e:
        print(f'Ошибка при отправке письма: {e}')
        return False

# ============ ОБРАБОТЧИКИ КОМАНД ============

@bot.message_handler(commands=['start'])
def start_handler(message):
    """Приветствие — показывает кнопки, но не начинает заявку."""
    user_id = message.chat.id
    # Сбрасываем состояние, если было
    if user_id in user_states:
        del user_states[user_id]

    bot.send_message(
        user_id,
        '👋 Привет! Я бот для сбора заявок.\n\n'
        'Нажмите «📝 Оставить заявку», чтобы начать.',
        reply_markup=get_main_keyboard()
    )

@bot.message_handler(commands=['cancel'])
def cancel_handler(message):
    """Отменяет заполнение заявки."""
    user_id = message.chat.id

    if user_id in user_states:
        del user_states[user_id]

    bot.send_message(
        user_id,
        '❌ Заполнение заявки отменено.\n\n'
        'Нажмите «📝 Оставить заявку», чтобы начать заново.',
        reply_markup=get_main_keyboard()
    )

# ============ ОБРАБОТЧИК ВСЕХ СООБЩЕНИЙ ============

@bot.message_handler(func=lambda message: True)
def handle_text(message):
    """Обрабатывает текст и кнопки."""
    user_id = message.chat.id
    text = message.text

    # --- Кнопка: Оставить заявку ---
    if text == '📝 Оставить заявку':
        # Начинаем заявку — сразу спрашиваем имя
        user_states[user_id] = {'step': 1}
        bot.send_message(
            user_id,
            '📝 Начинаем заполнение заявки.\n\n'
            'Шаг 1 из 3: Как вас зовут?'
        )
        return

    # --- Кнопка: Отменить ---
    if text == '❌ Отменить':
        cancel_handler(message)
        return

    # --- Если пользователь ещё не начал — показываем кнопки ---
    if user_id not in user_states:
        bot.send_message(
            user_id,
            'Чтобы оставить заявку, нажмите кнопку «📝 Оставить заявку».',
            reply_markup=get_main_keyboard()
        )
        return

    state = user_states[user_id]
    step = state['step']

    # --- ШАГ 1: Имя ---
    if step == 1:
        state['name'] = text
        state['step'] = 2
        bot.send_message(
            user_id,
            f'Отлично, {text}! ✍️\n\n'
            'Шаг 2 из 3: Введите ваш номер телефона.'
        )
        return

    # --- ШАГ 2: Телефон ---
    if step == 2:
        state['phone'] = text
        state['step'] = 3
        bot.send_message(
            user_id,
            'Шаг 3 из 3: Какая услуга вас интересует?'
        )
        return

    # --- ШАГ 3: Услуга ---
    if step == 3:
        state['service'] = text
        state['date'] = datetime.now().strftime('%Y-%m-%d %H:%M')
        state['user_id'] = user_id

        # Сохраняем в Excel
        save_to_excel(state)

        # Отправляем письмо
        email_sent = send_email_notification(state)

        # Формируем сообщение пользователю
        email_status = '✅ Уведомление отправлено!' if email_sent else '⚠️ Уведомление не отправлено (проверьте настройки почты).'

        bot.send_message(
            user_id,
            '✅ Заявка успешно сохранена!\n\n'
            f'📋 Имя: {state["name"]}\n'
            f'📞 Телефон: {state["phone"]}\n'
            f'🔧 Услуга: {state["service"]}\n\n'
            f'{email_status}\n'
            'Спасибо! Мы с вами свяжемся. 🙏\n\n'
            'Нажмите «📝 Оставить заявку», чтобы отправить ещё одну.',
            reply_markup=get_main_keyboard()
        )

        # Удаляем состояние — заявка заполнена
        del user_states[user_id]
        return

# ===================== ЗАПУСК =====================
if __name__ == '__main__':
    create_excel_if_not_exists()
    print('🤖 Бот для сбора заявок запущен!')
    print('📧 Email-уведомления включены.')
    print('Нажми Ctrl+C, чтобы остановить.')
    bot.infinity_polling()
