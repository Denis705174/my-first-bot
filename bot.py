# Импортируем библиотеку для работы с Telegram
import telebot

# Сюда нужно вставить токен, который выдаст BotFather (см. шаг 1)
TOKEN = 'ВАШ_ТОКЕН'

# Создаём объект бота
bot = telebot.TeleBot(TOKEN)


# Обработчик команды /start — срабатывает, когда пользователь запускает бота
@bot.message_handler(commands=['start'])
def send_welcome(message):
    # message.chat.id — это идентификатор чата, куда отвечать
    # message.from_user.first_name — имя пользователя
    bot.send_message(
        message.chat.id,
        f'Привет, {message.from_user.first_name}! 👋\n'
        'Я твой первый Telegram-бот. Рад тебя видеть!'
    )


# Обработчик команды /help — подсказка по командам
@bot.message_handler(commands=['help'])
def send_help(message):
    bot.send_message(
        message.chat.id,
        'ℹ️ Справка по командам:\n'
        '/start — приветствие\n'
        '/help — показать эту справку\n'
        'Можешь просто написать мне любое сообщение!'
    )


# Обработчик любых остальных сообщений (на случай если пользователь
# написал что-то, не используя команду)
@bot.message_handler(func=lambda message: True)
def echo_all(message):
    bot.send_message(
        message.chat.id,
        f'Ты написал: "{message.text}"\n'
        'Используй /help, чтобы узнать, что я умею.'
    )


# Запускаем бота — он будет постоянно слушать новые сообщения
if __name__ == '__main__':
    print('Бот запущен! Нажми Ctrl+C, чтобы остановить.')
    bot.infinity_polling()
