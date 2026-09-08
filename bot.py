import os
import threading
import telebot
import pandas as pd
from flask import Flask

# Замени на свой реальный токен бота от BotFather
TOKEN = "8965573915:AAGCvsCZpYnqE50wr05lzxxYFh8AQVYpfBQ"
bot = telebot.TeleBot(TOKEN)

app = Flask(__name__)


@app.route("/")
def home():
  return "Bot is alive and running!"


# Обработка команды /start
@bot.message_handler(commands=["start"])
def send_welcome(message):
  bot.reply_to(
      message,
      "Привет! Я рабочий бот на pyTelegramBotAPI и Flask. Всё настроено!",
  )


# Обработка документов (для работы с Excel через pandas и openpyxl)
@bot.message_handler(content_types=["document"])
def handle_docs(message):
  try:
    # Пример логики сохранения и чтения файла
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    local_filename = message.document.file_name
    with open(local_filename, "wb") as new_file:
      new_file.write(downloaded_file)

    # Если это Excel файл, проверяем чтение через pandas
    if local_filename.endswith((".xlsx", ".xls")):
      df = pd.read_excel(local_filename)
      bot.reply_to(
          message,
          f"Файл успешно принят и прочитан! Количество строк: {len(df)}",
      )
    else:
      bot.reply_to(
          message, "Файл принят, но это не Excel-формат (.xlsx / .xls)."
      )

  except Exception as e:
    bot.reply_to(message, f"Произошла ошибка при обработке файла: {e}")


if __name__ == "__main__":
  # Запуск бота в фоновом потоке, чтобы он не блокировал основной процесс Flask
  bot_thread = threading.Thread(
      target=bot.infinity_polling, kwargs={"skip_pending": True}
  )
  bot_thread.daemon = True
  bot_thread.start()

  # Запуск Flask-сервера на порту, который требует Render (переменная окружения PORT)
  port = int(os.environ.get("PORT", 10000))
  app.run(host="0.0.0.0", port=port)
