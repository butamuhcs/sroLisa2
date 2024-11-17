import asyncio
import smtplib
import time
import pymysql
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from aiogram import Bot, Dispatcher, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram.filters import Command, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

class OrderFood(StatesGroup):
    login = State()
    password = State()

# Токен вашего бота
BOT_TOKEN = "7819602602:AAHcn3arxa5runjj9nYRT8cC2KeS3-BpHKI"

# Глобальная переменная для блокировки
sending_in_progress = False

# Настройки MySQL
DB_HOST = 'VH290.spaceweb.ru'
DB_USER = 'butamuhcsy_Lisa'
DB_PASSWORD = 'Samreg25'
DB_NAME = 'butamuhcsy_Lisa'

# SMTP-сервер
SMTP_SERVER = 'smtp.mail.ru'

# Инициализация бота и диспетчера
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# Клавиатура главного меню
main_menu = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="Авторизоваться"), KeyboardButton(text="Тестовое письмо")],
        [KeyboardButton(text="Начать рассылку")]
    ],
    resize_keyboard=True
)

# Подключение к базе данных
def get_companies(date: str):
    connection = pymysql.connect(host=DB_HOST, user=DB_USER, password=DB_PASSWORD, database=DB_NAME)
    with connection.cursor() as cursor:
        query = "SELECT email, company_name FROM companies WHERE date = %s"
        cursor.execute(query, (date,))
        result = cursor.fetchall()
    connection.close()
    return result

# Функция для отправки письма
async def send_email(smtp_server, login, password, recipient, subject, body):
    try:
        msg = MIMEMultipart()
        msg['From'] = login
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))

        with smtplib.SMTP(smtp_server, 587) as server:
            server.starttls()
            server.login(login, password)
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"Ошибка при отправке письма: {e}")
        return False

# Обработчик команды /start
@dp.message(Command("start"))
async def start_command(message: types.Message):
    await message.answer("Добро пожаловать! Выберите действие:", reply_markup=main_menu)

# Обработчик авторизации
@dp.message(lambda message: message.text == "Авторизоваться")
async def authorize_user(message: types.Message):
    await message.answer(
        "Введите логин и пароль от почты через пробел (например: myemail@mail.ru mypassword)."
    )

@dp.message(StateFilter(None),lambda message: "@" in message.text and " " in message.text)
async def handle_login(message: types.Message, state: FSMContext):
    login, password = message.text.split(maxsplit=1)
    await state.update_data(login=login)
    await state.update_data(password=password)
    #message.from_user.login = login
    #message.from_user.password = password
    await message.answer("Авторизация прошла успешно!")

# Обработчик тестового письма
@dp.message(lambda message: message.text == "Тестовое письмо")
async def test_email(message: types.Message, state: FSMContext):
    user_data = await state.get_data()
    if not user_data['login'] or not user_data['password']:
        await message.answer("Сначала выполните авторизацию, используя 'Авторизоваться'.")
        return

    login = user_data['login']
    password = user_data['password']
    recipient = user_data['login']
    subject = "Тестовое письмо"
    body = "<html><body><h1>Ваше тестовое письмо</h1><p>Проверьте, всё ли отображается корректно.</p></body></html>"

    await message.answer("Отправляем тестовое письмо...")
    if await send_email(SMTP_SERVER, login, password, recipient, subject, body):
        await message.answer("Тестовое письмо отправлено! Проверьте ваш почтовый ящик.")
    else:
        await message.answer("Ошибка при отправке тестового письма. Проверьте логин и пароль.")

# Обработчик рассылки
@dp.message(lambda message: message.text == "Начать рассылку")
async def start_sending(message: types.Message, state: FSMContext):
    global sending_in_progress
    if sending_in_progress:
        await message.answer("Идёт отправка писем. Пожалуйста, дождитесь завершения.")
        return

    user_data = await state.get_data()
    if not user_data['login'] or not user_data['password']:
        await message.answer("Сначала выполните авторизацию, используя 'Авторизоваться'.")
        return

    await message.answer("Введите дату компаний для отправки писем (в формате YYYY-MM-DD).")

@dp.message(lambda message: len(message.text) == 10 and "-" in message.text)
async def handle_date(message: types.Message, state: FSMContext):
    global sending_in_progress
    sending_in_progress = True
    date = message.text.strip()
    companies = get_companies(date)

    if not companies:
        await message.answer(f"Нет компаний с датой {date}.")
        sending_in_progress = False
        return

    total_emails = len(companies)
    estimated_time = (10 + 15) / 2 * total_emails / 60
    await message.answer(
        f"Начата отправка {total_emails} писем. Примерное время отправки: {estimated_time:.1f} минут."
    )

    successful = 0
    skipped = 0
    login = message.from_user.login
    password = message.from_user.password

    for company in companies:
        email, company_name = company
        if not email:
            skipped += 1
            continue

        subject = "Ваше коммерческое предложение"
        body = "<html><body><h1>Коммерческое предложение</h1><p>Текст вашего письма.</p></body></html>"

        if await send_email(SMTP_SERVER, login, password, email, subject, body):
            successful += 1
        await asyncio.sleep(10)  # Задержка в 10 секунд

    await message.answer(
        f"Успешно отправлено {successful} писем. Пропущено {skipped} компаний (нет e-mail)."
    )
    sending_in_progress = False

# Запуск бота
async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())