import asyncio
import smtplib
import re
import time
from premailer import transform
import pymysql
import openpyxl
import dns.resolver
from openpyxl.styles import Font, Alignment
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from aiogram import Bot, Dispatcher, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram.filters import Command, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import FSInputFile

class OrderFood(StatesGroup):
    login = State()
    password = State()

class BotStates(StatesGroup):
    choose_date_for_sending = State()
    choose_date_for_excel = State()

# Чтение HTML-шаблона из файла
with open("kpLiza.html", "r", encoding="utf-8") as file:
    html_content = file.read()

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
        [KeyboardButton(text="Начать рассылку")],
        [KeyboardButton(text="Сформировать Excel")]
    ],
    resize_keyboard=True
)

def is_valid_email(email):
    # Проверка синтаксиса email
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z]{2,}$'
    if not re.match(email_regex, email):
        return False

    # Извлечение доменной части
    domain = email.split('@')[-1]
    try:
        # Проверка MX-записей для домена
        dns.resolver.resolve(domain, 'MX')
        return True
    except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN):
        return False

# Подключение к базе данных
def get_companies(date: str):
    connection = pymysql.connect(host=DB_HOST, user=DB_USER, password=DB_PASSWORD, database=DB_NAME)
    with connection.cursor() as cursor:
        query = "SELECT email FROM companies WHERE reg_date = %s"
        cursor.execute(query, (date,))
        result = cursor.fetchall()
    connection.close()
    return result

# Функция для отправки письма
async def send_email(smtp_server, login, password, recipient, subject):
    try:
        if is_valid_email(recipient):
            msg = MIMEMultipart()
            msg['From'] = login
            msg['To'] = recipient
            msg['Subject'] = subject
            msg.attach(MIMEText(html_content, 'html'))

            with smtplib.SMTP(smtp_server, 587) as server:
                server.starttls()
                server.login(login, password)
                server.send_message(msg)
            return True
    except smtplib.SMTPResponseException as e:
        print(f"Ошибка при отправке письма: {e.smtp_code}, {e.smtp_error}")
        if e.smtp_code == 451:
            print("Превышен лимит отправки писем. Пауза на 31 минуту.")
            await asyncio.sleep(31 * 60)  # Ожидание 31 минуты
        return False
    except Exception as e:
        print(f"Ошибка при отправке письма: {e}")
        return False

# Обработчик формирования Excel-файла
@dp.message(lambda message: message.text == "Сформировать Excel")
async def generate_excel_file(message: types.Message, state: FSMContext):
    await message.answer("Введите дату компаний для формирования Excel-файла (в формате YYYY-MM-DD).")
    await state.set_state(BotStates.choose_date_for_excel)

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

    await message.answer("Авторизация прошла успешно!")

# Обработчик тестового письма
@dp.message(lambda message: message.text == "Тестовое письмо")
async def test_email(message: types.Message, state: FSMContext):
    user_data = await state.get_data()
    if not user_data:
        await message.answer("Сначала выполните авторизацию, используя 'Авторизоваться'.")
        return

    login = user_data['login']
    password = user_data['password']
    recipient = user_data['login']
    subject = "Вступление в самую надёжную СРО! Премия от 10000 рублей"
    body = html_content

    await message.answer("Отправляем тестовое письмо...")
    if await send_email(SMTP_SERVER, login, password, recipient, subject):
        await message.answer("Тестовое письмо отправлено! Проверьте ваш почтовый ящик.")
    else:
        await message.answer("Ошибка при отправке тестового письма. Проверьте логин и пароль.")

# Обработчик начала рассылки
@dp.message(lambda message: message.text == "Начать рассылку")
async def start_sending(message: types.Message, state: FSMContext):
    global sending_in_progress
    if sending_in_progress:
        await message.answer("Идёт отправка писем. Пожалуйста, дождитесь завершения.")
        return

    user_data = await state.get_data()
    if not user_data:
        await message.answer("Сначала выполните авторизацию, используя 'Авторизоваться'.")
        return

    await message.answer("Введите дату компаний для отправки писем (в формате YYYY-MM-DD).")
    await state.set_state(BotStates.choose_date_for_sending)

@dp.message(StateFilter(BotStates.choose_date_for_sending))
async def handle_date(message: types.Message, state: FSMContext):
    global sending_in_progress
    sending_in_progress = True
    date = message.text.strip()
    companies = get_companies_with_email(date)

    if not companies:
        await message.answer(f"Нет компаний с датой {date}.")
        sending_in_progress = False
        await state.clear()
        return

    total_emails = len(companies)
    delay = 0.5  # Время задержки между письмами в минутах
    emails_per_group = 30
    pause_time = 31  # Пауза в минутах

    groups = (total_emails + emails_per_group - 1) // emails_per_group
    group_time = emails_per_group * delay
    estimated_time = group_time * groups + pause_time * (groups - 1)

    await message.answer(
        f"Начата отправка {total_emails} писем. Примерное время отправки: {estimated_time:.1f} минут."
    )

    successful = 0
    skipped = 0

    user_data = await state.get_data()
    login = user_data['login']
    password = user_data['password']

    for company in companies:
        print(company[1])
        email = company[1]
        if not email or not is_valid_email(email):
            skipped += 1
            print(f"Пропускаем некорректный email: {email}")
            continue

        subject = "Вступление в самую надёжную СРО! Премия от 10000 рублей"

        if await send_email(SMTP_SERVER, login, password, email, subject):
            successful += 1
            await asyncio.sleep(30)
        else:
            skipped += 1
        print(f'Отправлено {successful} писем. Пропущено {skipped}. Всего {len(companies)}')


    await message.answer(
        f"Успешно отправлено {successful} писем. Пропущено {skipped} компаний (нет e-mail)."
    )
    sending_in_progress = False

# Обработчик формирования Excel-файла
@dp.message(lambda message: message.text == "Сформировать Excel")
async def generate_excel_file(message: types.Message, state: FSMContext):
    await message.answer("Введите дату компаний для формирования Excel-файла (в формате YYYY-MM-DD).")
    await state.set_state(BotStates.choose_date_for_excel)

# Подключение к базе данных с фильтром по email
def get_companies_with_email(date: str):
    connection = pymysql.connect(host=DB_HOST, user=DB_USER, password=DB_PASSWORD, database=DB_NAME)
    with connection.cursor() as cursor:
        query = """
            SELECT name, email, reg_date 
            FROM companies 
            WHERE reg_date = %s AND email IS NOT NULL AND email != ''
        """
        cursor.execute(query, (date,))
        result = cursor.fetchall()
    connection.close()
    return result


@dp.message(lambda message: message.text == "Сформировать Excel")
async def generate_excel(message: types.Message):
    await message.answer("Введите дату для формирования файла (в формате YYYY-MM-DD).")


@dp.message(StateFilter(BotStates.choose_date_for_excel))
async def handle_excel_date(message: types.Message):
    date = message.text.strip()
    companies = get_companies_with_email(date)

    if not companies:
        await message.answer(f"Нет компаний с датой {date}.")
        return

    file_name = f"companies_with_email_{date}.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = f"Companies {date}"

    # Заголовки
    sheet.append(["Название", "Email"])
    for company in companies:
        print(company)
        email = company[1]  # Индекс 0 для email
        name = company[0]  # Индекс 1 для названия (замените по вашей БД)

        if email and is_valid_email(email):
            sheet.append([name, email])

    workbook.save(file_name)
    workbook.close()

    # Отправляем файл
    await message.answer_document(document=FSInputFile(file_name), caption="Вот ваш файл с компаниями.")


# Запуск бота
async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())