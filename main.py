import asyncio
import os

import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from aiogram import Bot, Dispatcher, types
from aiogram.types import InputFile
from aiogram.utils.executor import start_polling

API_TOKEN = os.getenv('BOT_TOKEN')  # Замените на ваш токен бота
LOGIN = os.getenv('LOGIN')
PASSWORD = os.getenv('PASSWORD')
print(LOGIN)
print(PASSWORD)
# Инициализация бота и диспетчера
bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)

# Хранение последнего сообщения
last_message_id = None


# Функция для авторизации и получения куки
def get_cookies_and_token():
    session = requests.Session()
    url = 'https://office.vilavi.com/'
    session.get(url)

    login_page = session.get('https://office.vilavi.com/Account/Login?ReturnUrl=%2F')
    soup = BeautifulSoup(login_page.content, 'html.parser')
    token = soup.find('input', {'name': '__RequestVerificationToken'})['value']

    login_data = {
        'Login': LOGIN,
        'Password': PASSWORD,
        '__RequestVerificationToken': token,
        'IsClient': 'false'
    }

    session.post('https://office.vilavi.com/Account/Login?returnurl=%2F', data=login_data)

    cookies = session.cookies.get_dict()
    return cookies, session


# Функция для получения и сохранения отчета
def fetch_and_save_report(cookies, session):
    url = 'https://store.vilavi.com/ConsignmentStockBalance'
    headers = {
        "Host": "store.vilavi.com"
    }
    cookies['Region'] = 'ru'
    response = session.get(url, headers=headers, cookies=cookies)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        table = soup.find('table', {'id': 'balanceTable'})

        headers = [header.text.strip() for header in table.find_all('th')]
        headers = [headers[1], headers[2], headers[4], headers[5]]  # Убираем столбец "Всего (PV)"

        rows = table.find_all('tr')
        data = []
        for row in rows[1:]:
            cols = row.find_all('td')
            if len(cols) >= 6:  # Проверяем, что есть как минимум 6 столбцов
                cols = [ele.text.strip() for ele in cols]
                cols = [cols[1], cols[2], cols[4], cols[5]]  # Убираем столбец "Всего (PV)"
                row_data = [cols[0]] + list(map(int, cols[1:]))
                data.append(row_data)

        # Преобразуем данные в DataFrame и сортируем по столбцу "Забронировано"
        df = pd.DataFrame(data, columns=headers)
        df = df.sort_values(by='Забронировано', ascending=False)

        output_file_path = 'ConsignmentStockBalance.xlsx'
        df.to_excel(output_file_path, index=False)

        workbook = openpyxl.load_workbook(output_file_path)
        worksheet = workbook.active

        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value)) / 2
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column].width = adjusted_width

        workbook.save(output_file_path)
        return output_file_path
    else:
        raise Exception(f"Ошибка при выполнении запроса: {response.status_code}")


# Динамическое обновление сообщения
async def update_message(message: types.Message, status: str, icons: list, delay: float):
    i = 0
    while True:
        await message.edit_text(f"{status} {icons[i % len(icons)]}")
        await asyncio.sleep(delay)
        i += 1


# Переменная для хранения ID последнего отправленного сообщения с отчетом
last_report_message_id = None


# Хэндлер на команду /stocks
@dp.message_handler(commands=['stocks'])
async def send_report(message: types.Message):
    global last_report_message_id

    status_message = await message.reply("Проверяю наличие 🕐")

    try:
        # Динамическое обновление сообщения
        icons_checking = ["🕐", "🕒", "🕕", "🕘", "🕛"]
        checking_task = asyncio.create_task(update_message(status_message, "Проверяю наличие", icons_checking, 2))

        cookies, session = await asyncio.to_thread(get_cookies_and_token)
        file_path = await asyncio.to_thread(fetch_and_save_report, cookies, session)
        checking_task.cancel()

        await status_message.delete()
        if last_report_message_id:
            try:
                await bot.delete_message(message.chat.id, last_report_message_id)
            except:
                pass

        sent_message = await bot.send_document(message.chat.id, InputFile(file_path), caption="Отчет о наличии")
        last_report_message_id = sent_message.message_id

    except Exception as e:
        await status_message.edit_text(f"Произошла ошибка: {e}")


# Запуск бота
if __name__ == '__main__':
    start_polling(dp, skip_updates=True)
