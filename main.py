import os
import smtplib

import paths
from time import sleep
import xlsxwriter

from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

url = 'https://www.moex.com/'
driver = webdriver.Chrome()
driver.get(url)

# путешествуем до индикативных курсов
# клик на меню
driver.find_element(By.CLASS_NAME, paths.menu_1).click()
sleep(6)

driver.find_element(By.LINK_TEXT, 'Срочный рынок').click()
sleep(6)

if driver.find_element(By.LINK_TEXT, "Согласен"):
    driver.find_element(By.LINK_TEXT, 'Согласен').click()
sleep(6)

# клик на меню
driver.find_element(By.CLASS_NAME, paths.menu_2).click()
sleep(6)

driver.find_element(By.LINK_TEXT, 'Индикативные курсы').click()
sleep(6)


def fetch_data(pair):
    """Получаем пару валюты для установки этой пары и даты"""
    driver.find_element(By.CLASS_NAME, 'ui-select__activator').click()
    driver.find_element(By.XPATH, f'//a[contains(text(), "{pair}")]').click()

    # клик по календарю для начального периода
    driver.find_element(By.XPATH, paths.start_period).click()
    sleep(6)

    # клик по месяцу начального периода
    driver.find_element(By.XPATH, paths.start_months).click()
    sleep(6)

    # смена на нужный месяц
    driver.find_element(By.XPATH, paths.change_to_start_month).click()
    sleep(6)

    # выбор нужной даты
    driver.find_element(By.XPATH, paths.change_to_start_date).click()
    sleep(6)

    # выбор календаря конца периода
    driver.find_element(By.XPATH, paths.end_period).click()
    sleep(6)

    # клик по месяцу конечного периода
    driver.find_element(By.XPATH, paths.end_months).click()
    sleep(6)

    # выбор нужного месяца конечного периода
    driver.find_element(By.XPATH, paths.change_to_end_month).click()
    sleep(6)

    # выбор даты конечного периода
    driver.find_element(By.XPATH, paths.change_to_end_date).click()

    # нажатие на кнопку для применения выбранного календарного периода
    driver.find_element(By.XPATH, paths.apply_period).click()
    sleep(6)

    # получаем код страницы
    html = driver.page_source

    return html


def find_table(pair):
    """Парсинг данных из таблицы"""

    soup = BeautifulSoup(fetch_data(pair), 'html.parser')
    table_div = soup.find('div', class_='ui-table__container')
    table = table_div.find('table')
    tbody = table.find('tbody')
    rows = tbody.find_all('tr', class_='ui-table-row -interactive')

    return rows


def create_list_of_data(pair):
    """создание списка списков с нужными данными (дата, значение, время)"""

    global headers_width
    data = [[f'Дата {pair}', f'Курс {pair}', f'Время {pair}']]
    headers_width.extend(data[0])

    test = find_table(pair)
    # Перебираем строки спаршенных данных из таблицы
    for row in test:
        cells = row.find_all('td', class_='ui-table-cell')
        data.append([cells[0].text, cells[3].text, cells[4].text])

    return data


def data_to_excel(pair):
    """внесение запаршенных данных в эксель-таблицу"""

    global counter_columns
    global counter_rows

    data_list = create_list_of_data(pair)
    financial_format = workbook.add_format({'num_format': '₽#,##0.00'})

    for row_num, row_data in enumerate(data_list):
        for col_num, cell_data in enumerate(row_data):
            if cell_data.count('.') == 1:
                worksheet.write(row_num, col_num + counter_columns, float(cell_data))
            else:
                worksheet.write(row_num, col_num + counter_columns, cell_data)

        res = f'={"B" + str(2 + row_num)}/{"E" + str(2 + row_num)}'

        if row_num == len(data_list) - 1:
            continue

        worksheet.write(f'{'G' + str(row_num + 2)}', res, financial_format)

    counter_columns += len(data_list[0])
    if row_num > counter_rows:
        counter_rows = row_num + 1


def send_email(file_name):
    sender = 'samofalov.zhenya@mail.ru'
    password = os.getenv('MY_PASSWORD')

    # Настройки для отправки письма
    smtp_server = 'smtp.mail.ru'
    smtp_port = 587
    username = sender
    password = password

    # Получатели и тема письма
    from_email = sender
    to_email = sender
    subject = 'Support RPA'

    # Создание письма
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # Текст письма
    body = line_declination(counter_rows)
    msg.attach(MIMEText(body, 'plain'))

    # Прикрепление Excel файла к письму
    filename = file_name
    attachment = open(filename, 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={filename}')
    msg.attach(part)

    # Отправка письма
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(username, password)
    text = msg.as_string()
    server.sendmail(from_email, to_email, text)
    server.quit()


def line_declination(num_lines):
    """склонение строк"""
    if num_lines == 1:
        return "1 строка"
    elif num_lines % 10 == 1 and num_lines % 100 != 11:
        return f"{num_lines} строка"
    elif 2 <= num_lines % 10 <= 4 and (num_lines % 100 < 10 or num_lines % 100 >= 20):
        return f"{num_lines} строки"
    else:
        return f"{num_lines} строк"


def main(pair):
    data_to_excel(pair)


# счетчик количества занятых столбцов и строк
counter_columns = 0
counter_rows = 0

# Сохранение названий столбцов для изменения ширины ячеек
headers_width = []

# Инициализация файла excel
file_name = 'Test.xlsx'
workbook = xlsxwriter.Workbook(file_name)
worksheet = workbook.add_worksheet()


def width_cells():
    # задание ячейкам ширины
    for col_num, header in enumerate(headers_width):
        column_width = len(header) + 2
        worksheet.set_column(col_num, col_num, column_width)


# Проверка на числовой формат
worksheet.write('J2', "Проверка")
worksheet.write("J3", '=SUM(G:G)')

main('USD/RUB')
main('JPY/RUB')

# Добавление столбца "результат"
worksheet.write('G1', 'Результат')
headers_width.append("Результат")

width_cells()

# отправка на почту
send_email(file_name)

driver.quit()
workbook.close()
