import os
from imbox import Imbox
import traceback
from openpyxl import load_workbook
import requests
import zipfile
from loguru import logger
from prettytable import PrettyTable

host = "imap.gmail.com"
username = "mailforsber@gmail.com"
password = 'SZF-in6-8hu-M9c'
sender = 'sbbol@sberbank.ru'
download_folder = "data"
maked_folder = 'maked'

logger.add('logs\\logs.txt', format="{time: DD-MM-YY  HH:mm:ss} {level} "
                                    "{module}:{function}:{line} - {message}")


def send_telegram(text: str):
    logger.info('###### Отправляем в телегу ######')
    token = "1907792666:AAHd1zArB4V8tmG8ek8UveFKe12FtjvACeI"
    url = "https://api.telegram.org/bot"
    channel_id = "667290393"
    url += token
    method = url + "/sendMessage"

    r = requests.post(method, data={
        "chat_id": channel_id,
        "text": text,
        "parse_mode": "html"
    })

    if r.status_code != 200:
        logger.error('Ошибка отправки')
        raise Exception("post_text error")


def making_text_for_tg(ls: list):
    # ------ новое форматирование ------------
    mytable = PrettyTable()
    mytable.field_names = mytable.field_names = ["Р/C", "%", "Сверка"]
    it_summ = 0
    it_fee = 0
    date_sravnenie = ls[0]['date']
    text = f'Данные по эквайрингу за {date_sravnenie}:\n'
    for item in ls:
        date_operation = item['date']
        summ = item['summ']
        fee = item['fee']
        if date_operation == date_sravnenie:
            mytable.add_row([summ, fee, round(summ + fee, 2)])
            it_summ += summ
            it_fee += fee
            #text += f'{count}. На Р/С: {summ}, комиссия: {fee}, сверка: {round(summ + fee, 2)} \n'
        else:
            text = text + '<pre>' + mytable.get_string() + '</pre>'
            percent = round(it_fee * 100 / it_summ, 2)
            text += f'\nИтого поступило на Р/С: {round(it_summ, 2)} руб. \n'
            text += f'Итого комиссия: {round(it_fee, 2)} руб. ({percent}%) \n\n'
            mytable.clear_rows()
            it_summ = summ
            it_fee = fee
            date_sravnenie = date_operation
            mytable.add_row([summ, fee, round(summ + fee, 2)])
            text = text + f'Данные по эквайрингу за {date_sravnenie}:\n'
    else:
        text = text + '<pre>' + mytable.get_string() + '</pre> '
        percent = round(it_fee * 100 / it_summ, 2)
        text += f'\nИтого поступило на Р/С: {round(it_summ, 2)} руб. \n'
        text += f'Итого комиссия: {round(it_fee, 2)} руб. ({percent}%) \n'


    text2 = text

    return text2


def conect_read_download():
    logger.info('###### Старт функции connect_read_download ######')
    mail = Imbox(host, username=username, password=password, ssl=True, ssl_context=None, starttls=False)
    # messages = mail.messages() # defaults to inbox
    logger.info('Получаем письма')
    messages = mail.messages(unread=True, sent_from=sender)

    for (uid, message) in messages:
        mail.mark_seen(uid)  # optional, mark message as read
        url = str(message).split('n<a href="')[1].split('" style="text-decoration: none')[0].strip()
        logger.info(f'Переходим по URL: {url}')
        r = requests.get(url)
        with open('file.zip', 'wb') as f:
            f.write(r.content)
        archive = zipfile.ZipFile('file.zip', 'r')
        logger.info('Распаковываем в папку data')
        archive.extractall('data')
        # os.remove('file.zip')
        # for idx, attachment in enumerate(message.attachments):
        #     try:
        #         att_fn = attachment.get('filename')
        #         download_path = f"{download_folder}/{att_fn}"
        #         print(download_path)
        #         with open(download_path, "wb") as fp:
        #             fp.write(attachment.get('content').read())
        #     except:
        #         print(traceback.print_exc())

    mail.logout()


def parsexl_movexl():
    logger.info('###### Начало парсинга файла ######')
    files = os.listdir('data')
    ls = []
    date = '@нет данных@'
    if len(files) != 0:
        wb = load_workbook(filename=f'{download_folder}/{files[0]}')
        logger.info('Загружаем книгу')
        wb.active = 0
        sheet = wb.active
        date = str(sheet['M8'].value).split('счету ')[1].strip()[:-1]
        for row in range(12, 20):
            cell1 = f'N{row}'
            cell2 = f'U{row}'
            if sheet[cell1].value is None:
                break
            if str(sheet[cell2].value).find('Комиссия') == -1:
                continue

            summ = float(sheet[cell1].value)
            fee = float(sheet[cell2].value.split('Комиссия')[1].split('Возврат')[0].strip()[:-1].replace(',', ''))
            date_operation = sheet[cell2].value.split('Дата реестра ')[1].split('. Комиссия')[0].strip()
            ls.append(
                {'summ': summ,
                 'fee': fee,
                 'date': date_operation
                 }
            )
        logger.info('Перемещаем файл')
        os.replace(f'{download_folder}/{files[0]}', f'{maked_folder}/{files[0]}')
    return ls


def main():
    if not os.path.isdir(download_folder):
        os.makedirs(download_folder, exist_ok=True)
    if not os.path.isdir(maked_folder):
        os.makedirs(maked_folder, exist_ok=True)

    conect_read_download()
    # input()
    ls = parsexl_movexl()
    if len(ls) == 0:
        print('Нет данных')
        send_telegram('Не удалось получить данные по расчетному счету')
    else:
        print('OK!!')
        text = making_text_for_tg(ls)
        # print(text)
        send_telegram(text)


if __name__ == "__main__":
    main()
