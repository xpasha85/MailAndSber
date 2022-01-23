import os
from imbox import Imbox
import traceback
from openpyxl import load_workbook
import requests

host = "imap.gmail.com"
username = "mailforsber@gmail.com"
password = 'SZF-in6-8hu-M9c'
download_folder = "data"
maked_folder = 'maked'


def send_telegram(text: str):
    token = "1907792666:AAHd1zArB4V8tmG8ek8UveFKe12FtjvACeI"
    url = "https://api.telegram.org/bot"
    channel_id = "667290393"
    url += token
    method = url + "/sendMessage"

    r = requests.post(method, data={
        "chat_id": channel_id,
        "text": text
    })

    if r.status_code != 200:
        raise Exception("post_text error")


def making_text_for_tg(ls: list, date_statement='1 января 1970'):
    text = f'Данные по эквайрингу за {date_statement}:\n'
    count = 1
    it_summ = 0
    it_fee = 0
    for item in ls:
        summ = item['summ']
        fee = item['fee']
        it_summ += summ
        it_fee += fee
        text += f'{count}. На Р/С: {summ}, комиссия: {fee}, сверка: {round(summ + fee, 2)} \n'
        count += 1

    text += f'\nИтого поступило на Р/С: {round(it_summ, 2)} руб. \n'
    text += f'Итого комиссия: {round(it_fee, 2)} руб.'
    return text


if not os.path.isdir(download_folder):
    os.makedirs(download_folder, exist_ok=True)
if not os.path.isdir(download_folder):
    os.makedirs(maked_folder, exist_ok=True)


def conect_read_download():
    mail = Imbox(host, username=username, password=password, ssl=True, ssl_context=None, starttls=False)
    # messages = mail.messages() # defaults to inbox
    messages = mail.messages(unread=True, sent_from='xpasha85@gmail.com')

    for (uid, message) in messages:
        mail.mark_seen(uid)  # optional, mark message as read

        for idx, attachment in enumerate(message.attachments):
            try:
                att_fn = attachment.get('filename')
                download_path = f"{download_folder}/{att_fn}"
                print(download_path)
                with open(download_path, "wb") as fp:
                    fp.write(attachment.get('content').read())
            except:
                print(traceback.print_exc())

    mail.logout()


def parsexl_movexl():
    files = os.listdir('data')
    ls = []
    date = '@нет данных@'
    if len(files) != 0:
        wb = load_workbook(filename=f'{download_folder}/{files[0]}')
        wb.active = 0
        sheet = wb.active
        date = str(sheet['M8'].value).split('счету ')[1].strip()[:-1]
        for row in range(12, 20):
            cell1 = f'N{row}'
            cell2 = f'U{row}'
            if sheet[cell1].value is None:
                break

            summ = float(sheet[cell1].value)
            fee = float(sheet[cell2].value.split('Комиссия')[1].split('Возврат')[0].strip()[:-1].replace(',', ''))
            ls.append(
                {'summ': summ,
                 'fee': fee
                 }
            )
        os.replace(f'{download_folder}/{files[0]}', f'{maked_folder}/{files[0]}')
    return ls, date


def main():
    conect_read_download()
    # input()
    ls, date = parsexl_movexl()
    if len(ls) == 0:
        print('Нет данных')
    else:
        print('OK')
        text = making_text_for_tg(ls, date_statement=date)
        # print(text)
        send_telegram(text)


if __name__ == "__main__":
    main()