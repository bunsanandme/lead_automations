import email
import imaplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from smtplib import *

# Важные константы для проверки и отправки писем
# Есть возможность для привязки разных почт (допустим, менеджер делает рассылку)

USERNAME = "#####@mail.com"
PASSWORD = "#######"

SEND_USERNAME = "#####@mail.com"
SEND_PASSWORD = "#######"


def get_headers_email(email_numbers):
    """
    Функция при помощи IMAP-сервера подключается к почтовому ящику
    Проверяет заголовки и даты их выводит
    """

    mail = imaplib.IMAP4_SSL('imap.yandex.ru')
    mail.login(USERNAME, PASSWORD)
    mail.list()
    mail.select("inbox")
    result, data = mail.search(None, "ALL")
    ids = data[0]
    id_list = ids.split()
    emails_list = id_list[len(id_list) - email_numbers:]

    headers_list = []
    for item in range(len(emails_list)):
        result, data = mail.fetch(emails_list[item], "(RFC822)")
        raw_email = data[0][1]
        raw_email_string = raw_email.decode('cp1251')
        email_message = email.message_from_string(raw_email_string)

        temp_list = {"Email": email.utils.parseaddr(email_message['From'])[1], "Date": email_message["Date"]}
        headers_list.append(temp_list)

    return headers_list


def get_email_addresses(emails_numbers):
    """
    Функция для вывода только имейлов, получаемых для сообщений
    TODO: Сделать в виде декоратора
    """

    raw_data = get_headers_email(emails_numbers)
    email_list = []
    for item in raw_data:
        email_list.append(item["Email"])

    return list(set(email_list))


def send_message_M(client_email):
    """
    Функция для отправки сообщений через SMTP-клиент лидам категории M
    Сообщение составляется из адресата, адресанта, темы и самого тела, а также вложений
    TODO: Сделать общую функция для всех клиентов и передавать флаг
    """

    server = SMTP("smtp.gmail.com", 587)
    server.ehlo()
    server.starttls()

    msg = MIMEMultipart()
    msg['Subject'] = 'Документы на авторизацию'
    msg['From'] = SEND_USERNAME
    msg['To'] = client_email

    attach = MIMEApplication(open("ФИО.docx", 'rb').read())
    attach.add_header('Content-Disposition', 'attachment', filename='doc.docx')
    msg.attach(attach)
    server.login(SEND_USERNAME, SEND_PASSWORD)

    server.sendmail(SEND_USERNAME, client_email, msg.as_string())
    server.quit()
