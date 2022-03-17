import pandas as pd
from email_reader import *

ABSOLUTE_PATH_CSV = "clients.csv"
EMAIL_COLUMN = "E-mail"
LEAD_TYPE_COLUMN = "Потенциал"
NAME_COLUMN = "Контактное лицо"


def get_client_data(client_email):

    """
    Функция по email ищет в csv файле нужные данные для дальнейшей обработки

    Выводит имя, статус лида и его имейл
    В случае ненахождения возвращает пустые переменные
    """
    data = pd.read_csv(ABSOLUTE_PATH_CSV,
                       delimiter=';',
                       encoding='cp1251',
                       skiprows=1,
                       index_col=5)
    try:
        client_data = dict(data.loc[client_email])
    except KeyError:
        return None, None, None
    return client_data[NAME_COLUMN], client_data[LEAD_TYPE_COLUMN], client_email


def check_client_messages(client_emails):

    """
    Функция является основной  этом скрипте, она проверяет входящие письма,
    Если есть письмо от лида и он категории M, отправляет письмо по форме
    """

    for item in client_emails:
        if get_client_data(item) != (None, None, None):
            name, lead, email = get_client_data(item)
            if lead == "M":
                send_message_M(email)
                print("Success!")


if __name__ == "__main__":
    check_client_messages(get_email_addresses(5))
