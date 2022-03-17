import pandas as pd
from email_reader import *
from pprint import pprint

ABSOLUTE_PATH_CSV = "clients.csv"
EMAIL_COLUMN = "E-mail"
LEAD_TYPE_COLUMN = "Потенциал"
NAME_COLUMN = "Контактное лицо"


def get_client_data(client_email):
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
    for item in client_emails:
        if get_client_data(item) != (None, None, None):
            name, lead, email = get_client_data(item)
            if lead == "M":
                send_message_M(email)
                print("Success!")


if __name__ == "__main__":
    check_client_messages(get_email_addresses(5))
