import sqlite3
import yaml
import os

with open("config.yaml") as file:
    config_list = list(yaml.load(file, Loader=yaml.FullLoader).values())

for config_item in config_list[2:]:
    assert config_item, "Значение не может быть пустым"

LOG_SYMBOL, DEBUG, OUR_USERNAME, OUR_PASSWORD, SEND_USERNAME, SEND_PASSWORD, DATABASE_PATH = config_list[:7]
if not os.path.exists(DATABASE_PATH):
    print("БД не существует, будет создана новая БД")


def query_executor_select(query):
    sqlite_connection = None
    try:
        sqlite_connection = sqlite3.connect(DATABASE_PATH)
        cursor = sqlite_connection.cursor()
        cursor.execute(query)
        rows = cursor.fetchall()
        cursor.close()
        return rows
    except sqlite3.Error as error:
        if DEBUG:
            print(LOG_SYMBOL + "Ошибка при подключении к sqlite")
            print(LOG_SYMBOL + "Ошибка:", error)
        return None
    finally:
        if sqlite_connection:
            sqlite_connection.close()


def query_executor_insert(query):
    sqlite_connection = sqlite3.connect(DATABASE_PATH)
    cursor = sqlite_connection.cursor()
    cursor.execute(query)
    sqlite_connection.commit()
    return cursor.lastrowid


def get_client_data(client_email):
    sqlite_connection = None
    try:

        sqlite_connection = sqlite3.connect(DATABASE_PATH)
        cursor = sqlite_connection.cursor()
        if DEBUG:
            print(LOG_SYMBOL + "Подключение прошло успешно!")
        cursor.execute("SELECT * FROM Clients WHERE Email = \"{}\"".format(client_email))
        if DEBUG:
            print(LOG_SYMBOL + "Выполнен запрос: успешно!")
        try:
            rows = cursor.fetchall()[0]
        except IndexError:
            if DEBUG:
                print(LOG_SYMBOL + "ВНИМАНИЕ: Запрос пустой!")
            return None
        cursor.close()
        return rows[3], rows[7]

    except sqlite3.Error as error:
        if DEBUG:
            print(LOG_SYMBOL + "Ошибка при подключении к sqlite")
            print(LOG_SYMBOL + "Ошибка:", error)
    finally:
        if sqlite_connection:
            sqlite_connection.close()
            if DEBUG:
                print(LOG_SYMBOL + "Отключение от БД...")
