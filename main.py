import time
from pprint import pprint
import smtp
import imap
import sql

if __name__ == "__main__":
    print("ПРОВЕРКА СБОРКИ")
    try:
        print(">> МОДУЛЬ SQL")
    except:
        print("Что-то не работает...")

