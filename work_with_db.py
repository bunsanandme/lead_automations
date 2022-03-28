import sqlite3

DATABASE_FILENAME = "T:\\my\\WORK\\automation\\firm.db"

def unpack_tuple(query_tuple):
    for item in query_tuple:
        print(item, end=", ")

try:
    sqlite_connection = sqlite3.connect(DATABASE_FILENAME)
    cursor = sqlite_connection.cursor()
    print("База данных подключена к SQLite\n")

    cursor.execute("SELECT * FROM Meetings")
    rows = cursor.fetchall()
    for row in rows:
        unpack_tuple(row)
        print("\n")

    cursor.close()

except sqlite3.Error as error:
    print("\nОшибка при подключении к sqlite", error)
finally:
    if sqlite_connection:
        sqlite_connection.close()
        print("\nСоединение с SQLite закрыто")
