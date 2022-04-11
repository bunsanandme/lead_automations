import smtp
import imap
import sql
import time

if __name__ == "__main__":
        for i in sql.query_executor_select("SELECT Email FROM Clients"):
                smtp.send_message_meetings(i[0])

        while True:
                some_email = "".join(imap.get_headers_email(1, "E_ONLY"))
                print("Полученный адрес: ",some_email)
                if sql.get_client_data(some_email):
                    smtp.send_message_clients(some_email)
                print("Ожидание: 10 минут...zzz")
                time.sleep(600)
