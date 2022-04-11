import smtp
import imap
import sql
import time

if __name__ == "__main__":
    while True:
        some_email = "".join(imap.get_headers_email(1, "E_ONLY"))
        print("Полученный адрес: ",some_email)
        if sql.get_client_data(some_email):
            smtp.send_message_clients(some_email)
        time.sleep(600)
