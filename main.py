import email
import imaplib
import os
import smtplib
import subprocess
from email.header import decode_header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from retrying import retry

import config

mail = imaplib.IMAP4_SSL(config.imap_host)
mail.login(config.email, config.password)


save_dir = os.getcwd()


def send_book(sender_email, book_path, book_name):
    acct = smtplib.SMTP_SSL(config.smtp_host)
    acct.login(config.email, config.password)

    msg = MIMEMultipart()
    msg["Subject"] = "Converted " + book_name
    msg["From"] = config.email
    msg["To"] = config.kindle_email
    with (open(book_path, "rb")) as fp:
        attachment = MIMEApplication(fp.read())
        attachment.add_header(
            "Content-Disposition", "attachment", filename=book_name + ".mobi"
        )
    msg.attach(attachment)

    try:
        acct.sendmail(config.email, config.kindle_email, msg.as_string())
        print("Sent", book_path)

    except smtplib.SMTPSenderRefused as e:
        print("Could not send book")
        msg = MIMEText(e.args[1].decode("utf-8"))
        msg["Subject"] = "Conversion ERROR"
        msg["To"] = sender_email
        acct.sendmail(config.email, sender_email, msg.as_string())


@retry
def main_loop():
    while True:
        mail.list()
        mail.select("inbox")
        retcode, messages = mail.search(None, "(UNSEEN)")

        if retcode == "OK":
            for num in messages[0].split():
                typ, data = mail.fetch(num, "(RFC822)")

                for response_part in data:
                    if isinstance(response_part, tuple):
                        m = email.message_from_bytes(response_part[1])
                        print(m["From"])
                        mail.store(num, "+FLAGS", "\\Seen")

                        for part in m.walk():

                            if part.get_content_maintype() == "multipart":
                                continue

                            if part.get_content_maintype() == "text":
                                continue

                            if part.get("Content-Disposition") is None:
                                continue

                            file_name = decode_header(part.get_filename())[0][0]

                            if not isinstance(file_name, str):
                                file_name = file_name.decode("utf-8")
                            # file_name=part.get_filename()

                            if file_name is not None:
                                print("CONVERSION STARTED", file_name)
                                save_path = os.path.join(save_dir, file_name)
                                book_name = os.path.splitext(os.path.basename(save_path))[0]
                                conv_path = os.path.join(save_dir, book_name + ".mobi")
                                sender = m["From"]

                                print("Book is ", book_name)
                                print("file is ", file_name)
                                try:
                                    os.remove(save_path)
                                    os.remove(conv_path)
                                except Exception as e:
                                    print(e)
                                fp = open(save_path, "wb")
                                fp.write(part.get_payload(decode=True))
                                fp.close()
                                print("saved file", file_name)
                                cmd = (
                                    "ebook-convert "
                                    + '"'
                                    + file_name
                                    + '" '
                                    + book_name
                                    + ".mobi -v"
                                )
                                print("Running", cmd)
                                subprocess.call(cmd)
                                send_book(sender, conv_path, book_name)
                                try:
                                    os.remove(save_path)
                                    os.remove(conv_path)
                                except Exception as e:
                                    print(e)


if __name__ == "__main__":
    main_loop()
