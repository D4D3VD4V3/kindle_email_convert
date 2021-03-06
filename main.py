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

CONVERT_EXT = ".mobi"


def generate_email(sender_email, book_path, book_name):
    msg = MIMEMultipart()
    msg["Subject"] = "Converted " + book_name
    msg["From"] = config.email
    msg["To"] = config.kindle_email

    with (open(book_path, "rb")) as fp:
        attachment = MIMEApplication(fp.read())

    attachment.add_header(
        "Content-Disposition", "attachment", filename=book_name + CONVERT_EXT
    )
    msg.attach(attachment)

    return msg


def cleanup_files(paths):
    try:
        for path in paths:
            os.remove(path)
    except Exception as e:
        print(e)


def parse_emails(messages):
    for num in messages[0].split():
        _, data = mail.fetch(num, "(RFC822)")

        for response_part in data:
            if isinstance(response_part, tuple):
                m = email.message_from_bytes(response_part[1])
                sender = m["From"]
                print(sender)

                mail.store(num, "+FLAGS", "\\Seen")

                for part in m.walk():

                    if part.get_content_maintype() in ("multipart", "text") or part.get("Content-Disposition") is None:
                        continue

                    file_name = decode_header(part.get_filename())[0][0]

                    if not isinstance(file_name, str):
                        file_name = file_name.decode("utf-8")

                    return file_name, part, sender


@retry
def main_loop():

    mail.list()
    mail.select("inbox")

    retcode, messages = mail.search(None, "(UNSEEN)")

    if retcode == "OK":
        file_name, part, sender_email = parse_emails(messages)

        save_dir = os.getcwd()
        print("Conversion started", file_name)
        save_path = os.path.join(save_dir, file_name)
        book_name = os.path.splitext(os.path.basename(save_path))[0]
        conv_path = os.path.join(save_dir, book_name + CONVERT_EXT)
        cleanup_paths = [save_path, conv_path]

        cleanup_files(cleanup_paths)

        with open(save_path, "wb") as fp:
            fp.write(part.get_payload(decode=True))
            print("Saved file", file_name)

        cmd = ["ebook-convert", file_name, book_name + CONVERT_EXT, "-v"]
        subprocess.call(cmd)

        acct = smtplib.SMTP_SSL(config.smtp_host)
        acct.login(config.email, config.password)

        msg = generate_email(sender_email, conv_path, book_name)

        try:
            acct.sendmail(config.email, config.kindle_email, msg.as_string())
            print("Sent", conv_path)

        except smtplib.SMTPSenderRefused as e:
            cleanup_files(cleanup_paths)
            print("Could not send book")
            msg = MIMEText(e.args[1].decode("utf-8"))
            msg["Subject"] = "Conversion ERROR"
            msg["To"] = sender_email
            acct.sendmail(config.email, sender_email, msg.as_string())


if __name__ == "__main__":
    mail = imaplib.IMAP4_SSL(config.imap_host)
    mail.login(config.email, config.password)
    main_loop()
