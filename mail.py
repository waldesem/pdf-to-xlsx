import os
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText

from dotenv import load_dotenv

from config import base_dir


load_dotenv(os.path.join(base_dir, ".env"))

email_sender = os.getenv("EMAIL_SENDER")
email_password = os.getenv("EMAIL_PASSWORD")
smtp_server = os.getenv("SMTP_SERVER")
smtp_port = os.getenv("SMTP_PORT")

"""Sends an email.

Loads the email body content from a text file, sets the subject, 
sender, and recipient, connects to the SMTP server, logs in, 
and sends the message.
"""


def send_email(email_receiver):
    textfile = "email.txt"

    with open(textfile, "rb") as fp:
        msg = MIMEText(fp.read())

        subject = "Test email"
        msg = f"Subject: {subject}\n\n{msg}"
        em = EmailMessage()
        em["From"] = email_sender
        em["To"] = email_receiver
        em["Subject"] = subject
        em.set_content(msg)

    with smtplib.SMTP_SSL(smtp_server, smtp_port) as smtp:
        smtp.login(email_sender, email_password)
        smtp.send_message(em)


if __name__ == "__main__":
    send_email("<EMAIL>")