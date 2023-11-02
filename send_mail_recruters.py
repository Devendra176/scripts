import argparse
import dotenv
import os
import copy
import openpyxl
import smtplib
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

from global_var import SUBJECT, BODY, ATTACHMENT, EMAIL_FILE


class SendMail:
    def __init__(self):
        self.sender_email = sender_email
        self.password = password
        self.host = host
        self.port = port
        self.emails = []
        self.sheet = None

    def read_file(self):
        file_path = EMAIL_FILE
        workbook = openpyxl.load_workbook(file_path)
        self.sheet = workbook.active
        return self.sheet

    def read_emails_from_file(self):
        for row in self.sheet.iter_rows(values_only=True):
            recipient_email = row[0]
            pattern = r"^(?!.*(?:deqode\.com|cis\.com)).*$"
            if re.match(pattern, row[0]):
                self.send_email(recipient_email)
                self.emails.append(row[0])

        return self.emails

    @staticmethod
    def get_message_attr(recipient_email):
        attachment_resume_obj = AttachResume()
        attachment_resume_obj.set_recipent_email(recipient_email)
        return attachment_resume_obj.message

    def send_email(self, recipient_email):
        message = SendMail.get_message_attr(recipient_email)
        try:
            server = smtplib.SMTP(self.host, self.port)
            server.starttls()
            server.login(self.sender_email, self.password)
            server.sendmail(self.sender_email, recipient_email, message.as_string())
            server.quit()
            print(f"Email sent to {recipient_email}")
        except Exception as e:
            print(f"Failed to send email to {recipient_email}. Error: {str(e)}")

    def run(self):
        self.read_file()
        self.read_emails_from_file()


# Trying to Implement Prototype Design Pattern
class AttachResume:
    __instance = None

    def __init__(self):
        self._sender_email = sender_email
        self._sender_password = password
        self._subject = SUBJECT
        self._body = BODY
        self._attachment_path = ATTACHMENT
        self._recipient_email = None
        self._message = None
        self.get_message_attr()

    def clone(self):
        AttachResume.__instance = copy.copy(self)
        return AttachResume.__instance

    @property
    def message(self):
        return self._message

    def set_recipent_email(self, recipent):
        self._message["To"] = recipent
        return recipent

    def get_message_attr(self):
        message = MIMEMultipart()
        message["From"] = self._sender_email
        message["Subject"] = self._subject

        message.attach(MIMEText(self._body, "plain"))

        with open(self._attachment_path, "rb") as file:
            attach = MIMEApplication(file.read(), _subtype="pdf")
            attach.add_header(
                "Content-Disposition", "attachment", filename=self._attachment_path
            )
            message.attach(attach)

        self._message = message
        return self._message


if __name__ == "__main__":
    dotenv.load_dotenv(".env")

    parser = argparse.ArgumentParser(description="Request Parser options")
    parser.add_argument(
        "--prod",
        action="store_true",
        help="Script will take actual cred and this may  trigger the mails",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="Script will take dummy cred and for testing purpose see.. mailtrap",
    )

    args = parser.parse_args()

    sender_email = None
    password = None
    host = None
    port = None

    if args.test:
        sender_email = os.getenv("TEST_SENDER")
        password = os.getenv("TEST_PASSWORD")
        host = os.getenv("TEST_HOST")
        port = os.getenv("TEST_PORT")

    if args.prod:
        sender_email = os.getenv("SENDER")
        password = os.getenv("PASSWORD", "")
        host = os.getenv("HOST")
        port = os.getenv("PORT")

    sendmail = SendMail()
    sendmail.run()
