#!/usr/bin/python3
import smtplib
import csv
import time
import os
import re
import logging
import traceback
from getpass import getpass
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header

MAIL_HOST = "mail.pku.edu.cn"
ME        = "chengke<chengke@pku.edu.cn>"
SENDER    = 'chengke@pku.edu.cn'
DATA_FILE = 'a.csv'
MAIL_FILE = 'document.html'
SUBJECT   = 'Test Document'

logging.basicConfig(level=logging.DEBUG)

def render_content(s, fields):
    def replacer(match):
        key = match.group(1)
        return fields[key]
    s = re.sub(r"{{\s*(\S+)\s*}}", replacer, s)
    return s


def get_file_ext(s):
    filename, file_extension = os.path.splitext(s)
    return file_extension


class MailContent:
    def __init__(self, filename, fields, from_, to, subject):
        with open(filename) as f:
            body = f.read()

        # root > alt > body

        self.root = MIMEMultipart('related')
        self.root['From']     = from_
        self.root['To']       = to
        self.root['Subject']  = Header(subject, 'utf-8')

        self.alt = MIMEMultipart('alternative')
        self.root.attach(self.alt)

        self.body = MIMEText(render_content(body, fields), 'html', 'utf-8')
        self.alt.attach(self.body) 

        self.img_id = 0

    def as_string(self):
        return self.root.as_string()

    def attach_image(self, filename):
        img_id = self.img_id
        self.img_id += 1

        with open(filename, 'rb') as f:
            img = f.read()
        m = MIMEBase('image', get_file_ext(filename), filename=filename)
        m.add_header('Content-Disposition', 'attachment', filename=filename)
        m.add_header('X-Attachment-Id', str(img_id))
        m.add_header('Content-Id', str(img_id))
        m.set_payload(img)
        encoders.encode_base64(m)
        self.body.attach(m)


class Mailer:
    def __init__(self, host, user):
        self.server = smtplib.SMTP()
        self.host   = host
        self.user   = user

    def __enter__(self):
        self.server.connect(self.host)
        passwd = getpass()
        self.server.login(self.user, passwd)
        return self

    def __exit__(self, exc_type, exc_value, tb):
        if exc_type is not None:
            logging.error(exc_value)
            traceback.print_exception(exc_type, exc_value, tb)
        self.server.quit()

    def send(self, from_, to, content: MailContent):
        self.server.sendmail(from_, to, content.as_string())


def start_sending_email(mailer: Mailer, mailfile, datafile, from_, subject):
    with open(datafile) as csvfile:
        reader = csv.DictReader(csvfile)
        for fields in reader:
            logging.info('try: ' + str(fields))
            to = fields['邮箱']
            mail_content = MailContent(mailfile, fields, from_, to, subject)
            mailer.send(from_, to, mail_content)
            logging.info('success: ' + to)


if __name__ == "__main__":
    with Mailer(MAIL_HOST, SENDER) as mailer:
        start_sending_email(mailer, MAIL_FILE, DATA_FILE, ME, SUBJECT)


