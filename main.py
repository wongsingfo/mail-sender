#!/usr/bin/python3
import smtplib
import csv
import time
import os
import re
import logging
import traceback
import socks
from getpass import getpass
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header

# ENCODING      = 'ISO-8859-1'
ENCODING      = 'utf-8'
CSV_DELIMITER = '\t'

# MAIL_HOST = 'smtp.gmail.com:587'
MAIL_HOST = "mail.pku.edu.cn"
ME        = "chengke<chengke@pku.edu.cn>"
SENDER    = 'chengke@pku.edu.cn'

DATA_FILE = 'a.csv'
MAIL_FILE = 'document.html'
SUBJECT   = 'Test Document'

# socks.setdefaultproxy(socks.SOCKS5, '127.0.0.1', 7890)
# socks.wrapmodule(smtplib)

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
        self.root['Subject']  = Header(subject, ENCODING)

        self.alt = MIMEMultipart('alternative')
        self.root.attach(self.alt)

        content = render_content(body, fields)
        self.body = MIMEText(content, 'html', ENCODING)
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
        self.server.ehlo()
        self.server.starttls()
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


def start_sending_email(mailer: Mailer, mailfile, datafile, from_):
    with open(datafile, encoding=ENCODING) as csvfile:
        reader = csv.DictReader(csvfile, delimiter=CSV_DELIMITER)
        for fields in reader:
            logging.info('try: ' + str(fields))
            to = fields['emails']
            subject = f'[{ fields["acronym"] }] Problem with #{ fields["ID"] }: { fields["title"] }'
            mail_content = MailContent(mailfile, fields, from_, to, subject)
            
            mailer.send(from_, to.split(';'), mail_content)
            logging.info(f'success: {to}')

if __name__ == "__main__":
    with Mailer(MAIL_HOST, SENDER) as mailer:
        start_sending_email(mailer, MAIL_FILE, DATA_FILE, ME)


