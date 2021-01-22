import imaplib
import email
import os
import datetime
from random import randint
from dateutil.parser import parse
from pytz import timezone
import pdfkit
import re

config = pdfkit.configuration(wkhtmltopdf='/usr/bin/wkhtmltopdf')
download_folder = "new_attach"
if not os.path.exists(download_folder):
    os.mkdir(download_folder)

logs_folder = "logs"
if not os.path.exists(logs_folder):
    os.mkdir(logs_folder)

def remove_img_tags(data):
    p = re.compile(r'<img.*?>')
    return p.sub('', data)

def file_no(len):
    return str(randint((10 ** (len - 1)), 10 ** len)) + '_'

def format_date(date):
    date = date.split(',')[-1].strip()
    format = '%d %b %Y %H:%M:%S %z'
    if '(' in date:
        date = date.split('(')[0].strip()
    try:
        date = datetime.strptime(date, format)
    except:
        try:
            date = parse(date)
        except:
            with open('logs/date_err.log', 'a') as fp:
                print(date, file=fp)
            raise Exception
    date = date.astimezone(timezone('Asia/Kolkata')).replace(tzinfo=None)
    format1 = '%d/%m/%Y %H:%M:%S'
    date = date.strftime(format1)
    return date


def file_blacklist(filename):
    fp = filename
    filename, file_extension = os.path.splitext(fp)
    ext = ['.pdf', '.htm', '.html']
    if file_extension not in ext:
        return False
    if fp.find('ATT00001') != -1:
        return False
    if (fp.find('MDI') != -1) and (fp.find('Query') == -1):
        return False
    if (fp.find('knee') != -1):
        return False
    if (fp.find('KYC') != -1):
        return False
    if fp.find('image') != -1:
        return False
    if (fp.find('DECLARATION') != -1):
        return False
    if (fp.find('Declaration') != -1):
        return False
    if (fp.find('notification') != -1):
        return False
    if (fp.find('CLAIMGENIEPOSTER') != -1):
        return False
    if (fp.find('declar') != -1):
        return False
    if (fp.find('PAYMENT_DETAIL') != -1):
        return False
    return True

def save_attachment(msg):
    print("====================")
    """
    Given a message, save its attachments to the specified
    download folder (default is /tmp)

    return: file path to attachment
    """
    att_path = []
    flag = 0
    for part in msg.walk():
        print(part.get_content_maintype())
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        flag = 1
        filename = part.get_filename()
        if filename is not None and file_blacklist(filename):
            if not os.path.isfile(filename):
                fp = open(os.path.join(download_folder, file_no(4) + filename), 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
                att_path.append(os.path.join(download_folder, file_no(4) + filename))
    if flag == 0 or filename is None or len(att_path) == 0:
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                filename = 'text.txt'
                fp = open(os.path.join(download_folder, filename), 'wb')
                data = part.get_payload(decode=True)
                fp.write(data)
                fp.close()
                att_path = os.path.join(download_folder, filename)
            if part.get_content_type() == 'text/html':
                filename = 'text.html'
                fp = open(os.path.join(download_folder, filename), 'wb')
                data = part.get_payload(decode=True)
                fp.write(data)
                fp.close()
                with open(os.path.join(download_folder, filename), 'r') as fp:
                    data = fp.read()
                data = remove_img_tags(data)
                with open(os.path.join(download_folder, filename), 'w') as fp:
                    fp.write(data)
                att_path = os.path.join(download_folder, filename)
                pass
    return att_path
# Connect to inbox
imap_server = imaplib.IMAP4_SSL(host="gptgroup.icewarpcloud.in")
imap_server.login('mediclaim.ils.howrah@gptgroup.co.in', 'Gpt@2019')
# imap_server.login('billing.ils.agartala@gptgroup.co.in', 'Gpt@2019')
imap_server.select(readonly=True)  # Default is `INBOX`
# Find all emails in inbox and print out the raw email data
# _, message_numbers_raw = imap_server.search(None, 'ALL')
_, message_numbers_raw = imap_server.search(None, '(SINCE "22-Jan-2021")')
for message_number in message_numbers_raw[0].split():
    _, msg = imap_server.fetch(message_number, '(RFC822)')
    message = email.message_from_bytes(msg[0][1])
    sender = message['from']
    date = format_date(message['Date'])
    subject = message['Subject']
    mid = int(message_number)
    print(mid, date)
    a = save_attachment(message)
    if not isinstance(a, list):
        filename = 'new_attach/' + file_no(8) + '.pdf'
        pdfkit.from_file(a, filename, configuration=config)
        print(filename)

    else:
        print(a[-1])
    # print(subject, date, sender, a, sep='||')