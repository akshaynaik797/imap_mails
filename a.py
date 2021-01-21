import imaplib
import email
import os

def save_attachment(msg):
    print("====================")
    """
    Given a message, save its attachments to the specified
    download folder (default is /tmp)

    return: file path to attachment
    """
    download_folder="new_attach"
    if not os.path.exists(download_folder):
        os.mkdir(download_folder)
    att_path = "No attachment found."
    for part in msg.walk():
        # print(part.get_content_maintype())
        if part.get_content_maintype() == 'text':
            print(part.get_content_type())
            z =1
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()
        print(filename)
        if filename is not None:
            att_path = os.path.join(download_folder, filename)
        else:
            print('not found')

        if not os.path.isfile(att_path):
            fp = open(att_path, 'wb')
            fp.write(part.get_payload(decode=True))
            fp.close()
    return att_path
# Connect to inbox
imap_server = imaplib.IMAP4_SSL(host="gptgroup.icewarpcloud.in")
# imap_server.login('mediclaim.ils.howrah@gptgroup.co.in', 'Gpt@2019')
imap_server.login('billing.ils.agartala@gptgroup.co.in', 'Gpt@2019')
imap_server.select(readonly=True)  # Default is `INBOX`
# Find all emails in inbox and print out the raw email data
# _, message_numbers_raw = imap_server.search(None, 'ALL')
_, message_numbers_raw = imap_server.search(None, '(SINCE "21-Jan-2021")')
for message_number in message_numbers_raw[0].split():
    _, msg = imap_server.fetch(message_number, '(RFC822)')
    message = email.message_from_bytes(msg[0][1])
    sender = message['from']
    date = message['Date']
    subject = message['Subject']
    mid = int(message_number)
    a = save_attachment(message)
    # print(subject, date, sender, a, sep='||')