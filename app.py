# Brian Lesko
# 5/10/24
# Fetch emails from your account and gather the data on vacation requests and approvals

from mysecrets import SENDER_EMAIL, SENDER_PASSWORD
import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta

class EmailRetriever:

    def __init__(self,SENDER_EMAIL, SENDER_PASSWORD):
        self.email = SENDER_EMAIL
        self.password = SENDER_PASSWORD
        self.inbox = None
        self.sent = None
        self.mail = mail = imaplib.IMAP4_SSL('imap.gmail.com')
        self.login()

    def login(self):
        self.mail.login(self.email, self.password)

    def fetch_inbox(self):
        self.inbox = self.fetch_box('inbox')
        return self.inbox

    def fetch_sent(self):
        self.sent = self.fetch_box('"[Gmail]/Sent Mail"')
        return self.sent

    def decode_header(self, header_value):
        decoded_header = decode_header(header_value)
        header_parts = [str(part[0], part[1] or 'utf-8') if isinstance(part[0], bytes) else str(part[0]) for part in decoded_header]
        return ''.join(header_parts)

    def fetch_box(self, box):
        self.mail.select(box)
        _, data = self.mail.uid('search', None, "ALL")
        email_ids = data[0].split()
        def fetch_email(e_id):
            _, email_data_raw = self.mail.uid('fetch', e_id, '(BODY.PEEK[])')
            raw_email = email_data_raw[0][1]
            email_message = email.message_from_bytes(raw_email)
            date_tuple = email.utils.parsedate_tz(email_message['Date'])
            date_str = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple)).strftime("%a, %d %b %Y %H:%M:%S") if date_tuple else "Unknown"
            return {
                "From": self.decode_header(email_message['From']),
                "To": self.decode_header(email_message['To']),
                "Subject": self.decode_header(email_message['Subject']),
                "Date": date_str
            }
        return [fetch_email(e_id) for e_id in email_ids]

emr = EmailRetriever(SENDER_EMAIL, SENDER_PASSWORD)


