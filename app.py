# Brian Lesko
# 5/10/24
# Fetch emails from your account and gather the data on vacation requests and approvals

import imaplib
import streamlit as st
import pandas as pd
from mysecrets import SENDER_EMAIL, SENDER_PASSWORD
import email
from email.header import decode_header
import datetime

class EmailRetriever:
    def __init__(self,SENDER_EMAIL, SENDER_PASSWORD):
        self.email = SENDER_EMAIL
        self.password = SENDER_PASSWORD

    def fetch_inbox(self,email_address, email_password):
        # Connect to the email server
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(SENDER_EMAIL, SENDER_PASSWORD)
        mail.select('inbox')

        # Fetch emails
        result, data = mail.uid('search', None, "ALL")
        email_ids = data[0].split()
        email_data = []

        for e_id in email_ids:
            _, email_data_raw = mail.uid('fetch', e_id, '(BODY.PEEK[])')
            raw_email = email_data_raw[0][1].decode("utf-8")
            email_message = email.message_from_string(raw_email)

            # Get the details
            date_tuple = email.utils.parsedate_tz(email_message['Date'])
            if date_tuple:
                local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
                local_message_date = "%s" %(str(local_date.strftime("%a, %d %b %Y %H:%M:%S")))
            email_from = str(decode_header(email_message['From'])[0][0])
            email_to = str(decode_header(email_message['To'])[0][0])
            subject = str(decode_header(email_message['Subject'])[0][0])

            email_data.append({"Date": local_message_date, "From": email_from, "To": email_to, "Subject": subject})

        # Create DataFrame
        df = pd.DataFrame(email_data)
        self.inbox = df
        return df
    
    def fetch_sent(self):
        # Connect to the email server
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(self.email, self.password)
        mail.select('"[Gmail]/Sent Mail"')  # Select the Sent Mail box

        # Fetch emails
        result, data = mail.uid('search', None, "ALL")
        email_ids = data[0].split()
        email_data = []

        for e_id in email_ids:
            _, email_data_raw = mail.uid('fetch', e_id, '(BODY.PEEK[])')
            raw_email = email_data_raw[0][1].decode("utf-8")
            email_message = email.message_from_string(raw_email)

            # Get the details
            date_tuple = email.utils.parsedate_tz(email_message['Date'])
            if date_tuple:
                local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
                local_message_date = "%s" %(str(local_date.strftime("%a, %d %b %Y %H:%M:%S")))
            email_from = str(decode_header(email_message['From'])[0][0])
            email_to = str(decode_header(email_message['To'])[0][0])
            subject = str(decode_header(email_message['Subject'])[0][0])

            email_data.append({"Date": local_message_date, "From": email_from, "To": email_to, "Subject": subject})

        # Create DataFrame
        df = pd.DataFrame(email_data)
        self.sent = df
        return df

emr = EmailRetriever(SENDER_EMAIL, SENDER_PASSWORD)
df = emr.fetch_inbox(SENDER_EMAIL, SENDER_PASSWORD)
st.header("Inbox")
st.write(df)
st.header("Sent")
df = emr.fetch_sent()
st.write(df)

