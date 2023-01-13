"""
Tool, which sends mail responses on filled Google Forms
Copyright (C) 2021  Dmytro Kolupaiev

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
"""

import time
import datetime
import smtplib
import codecs

import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials


CREDENTIALS_FILE = 'creds.json'
SPREADSHEET_ID = '1MEsqTX4rYJTSvKRAgTizUchZ7PNuPLFMutNrFArf6eM'


def get_data_from_spreadsheet():
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        CREDENTIALS_FILE,
        ['https://www.googleapis.com/auth/spreadsheets',
         'https://www.googleapis.com/auth/drive']
    )
    httpAuth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

    values = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='A1:D100',
        majorDimension='ROWS'
    ).execute()

    return values


def send_mail(recipient, text):
    smtp_obj = smtplib.SMTP('smtp.gmail.com', 587)
    smtp_obj.starttls()
    with open("gmail_password.txt") as f:
        password = f.readline()
    smtp_obj.login('csoprocom.sa@gmail.com', password)
    smtp_obj.sendmail('csoprocom.sa@gmail.com', recipient, text)
    smtp_obj.quit()


current_recipient_number = 0
FORM_VALIDITY_END_DATE = datetime.date(2021, 6, 17)
while datetime.date.today() < FORM_VALIDITY_END_DATE:
    recipients_mails = []

    recipients_data = get_data_from_spreadsheet()
    number_of_recipients = len(recipients_data['values'])
    for i in range(current_recipient_number, number_of_recipients):
        recipients_mails.append(recipients_data['values'][i][3])

    with codecs.open('letter_text.txt', 'r', 'utf_8') as file:
        message_text = file.read()

    text = (
        "Subject: Вебінар \"Розробка телеграм-боту на Python\"" + message_text
    ).encode('utf-8')

    for recipient_mail in recipients_mails:
        send_mail(recipient_mail, text)

    time.sleep(60*10)
