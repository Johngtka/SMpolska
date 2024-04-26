import os
import json
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

buffer = []
rowCounter = 1

indexKRS = pd.read_excel('SMpolska.xls', sheet_name='Index', dtype={
    'Członkowie Zarządu': str})

boardMembersKRS = pd.read_excel('SMpolska.xls', sheet_name='CzlonkowieZarzadu', dtype={
    'Członkowie Zarządu': str})

relationKRS = set(indexKRS['Członkowie Zarządu']).intersection(
    boardMembersKRS['Członkowie Zarządu'])


def send_email(subject, message, to):
    msg = MIMEMultipart()
    msg['From'] = 'sender-email@gmail.com'
    msg['To'] = to
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    server = smtplib.SMTP('smtp-domain.com', 000)
    server.starttls()
    server.login(msg['From'], 'sender-password')
    server.send_message(msg)
    server.quit()


if os.path.exists('bufor.json'):
    with open('bufor.json', 'r', encoding='utf-8') as bufferFileRead:
        temp = json.load(bufferFileRead)

for numberKRS in relationKRS:
    rows = boardMembersKRS[boardMembersKRS['Członkowie Zarządu'] == numberKRS]

    for _, row in rows.iterrows():
        rowDict = row.to_dict()
        print(f'Wiersz {rowCounter}:')
        print(row)
        print('\n')
        rowCounter += 1
        buffer.append(rowDict)

with open('bufor.json', 'a', encoding='utf-8') as bufferFileSave:
    json.dump(buffer, bufferFileSave, ensure_ascii=False, indent=4)
