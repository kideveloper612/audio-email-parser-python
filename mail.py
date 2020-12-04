import imaplib
import email
from email.header import decode_header
import re
import dotenv
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
import time

dotenv.load_dotenv()
username = os.environ.get('EMAIL')
password = os.environ.get('PASSWORD')
source = os.environ.get('FROM')

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SAMPLE_SPREADSHEET_ID_input = '1OFPUH6ofB7a4RocJ2LdtB_MKbyfY8ZID1D-FD4CLEGo'
SAMPLE_RANGE_NAME = 'A1:AA1000000'


def write_sheet(records):
    creeds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creeds = pickle.load(token)
    if not creeds or not creeds.valid:
        if creeds and creeds.expired and creeds.refresh_token:
            creeds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creeds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creeds, token)

    service = build('sheets', 'v4', credentials=creeds)

    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_input, range=SAMPLE_RANGE_NAME).execute()
    values_input = result_input.get('values', [])

    column = ['Email TimeStamp', 'Name', 'Phone', 'DOB', 'Location', 'Note', 'CID']
    if len(values_input) == 0:
        records.insert(0, column)

    sheet.values().append(
        spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
        valueInputOption='RAW',
        range=SAMPLE_RANGE_NAME,
        body=dict(
            majorDimension='ROWS',
            values=records)
    ).execute()


def main():
    global source

    while True:
        print('============ START NEW LOOP =============')
        imap = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        imap.login(username, password)
        imap.select("INBOX")

        sub_status, messages = imap.search(None, '(UNSEEN FROM {})'.format(source))
        lines = []
        if sub_status == 'OK':
            for message in messages[0].split():
                res, msg = imap.fetch(message, "(RFC822)")
                for response in msg:
                    try:
                        if isinstance(response, tuple):
                            msg = email.message_from_bytes(response[1])
                            subject = decode_header(msg["Subject"])[0][0]
                            if isinstance(subject, bytes):
                                subject = subject.decode()
                            From, encoding = decode_header(msg.get("From"))[0]
                            if isinstance(From, bytes):
                                From = From.decode(encoding)
                            if msg.is_multipart():
                                for part in msg.walk():
                                    payload = part.get_payload(decode=True)
                                    if payload is not None:
                                        text = payload.decode('utf-8').replace('\r', '').replace('\n', ' ')
                                        date = re.search('=================(.*)======================================', text).group(1).strip()
                                        name = re.search('Name:(.*)Ph#:', text).group(1).strip()
                                        phone = re.search('Ph#:(.*)DOB:', text).group(1).strip()
                                        birthday = re.search('DOB:(.*)Which office were you seen at', text).group(1).strip()
                                        office = re.search('Which office were you seen at(.*)Msg:', text).group(1)[2:].strip()
                                        note = re.search('Msg(.*)CID:', text).group(1)[1:].strip()
                                        cid = re.search('CID:(.*)--------------------------------------', text).group(1).strip()
                                        line = [date, name, phone, birthday, office, note, cid]
                                        print(line)
                                        lines.append(line)
                                        break
                            else:
                                payload = msg.get_payload(decode=True)
                                if payload is not None:
                                    text = payload.decode('utf-8').replace('\r', '').replace('\n', ' ')
                                    date = re.search('=================(.*)======================================',
                                                     text).group(1).strip()
                                    name = re.search('Name:(.*)Ph#:', text).group(1).strip()
                                    phone = re.search('Ph#:(.*)DOB:', text).group(1).strip()
                                    birthday = re.search('DOB:(.*)Which office were you seen at', text).group(1).strip()
                                    office = re.search('Which office were you seen at(.*)Msg:', text).group(1)[2:].strip()
                                    note = re.search('Msg(.*)CID:', text).group(1)[1:].strip()
                                    cid = re.search('CID:(.*)--------------------------------------', text).group(1).strip()
                                    line = [date, name, phone, birthday, office, note, cid]
                                    print(line)
                                    lines.append(line)
                    except:
                        continue
                imap.store(message, '+FLAGS', '\\Seen')

        imap.close()
        imap.logout()

        if len(lines) > 0:
            write_sheet(records=lines)

        time.sleep(4)


if __name__ == '__main__':
    main()

