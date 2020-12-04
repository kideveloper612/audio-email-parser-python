import imaplib
import email
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

SAMPLE_SPREADSHEET_ID_input = '1aTeq8ndBQxBpvgyx8GES4tnW_Di_iFyXTkBpinDTCgk'
SAMPLE_RANGE_NAME = 'A1:AA1000000'


def get_rows_from_sheet():
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
    return len(values_input)


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

    column = ['Date', 'Name', 'Phone Number', 'Voicemail Duration', 'Transcribed Voicemail', 'Misc.']
    if len(values_input) == 0:
        records.insert(0, column)
    print(len(values_input))

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

    if not os.path.isdir(fileDirectory):
        os.mkdir(fileDirectory)
    while True:
        print('============ START NEW LOOP =============')
        imap = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        imap.login(username, password)
        imap.select("INBOX")
        sub_status, messages = imap.search(None, '(SEEN FROM {})'.format(source))
        lines = []
        if sub_status == 'OK':
            records_count = get_rows_from_sheet()
            print(type(records_count), records_count)
            for message in messages[0].split():
                res, msg = imap.fetch(message, "(RFC822)")
                for response in msg:
                    try:
                        if isinstance(response, tuple):
                            msg = email.message_from_bytes(response[1])
                            for part in msg.walk():
                                payload = part.get_payload(decode=True)
                                if payload is not None:
                                    text = payload.decode('utf-8').replace('\r', '').replace('\n', ' ')
                                    date = re.search('Time:(.*)From:', text).group(1).strip()
                                    namePhone = re.search('Time:(.*)Duration:', text).group(1).replace('From:', '').replace(date, '').strip()
                                    name = namePhone.split('(')[0].strip()
                                    phone = namePhone.replace(name, '').strip()
                                    duration = re.search('Duration:(.*)Transcript:', text).group(1).strip()
                                    transcript = re.search('Transcript:(.*)Voicemail box:', text).group(1).split('Rate this transcript')[0].strip()
                                    misc = re.search('Voicemail box:(.*)Sincerely', text).group(1).strip()
                                    line = [date, name, phone, duration, transcript, misc]
                                    print(line)
                                    records_count += 1
                                    lines.append(line)
                                    break
                            for part in msg.walk():
                                if part.get_content_maintype() == 'multipart':
                                    continue
                                if part.get('Content-Disposition') is None:
                                    continue
                                fileName = part.get_filename()
                                if bool(fileName):
                                    filePath = os.path.join(fileDirectory, '{}.mp3'.format(records_count))
                                    if not os.path.isfile(filePath):
                                        fp = open(filePath, 'wb')
                                        fp.write(part.get_payload(decode=True))
                                        fp.close()
                    except Exception as e:
                        print(e)
                        continue
                imap.store(message, '+FLAGS', '\\Seen')

        imap.close()
        imap.logout()

        if len(lines) > 0:
            write_sheet(records=lines)

        time.sleep(4)


if __name__ == '__main__':
    fileDirectory = 'files'
    main()

