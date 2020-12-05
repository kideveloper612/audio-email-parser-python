import imaplib
import email
import re
import glob
import dotenv
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
import time
from dateutil.parser import parse
from googleapiclient.http import MediaFileUpload

dotenv.load_dotenv()
username = os.environ.get('EMAIL')
password = os.environ.get('PASSWORD')
source = os.environ.get('FROM')

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

SAMPLE_SPREADSHEET_ID_input = '1aTeq8ndBQxBpvgyx8GES4tnW_Di_iFyXTkBpinDTCgk'
SAMPLE_RANGE_NAME = 'A1:AA1000000'


def write_sheet(records, filenames):
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

    column = ['Date', 'Name', 'Phone Number', 'Voicemail Duration', 'Transcribed Voicemail', 'Misc.', 'File']
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

    serviceDrive = build('drive', 'v3', credentials=creeds)

    results = serviceDrive.files().list(
        pageSize=10, fields="nextPageToken, files(id, name)").execute()
    items = [item['name'] for item in results.get('files', [])]

    for filename in filenames:
        if filename in items:
            continue
        file_metadata = {'name': filename}
        media = MediaFileUpload('{}'.format(filename), mimetype='audio/mpeg')
        serviceDrive.files().create(body=file_metadata, media_body=media, fields='id').execute()


def getPayload(msg):
    if msg.is_multipart():
        return getPayload(msg.get_payload(0))
    return msg.get_payload(None, True)


def main():
    global source
    while True:
        print('============ START NEW LOOP =============')
        for file in glob.glob("./*.mp3"):
            os.remove(file)
        imap = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        imap.login(username, password)
        imap.select("INBOX")
        sub_status, messages = imap.search(None, '(SEEN FROM {})'.format(source))
        lines = []
        filenames = []
        if sub_status == 'OK':
            for message in messages[0].split():
                res, msg = imap.fetch(message, "(RFC822)")
                for response in msg:
                    try:
                        fileName = ''
                        if isinstance(response, tuple):
                            msg = email.message_from_bytes(response[1])
                            payload = getPayload(msg)
                            if payload is not None:
                                text = payload.decode('utf-8').replace('\r', '').replace('\n', ' ')
                                date = re.search('Time:(.*)From:', text).group(1).strip()
                                namePhone = re.search('Time:(.*)Duration:', text).group(1).replace('From:', '').replace(date, '').strip()
                                name = namePhone.split('(')[0].strip()
                                phone = namePhone.replace(name, '').strip()
                                duration = re.search('Duration:(.*)Transcript:', text).group(1).strip()
                                transcript = re.search('Transcript:(.*)Voicemail box:', text).group(1).split('Rate this transcript')[0].strip()
                                misc = re.search('Voicemail box:(.*)Sincerely', text).group(1).strip()
                                fileDate = parse(date).date()
                                fileTime = str(parse(date).time()).replace(':', '~')
                                fileName = '{}_{}_{}.mp3'.format(name, fileDate, fileTime)
                                line = [date, name, phone, duration, transcript, misc, fileName]
                                print(line)
                                lines.append(line)
                            for part in msg.walk():
                                if part.get_content_maintype() == 'multipart':
                                    continue
                                if part.get('Content-Disposition') is None:
                                    continue
                                fileFlag = part.get_filename()
                                if bool(fileFlag):
                                    if not os.path.isfile(fileName):
                                        fp = open(fileName, 'wb')
                                        fp.write(part.get_payload(decode=True))
                                        fp.close()
                                    filenames.append(fileName)
                    except Exception as e:
                        print(e)
                        continue
                imap.store(message, '+FLAGS', '\\Seen')

        imap.close()
        imap.logout()

        if len(lines) > 0:
            write_sheet(records=lines, filenames=filenames)

        time.sleep(4)


if __name__ == '__main__':
    main()

