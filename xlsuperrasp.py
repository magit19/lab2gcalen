# coding: utf-8
from __future__ import print_function
import datetime
import pickle
import os.path
import xlrd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']

book = xlrd.open_workbook('imi2019.xls')

start = ['08:00', '09:50', '11:40', '14:00', '15:50', '17:40']
end = ['09:35', '11:25', '13:15', '15:35', '17:25', '19:15'] 

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_id.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    # Call the Calendar API

    sh = book.sheets()[9]
    for r in range(3,39):
        para = sh.cell(r,23).value
        room = sh.cell(r,25).value
        l_pr = sh.cell(r,24).value
        day = 2+(r-3)//6
        if para != "":
            para = para.replace("\n", " ")

            print(day,"сент", start[(r-3)%6]+"-"+
                end[(r-3)%6], para, room)
            event = {
              'summary': para,
              'location': room,
              'description': l_pr,
              'start': {
                'dateTime': '2019-09-0'+str(day)+'T'+start[(r-3)%6]+':00+09:00',
                'timeZone': 'Asia/Yakutsk',
              },
              'end': {
                'dateTime': '2019-09-0'+str(day)+'T'+end[(r-3)%6]+':00+09:00',
                'timeZone': 'Asia/Yakutsk',
              },
              'recurrence': [
                'RRULE:FREQ=WEEKLY;COUNT=12'
              ],
              'reminders': {
              }
            }
            event = service.events().insert(calendarId='kbklc9bmhm1ml5h4emb0q4osmc@group.calendar.google.com', body=event).execute()
            print('Event created: %s' % (event.get('htmlLink')))

if __name__ == '__main__':
    main()