# Copyright 2018 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# [START calendar_quickstart]
from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']


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
	now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
	print('Getting the upcoming 10 events')
	events_result = service.events().list(calendarId='primary', timeMin=now,
										maxResults=10, singleEvents=True,
										orderBy='startTime').execute()
	events = events_result.get('items', [])

	if not events:
		print('No upcoming events found.')
	for event in events:
		start = event['start'].get('dateTime', event['start'].get('date'))
		print(start, event['summary'])


	#add event
	import xlrd
	book = xlrd.open_workbook('imi2019.xls')
	mag = book.sheets()[9]

	start = ['08:00', '09:50', '11:40', '14:00', '15:45', '17:30']
	end = ['09:35', '11:25', '13:15', '15:35', '17:20', '19:05']  


	for i in range(3, 39):
		day = 2 + (i-3)//6
		if mag.cell(i, 23).value == "":                            
			continue                                              
		para = mag.cell(i, 23).value                               
		l_pr = mag.cell(i, 24).value                               
		room = mag.cell(i, 25).value 
		print(day, "сент", start[(i-3)%6], end[(i-3)%6], para, l_pr, room)
		
		event = {
		  'summary': para,
		  'location': room,
		  'description': l_pr,
		  'start': {
			'dateTime': '2019-09-0'+str(day)+'T'+start[(i-3)%6]+':00+09:00',
			'timeZone': 'Asia/Yakutsk',
		  },
		  'end': {
			'dateTime': '2019-09-0'+str(day)+'T'+end[(i-3)%6]+':00+09:00',
			'timeZone': 'Asia/Yakutsk',
		  },
		  'recurrence': [
			'RRULE:FREQ=WEEKLY;COUNT=12'
		  ],
		  'reminders': {
		  }
		}

		event = service.events().insert(calendarId='tatdr73n8g0pttaobo46317k9k@group.calendar.google.com', body=event).execute()
		print('Event created: %s' % (event.get('htmlLink')))

if __name__ == '__main__':
	main()