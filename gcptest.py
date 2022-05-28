from __future__ import print_function

from pickle import TRUE

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import datetime
import time

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/calendar.events.readonly']

SERVICE_ACCOUNT_FILE = 'GCP test\credentials.json'

SAMPLE_SPREADSHEET_ID = '1dQViypBfyV8xM-489rnebxI8hbi_84UPQ6JSWzxyC_s'
SAMPLE_RANGE_NAME = 'Event storage!A2:E'
VALUE_INPUT_OPTION = 'USER_ENTERED'
INSERT_INPUT_OPTION = 'INSERT_ROWS'


creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


def main():
   
   continue_ = TRUE

   while(continue_) :
        try : 
            
            service_sheets = build('sheets', 'v4', credentials = creds)
            sheet = service_sheets.spreadsheets()


            service_calendar = build('calendar', 'v3', credentials = creds)

            now = datetime.datetime.utcnow()
            dateMin = now - datetime.timedelta(minutes=20, microseconds=1)
            print("Getting the events created during the last 15 minutes")
            events_result = service_calendar.events().list(calendarId = 'qlenestour@gmail.com', updatedMin = dateMin.isoformat() + 'Z',
                                                orderBy = 'startTime', singleEvents =True ).execute()
            events = events_result.get('items', [])

            if not events:
                print('No event added during the last 15 minutes')
                return

            # Prints the start and name of the next 10 events
            for event in events:
                if event['status'] == 'confirmed' :
                    start = event['start'].get('dateTime', event['start'].get('date'))
                    print(start, event['summary'], event['description'])

                    VALUE_RANGE_BODY = {
                        "range" : 'Event storage!A2:E',
                        "majorDimension" : 'ROWS',
                        "values" : [
                            [
                                event['summary'],
                                start,
                                event['description'],
                                event['updated'],
                            ]
                        ]
                    }
                    request =  sheet.values().append(spreadsheetId = SAMPLE_SPREADSHEET_ID, range = SAMPLE_RANGE_NAME, 
                                        valueInputOption = VALUE_INPUT_OPTION, 
                                        insertDataOption = INSERT_INPUT_OPTION, body = VALUE_RANGE_BODY)
                    response = request.execute()
                    print(response)

        except HttpError as error:
            print('An error occurred: %s' % error)

        time.sleep(1200)


main()

