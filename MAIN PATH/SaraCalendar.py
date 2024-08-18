import datetime
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from Monitoring import get_data

# If modifying these scopes, delete the file token.json.
CS = ["https://www.googleapis.com/auth/calendar"]



"""Shows basic usage of the Google Calendar API.
Prints the start and name of the next 10 events on the user's calendar.
"""
creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", CS)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
else:
    flow = InstalledAppFlow.from_client_secrets_file(
        "FEATURES/CALENDAR.json", CS
    )
    creds = flow.run_local_server(port=0)
# Save the credentials for the next run
with open("token.json", "w") as token:
    token.write(creds.to_json())

try:
    service = build("calendar", "v3", credentials=creds)

    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time
    print("Getting the upcoming 10 events")
    events_result = (
        service.events()
        .list(
            calendarId="primary",
            timeMin=now,
            maxResults=10,
            singleEvents=True,
            orderBy="startTime",
        )
        .execute()
    )
    events = events_result.get("items", [])


    if not events:
        print("No upcoming events found.")

    # Prints the start and name of the next 10 events
    for event in events:
        start = event["start"].get("dateTime", event["start"].get("date"))
        print(start, event["summary"])

    # Refer to the Python quickstart on how to setup the environment:
    # https://developers.google.com/calendar/quickstart/python
    # Change the scope to 'https://www.googleapis.com/auth/calendar' and delete any
    # stored credentials.
    event = {
        'summary': 'Google I/O 2024',
        'location': 'Assembleia de Deus Quadra 12',
        'description': 'Escala | Átrio Music',
        'start': {
            'dateTime': '2024-08-18T09:00:00-03:00',  # Horário de início com fuso horário de São Paulo
            'timeZone': 'America/Sao_Paulo',
        },
        'end': {
            'dateTime': '2024-08-18T17:00:00-03:00',  # Horário de término com fuso horário de São Paulo
            'timeZone': 'America/Sao_Paulo',
        },
        'recurrence': [
            'RRULE:FREQ=DAILY;COUNT=2'  # Recorrência diária por 2 dias
        ],
        'attendees': [
            {'email': 'contatoeuvictoremmanoel@gmail.com'},
            {'email': 'soloffterminato@gmail.com'},
        ],
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 24 * 60},  # Lembrete por e-mail 24 horas antes
                {'method': 'popup', 'minutes': 10},       # Lembrete pop-up 10 minutos antes
            ],
        },
    }

    event = service.events().insert(calendarId='primary', body=event).execute()

except HttpError as error:
    print(f"An error occurred: {error}")


if __name__ == "__main__":
  main()