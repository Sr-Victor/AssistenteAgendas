import os.path
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
import pyttsx3 as sx

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
CS = ["https://www.googleapis.com/auth/calendar"]
SAMPLE_SPREADSHEET_ID = "1SjRO_pxgDIq25SZhdIq0wBZb4CQZdvuDLBzJfs-n0Y4"
RANGES = ["Átrio ' 24 | Escala!A1:C44"]

def speak(text):
    a = sx.init()
    a.say(text)
    a.runAndWait()

def get_data():
    creds = None
    if os.path.exists("FEATURES/Mon Features/token.json"):
        creds = Credentials.from_authorized_user_file("FEATURES/Mon Features/token.json", SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "FEATURES/Mon Features/credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        if creds:
            with open("FEATURES/Mon Features/token.json", "w") as token:
                token.write(creds.to_json())
    
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()

        all_data_atrio = []
        all_data_connect = []

        for range_name in RANGES:
            result = (
                sheet.values()
                .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name)
                .execute()
            )
            values = result.get("values", [])

            for row in values:
                if len(row) >= 3:
                    date_str = row[0] if len(row) > 1 else 'Sem escala'
                    equipe = row[1] if len(row) > 1 else 'Sem equipe'
                    culto = row[2] if len(row) > 2 else 'Sem culto'
                    print(f"Data lida: {date_str}, Equipe: {equipe}, Culto: {culto}")
                    if "Átrio Music" in equipe:
                        all_data_atrio.append({"data": date_str, "equipe": equipe, "culto": culto})
                    elif "Connect" in equipe:
                        all_data_connect.append({"data": date_str, "equipe": equipe, "culto": culto})

        return {"Átrio Music": all_data_atrio, "Connect": all_data_connect}

    except HttpError as err:
        print(err)
        return []

# Adiciona as agendas encontradas no Google Calendar
def add_event_to_calendar(event_data, equipe):
    creds = None
    if os.path.exists("FEATURES/CalendarF/calendar_token.json"):
        creds = Credentials.from_authorized_user_file("FEATURES/CalendarF/calendar_token.json", CS)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "FEATURES/CalendarF/CALENDAr.json", CS
            )
            creds = flow.run_local_server(port=0)
        if creds:
            with open("FEATURES/CalendarF/token.json", "w") as token:
                token.write(creds.to_json())
    
    try:
        # Cria conexão com o calendário e envia para o Google Calendar
        service = build("calendar", "v3", credentials=creds)

        for entry in event_data:
            start_time = datetime.strptime(entry['data'], "%d/%m/%Y").strftime("%Y-%m-%dT18:30:00")
            end_time = datetime.strptime(entry['data'], "%d/%m/%Y").strftime("%Y-%m-%dT21:00:00")
            attendees = []

            if equipe == "Átrio Music":
                attendees = [
                    {'email': 'contatoeuvictoremmanoel@gmail.com'},
                    {'email': 'soloffterminato@gmail.com'},
                    {'email': 'gabrielian387@gmail.com'},
                    {'email': 'reblima13042003@gmail.com'},
                    {'email': 'anavitoriactt@gmail.com'},
                ]
            elif equipe == "Connect":
                attendees = [
                    {'email': 'paulomiqueiasserrabastos2003@gmail.com'},
                ]

            event = {
                'summary': f'Escala {equipe}',
                'location': 'Assembleia de Deus Quadra 12',
                'description': f"{entry['culto']} - Equipe: {entry['equipe']}",
                'start': {
                    'dateTime': start_time,
                    'timeZone': 'America/Sao_Paulo',
                },
                'end': {
                    'dateTime': end_time,
                    'timeZone': 'America/Sao_Paulo',
                },
                'attendees': attendees,
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                        {'method': 'email', 'minutes': 24 * 60},
                        {'method': 'popup', 'minutes': 10},
                    ],
                },
            }
            event = service.events().insert(calendarId='primary', body=event).execute()
            print(f"Evento criado: {event.get('htmlLink')}")
            speak("Agenda criada com sucesso e enviada para os líderes e participantes do grupo.")

    except HttpError as error:
        print(f"An error occurred: {error}")

def find_next_scale(scales):
    today = datetime.now()
    upcoming_scales = [scale for scale in scales if datetime.strptime(scale['data'], "%d/%m/%Y") >= today]

    if not upcoming_scales:
        return None

    next_scale = min(upcoming_scales, key=lambda x: datetime.strptime(x['data'], "%d/%m/%Y"))
    return next_scale

def main():
    data = get_data()
    if data:
        next_scale_atrio = find_next_scale(data["Átrio Music"])
        next_scale_connect = find_next_scale(data["Connect"])

        if next_scale_atrio:
            print(f"\nPróxima Escala Átrio Music:\nData: {next_scale_atrio['data']}, Equipe: {next_scale_atrio['equipe']}, Culto: {next_scale_atrio['culto']}")
            add_event_to_calendar([next_scale_atrio], "Átrio Music")
        
        if next_scale_connect:
            print(f"\nPróxima Escala Connect:\nData: {next_scale_connect['data']}, Equipe: {next_scale_connect['equipe']}, Culto: {next_scale_connect['culto']}")
            add_event_to_calendar([next_scale_connect], "Connect")
    else:
        print("Nenhum dado encontrado.")
        speak("Nenhum dado encontrado.")

if __name__ == "__main__":
    main()
