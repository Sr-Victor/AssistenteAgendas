# Para diminuir o seu processo de leituro: Este código ele checa se você possui alguma escala hoje.

import os.path
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
import pyttsx3 as sx

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

def speak(text):
    a = sx.init()
    a.say(text)
    a.runAndWait()

def get_data():
    creds = None
    if os.path.exists("FEATURES/CalendarF/CHECK/token.json"):
        creds = Credentials.from_authorized_user_file("FEATURES/CalendarF/CHECK/token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("FEATURES/CalendarF/CHECK/calendar.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("FEATURES/CalendarF/CHECK/token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()

        result = sheet.values().get(
            spreadsheetId="1SjRO_pxgDIq25SZhdIq0wBZb4CQZdvuDLBzJfs-n0Y4",
            range="Átrio ' 24 | Escala!A1:C44"
        ).execute()
        values = result.get("values", [])

        return values

    except HttpError as err:
        print(err)
        return []

def check_today_schedule(data):
    today = datetime.now().strftime("%d/%m/%Y")
    for row in data:
        if len(row) >= 3 and row[0] == today:
            print(f"Hoje tem agenda: Data: {row[0]}, Equipe: {row[1]}, Culto: {row[2]}")
            speak(f"Hoje tem agenda com a equipe {row[1]} para o culto {row[2]}.")
            return True
    return False

def main():
    data = get_data()
    if not data:
        print("Nenhum dado encontrado.")
        speak("Nenhum dado encontrado.")
        return

    if not check_today_schedule(data):
        print("Não há agenda para hoje.")
        speak("Não há agenda para hoje.")

if __name__ == "__main__":
    main()
