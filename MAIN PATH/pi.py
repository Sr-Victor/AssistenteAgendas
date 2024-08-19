import os.path
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
import pyttsx3 as sx

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SAMPLE_SPREADSHEET_ID = "1SjRO_pxgDIq25SZhdIq0wBZb4CQZdvuDLBzJfs-n0Y4"
RANGES = ["Equipes!A1:C19"]

# Dicionários contendo os emails dos membros de cada equipe
integrantes_atrio = {
    "Navi": "navi@example.com",
    "Rebecka": "rebecka@example.com",
    "Andressa": "andressa@example.com",
    "Sara Naves": "sara@example.com",
    "Arthur Tavares": "arthur@example.com",
    "Ana Júlia": "ana@example.com",
    "Thayná": "thayna@example.com",
    "Jessé": "jesse@example.com",
    "Ian": "ian@example.com",
    "André": "andre@example.com",
    "Victor": "victor@example.com",
    "Jesse": "jesse@example.com",
    "Ezequias": "ezequias@example.com",
    "Samuel": "samuel@example.com"
}

integrantes_connect = {
    "Paulo": "paulo@example.com",
    "Larissa M": "larissa@example.com",
    "Letícia C": "leticia@example.com",
    "Hilel": "hilel@example.com",
    "Layane": "layane@example.com",
    "Daniel": "daniel@example.com",
    "Paulo Miquéias": "paulomiqueias@example.com",
    "Joás": "joas@example.com",
    "Arthur (setor)": "arthursetor@example.com",
    "Davi": "davi@example.com",
    "Guilherme": "guilherme@example.com",
    "Gabriel Akira": "gabriel@example.com",
    "Jesse (GC)": "jessegc@example.com",
    "Pedro": "pedro@example.com",
    "Seixas": "seixas@example.com",
    "Luís": "luis@example.com"
}

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

        all_data = []
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
                    print(f"Vocais: {date_str}")
                    print("===================================================\n")
                    print(f" {equipe}")
                    print("===================================================\n")
                    print(f"{culto}")
                    if "Átrio Music" in equipe:
                        all_data.append({"data": date_str, "equipe": equipe, "culto": culto, "emails": list(integrantes_atrio.values())})
                    elif "Connect Band" in equipe:
                        all_data.append({"data": date_str, "equipe": equipe, "culto": culto, "emails": list(integrantes_connect.values())})

        return all_data

    except HttpError as err:
        print(err)
        return []

def main():
    data = get_data()
    
    if data:
        for entry in data:
            print(f"Data: {entry['data']}, Equipe: {entry['equipe']}, Culto: {entry['culto']}")
            create_event(entry)
    else:
        print("Nenhum dado encontrado.")
        speak("Nenhum dado encontrado.")

def create_event(entry):
    creds = None
    if os.path.exists("FEATURES/Mon Features/token.json"):
        creds = Credentials.from_authorized_user_file("FEATURES/Mon Features/token.json", CS)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "FEATURES/Mon Features/credentials.json", CS
            )
            creds = flow.run_local_server(port=0)
        if creds:
            with open("FEATURES/Mon Features/token.json", "w") as token:
                token.write(creds.to_json())

if __name__ == "__main__":
    main()
