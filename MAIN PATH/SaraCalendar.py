import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pyttsx3 as sx
# Se modificar esses escopos, delete o arquivo token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# ID da planilha e os ranges de exemplo.
SAMPLE_SPREADSHEET_ID = "1SjRO_pxgDIq25SZhdIq0wBZb4CQZdvuDLBzJfs-n0Y4"
RANGES = [
    "Átrio ' 24 | Escala!A1:C44",  # Range da primeira aba
]

def speak(text):
    a = sx.init()
    a.say(text)
    a.runAndWait()
def main():
    creds = None
    if os.path.exists("FEATURES/token.json"):
        creds = Credentials.from_authorized_user_file("FEATURES/token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "FEATURES/credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("FEATURES/token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()

        for range_name in RANGES:
            result = (
                sheet.values()
                .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name)
                .execute()
            )
            values = result.get("values", [])

            if not values:
                print(f"No data found in range {range_name}.")
                continue

            print(f"Data from range {range_name}:")
            for row in values:
                # Supondo que "EQUIPE" esteja na coluna 2 (índice 1) e "CULTO" na coluna 3 (índice 2)
                dia = row[0] if len(row) > 1 else 'Sem escala'
                equipe = row[1] if len(row) > 1 else 'Sem equipe'
                culto = row[2] if len(row) > 2 else 'Sem culto'
                if "Átrio Music" in equipe:
                  print('-----------------------------------------')
                  print(f"DIA: {dia}, EQUIPE: {equipe}, CULTO: {culto}")

    except HttpError as err:
        print(err)

if __name__ == "__main__":
    main()
