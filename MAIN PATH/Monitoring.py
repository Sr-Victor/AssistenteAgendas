import os.path
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
import json
from flask import *
import pyttsx3 as sx
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SAMPLE_SPREADSHEET_ID = "1SjRO_pxgDIq25SZhdIq0wBZb4CQZdvuDLBzJfs-n0Y4"
RANGES = ["Átrio ' 24 | Escala!A1:C44"]



def speak(text):
    a = sx.init()
    a.say(text)
    a.runAndWait()

def get_data():
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
                    date_str = row[0] if len(row) > 1 else 'Sem escala'  # Assume que a data está na primeira coluna
                    equipe = row[1] if len(row) > 1 else 'Sem equipe'
                    culto = row[2] if len(row) > 2 else 'Sem culto'
                    print(f"Data lida: {date_str}, Equipe: {equipe}, Culto: {culto}")  # Depuração
                    if "Átrio Music" in equipe:
                        all_data.append({"data": date_str, "equipe": equipe, "culto": culto})

        return all_data

    except HttpError as err:
        print(err)
        return []
# Envia os dados para uma API Json
def send_to_api(data):
    app = Flask(__name__)
    @app.route('/')
    def index():
        return json.dumps({'name': 'alice',
                        'email': 'alice@outlook.com'})
    app.run()


def find_next_scale(scales):
    today = datetime.now()
    upcoming_scales = []

    for scale in scales:
        scale_date = datetime.strptime(scale['data'], "%d/%m/%Y")  # Ajuste o formato da data se necessário
        print(f"Comparando com a data de hoje: {scale_date} >= {today}")  # Depuração
        if scale_date >= today:
            upcoming_scales.append(scale)

    if not upcoming_scales:
        return None  # Nenhuma escala futura encontrada

    next_scale = min(upcoming_scales, key=lambda x: datetime.strptime(x['data'], "%d/%m/%Y"))
    time_remaining = (datetime.strptime(next_scale['data'], "%d/%m/%Y") - today).days

    return next_scale, time_remaining

def main():
    data = get_data()  # Obtém os dados do Google Sheets

    # Encontra a próxima escala e o tempo restante
    next_scale, time_remaining = find_next_scale(data) if data else (None, None)

    # Saída no console
    speak("Exibindo as próximas escalas de Átrio Music")
    print("---------------------------------")
    
    # Mostra as escalas
    if data:
        for entry in data:
            speak(f"Data: {entry['data']}, Equipe: {entry['equipe']}, Culto: {entry['culto']}")
    else:
        speak("Nenhum dado encontrado.")

    # Adiciona aviso sobre a próxima escala
    if next_scale:
        speak("\nPróxima Escala:")
        speak(f"Data: {next_scale['data']}")
        speak(f"Equipe: {next_scale['equipe']}")
        speak(f"Culto: {next_scale['culto']}")
        speak(f"Faltam {time_remaining} dias para a próxima escala.")
    else:
        speak("\nNenhuma escala futura encontrada.")
    datas = f"Data: {entry['data']}, Equipe: {entry['equipe']}, Culto: {entry['culto']}"
    send_to_api(datas)

if __name__ == "__main__":
    main()
