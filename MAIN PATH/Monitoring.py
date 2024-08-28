import sys
import json
import csv
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QLineEdit, QLabel, QHBoxLayout, QMessageBox, QFileDialog, QTableWidget, QTableWidgetItem, QComboBox
)
from PyQt5.QtCore import Qt
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
import pyttsx3

# Arquivo de configuração para emails
EMAIL_CONFIG_FILE = "emails_config.json"
# Configurações do Google Sheets e Calendar
config = {
    "SAMPLE_SPREADSHEET_ID": "1SjRO_pxgDIq25SZhdIq0wBZb4CQZdvuDLBzJfs-n0Y4",
    "RANGES": ["Átrio ' 24 | Escala!A1:C44"]
}
# Path dos arquivos de credenciais do Google
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'

class ConfigPage(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Configurações")
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # Campo para alterar SAMPLE_SPREADSHEET_ID
        self.spreadsheet_id_input = QLineEdit(self)
        self.spreadsheet_id_input.setPlaceholderText("SAMPLE_SPREADSHEET_ID")
        layout.addWidget(QLabel("Alterar Spreadsheet ID:"))
        layout.addWidget(self.spreadsheet_id_input)

        # Campo para alterar RANGES
        self.range_input = QLineEdit(self)
        self.range_input.setPlaceholderText("RANGES")
        layout.addWidget(QLabel("Alterar Ranges:"))
        layout.addWidget(self.range_input)

        # Botão para salvar alterações
        save_button = QPushButton("Salvar Alterações", self)
        save_button.clicked.connect(self.save_config)
        layout.addWidget(save_button)

        # Adicionar Email
        self.email_input = QLineEdit(self)
        self.email_input.setPlaceholderText("Adicionar Email")
        layout.addWidget(self.email_input)

        add_email_button = QPushButton("Adicionar Email", self)
        add_email_button.clicked.connect(self.add_email)
        layout.addWidget(add_email_button)

        # Excluir Email
        self.email_delete_input = QLineEdit(self)
        self.email_delete_input.setPlaceholderText("Excluir Email")
        layout.addWidget(self.email_delete_input)

        delete_email_button = QPushButton("Excluir Email", self)
        delete_email_button.clicked.connect(self.delete_email)
        layout.addWidget(delete_email_button)

        # Botão de Voltar
        back_button = QPushButton("Voltar", self)
        back_button.clicked.connect(self.close)
        layout.addWidget(back_button)

        self.setLayout(layout)

    def save_config(self):
        spreadsheet_id = self.spreadsheet_id_input.text()
        ranges = self.range_input.text()

        if spreadsheet_id and ranges:
            config["SAMPLE_SPREADSHEET_ID"] = spreadsheet_id
            config["RANGES"] = [ranges]
            with open("config.json", "w") as config_file:
                json.dump(config, config_file)
            QMessageBox.information(self, "Sucesso", "Configurações salvas com sucesso!")
        else:
            QMessageBox.warning(self, "Erro", "Por favor, preencha todos os campos.")

    def add_email(self):
        email = self.email_input.text()
        if email:
            if os.path.exists(EMAIL_CONFIG_FILE):
                with open(EMAIL_CONFIG_FILE, "r") as file:
                    emails = json.load(file)
            else:
                emails = []

            if email not in emails:
                emails.append(email)
                with open(EMAIL_CONFIG_FILE, "w") as file:
                    json.dump(emails, file)
                QMessageBox.information(self, "Sucesso", "Email adicionado com sucesso!")
            else:
                QMessageBox.warning(self, "Erro", "Este email já está na lista.")
        else:
            QMessageBox.warning(self, "Erro", "Por favor, insira um email válido.")

    def delete_email(self):
        email = self.email_delete_input.text()
        if email:
            if os.path.exists(EMAIL_CONFIG_FILE):
                with open(EMAIL_CONFIG_FILE, "r") as file:
                    emails = json.load(file)

                if email in emails:
                    emails.remove(email)
                    with open(EMAIL_CONFIG_FILE, "w") as file:
                        json.dump(emails, file)
                    QMessageBox.information(self, "Sucesso", "Email removido com sucesso!")
                else:
                    QMessageBox.warning(self, "Erro", "Email não encontrado na lista.")
            else:
                QMessageBox.warning(self, "Erro", "Nenhum email encontrado para deletar.")
        else:
            QMessageBox.warning(self, "Erro", "Por favor, insira um email válido.")

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerenciador de Escalas")
        self.initUI()
        self.init_google_services()
        self.init_speech_engine()

    def initUI(self):
        layout = QVBoxLayout()

        # Tabela para mostrar escalas
        self.table = QTableWidget(self)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Data", "Equipe", "Músicas"])
        layout.addWidget(self.table)

        # Filtros por equipe
        self.team_filter = QComboBox(self)
        self.team_filter.addItems(["Todas", "Átrio Music", "Connect Band"])
        self.team_filter.currentIndexChanged.connect(self.load_schedule_data)
        layout.addWidget(self.team_filter)

        # Botão para enviar escala
        send_schedule_button = QPushButton("Enviar Escala", self)
        send_schedule_button.clicked.connect(self.send_schedule)
        layout.addWidget(send_schedule_button)

        # Botão para criar backup em CSV
        create_csv_button = QPushButton("Criar Backup em CSV", self)
        create_csv_button.clicked.connect(self.create_csv_backup)
        layout.addWidget(create_csv_button)

        # Botão para criar nova escala
        create_new_schedule_button = QPushButton("Criar Nova Escala", self)
        create_new_schedule_button.clicked.connect(self.create_new_schedule)
        layout.addWidget(create_new_schedule_button)

        # Botão para abrir página de configurações
        config_button = QPushButton("Configurações", self)
        config_button.clicked.connect(self.open_config_page)
        layout.addWidget(config_button)

        self.setLayout(layout)
        self.load_schedule_data()

    def init_google_services(self):
        # Configurar autenticação do Google Sheets e Calendar
        self.creds = None
        if os.path.exists(TOKEN_FILE):
            self.creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                self.creds = flow.run_local_server(port=0)
            with open(TOKEN_FILE, 'w') as token:
                token.write(self.creds.to_json())

        self.service = build('sheets', 'v4', credentials=self.creds)
        self.calendar_service = build('calendar', 'v3', credentials=self.creds)

    def init_speech_engine(self):
        # Inicializar mecanismo de síntese de fala
        self.engine = pyttsx3.init()

    def load_schedule_data(self):
        # Carregar dados da planilha e mostrar na tabela
        sheet = self.service.spreadsheets()
        result = sheet.values().get(spreadsheetId=config["SAMPLE_SPREADSHEET_ID"], range=config["RANGES"][0]).execute()
        values = result.get('values', [])

        team_filter = self.team_filter.currentText()
        self.table.setRowCount(0)
        for row in values:
            if len(row) < 3:
                continue
            date, team, songs = row[0], row[1], row[2]
            if team_filter == "Todas" or team_filter == team:
                self.add_table_row(date, team, songs)

    def add_table_row(self, date, team, songs):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        self.table.setItem(row_position, 0, QTableWidgetItem(date))
        self.table.setItem(row_position, 1, QTableWidgetItem(team))
        self.table.setItem(row_position, 2, QTableWidgetItem(songs))

    def send_schedule(self):
        # Enviar escala para os membros via Google Calendar e Email
        team_filter = self.team_filter.currentText()
        if team_filter == "Todas":
            self.speak("Por favor, selecione uma equipe para enviar a escala.")
            return

        # Carregar emails da configuração
        if os.path.exists(EMAIL_CONFIG_FILE):
            with open(EMAIL_CONFIG_FILE, "r") as file:
                emails = json.load(file)
        else:
            self.speak("Nenhum email encontrado para enviar a escala.")
            return

        # Lógica para enviar a escala aos emails (simplificado)
        for email in emails:
            self.send_email(email)
        self.speak(f"Escala enviada com sucesso para a equipe {team_filter}.")

    def send_email(self, email):
        # Lógica para enviar email (esqueleto)
        print(f"Enviando email para {email}")

    def create_csv_backup(self):
        # Criar backup dos dados em CSV
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Salvar Backup em CSV", "", "CSV Files (*.csv)", options=options)
        if file_name:
            with open(file_name, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Data", "Equipe", "Músicas"])
                for row in range(self.table.rowCount()):
                    writer.writerow([
                        self.table.item(row, 0).text(),
                        self.table.item(row, 1).text(),
                        self.table.item(row, 2).text()
                    ])
            self.speak("Backup criado com sucesso.")

    def create_new_schedule(self):
        # Lógica para criar uma nova escala no Google Sheets
        sheet = self.service.spreadsheets()
        date = "Nova Data"
        team = "Nova Equipe"
        songs = "Novas Músicas"
        values = [[date, team, songs]]
        body = {
            'values': values
        }
        result = sheet.values().append(
            spreadsheetId=config["SAMPLE_SPREADSHEET_ID"],
            range=config["RANGES"][0],
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body=body
        ).execute()
        self.speak("Nova escala criada com sucesso.")
        self.load_schedule_data()

    def open_config_page(self):
        self.config_page = ConfigPage()
        self.config_page.show()

    def speak(self, text):
        self.engine.say(text)
        self.engine.runAndWait()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
