# Documentação do Projeto: Assistente de Agenda para Escalas do Ministério de Louvor

Este projeto visa desenvolver um assistente de agenda em Python que ajuda a gerenciar as escalas do Ministério de Louvor. O programa se conecta a uma planilha do Google Sheets, onde as informações das escalas estão armazenadas, e fornece dados sobre as próximas escalas, incluindo a equipe responsável e o culto associado.

# Requisitos
Python 3.x
Bibliotecas Python:
google-auth
google-auth-oauthlib
google-api-python-client
pyttsx3
Instalação
Instale as bibliotecas necessárias usando pip:

´´´bash
pip install google-auth google-auth-oauthlib google-api-python-client pyttsx3
´´´
# Credenciais do Google API:

Crie um projeto no Google Cloud Console.
Ative a Google Sheets API.
Crie credenciais do tipo "OAuth 2.0 Client ID" e baixe o arquivo JSON (denominado credentials.json).
Salve este arquivo no diretório do seu projeto.
Token de Autenticação:

Na primeira execução do programa, um arquivo token.json será gerado, armazenando as credenciais de autenticação.
Estrutura do Código
python
```
import os.path
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pyttsx3 as sx

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SAMPLE_SPREADSHEET_ID = "1SjRO_pxgDIq25SZhdIq0wBZb4CQZdvuDLBzJfs-n0Y4"
RANGES = ["Átrio ' 24 | Escala!A1:C44"]

def speak(text):
    a = sx.init()
    a.say(text)
    a.runAndWait()

def get_data():
    ...
    
def find_next_scale(scales):
    ...

def main():
    ...
    
if __name__ == "__main__":
    main()
```
# Funções Principais
```speak(text)```: Utiliza a biblioteca pyttsx3 para converter texto em fala, anunciando informações no console.
```get_data()```: Obtém dados das escalas do Google Sheets. Filtra as informações da equipe "Átrio Music".
```find_next_scale(scales)```: Compara a data atual com as escalas futuras e determina a próxima escala e quanto tempo falta para ela.
```main()```: Função principal que orquestra a execução do programa, chamando as outras funções e exibindo as informações no terminal.


# Funcionalidades


Conexão com Google Sheets: O programa lê as escalas diretamente de uma planilha do Google.
Filtragem de Dados: Filtra as escalas para a equipe "Átrio Music".
Voz Sintetizada: Anuncia as próximas escalas através de síntese de voz.
Cálculo do Tempo Restante: Informa quantos dias faltam para a próxima escala.
Como Executar o Projeto
Salve o código em um arquivo Python, por exemplo, assistente_agenda.py.
No terminal, navegue até o diretório onde o arquivo está salvo e execute:

```bash
python assistente_agenda.py
```
Siga as instruções no terminal para autenticar sua conta do Google se necessário.
Exemplos de Uso
Ao executar o programa, ele anunciará as próximas escalas de "Átrio Music" e dirá quantos dias faltam para a próxima escala.
Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para enviar um pull request ou relatar problemas no repositório.

Licença
Este projeto está licenciado sob a MIT License - veja o arquivo LICENSE para mais detalhes.