### **Documentação do Projeto: Assistente de Agendas para o Ministério de Louvor**

---

## **Visão Geral**

Este projeto tem como objetivo criar uma ferramenta de assistente de agenda para um ministério de louvor. O assistente acessa uma planilha do Google Sheets com as escalas do ministério, identifica a próxima escala e informa o tempo restante até o próximo evento. O projeto utiliza Python para acessar os dados do Google Sheets, processar essas informações e comunicar os detalhes via console, utilizando o módulo `pyttsx3` para síntese de voz.

---

## **Instalação**

### **1. Requisitos**

Antes de começar, você precisa garantir que o Python esteja instalado em sua máquina. Este projeto foi desenvolvido e testado com Python 3.10, mas versões anteriores (Python 3.x) também devem funcionar.

### **2. Clonando o Repositório**

Clone o repositório do projeto para sua máquina local:
```bash
git clone <URL do repositório>
```

### **3. Instalando Dependências**

Navegue até o diretório do projeto e instale as dependências necessárias utilizando `pip`:
```bash
cd assistente-de-agenda
pip install -r requirements.txt
```

O arquivo `requirements.txt` deve conter as seguintes bibliotecas:
- `google-auth`
- `google-auth-oauthlib`
- `google-auth-httplib2`
- `google-api-python-client`
- `pyttsx3`

### **4. Configuração das Credenciais do Google Sheets**

Para acessar o Google Sheets, é necessário configurar as credenciais:

1. Acesse o [Google Cloud Console](https://console.cloud.google.com/).
2. Crie um novo projeto ou utilize um existente.
3. Habilite a API do Google Sheets.
4. Na seção “Credenciais”, crie uma credencial de tipo `OAuth 2.0 Client ID`.
5. Baixe o arquivo `credentials.json` e mova-o para o diretório do projeto.

### **5. Primeiro Uso: Gerando o Token**

Na primeira execução, o programa irá gerar um arquivo `token.json` para armazenar as credenciais de acesso. Execute o script para iniciar o processo de autenticação:

```bash
python Monitoring.py
```

Isso abrirá uma janela de navegador solicitando que você faça login na sua conta do Google e permita o acesso ao Google Sheets. Após a autenticação, o arquivo `token.json` será criado.

---

## **Estrutura do Projeto**

```
assistente-de-agenda/
│
├── Monitoring.py        # Arquivo principal do projeto
├── credentials.json     # Credenciais de acesso ao Google API
└── token.json           # Token de autenticação (gerado automaticamente)
```

### **Monitoring.py**

O arquivo `Monitoring.py` contém toda a lógica do projeto. Abaixo está uma explicação detalhada de cada parte do código.

### **Imports Necessários**

```python
import os.path
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pyttsx3 as sx
```

- **`os.path`**: Para manipulação de caminhos e verificação da existência de arquivos.
- **`datetime`**: Para manipulação e comparação de datas.
- **`google.auth` e `googleapiclient`**: Para autenticação e acesso ao Google Sheets.
- **`pyttsx3`**: Para síntese de voz.

### **Constantes**

```python
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SAMPLE_SPREADSHEET_ID = "1SjRO_pxgDIq25SZhdIq0wBZb4CQZdvuDLBzJfs-n0Y4"
RANGES = ["Átrio ' 24 | Escala!A1:C44"]
```

- **`SCOPES`**: Define o escopo de acesso à API do Google Sheets, limitado à leitura.
- **`SAMPLE_SPREADSHEET_ID`**: ID da planilha a ser acessada.
- **`RANGES`**: Intervalo de células na planilha que contém as escalas do ministério.

### **Função `speak`**

```python
def speak(text):
    a = sx.init()
    a.say(text)
    a.runAndWait()
```

- **`speak`**: Utiliza o módulo `pyttsx3` para sintetizar e pronunciar o texto passado como argumento.

### **Função `get_data`**

```python
def get_data():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
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
                    if "Átrio Music" in equipe:
                        all_data.append({"data": date_str, "equipe": equipe, "culto": culto})

        return all_data

    except HttpError as err:
        print(err)
        return []
```

- **Autenticação**: Carrega as credenciais do arquivo `token.json` ou solicita uma nova autenticação se necessário.
- **Acesso ao Google Sheets**: Utiliza as credenciais para acessar os dados do Google Sheets.
- **Processamento de Dados**: Filtra as linhas da planilha que correspondem à equipe “Átrio Music” e retorna uma lista de dicionários com as informações relevantes.

### **Função `find_next_scale`**

```python
def find_next_scale(scales):
    today = datetime.now()
    upcoming_scales = []

    for scale in scales:
        scale_date = datetime.strptime(scale['data'], "%d/%m/%Y")
        if scale_date >= today:
            upcoming_scales.append(scale)

    if not upcoming_scales:
        return None

    next_scale = min(upcoming_scales, key=lambda x: datetime.strptime(x['data'], "%d/%m/%Y"))
    time_remaining = (datetime.strptime(next_scale['data'], "%d/%m/%Y") - today).days

    return next_scale, time_remaining
```

- **`find_next_scale`**: Compara as datas das escalas com a data atual para encontrar a próxima escala. Calcula e retorna o tempo restante até essa escala.

### **Função `main`**

```python
def main():
    data = get_data()

    next_scale, time_remaining = find_next_scale(data) if data else (None, None)

    speak("Exibindo as próximas escalas de Átrio Music")
    print("---------------------------------")
    
    if data:
        for entry in data:
            speak(f"Data: {entry['data']}, Equipe: {entry['equipe']}, Culto: {entry['culto']}")
    else:
        speak("Nenhum dado encontrado.")

    if next_scale:
        speak("\nPróxima Escala:")
        speak(f"Data: {next_scale['data']}")
        speak(f"Equipe: {next_scale['equipe']}")
        speak(f"Culto: {next_scale['culto']}")
        speak(f"Faltam {time_remaining} dias para a próxima escala.")
    else:
        speak("\nNenhuma escala futura encontrada.")
```

- **`main`**: Integra as funções anteriores para obter os dados, identificar a próxima escala e fornecer feedback via console e síntese de voz.

### **Execução do Programa**

```python
if __name__ == "__main__":
    main()
```

- Este bloco garante que o código seja executado apenas se o script for executado diretamente.

---

## **Conclusão**

Este projeto fornece uma solução simples e eficaz para gerenciar e monitorar as escalas do ministério de louvor diretamente no terminal. A ferramenta utiliza Python para acessar dados em tempo real do Google Sheets e comunica as informações de forma clara e acessível.

---

## **Futuras Melhorias**

- **Interface Gráfica**: Implementar uma interface gráfica para visualização e interação mais amigável.
- **Integração com o Google Calendar**: Adicionar a funcionalidade para sincronizar automaticamente as escalas com o Google Calendar.
- **Notificações**: Incluir alertas via e-mail ou notificações móveis para avisar sobre escalas futuras.

---

**Autor:** [Seu Nome]  
**Contato:** [Seu E-mail]  
**Data:** [Data de Criação]