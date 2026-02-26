import os
import re
import time
from datetime import datetime

from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

from google.oauth2 import service_account
from googleapiclient.discovery import build

from docx import Document
from dotenv import load_dotenv

# ========================
# CONFIG
# ========================

load_dotenv()

SLACK_BOT_TOKEN = os.getenv("SLACK_BOT_TOKEN")
GOOGLE_FOLDER_ID = "1gfPpYmWNc8cnw8D-aRLcQHFGZCarpNgM"

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

MESSAGE_TEXT = (
    "Oiii, tudo bem? 😊\n\n"
    "Estamos verificando o seu ponto e identifiquei lançamentos pendentes.\n"
    "Poderia verificar e, se necessário, realizar os lançamentos?\n\n"
    "Lembrando que as batidas pendentes dessa semana deve ser lançada hoje!\n"
    "Se você já ajustou, pode desconsiderar a mensagem, nesse caso, fica pendente apenas a aprovação do seu gestor.\n"
    "Para dúvidas, entre em contato com o time de DP.\n\n"
    "Não responda esta mensagem — digite *DP* para acessar o menu de perguntas."
)

PROCESSED_FILE = "processed_files.txt"

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")

# ========================
# GOOGLE DRIVE
# ========================

def drive_service():
    creds = service_account.Credentials.from_service_account_file(
        "credentials.json", scopes=SCOPES
    )
    return build("drive", "v3", credentials=creds)


def list_files(service):
    results = service.files().list(
        q=f"'{GOOGLE_FOLDER_ID}' in parents and trashed=false",
        fields="files(id, name, mimeType)"
    ).execute()
    return results.get("files", [])


def download_docx(service, file_id, filename):
    request = service.files().get_media(fileId=file_id)
    with open(filename, "wb") as f:
        f.write(request.execute())


# ========================
# FILE CONTROL
# ========================

def load_processed():
    if not os.path.exists(PROCESSED_FILE):
        return set()
    with open(PROCESSED_FILE, "r") as f:
        return set(line.strip() for line in f)


def mark_processed(file_id):
    with open(PROCESSED_FILE, "a") as f:
        f.write(file_id + "\n")


# ========================
# EMAIL EXTRACTION
# ========================

def extract_emails_from_docx(path):
    doc = Document(path)
    text = "\n".join(p.text for p in doc.paragraphs)
    return list(set(EMAIL_REGEX.findall(text)))


# ========================
# SLACK
# ========================

slack = WebClient(token=SLACK_BOT_TOKEN)


def find_user_by_email(email):
    try:
        resp = slack.users_lookupByEmail(email=email)
        return resp["user"]["id"]
    except SlackApiError:
        return None


def send_dm(user_id, text):
    channel = slack.conversations_open(users=user_id)["channel"]["id"]
    slack.chat_postMessage(channel=channel, text=text)


# ========================
# MAIN LOGIC
# ========================

def run():
    print("🔍 Iniciando varredura...")
    drive = drive_service()
    processed = load_processed()
    files = list_files(drive)

    for file in files:
        file_id = file["id"]
        name = file["name"]
        mime = file["mimeType"]

        if file_id in processed:
            continue

        print(f"📄 Processando: {name}")

        emails = []

        if mime == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            os.makedirs("tmp", exist_ok=True)
            tmp = os.path.join("tmp", f"{file_id}.docx")
            download_docx(drive, file_id, tmp)
            emails = extract_emails_from_docx(tmp)

        elif mime == "application/vnd.google-apps.document":
            export = drive.files().export(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ).execute()

            os.makedirs("tmp", exist_ok=True)
            tmp = os.path.join("tmp", f"{file_id}.docx")

            with open(tmp, "wb") as f:
                f.write(export)

                emails = extract_emails_from_docx(tmp)


        print(f"📧 Emails encontrados: {emails}")

        for email in emails:
            user_id = find_user_by_email(email)
            if not user_id:
                print(f"❌ Usuário não encontrado no Slack: {email}")
                continue

            try:
                send_dm(user_id, MESSAGE_TEXT)
                print(f"✅ Mensagem enviada para {email}")
                time.sleep(1)
            except SlackApiError as e:
                print(f"⚠️ Erro ao enviar para {email}: {e}")

        mark_processed(file_id)

    print("✅ Ciclo finalizado", datetime.now())


# ========================
# LOOP (RENDER)
# ========================

if __name__ == "__main__":
    while True:
        run()
        time.sleep(60)
