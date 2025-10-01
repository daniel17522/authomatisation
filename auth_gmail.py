# auth_gmail.py — единоразовая локальная авторизация и сохранение токена *.pkl
import os, pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/calendar.events",
]

BASE_DIR = os.path.dirname(__file__)
CREDENTIALS_PATH = os.path.join(BASE_DIR, "credentials.json")
TOKENS_DIR = os.path.join(BASE_DIR, "tokens")
os.makedirs(TOKENS_DIR, exist_ok=True)

def main():
    creds = None
    token_path = os.path.join(TOKENS_DIR, "token_TEMP.pkl")  # временно
    if os.path.exists(token_path):
        with open(token_path, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Откроет локальный браузер — это делаем ТОЛЬКО локально, не в Actions
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_path, "wb") as f:
            pickle.dump(creds, f)

    # Узнаем реальный email и переименуем файл под него
    gmail = build("gmail", "v1", credentials=creds)
    profile = gmail.users().getProfile(userId="me").execute()
    email_addr = profile.get("emailAddress", "").strip()
    final = os.path.join(TOKENS_DIR, f"token_{email_addr.replace('@','_at_')}.pkl")
    if final != token_path:
        os.replace(token_path, final)
    print("✅ Token saved:", final)

if __name__ == "__main__":
    main()
