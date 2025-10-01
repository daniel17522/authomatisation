# auth_gmail.py
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

flow = InstalledAppFlow.from_client_secrets_file(
    CREDENTIALS_PATH, SCOPES
)
creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")

gmail = build("gmail", "v1", credentials=creds)
email_addr = gmail.users().getProfile(userId="me").execute()["emailAddress"]
final = os.path.join(TOKENS_DIR, f"token_{email_addr.replace('@','_at_')}.pkl")
with open(final, "wb") as f:
    pickle.dump(creds, f)
print("✅ Saved:", final)
