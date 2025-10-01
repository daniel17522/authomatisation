# school_autho_egor.py — ALL-IN-ONE, чистая версия

import os, re, json, base64, pickle, requests, builtins, logging, logging.handlers, sys
from pathlib import Path
from io import BytesIO
from datetime import datetime, timedelta

# ---------- ENV ----------
from dotenv import load_dotenv

BASE_DIR = Path(__file__).parent
load_dotenv(BASE_DIR / ".env")  # грузим .env рядом со скриптом (если есть)

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
CHAT_IDS       = [s.strip() for s in os.getenv("CHAT_ID", "").split(",") if s.strip()]
TARGET_GOOGLE_ACCOUNT = os.getenv("TARGET_GOOGLE_ACCOUNT", "").strip()
LOOKBACK_DAYS  = int(os.getenv("LOOKBACK_DAYS", "7"))
HEADLESS       = os.getenv("HEADLESS", "0") == "1"

REPROCESS_ALL   = os.getenv("REPROCESS_ALL", "0") == "1"
RESET_PROCESSED = os.getenv("RESET_PROCESSED","0") == "1"
MAX_MESSAGES    = int(os.getenv("MAX_MESSAGES", "0"))

TOKEN_DIR = BASE_DIR / "tokens"
TOKEN_DIR.mkdir(exist_ok=True)

# ---------- LOGGING (ротация + перехват print) ----------
LOG_DIR   = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE  = LOG_DIR / "school.log"
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()

def _setup_logging():
    logger = logging.getLogger("school")
    logger.setLevel(getattr(logging, LOG_LEVEL, logging.INFO))
    fmt = logging.Formatter("%(asctime)s | %(levelname)-8s | %(name)s | %(message)s", "%Y-%m-%d %H:%M:%S")
    fh = logging.handlers.TimedRotatingFileHandler(LOG_FILE, when="midnight", interval=1, backupCount=14, encoding="utf-8")
    fh.setFormatter(fmt); logger.addHandler(fh)
    ch = logging.StreamHandler(); ch.setFormatter(fmt); logger.addHandler(ch)
    builtins._orig_print = print
    def _dual_print(*args, **kwargs):
        msg = " ".join(str(a) for a in args)
        logger.info(msg); builtins._orig_print(*args, **kwargs)
    builtins.print = _dual_print
    return logger
logger = _setup_logging()

def _tail(s: str, n: int = 6) -> str: return s[-n:] if s else ""
def validate_env():
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY пуст — добавь в .env или Env Variables")
    print(f"🔑 OPENAI_API_KEY: ...{_tail(OPENAI_API_KEY)}")
    if TELEGRAM_TOKEN: print(f"🤖 TELEGRAM_TOKEN: ...{_tail(TELEGRAM_TOKEN)}")
    if CHAT_IDS: print(f"📨 CHAT_ID(s): {', '.join(CHAT_IDS)}")
validate_env()

# ---------- Gmail/Calendar scopes ----------
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/calendar.events",
]
CREDENTIALS_PATH = BASE_DIR / "credentials.json"
PROCESSED_IDS_PATH = BASE_DIR / "processed_ids.json"

# ---------- School filters ----------
SCHOOL_DOMAINS = [
    "schools.nyc.gov",
    "ps53.org",
]
SCHOOL_KEYWORDS = [
    "p.s. 53", "ps 53", "public school 53",
    "newcomer", "newcomer's tea", "town hall", "breakfast bowl",
    "kindergarten", "pk", "enl", "pta",
    "principal", "assistant principal", "teacher",
]

# ---------- OpenAI ----------
from openai import OpenAI
oai = OpenAI(api_key=OPENAI_API_KEY)
OPENAI_CHAT_MODEL = "gpt-4o-mini"

# ---------- Google API ----------
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ---------- Attachments parsing ----------
from PyPDF2 import PdfReader
from docx import Document
from PIL import Image
import pytesseract

try:
    from pdf2image import convert_from_bytes
    HAS_PDF2IMAGE = True
except Exception:
    HAS_PDF2IMAGE = False

# Windows zoneinfo fallback
try:
    from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
    def get_tz(tz_name: str):
        try:
            return ZoneInfo(tz_name)
        except ZoneInfoNotFoundError:
            from dateutil.tz import gettz
            return gettz(tz_name)
except Exception:
    from dateutil.tz import gettz as _gettz
    def get_tz(tz_name: str):
        return _gettz(tz_name)

# ---------- Helpers ----------
USE_VISION_FALLBACK = False
MIN_TEXT_LEN_FOR_OK = 50
VISION_MODEL_FALLBACK = OPENAI_CHAT_MODEL

def _b64url_decode(s: str) -> bytes:
    s += "=" * (-len(s) % 4); return base64.urlsafe_b64decode(s.encode("utf-8"))

def _html_to_text(html: str) -> str:
    html = re.sub(r"(?is)<(script|style).*?>.*?</\1>", "", html)
    html = re.sub(r"(?is)<br\s*/?>", "\n", html)
    html = re.sub(r"(?is)</p\s*>", "\n\n", html)
    text = re.sub(r"(?is)<.*?>", "", html)
    return re.sub(r"[ \t]+\n", "\n", text).strip()

def _extract_text_from_payload(payload) -> str | None:
    mime = payload.get("mimeType", "") or ""
    body = payload.get("body", {}) or {}
    data = body.get("data")

    if mime.startswith("text/plain") and data:
        return _b64url_decode(data).decode("utf-8", errors="replace")
    if mime.startswith("text/html") and data:
        html = _b64url_decode(data).decode("utf-8", errors="replace")
        return _html_to_text(html)
    for part in payload.get("parts", []) or []:
        txt = _extract_text_from_payload(part)
        if txt: return txt
    return None

# ---------- Gmail ----------
def _token_path_for(email_hint: str | None) -> Path:
    if email_hint:
        safe = email_hint.replace("@", "_at_")
        return TOKEN_DIR / f"token_{safe}.pkl"
    return BASE_DIR / "token.pkl"

def gmail_auth():
    token_path = _token_path_for(TARGET_GOOGLE_ACCOUNT or None)

    creds = None
    if os.path.exists(token_path):
        with open(token_path, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # В CI/Actions НЕЛЬЗЯ начинать интерактивный OAuth — там не будет браузера.
            if os.environ.get("HEADLESS") == "1":
                raise RuntimeError(
                    f"No valid Gmail token at {token_path}. "
                    "Upload your pickled token (*.pkl) via GMAIL_TOKEN_B64 secret."
                )
            # Локальный интерактивный OAuth (только на своей машине)
            from google_auth_oauthlib.flow import InstalledAppFlow
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)

        # на всякий случай сохраняем обновлённые креды
        os.makedirs(os.path.dirname(token_path), exist_ok=True)
        with open(token_path, "wb") as f:
            pickle.dump(creds, f)

    return build("gmail", "v1", credentials=creds)


def build_school_gmail_query() -> str:
    after_date = (datetime.utcnow() - timedelta(days=LOOKBACK_DAYS)).strftime("%Y/%m/%d")
    dom_q  = " OR ".join([f'from:({d})' for d in SCHOOL_DOMAINS])
    kw_q   = " OR ".join([f'"{k}"' for k in SCHOOL_KEYWORDS])  # ищем везде (без subject:)
    q = f'in:anywhere after:{after_date} ({dom_q} OR ({kw_q}))'
    return q

def iter_message_ids(service, q: str, page_size: int = 50, max_pages: int = 10):
    page_token = None; pages = 0
    while True:
        resp = service.users().messages().list(userId="me", q=q, maxResults=page_size, pageToken=page_token).execute()
        for m in resp.get("messages", []) or []:
            yield m["id"]
        page_token = resp.get("nextPageToken"); pages += 1
        if not page_token or pages >= max_pages: break

def get_email_with_attachments_by_id(service, msg_id: str):
    msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
    payload = msg.get("payload", {}) or {}
    headers = {h["name"].lower(): h["value"] for h in payload.get("headers", [])}
    subject = headers.get("subject", ""); sender = headers.get("from", "")
    body_text = _extract_text_from_payload(payload) or msg.get("snippet","") or " "
    attachments = []
    def walk(p):
        for part in (p.get("parts") or []):
            filename = (part.get("filename") or "").strip()
            mime = part.get("mimeType") or ""
            body = part.get("body", {}) or {}; att_id = body.get("attachmentId")
            if filename and att_id:
                att = service.users().messages().attachments().get(userId="me", messageId=msg["id"], id=att_id).execute()
                data_b64 = att.get("data",""); data_b64 += "=" * (-len(data_b64) % 4)
                content = base64.urlsafe_b64decode(data_b64.encode("utf-8"))
                attachments.append({"filename": filename, "mime": mime, "content": content})
            if part.get("parts"): walk(part)
    walk(payload)
    return {"subject": subject, "from": sender, "body": body_text, "attachments": attachments}

# ---------- classify "school email" ----------
import email.utils
def _addr_domain(addr: str) -> str:
    _, email_addr = email.utils.parseaddr(addr)
    return email_addr.split("@",1)[1].lower() if "@" in email_addr else ""

def is_school_email(email_obj: dict) -> bool:
    sender = (email_obj.get("from") or "").lower()
    subject = (email_obj.get("subject") or "").lower()
    body    = (email_obj.get("body") or "").lower()
    dom = _addr_domain(sender)
    if dom and any(dom.endswith(d) for d in SCHOOL_DOMAINS): return True
    txt = subject + "\n" + body
    return any(k.lower() in txt for k in SCHOOL_KEYWORDS)

# ---------- attachments → text ----------
def ocr_tesseract(img_bytes: bytes, lang: str = "eng") -> str:
    try:
        img = Image.open(BytesIO(img_bytes)).convert("L")
        return pytesseract.image_to_string(img, lang=lang)
    except Exception:
        return ""

def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    try:
        reader = PdfReader(BytesIO(pdf_bytes))
        text = "\n".join([(p.extract_text() or "") for p in reader.pages]).strip()
        if text: return text
    except Exception:
        pass
    if HAS_PDF2IMAGE:
        try:
            images = convert_from_bytes(pdf_bytes)
            chunks = []
            for img in images:
                buf = BytesIO(); img.save(buf, format="PNG")
                chunks.append(ocr_tesseract(buf.getvalue(), lang="eng"))
            text = "\n".join(chunks).strip()
            if len(text) >= MIN_TEXT_LEN_FOR_OK: return text
        except Exception:
            pass
    return ""

def extract_text_from_attachment(att: dict) -> str:
    name = att["filename"].lower(); mime = att["mime"]; data = att["content"]
    if mime == "application/pdf" or name.endswith(".pdf"):
        return extract_text_from_pdf_bytes(data)
    if (mime.startswith("application/vnd.openxmlformats") or name.endswith(".docx") or mime == "application/msword"):
        try:
            doc = Document(BytesIO(data))
            return "\n".join(p.text for p in doc.paragraphs).strip()
        except Exception:
            return ""
    if mime.startswith("image/") or name.endswith((".png",".jpg",".jpeg",".webp")):
        return ocr_tesseract(data, lang="eng")
    return ""

def build_email_text(email_obj: dict) -> str:
    parts = []
    body = (email_obj.get("body") or "").strip()
    if body: parts.append(body)
    for att in email_obj.get("attachments", []):
        txt = extract_text_from_attachment(att)
        if txt: parts.append(f"[Attachment: {att['filename']}]\n{txt}")
    text = "\n\n".join(p for p in parts if p).strip()
    return text[:100_000] if text else " "

# ---------- OpenAI parse ----------
def parse_event(text: str) -> dict:
    system = (
        "Ты извлекаешь событие из письма/вложений. Верни ТОЛЬКО JSON-объект с ключами: "
        "title, place, desc (≤200), start, end, price. "
        "Время: YYYY-MM-DDTHH:MM[:SS] без таймзоны. Если чего-то нет — пустая строка."
    )
    user = f"Исходный текст:\n{text}"
    try:
        resp = oai.chat.completions.create(
            model=OPENAI_CHAT_MODEL,
            messages=[{"role":"system","content":system},{"role":"user","content":user}],
            temperature=0,
            response_format={"type":"json_object"},
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        logger.exception("OpenAI parse error")
        data = {}

    now = datetime.utcnow()
    def_start = (now + timedelta(days=1)).replace(hour=10, minute=0, second=0, microsecond=0)
    def_end   = def_start + timedelta(hours=2)

    def _s(v, d=""): return (v or d).strip() if isinstance(v,str) else d
    title=_s(data.get("title"),"Событие"); place=_s(data.get("place")); desc=_s(data.get("desc"))[:200]
    start_s=_s(data.get("start")); end_s=_s(data.get("end")); price=_s(data.get("price"))

    def _parse_dt(s):
        for f in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M"):
            try: return datetime.strptime(s,f)
            except ValueError: pass
        return None

    sdt = _parse_dt(start_s) or def_start
    edt = _parse_dt(end_s) or def_end
    return {"title":title,"place":place,"desc":desc,"start":sdt.isoformat(),"end":edt.isoformat(),"price":price}

# ---------- normalize dates ----------
def _parse_iso(dt_s: str):
    if not dt_s: return None
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M"):
        try: return datetime.strptime(dt_s, fmt)
        except ValueError: pass
    try: return datetime.fromisoformat(dt_s)
    except Exception: return None

def normalize_event_times(event_data: dict, default_hours: int = 2) -> dict:
    now = datetime.utcnow()
    start = _parse_iso(event_data.get("start")) or (now + timedelta(days=1)).replace(hour=10,minute=0,second=0,microsecond=0)
    end   = _parse_iso(event_data.get("end"))
    if (not end) or (end <= start) or ((end - start).total_seconds() > 7*24*3600):
        end = start + timedelta(hours=default_hours)
    event_data["start"]=start.isoformat(); event_data["end"]=end.isoformat()
    return event_data

# ---------- Calendar ----------
def event_exists(calendar_service, summary: str, start_iso: str) -> bool:
    s = datetime.fromisoformat(start_iso)
    time_min = (s - timedelta(minutes=15)).isoformat() + "Z"
    time_max = (s + timedelta(minutes=15)).isoformat() + "Z"
    resp = calendar_service.events().list(calendarId="primary", q=summary, timeMin=time_min,timeMax=time_max, singleEvents=True, orderBy="startTime").execute()
    for it in (resp.get("items") or []):
        if (it.get("summary","").strip() == summary.strip()):
            return True
    return False

def create_event(service, event_data):
    calendar_service = build("calendar", "v3", credentials=service._http.credentials)
    if event_exists(calendar_service, event_data["title"], event_data["start"]):
        print(f"↪️ Уже есть событие: {event_data['title']} @ {event_data['start']} — пропускаю")
        return
    event = {
        "summary": event_data["title"],
        "location": event_data.get("place",""),
        "description": event_data.get("desc",""),
        "start": {"dateTime": event_data["start"], "timeZone": "America/New_York"},
        "end":   {"dateTime": event_data["end"],   "timeZone": "America/New_York"},
    }
    calendar_service.events().insert(calendarId="primary", body=event).execute()
    print("✅ Событие добавлено в Google Calendar")

# ---------- Telegram ----------
_RU_WD  = ["Пн","Вт","Ср","Чт","Пт","Сб","Вс"]
_RU_MON = ["янв","фев","мар","апр","май","июн","июл","авг","сен","окт","ноя","дек"]
def _part_of_day(hour: int) -> tuple[str, str]:
    if 5 <= hour < 12:  return ("утро","🌅")
    if 12 <= hour < 17: return ("днём","☀️")
    if 17 <= hour < 22: return ("вечер","🌆")
    return ("ночью","🌙")

def _pretty_dt_range(start_iso: str, end_iso: str, tz_name: str = "America/New_York") -> str:
    tz = get_tz(tz_name)
    s = datetime.fromisoformat(start_iso); e = datetime.fromisoformat(end_iso)
    s = s.replace(tzinfo=tz) if s.tzinfo is None else s.astimezone(tz)
    e = e.replace(tzinfo=tz) if e.tzinfo is None else e.astimezone(tz)
    wd, mon = _RU_WD[s.weekday()], _RU_MON[s.month-1]
    tod, emoji = _part_of_day(s.hour)
    return f"{wd}, {s.day} {mon} {s.year}, {s:%H:%M}–{e:%H:%M} ({tod}) {emoji}"

from requests.exceptions import RequestException
def send_telegram(event_data):
    if not TELEGRAM_TOKEN or not CHAT_IDS:
        print("ℹ️ Telegram отключён (нет токена или chat_id)")
        return
    when_line = _pretty_dt_range(event_data["start"], event_data["end"], tz_name="America/New_York")
    place = (event_data.get("place") or "—").strip()
    desc  = (event_data.get("desc")  or "—").strip()
    if len(desc)>180: desc = desc[:177] + "…"
    price = (event_data.get("price") or "—").strip()
    text = f"📌 {event_data['title']}\n🗓 {when_line}\n📍 {place}\nℹ️ {desc}\n💵 {price}"
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    for chat_id in CHAT_IDS:
        payload = {"chat_id": chat_id, "text": text}
        try:
            r = requests.post(url, data=payload, timeout=15)
            if r.status_code == 200:
                print(f"✅ Telegram → {chat_id}")
            else:
                logger.warning("Telegram HTTP %s: %s", r.status_code, r.text[:200])
        except RequestException:
            logger.exception("Telegram: не удалось отправить сообщение (%s)", chat_id)

# ---------- processed_ids ----------
def load_processed_ids() -> set:
    try:
        return set(json.loads(PROCESSED_IDS_PATH.read_text(encoding="utf-8")))
    except Exception:
        return set()

def save_processed_ids(s: set):
    try:
        PROCESSED_IDS_PATH.write_text(json.dumps(sorted(list(s))[-5000:]), encoding="utf-8")
        logger.debug("💾 processed_ids: %d", len(s))
    except Exception:
        logger.exception("Не удалось сохранить processed_ids")

# ---------- отладочный список последних писем (по желанию) ----------
def debug_list_recent(service, limit: int = 15):
    resp = service.users().messages().list(userId="me", maxResults=limit).execute()
    ids = [m["id"] for m in resp.get("messages", [])]
    print(f"🧪 Последние {len(ids)} писем в ящике:")
    for mid in ids:
        msg = service.users().messages().get(userId="me", id=mid, format="metadata", metadataHeaders=["From","Subject","Date"]).execute()
        hdrs = {h["name"]: h["value"] for h in msg.get("payload", {}).get("headers", [])}
        print(f"- From: {hdrs.get('From','')} | Subject: {hdrs.get('Subject','')}")

# ---------- main (один проход) ----------
def run_once():
    processed_ids = load_processed_ids()
    if RESET_PROCESSED:
        processed_ids = set()
        logger.info("🧹 История processed_ids очищена по флагу")

    gmail_service = gmail_auth()
    query = build_school_gmail_query()
    logger.info("🔎 Gmail query: %s", query)

    # debug_list_recent(gmail_service, 15)  # при необходимости

    processed = 0; matched = 0; count = 0
    for msg_id in iter_message_ids(gmail_service, query, page_size=25, max_pages=10):
        if not REPROCESS_ALL and msg_id in processed_ids:
            continue
        email_obj = get_email_with_attachments_by_id(gmail_service, msg_id)
        if not is_school_email(email_obj):
            continue
        matched += 1

        text_for_llm = build_email_text(email_obj)
        event_data   = normalize_event_times(parse_event(text_for_llm))
        logger.info("📩 Распарсенные данные: %s", event_data)

        create_event(gmail_service, event_data)
        send_telegram(event_data)
        processed += 1

        processed_ids.add(msg_id)
        count += 1
        if MAX_MESSAGES and count >= MAX_MESSAGES:
            logger.info("⏸ Достигнут лимит MAX_MESSAGES=%d", MAX_MESSAGES)
            break

    save_processed_ids(processed_ids)

    if matched == 0:
        logger.info("ℹ️ Писем школы не найдено по текущим правилам.")
    else:
        logger.info("🏁 Найдено писем школы: %d. Создано/обновлено: %d.", matched, processed)

if __name__ == "__main__":
    try:
        # CLI-флажки (удобно иногда)
        for arg in sys.argv[1:]:
            if arg == "--reprocess": REPROCESS_ALL = True
            if arg == "--reset-processed": RESET_PROCESSED = True
            if arg.startswith("--max="): MAX_MESSAGES = int(arg.split("=",1)[1])

        run_once()
    except Exception:
        logger.exception("Фатальная ошибка выполнения скрипта")
        raise
