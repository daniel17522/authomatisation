import os, re, json, base64, pickle, requests, builtins, logging, logging.handlers, sys
from pathlib import Path
from io import BytesIO
from datetime import datetime, timedelta

# ============ ENV ============
from dotenv import load_dotenv

BASE_DIR = Path(__file__).parent
load_dotenv(BASE_DIR / ".env")  # .env рядом со скриптом
LAST_SENT_PATH = BASE_DIR / "last_sent_events.json"

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

# ============ LOGGING (rotating + hijack print) ============
LOG_DIR   = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE  = LOG_DIR / "school.log"
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()


def _setup_logging():
    logger = logging.getLogger("school")
    logger.setLevel(getattr(logging, LOG_LEVEL, logging.INFO))

    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        "%Y-%m-%d %H:%M:%S"
    )

    # файл с ротацией по дню
    fh = logging.handlers.TimedRotatingFileHandler(
        LOG_FILE, when="midnight", interval=1, backupCount=14, encoding="utf-8"
    )
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    # дублируем в консоль
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    # перехватываем print -> logger.info
    builtins._orig_print = print

    def _dual_print(*args, **kwargs):
        msg = " ".join(str(a) for a in args)
        logger.info(msg)
        builtins._orig_print(*args, **kwargs)

    builtins.print = _dual_print
    return logger


logger = _setup_logging()


def _tail(s: str, n: int = 6) -> str:
    return s[-n:] if s else ""


def validate_env():
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY пуст — добавь в .env или Env Variables")

    print(f"🔑 OPENAI_API_KEY: ...{_tail(OPENAI_API_KEY)}")

    if TELEGRAM_TOKEN:
        print(f"🤖 TELEGRAM_TOKEN: ...{_tail(TELEGRAM_TOKEN)}")

    if CHAT_IDS:
        print(f"📨 CHAT_ID(s): {', '.join(CHAT_IDS)}")

    print(
        f"⚙️  FLAGS — REPROCESS_ALL={REPROCESS_ALL}  "
        f"RESET_PROCESSED={RESET_PROCESSED}  "
        f"MAX_MESSAGES={MAX_MESSAGES}"
    )
    print(
        f"📄 PATHS — base={BASE_DIR}  "
        f"processed={BASE_DIR/'processed_ids.json'}  "
        f"state={BASE_DIR/'state.json'}"
    )


validate_env()

# ============ CONSTANTS / PATHS ============

SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",  # ← нужно для меток
    "https://www.googleapis.com/auth/calendar.events",
]

CREDENTIALS_PATH = BASE_DIR / "credentials.json"
PROCESSED_IDS_PATH = BASE_DIR / "processed_ids.json"
STATE_PATH = BASE_DIR / "state.json"

# используем Gmail-метку, чтобы исключать уже обработанные письма между запусками
PROCESSED_LABEL_NAME = "schoolbot/processed"

# фильтр писем школы
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

# ============ OpenAI ============
from openai import OpenAI
oai = OpenAI(api_key=OPENAI_API_KEY)
OPENAI_CHAT_MODEL = "gpt-4o-mini"

# ============ Google API ============
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ============ Attachments parsing ============
from PyPDF2 import PdfReader
from docx import Document
from PIL import Image
import pytesseract

try:
    from pdf2image import convert_from_bytes
    HAS_PDF2IMAGE = True
except Exception:
    HAS_PDF2IMAGE = False

# zoneinfo fallback (Windows compat)
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


# ============ Utils ============
USE_VISION_FALLBACK = False
MIN_TEXT_LEN_FOR_OK = 50
VISION_MODEL_FALLBACK = OPENAI_CHAT_MODEL


def _b64url_decode(s: str) -> bytes:
    s += "=" * (-len(s) % 4)
    return base64.urlsafe_b64decode(s.encode("utf-8"))


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
        if txt:
            return txt
    return None


# ============ Gmail auth / API helpers ============

def _token_path_for(email_hint: str | None) -> Path:
    # отдельный токен-файл на каждый гугл-акк
    if email_hint:
        safe = email_hint.replace("@", "_at_")
        return TOKEN_DIR / f"token_{safe}.pkl"
    return BASE_DIR / "token.pkl"


def gmail_auth():
    """
    Локально (HEADLESS=0):
      - если токен OK -> юзаем
      - если токена нет -> откроется браузер Google OAuth и создаст токен

    В GitHub Actions (HEADLESS=1):
      - токен должен быть уже положен в tokens/...pkl через secrets (мы его просто читаем)
      - если он просрочен и не рефрешится -> упадём (это норм, надо будет обновить токен локально
        и снова загрузить в secrets).
    """
    token_path = _token_path_for(TARGET_GOOGLE_ACCOUNT or None)

    creds = None
    if os.path.exists(token_path):
        with open(token_path, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            # есть refresh_token -> просто обновляем
            creds.refresh(Request())
        else:
            # Нет валидных creds
            if HEADLESS:
                # на сервере / github actions нельзя открыть браузер
                raise RuntimeError(
                    f"No valid Gmail token at {token_path}. "
                    "In headless mode you MUST upload a fresh token via CI secrets."
                )
            # локально запускаем интерактивный OAuth
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_PATH, SCOPES
            )
            creds = flow.run_local_server(port=0)

        # сохранить/обновить токен
        os.makedirs(os.path.dirname(token_path), exist_ok=True)
        with open(token_path, "wb") as f:
            pickle.dump(creds, f)

    # build gmail service
    svc = build("gmail", "v1", credentials=creds)

    # DEBUG: кто мы?
    profile = svc.users().getProfile(userId="me").execute()
    email_addr = profile.get("emailAddress", "")
    logger.info("👤 Google account: %s", email_addr)

    return svc


def get_or_create_label(service, name: str) -> str:
    """Возвращает ID метки Gmail, создаёт если нет."""
    labels = service.users().labels().list(userId="me").execute().get("labels", [])
    for lb in labels:
        if lb.get("name") == name:
            return lb["id"]
    body = {
        "name": name,
        "labelListVisibility": "labelShow",
        "messageListVisibility": "show",
    }
    return service.users().labels().create(userId="me", body=body).execute()["id"]


def add_label_to_message(service, msg_id: str, label_id: str):
    service.users().messages().modify(
        userId="me",
        id=msg_id,
        body={"addLabelIds": [label_id], "removeLabelIds": []},
    ).execute()


def build_school_gmail_query() -> str:
    """
    Ищем письма за последние LOOKBACK_DAYS
    Ищем по доменам ИЛИ по ключевым словам (в тексте).
    Ищем везде (входящие, пересланные и т.д.) — in:anywhere.
    Исключаем письма, уже помеченные нашей меткой.
    """
    after_date = (datetime.utcnow() - timedelta(days=LOOKBACK_DAYS)).strftime("%Y/%m/%d")

    dom_q  = " OR ".join([f'from:({d})' for d in SCHOOL_DOMAINS])
    kw_q   = " OR ".join([f'"{k}"' for k in SCHOOL_KEYWORDS])  # ищем слова где угодно

    q = (
        f'in:anywhere after:{after_date} -label:"{PROCESSED_LABEL_NAME}" '
        f'({dom_q} OR ({kw_q}))'
    )
    return q


def iter_message_ids(service, q: str, page_size: int = 50, max_pages: int = 10):
    """
    Возвращаем только id писем (остальное достанем отдельно).
    """
    page_token = None
    pages = 0
    while True:
        resp = service.users().messages().list(
            userId="me",
            q=q,
            maxResults=page_size,
            pageToken=page_token
        ).execute()

        for m in resp.get("messages", []) or []:
            yield m["id"]

        page_token = resp.get("nextPageToken")
        pages += 1
        if not page_token or pages >= max_pages:
            break


def get_email_with_attachments_by_id(service, msg_id: str):
    """
    Забираем тело письма + мету:
    - threadId
    - internalDate (ms since epoch)
    - subject/from/body
    - attachments as bytes
    """
    msg = service.users().messages().get(
        userId="me",
        id=msg_id,
        format="full"
    ).execute()

    payload = msg.get("payload", {}) or {}
    headers = {
        h["name"].lower(): h["value"]
        for h in payload.get("headers", [])
    }

    subject = headers.get("subject", "")
    sender  = headers.get("from", "")

    body_text = (
        _extract_text_from_payload(payload)
        or msg.get("snippet", "")
        or " "
    )

    thread_id     = msg.get("threadId")
    internal_date = int(msg.get("internalDate", "0"))  # ms since epoch

    attachments = []

    def walk_parts(p):
        for part in (p.get("parts") or []):
            filename = (part.get("filename") or "").strip()
            mime = part.get("mimeType") or ""
            body = part.get("body", {}) or {}
            att_id = body.get("attachmentId")

            if filename and att_id:
                att = service.users().messages().attachments().get(
                    userId="me",
                    messageId=msg["id"],
                    id=att_id
                ).execute()

                data_b64 = att.get("data", "")
                data_b64 += "=" * (-len(data_b64) % 4)
                content = base64.urlsafe_b64decode(data_b64.encode("utf-8"))

                attachments.append({
                    "filename": filename,
                    "mime": mime,
                    "content": content
                })

            if part.get("parts"):
                walk_parts(part)

    walk_parts(payload)

    return {
        "id": msg_id,
        "threadId": thread_id,
        "internalDate": internal_date,
        "subject": subject,
        "from": sender,
        "body": body_text,
        "attachments": attachments,
    }


# ============ classify "school email" ============

import email.utils


def _addr_domain(addr: str) -> str:
    _, email_addr = email.utils.parseaddr(addr)
    if "@" in email_addr:
        return email_addr.split("@", 1)[1].lower()
    return ""


def is_school_email(email_obj: dict) -> bool:
    sender = (email_obj.get("from") or "").lower()
    subject = (email_obj.get("subject") or "").lower()
    body    = (email_obj.get("body") or "").lower()

    dom = _addr_domain(sender)
    if dom and any(dom.endswith(d) for d in SCHOOL_DOMAINS):
        return True

    txt = subject + "\n" + body
    for k in SCHOOL_KEYWORDS:
        if k.lower() in txt:
            return True

    return False


# ============ Attachments -> text ============

def ocr_tesseract(img_bytes: bytes, lang: str = "eng") -> str:
    try:
        img = Image.open(BytesIO(img_bytes)).convert("L")
        return pytesseract.image_to_string(img, lang=lang)
    except Exception:
        return ""


def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    # сначала пытаемся вытащить текст как текст
    try:
        reader = PdfReader(BytesIO(pdf_bytes))
        text = "\n".join([(p.extract_text() or "") for p in reader.pages]).strip()
        if text:
            return text
    except Exception:
        pass

    # если это сканы -> попробуем через OCR
    if HAS_PDF2IMAGE:
        try:
            images = convert_from_bytes(pdf_bytes)
            chunks = []
            for img in images:
                buf = BytesIO()
                img.save(buf, format="PNG")
                chunks.append(ocr_tesseract(buf.getvalue(), lang="eng"))
            text = "\n".join(chunks).strip()
            if text and len(text) >= MIN_TEXT_LEN_FOR_OK:
                return text
        except Exception:
            pass

    return ""


def extract_text_from_attachment(att: dict) -> str:
    name = att["filename"].lower()
    mime = att["mime"]
    data = att["content"]

    if mime == "application/pdf" or name.endswith(".pdf"):
        return extract_text_from_pdf_bytes(data)

    if (
        mime.startswith("application/vnd.openxmlformats")
        or name.endswith(".docx")
        or mime == "application/msword"
    ):
        try:
            doc = Document(BytesIO(data))
            return "\n".join(p.text for p in doc.paragraphs).strip()
        except Exception:
            return ""

    if mime.startswith("image/") or name.endswith((".png", ".jpg", ".jpeg", ".webp")):
        return ocr_tesseract(data, lang="eng")

    return ""


def build_email_text(email_obj: dict) -> str:
    parts = []

    body = (email_obj.get("body") or "").strip()
    if body:
        parts.append(body)

    for att in email_obj.get("attachments", []):
        txt = extract_text_from_attachment(att)
        if txt:
            parts.append(f"[Attachment: {att['filename']}]\n{txt}")

    text = "\n\n".join(p for p in parts if p).strip()
    if not text:
        text = " "

    # safety limit for prompt size
    return text[:100_000]


# ============ OpenAI parse of event ============

def _normalize_year_if_past(dt: datetime) -> datetime:
    """
    Если дата оказалась в прошлом году (например 2023),
    но месяц/день говорят про будущее мероприятие школы сейчас,
    попробуем подтянуть год вперёд до текущего или следующего.
    """
    now = datetime.utcnow()
    # если дата более чем на 60 дней в прошлом от текущего момента -> подними год до сейчас
    if dt < now - timedelta(days=60):
        candidate = dt.replace(year=now.year)
        # если всё ещё сильно в прошлом (типа январь, а сейчас октябрь),
        # можно попробовать +1 год
        if candidate < now - timedelta(days=60):
            candidate = candidate.replace(year=now.year + 1)
        return candidate
    return dt


def parse_event(text: str) -> dict:
    """
    Гоним письмо/вложения в LLM -> ждём JSON с:
      title, place, desc (<=200), start, end, price
    Формат дат: "YYYY-MM-DDTHH:MM[:SS]" без таймзоны.
    """
    system_msg = (
        "Ты извлекаешь событие для календаря из текста письма/объявления. "
        "Ответь ТОЛЬКО валидным JSON с полями: "
        "title, place, desc, start, end, price. "
        "desc максимум 200 символов. "
        "Дата/время в формате YYYY-MM-DDTHH:MM (24h). "
        "Если чего-то нет — делай пустую строку."
    )

    user_msg = f"Исходный текст:\n{text}"

    try:
        resp = oai.chat.completions.create(
            model=OPENAI_CHAT_MODEL,
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user",   "content": user_msg},
            ],
            temperature=0,
            response_format={"type": "json_object"},
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        logger.exception("OpenAI parse error")
        data = {}

    now = datetime.utcnow()
    def_start = (
        now + timedelta(days=1)
    ).replace(hour=10, minute=0, second=0, microsecond=0)
    def_end = def_start + timedelta(hours=2)

    def _s(v, d=""):
        return (v or d).strip() if isinstance(v, str) else d

    title   = _s(data.get("title"), "Событие")
    place   = _s(data.get("place"))
    desc    = _s(data.get("desc"))[:200]
    start_s = _s(data.get("start"))
    end_s   = _s(data.get("end"))
    price   = _s(data.get("price"))

    def _parse_dt(s):
        for f in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M"):
            try:
                return datetime.strptime(s, f)
            except ValueError:
                pass
        return None

    sdt = _parse_dt(start_s) or def_start
    edt = _parse_dt(end_s) or def_end

    # поправка года если llm дал старый (2023 и т.п.)
    sdt = _normalize_year_if_past(sdt)
    edt = _normalize_year_if_past(edt)

    return {
        "title": title,
        "place": place,
        "desc": desc,
        "start": sdt.isoformat(),
        "end":   edt.isoformat(),
        "price": price,
    }


# ============ time normalization again (safety) ============

def _parse_iso(dt_s: str):
    if not dt_s:
        return None
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M"):
        try:
            return datetime.strptime(dt_s, fmt)
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(dt_s)
    except Exception:
        return None


def normalize_event_times(event_data: dict, default_hours: int = 2) -> dict:
    now = datetime.utcnow()

    start = (
        _parse_iso(event_data.get("start"))
        or (now + timedelta(days=1)).replace(
            hour=10, minute=0, second=0, microsecond=0
        )
    )
    end = _parse_iso(event_data.get("end"))

    if (
        not end
        or end <= start
        or ((end - start).total_seconds() > 7 * 24 * 3600)
    ):
        end = start + timedelta(hours=default_hours)

    event_data["start"] = start.isoformat()
    event_data["end"]   = end.isoformat()
    return event_data


# ============ Google Calendar ============

def event_exists(calendar_service, summary: str, start_iso: str) -> bool:
    s = datetime.fromisoformat(start_iso)
    time_min = (s - timedelta(minutes=15)).isoformat() + "Z"
    time_max = (s + timedelta(minutes=15)).isoformat() + "Z"

    resp = calendar_service.events().list(
        calendarId="primary",
        q=summary,
        timeMin=time_min,
        timeMax=time_max,
        singleEvents=True,
        orderBy="startTime",
    ).execute()

    for it in (resp.get("items") or []):
        if (it.get("summary", "").strip() == summary.strip()):
            return True
    return False


def create_event(gmail_service, event_data):
    calendar_service = build(
        "calendar",
        "v3",
        credentials=gmail_service._http.credentials
    )

    if event_exists(calendar_service, event_data["title"], event_data["start"]):
        print(
            f"↪️ Уже есть событие: {event_data['title']} @ {event_data['start']} — пропускаю"
        )
        return

    event = {
        "summary":     event_data["title"],
        "location":    event_data.get("place", ""),
        "description": event_data.get("desc", ""),
        "start": {
            "dateTime": event_data["start"],
            "timeZone": "America/New_York",
        },
        "end": {
            "dateTime": event_data["end"],
            "timeZone": "America/New_York",
        },
    }

    calendar_service.events().insert(
        calendarId="primary",
        body=event
    ).execute()

    print("✅ Событие добавлено в Google Calendar")


# ============ Telegram notify ============

_RU_WD  = ["Пн","Вт","Ср","Чт","Пт","Сб","Вс"]
_RU_MON = ["янв","фев","мар","апр","май","июн","июл","авг","сен","окт","ноя","дек"]


def _part_of_day(hour: int) -> tuple[str, str]:
    if 5 <= hour < 12:
        return ("утро", "🌅")
    if 12 <= hour < 17:
        return ("днём", "☀️")
    if 17 <= hour < 22:
        return ("вечер", "🌆")
    return ("ночью", "🌙")


def _pretty_dt_range(start_iso: str, end_iso: str, tz_name: str = "America/New_York") -> str:
    tz = get_tz(tz_name)

    s = datetime.fromisoformat(start_iso)
    e = datetime.fromisoformat(end_iso)

    # добавляем таймзону если её нет
    s = s.replace(tzinfo=tz) if s.tzinfo is None else s.astimezone(tz)
    e = e.replace(tzinfo=tz) if e.tzinfo is None else e.astimezone(tz)

    wd  = _RU_WD[s.weekday()]
    mon = _RU_MON[s.month - 1]
    tod, emoji = _part_of_day(s.hour)

    # например: "Ср, 1 окт 2025, 09:00–11:00 (утро) 🌅"
    return f"{wd}, {s.day} {mon} {s.year}, {s:%H:%M}–{e:%H:%M} ({tod}) {emoji}"


from requests.exceptions import RequestException


def send_telegram(event_data):
    if not TELEGRAM_TOKEN or not CHAT_IDS:
        print("ℹ️ Telegram отключён (нет токена или chat_id)")
        return

    # === антифлуд по содержимому ===
    last_sent = load_last_sent()
    fp = make_event_fingerprint(event_data)
    now_ts = int(datetime.utcnow().timestamp())

    if fp in last_sent:
        logger.info("⏭ Уже отправлялось в Telegram недавно, скипаю (fp=%s)", fp)
        return

    # формируем текст
    when_line = _pretty_dt_range(
        event_data["start"],
        event_data["end"],
        tz_name="America/New_York"
    )

    place = (event_data.get("place") or "—").strip()
    desc  = (event_data.get("desc")  or "—").strip()
    if len(desc) > 180:
        desc = desc[:177] + "…"
    price = (event_data.get("price") or "—").strip()

    text = (
        f"📌 {event_data['title']}\n"
        f"🗓 {when_line}\n"
        f"📍 {place}\n"
        f"ℹ️ {desc}\n"
        f"💵 {price}"
    )

    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"

    success_any = False
    for chat_id in CHAT_IDS:
        payload = {"chat_id": chat_id, "text": text}
        try:
            r = requests.post(url, data=payload, timeout=15)
            if r.status_code == 200:
                print(f"✅ Telegram → {chat_id}")
                success_any = True
            else:
                logger.warning(
                    "Telegram HTTP %s: %s",
                    r.status_code,
                    r.text[:200]
                )
        except RequestException:
            logger.exception(
                "Telegram: не удалось отправить сообщение (%s)",
                chat_id
            )

    # если хотя бы в один чат успешно ушло — помечаем как отправленное
    if success_any:
        last_sent[fp] = now_ts
        save_last_sent(last_sent)



# ============ processed_ids + state (anti-dup) ============

def load_processed_ids() -> set:
    try:
        return set(
            json.loads(
                PROCESSED_IDS_PATH.read_text(encoding="utf-8")
            )
        )
    except Exception:
        return set()


def save_processed_ids(s: set):
    try:
        # режем до последних 5000 чтобы не пухло
        data = sorted(list(s))[-5000:]
        PROCESSED_IDS_PATH.write_text(
            json.dumps(data, ensure_ascii=False),
            encoding="utf-8"
        )
        logger.debug("💾 processed_ids saved: %d", len(s))
    except Exception:
        logger.exception("Не удалось сохранить processed_ids")

def load_last_sent() -> dict:
    """
    Читаем историю отправленных в ТГ событий.
    Формат:
    {
      "<fingerprint>": 1698534000,  # unix time когда слали
      ...
    }
    """
    try:
        if LAST_SENT_PATH.exists():
            return json.loads(LAST_SENT_PATH.read_text(encoding="utf-8"))
    except Exception:
        logger.exception("Не удалось прочитать last_sent_events.json")
    return {}


def save_last_sent(d: dict):
    try:
        # зачистим старые (>24ч назад)
        now_ts = int(datetime.utcnow().timestamp())
        fresh = {
            k: v for (k, v) in d.items()
            if now_ts - v < 24 * 3600
        }
        LAST_SENT_PATH.write_text(
            json.dumps(fresh, ensure_ascii=False),
            encoding="utf-8"
        )
    except Exception:
        logger.exception("Не удалось сохранить last_sent_events.json")


def make_event_fingerprint(event_data: dict) -> str:
    """
    Устойчивый отпечаток события для антифлуда в Telegram.
    Берём title + start (до минут).
    """
    title = (event_data.get("title") or "").strip()
    start = (event_data.get("start") or "").strip()[:16]
    return f"{title}|{start}"



def load_state() -> dict:
    try:
        if STATE_PATH.exists():
            return json.loads(
                STATE_PATH.read_text(encoding="utf-8")
            )
    except Exception:
        logger.exception("Не удалось прочитать state.json")
    return {"last_internal_ts": 0}


def save_state(state: dict):
    try:
        STATE_PATH.write_text(
            json.dumps(state, ensure_ascii=False),
            encoding="utf-8"
        )
        logger.info(
            "🧭 last_internal_ts saved: %s",
            state.get("last_internal_ts")
        )
    except Exception:
        logger.exception("Не удалось сохранить state.json")


def debug_list_recent(service, limit: int = 15):
    resp = service.users().messages().list(
        userId="me", maxResults=limit
    ).execute()

    ids = [m["id"] for m in resp.get("messages", [])]
    print(f"🧪 Последние {len(ids)} писем в ящике:")
    for mid in ids:
        msg = service.users().messages().get(
            userId="me",
            id=mid,
            format="metadata",
            metadataHeaders=["From", "Subject", "Date"]
        ).execute()

        hdrs = {
            h["name"]: h["value"]
            for h in msg.get("payload", {}).get("headers", [])
        }
        print(
            f"- From: {hdrs.get('From','')} | "
            f"Subject: {hdrs.get('Subject','')}"
        )


# ============ main one-shot run ============

def run_once():
    processed_ids = load_processed_ids()
    if RESET_PROCESSED:
        processed_ids = set()
        logger.info("🧹 История processed_ids очищена по флагу")

    state = load_state()
    last_ts = int(state.get("last_internal_ts", 0))  # ms
    logger.info("🧭 last_internal_ts(current): %s", last_ts)

    gmail_service = gmail_auth()

    # убедимся, что метка есть (получим её id)
    label_id = get_or_create_label(gmail_service, PROCESSED_LABEL_NAME)

    query = build_school_gmail_query()
    logger.info("🔎 Gmail query: %s", query)

    # debug_list_recent(gmail_service, 15)  # включить если надо

    processed = 0
    matched   = 0
    count     = 0
    max_ts    = last_ts  # будем апдейтить, если найдём свежее

    for msg_id in iter_message_ids(
        gmail_service,
        query,
        page_size=25,
        max_pages=10
    ):
        logger.debug("🔎 message_id=%s", msg_id)

        # антидубль слой 1: уже обработан по id (локально)
        if not REPROCESS_ALL and msg_id in processed_ids:
            logger.debug("⏭ уже обработан по id: %s", msg_id)
            continue

        email_obj = get_email_with_attachments_by_id(
            gmail_service,
            msg_id
        )

        internal_ts = int(email_obj.get("internalDate") or 0)

        # антидубль слой 2: письмо старое (ts <= last_ts)
        if (
            not REPROCESS_ALL
            and internal_ts
            and internal_ts <= last_ts
        ):
            logger.debug(
                "⏭ старше/равно last_ts: id=%s ts=%s last_ts=%s",
                msg_id, internal_ts, last_ts
            )
            # пометим, чтоб не дёргать больше
            processed_ids.add(msg_id)
            continue

        if not is_school_email(email_obj):
            logger.debug(
                "🚫 не школа: %s — from=%s subject=%s",
                msg_id,
                email_obj.get("from"),
                email_obj.get("subject"),
            )
            # тоже пометим, чтоб не долбить одно и то же вечно
            processed_ids.add(msg_id)
            continue

        matched += 1

        text_for_llm = build_email_text(email_obj)
        event_data   = normalize_event_times(parse_event(text_for_llm))

        logger.info("📩 Распарсенные данные: %s", event_data)

        # выполняем действия и в любом случае помечаем письмо, чтобы больше не обрабатывать
        try:
            create_event(gmail_service, event_data)
            send_telegram(event_data)
        finally:
            try:
                add_label_to_message(gmail_service, email_obj["id"], label_id)
                logger.info("🏷️  Письмо помечено меткой %s", PROCESSED_LABEL_NAME)
            except Exception:
                logger.exception("Не удалось навесить метку на письмо %s", email_obj.get("id"))

        processed_ids.add(msg_id)
        save_processed_ids(processed_ids)  # сохраняем сразу на всякий

        if internal_ts > max_ts:
            max_ts = internal_ts

        processed += 1
        count     += 1

        if MAX_MESSAGES and count >= MAX_MESSAGES:
            logger.info(
                "⏸ Достигнут лимит MAX_MESSAGES=%d",
                MAX_MESSAGES
            )
            break

    # апдейтим last_internal_ts в state.json
    if not REPROCESS_ALL and max_ts > last_ts:
        state["last_internal_ts"] = max_ts
        save_state(state)

    if matched == 0:
        logger.info(
            "ℹ️ Писем школы не найдено по текущим правилам."
        )
    else:
        logger.info(
            "🏁 Найдено писем школы: %d. Создано/обновлено: %d.",
            matched,
            processed
        )


# ============ entrypoint ============

if __name__ == "__main__":
    try:
        # мини CLI-флаги при ручном запуске
        for arg in sys.argv[1:]:
            if arg == "--reprocess":
                REPROCESS_ALL = True
            if arg == "--reset-processed":
                RESET_PROCESSED = True
            if arg.startswith("--max="):
                MAX_MESSAGES = int(arg.split("=", 1)[1])

        run_once()
    except Exception:
        logger.exception("Фатальная ошибка выполнения скрипта")
        raise
