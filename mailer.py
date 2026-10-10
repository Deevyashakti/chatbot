import base64
import json
import os
import re
import smtplib
import threading
from email.message import EmailMessage
from email.utils import getaddresses, make_msgid

import claude
from config import model, email_address, email_password, smtp_host, smtp_port, DATA_DIR

CONTACTS_FILE = os.path.join(DATA_DIR, "contacts.json")
EMAIL_RE = re.compile(r"^[^@\s<>,;]+@[^@\s<>,;]+\.[^@\s<>,;]+$")
DATA_IMAGE_RE = re.compile(r'src="data:(image/[\w.+-]+);base64,([A-Za-z0-9+/=]+)"')
_contacts_lock = threading.Lock()


def parse_addresses(value):
    """Accepts a list or a comma/semicolon separated string. Returns (valid, invalid) address lists."""
    if not value:
        return [], []
    if isinstance(value, (list, tuple)):
        value = ",".join(value)
    valid, invalid = [], []
    for name, addr in getaddresses([value.replace(";", ",")]):
        addr = addr.strip()
        if not addr:
            continue
        (valid if EMAIL_RE.match(addr) else invalid).append(addr)
    return valid, invalid


def load_contacts():
    with _contacts_lock:
        try:
            with open(CONTACTS_FILE, encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return {}


def remember_contacts(pairs):
    """pairs: iterable of (name, email)"""
    with _contacts_lock:
        try:
            with open(CONTACTS_FILE, encoding="utf-8") as f:
                contacts = json.load(f)
        except (OSError, ValueError):
            contacts = {}
        changed = False
        for name, addr in pairs:
            addr = (addr or "").strip().lower()
            if not EMAIL_RE.match(addr):
                continue
            if addr not in contacts or (name and not contacts[addr]):
                contacts[addr] = name or ""
                changed = True
        if changed:
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(CONTACTS_FILE, "w", encoding="utf-8") as f:
                json.dump(contacts, f, indent=2)


def _attach_inline_images(msg, html):
    """Swap data: URL images in the HTML for cid: references so mail clients display them."""
    images = []

    def replace(match):
        cid = make_msgid(domain="jarvis.local")
        images.append((cid, match.group(1), base64.b64decode(match.group(2))))
        return f'src="cid:{cid[1:-1]}"'

    html = DATA_IMAGE_RE.sub(replace, html)
    msg.add_alternative(html, subtype="html")
    if images:
        html_part = msg.get_payload()[-1]
        for cid, mimetype, content in images:
            maintype, _, subtype = mimetype.partition("/")
            html_part.add_related(content, maintype=maintype, subtype=subtype, cid=cid)


def send_email(to, subject, body, attachments=None, cc=None, bcc=None, html=None,
               in_reply_to=None, references=None):
    """attachments: list of (filename, bytes, mimetype) tuples. html: optional rich version of body."""
    if not email_address or not email_password:
        return "Email is not configured. Set EMAIL_ADDRESS and EMAIL_APP_PASSWORD."

    to_list, bad_to = parse_addresses(to)
    cc_list, bad_cc = parse_addresses(cc)
    bcc_list, bad_bcc = parse_addresses(bcc)
    invalid = bad_to + bad_cc + bad_bcc
    if invalid:
        return "Failed to send email: invalid address " + ", ".join(invalid)
    if not to_list:
        return "Failed to send email: please add at least one recipient in To"

    msg = EmailMessage()
    msg["From"] = email_address
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    if bcc_list:
        msg["Bcc"] = ", ".join(bcc_list)  # smtplib strips this header and still delivers to these addresses
    msg["Subject"] = subject
    if in_reply_to:
        msg["In-Reply-To"] = in_reply_to
        msg["References"] = references or in_reply_to
    msg.set_content(body or "")
    if html:
        _attach_inline_images(msg, html)

    for filename, content, mimetype in attachments or []:
        maintype, _, subtype = (mimetype or "application/octet-stream").partition("/")
        msg.add_attachment(content, maintype=maintype, subtype=subtype or "octet-stream", filename=filename)

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=60) as server:
            server.starttls()
            server.login(email_address, email_password)
            server.send_message(msg)
    except (smtplib.SMTPAuthenticationError, smtplib.SMTPServerDisconnected) as e:
        print("Error: ", e)
        return ("Failed to send email: the mail server rejected the login. "
                "Use an app password for EMAIL_APP_PASSWORD, not your normal password.")
    except Exception as e:
        print("Error: ", e)
        return f"Failed to send email: {e}"

    remember_contacts(("", addr) for addr in to_list + cc_list + bcc_list)
    summary = f"Email sent to {', '.join(to_list)}"
    if cc_list or bcc_list:
        summary += f" (+{len(cc_list) + len(bcc_list)} in Cc/Bcc)"
    if attachments:
        summary += f" with {len(attachments)} attachment(s)"
    return summary


def polish_email(subject, body):
    data = {
        "model": model,
        "max_tokens": 1024,
        "system": "You turn the user's rough draft or short notes into a well-written, "
                  "natural-sounding professional email. Write a clear, specific subject line. "
                  "Write the body with a greeting, well-structured paragraphs that expand the "
                  "user's points into complete sentences, and a polite closing. "
                  "Fix grammar and spelling and keep the original meaning, intent and language. "
                  "Keep any @mentions of people exactly as written. "
                  "Do not invent facts, dates, numbers or names, and do not use placeholders "
                  "like [Name]; if the recipient's name is unknown use a neutral greeting such as "
                  "'Hello,', and end with 'Best regards,' without a name.",
        "messages": [{"role": "user", "content": f"Subject: {subject}\n\nBody:\n{body}"}],
        "tools": [{
            "name": "polished_email",
            "description": "Return the polished email",
            "input_schema": {
                "type": "object",
                "properties": {
                    "subject": {"type": "string"},
                    "body": {"type": "string"},
                },
                "required": ["subject", "body"],
            },
        }],
        "tool_choice": {"type": "tool", "name": "polished_email"},
    }
    result, error = claude.create_message(data)
    if error:
        return None

    for block in result.get("content", []):
        if block.get("type") == "tool_use":
            return block.get("input")
    return None
