# Emails queued with "Schedule send". Each one is a JSON file so it survives restarts.
import base64
import glob
import json
import os
import uuid
from datetime import datetime

import mailer
from config import DATA_DIR

SCHEDULE_DIR = os.path.join(DATA_DIR, "scheduled")
MAX_ATTEMPTS = 3


def schedule(send_at, to, cc, bcc, subject, body, html, attachments):
    """send_at: naive local datetime. attachments: list of (filename, bytes, mimetype)."""
    os.makedirs(SCHEDULE_DIR, exist_ok=True)
    item = {
        "id": uuid.uuid4().hex[:10],
        "send_at": send_at.strftime("%Y-%m-%d %H:%M"),
        "to": to, "cc": cc, "bcc": bcc,
        "subject": subject, "body": body, "html": html,
        "attachments": [
            {"filename": name, "mimetype": mimetype, "data": base64.b64encode(content).decode()}
            for name, content, mimetype in attachments
        ],
        "attempts": 0,
        "error": "",
    }
    with open(os.path.join(SCHEDULE_DIR, item["id"] + ".json"), "w", encoding="utf-8") as f:
        json.dump(item, f)
    return item


def _items():
    items = []
    for path in glob.glob(os.path.join(SCHEDULE_DIR, "*.json")):
        try:
            with open(path, encoding="utf-8") as f:
                items.append((path, json.load(f)))
        except (OSError, ValueError):
            continue
    return sorted(items, key=lambda pair: pair[1]["send_at"])


def list_scheduled():
    return [{
        "id": item["id"], "send_at": item["send_at"], "to": item["to"],
        "subject": item["subject"], "attachments": len(item["attachments"]), "error": item.get("error", ""),
    } for _, item in _items()]


def cancel(item_id):
    path = os.path.join(SCHEDULE_DIR, os.path.basename(item_id) + ".json")
    if os.path.exists(path):
        os.remove(path)
        return True
    return False


def send_due():
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    for path, item in _items():
        if item["send_at"] > now or item.get("attempts", 0) >= MAX_ATTEMPTS:
            continue
        attachments = [(a["filename"], base64.b64decode(a["data"]), a["mimetype"]) for a in item["attachments"]]
        outcome = mailer.send_email(item["to"], item["subject"], item["body"], attachments,
                                    cc=item["cc"], bcc=item["bcc"], html=item["html"])
        if outcome.startswith("Email sent"):
            os.remove(path)
            print("Scheduled email sent:", item["subject"])
        else:
            item["attempts"] = item.get("attempts", 0) + 1
            item["error"] = outcome
            with open(path, "w", encoding="utf-8") as f:
                json.dump(item, f)
