import os

ENV_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")


def load_env(path):
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            name, value = line.split("=", 1)
            # .env wins over variables left over in the terminal session
            os.environ[name.strip()] = value.strip().strip('"').strip("'")


load_env(ENV_FILE)

key = os.environ.get("ANTHROPIC_API_KEY", "")
workspace_id = os.environ.get("ANTHROPIC_WORKSPACE_ID", "")
model = os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-5")

claude_headers = {
    "x-api-key": key,
    "anthropic-version": "2023-06-01",
    "content-type": "application/json",
}
# user-scoped keys (sk-ant-usr-...) must name the workspace to bill
if workspace_id:
    claude_headers["anthropic-workspace-id"] = workspace_id

email_address = os.environ.get("EMAIL_ADDRESS", "")
email_sender_name = os.environ.get("EMAIL_SENDER_NAME", "")
email_password = os.environ.get("EMAIL_APP_PASSWORD", "").replace(" ", "")
smtp_host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
smtp_port = int(os.environ.get("SMTP_PORT", "587"))

imap_host = os.environ.get("IMAP_HOST", "imap.gmail.com")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Vercel's deployment folder is read-only; /tmp is writable but wiped between cold starts
DATA_DIR = os.environ.get("DATA_DIR") or (
    "/tmp/jarvis-data" if os.environ.get("VERCEL") else os.path.join(BASE_DIR, "data"))
tasks_file = os.environ.get("TASKS_FILE") or os.path.join(DATA_DIR, "tasks.xlsx")
tasks_sheet = os.environ.get("TASKS_SHEET", "")
task_managers = [
    addr.strip().lower()
    for addr in os.environ.get("TASK_MANAGERS", "").split(",")
    if addr.strip()
]
task_check_minutes = float(os.environ.get("TASK_CHECK_MINUTES", "5"))

headers = {
        "X-RapidAPI-Key": os.environ.get("RAPIDAPI_KEY", ""),
        "X-RapidAPI-Host": os.environ.get("RAPIDAPI_HOST", "yahoo-weather5.p.rapidapi.com")
    }
