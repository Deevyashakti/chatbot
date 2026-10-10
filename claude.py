import time

import requests
from config import claude_headers

API_URL = "https://api.anthropic.com/v1/messages"
# (connect, read) seconds. On this network some new TLS handshakes stall forever,
# so give up on a connection quickly and retry instead of waiting.
TIMEOUT = (5, 30)
ATTEMPTS = 3

# Reusing one keep-alive connection avoids a new handshake on every request.
session = requests.Session()
session.headers.update(claude_headers)


def create_message(data):
    """Send a Messages API request. Returns (json, None) on success or (None, error text)."""
    for attempt in range(1, ATTEMPTS + 1):
        try:
            response = session.post(API_URL, json=data, timeout=TIMEOUT)
        except (requests.ConnectionError, requests.Timeout) as e:
            print(f"Claude request attempt {attempt} failed: {type(e).__name__}")
            if attempt < ATTEMPTS:
                time.sleep(0.5)
                continue
            return None, "Couldn't reach Claude. Please check your internet connection and try again."

        if response.status_code in (429, 500, 502, 503, 529) and attempt < ATTEMPTS:
            print(f"Claude returned {response.status_code}, retrying")
            time.sleep(attempt)
            continue
        if response.status_code != 200:
            print(response.text)
            return None, f"Claude returned an error ({response.status_code})."
        return response.json(), None
