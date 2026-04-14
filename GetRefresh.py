"""
QBO One-Time Token Setup
========================
Run this script ONCE to get your initial refresh token and Realm ID.
It will write both directly into your .env file automatically.

Usage
-----
    python qbo_get_tokens.py

Requirements
------------
    pip install requests python-dotenv

Before running
--------------
Make sure your .env file has at least these two fields filled in:
    QBO_CLIENT_ID=your_client_id
    QBO_CLIENT_SECRET=your_client_secret

And make sure this redirect URI is added in the Intuit Developer Portal
under Settings -> Redirect URIs (Development tab):
    http://localhost:8080/callback
"""

import os
import re
import sys
import webbrowser
import urllib.parse
import secrets
from http.server import BaseHTTPRequestHandler, HTTPServer

import requests
from dotenv import load_dotenv

# ── Config ─────────────────────────────────────────────────────────────────────

ENV_PATH      = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
REDIRECT_URI  = "http://localhost:8080/callback"
AUTH_URL      = "https://appcenter.intuit.com/connect/oauth2"
TOKEN_URL     = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
SCOPE         = "com.intuit.quickbooks.accounting"

load_dotenv(ENV_PATH)

CLIENT_ID     = os.getenv("QBO_CLIENT_ID", "").strip()
CLIENT_SECRET = os.getenv("QBO_CLIENT_SECRET", "").strip()

# ── Validate ───────────────────────────────────────────────────────────────────

if not CLIENT_ID or not CLIENT_SECRET:
    print(
        "\n[ERROR] QBO_CLIENT_ID and QBO_CLIENT_SECRET must be set in your .env "
        "before running this script.\n"
    )
    sys.exit(1)

# ── .env writer ────────────────────────────────────────────────────────────────

def write_env_value(key: str, value: str) -> None:
    """Insert or update a single key=value line in the .env file."""
    if os.path.exists(ENV_PATH):
        with open(ENV_PATH, "r", encoding="utf-8") as f:
            contents = f.read()
    else:
        contents = ""

    if re.search(rf"^{key}\s*=", contents, flags=re.MULTILINE):
        contents = re.sub(
            rf"^({key}\s*=\s*).*$",
            rf"\g<1>{value}",
            contents,
            flags=re.MULTILINE,
        )
    else:
        contents = contents.rstrip("\n") + f"\n{key}={value}\n"

    tmp = ENV_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(contents)
    os.replace(tmp, ENV_PATH)


# ── Local callback server ──────────────────────────────────────────────────────

# Shared state between the HTTP handler and the main thread
_callback_data: dict = {}

class _CallbackHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)

        if parsed.path != "/callback":
            self._respond(404, "Not found")
            return

        if "error" in params:
            error = params["error"][0]
            self._respond(400, f"Authorization failed: {error}")
            _callback_data["error"] = error
            return

        code     = params.get("code",    [None])[0]
        realm_id = params.get("realmId", [None])[0]
        state    = params.get("state",   [None])[0]

        if not code:
            self._respond(400, "Missing authorization code.")
            return

        _callback_data["code"]     = code
        _callback_data["realm_id"] = realm_id
        _callback_data["state"]    = state

        self._respond(
            200,
            "<html><body style='font-family:sans-serif;padding:40px'>"
            "<h2>✅ Authorization successful!</h2>"
            "<p>You can close this tab and return to the terminal.</p>"
            "</body></html>",
            content_type="text/html",
        )

    def _respond(self, code: int, body: str, content_type: str = "text/plain"):
        encoded = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def log_message(self, format, *args):
        pass  # Suppress default request logging


# ── Main flow ──────────────────────────────────────────────────────────────────

def main():
    state = secrets.token_urlsafe(16)

    # Build authorization URL
    auth_params = urllib.parse.urlencode({
        "client_id":     CLIENT_ID,
        "response_type": "code",
        "scope":         SCOPE,
        "redirect_uri":  REDIRECT_URI,
        "state":         state,
    })
    auth_link = f"{AUTH_URL}?{auth_params}"

    print("\n" + "=" * 60)
    print("  QBO One-Time Token Setup")
    print("=" * 60)
    print("\nOpening your browser to authorize with QuickBooks...")
    print("If the browser doesn't open, visit this URL manually:\n")
    print(f"  {auth_link}\n")

    webbrowser.open(auth_link)

    # Start local server and wait for callback
    print("Waiting for QuickBooks to redirect back... (do not close this window)\n")
    server = HTTPServer(("localhost", 8080), _CallbackHandler)
    server.handle_request()  # Handles exactly one request then returns

    if "error" in _callback_data:
        print(f"\n[ERROR] Authorization was denied: {_callback_data['error']}")
        sys.exit(1)

    code     = _callback_data.get("code")
    realm_id = _callback_data.get("realm_id")
    returned_state = _callback_data.get("state")

    # Verify state to prevent CSRF
    if returned_state != state:
        print("\n[ERROR] State mismatch — possible CSRF. Aborting.")
        sys.exit(1)

    if not code:
        print("\n[ERROR] No authorization code received.")
        sys.exit(1)

    # Exchange auth code for tokens
    print("Authorization code received. Exchanging for tokens...")

    resp = requests.post(
        TOKEN_URL,
        data={
            "grant_type":   "authorization_code",
            "code":         code,
            "redirect_uri": REDIRECT_URI,
        },
        auth=(CLIENT_ID, CLIENT_SECRET),
        headers={"Accept": "application/json"},
        timeout=15,
    )

    if not resp.ok:
        print(f"\n[ERROR] Token exchange failed: {resp.status_code} {resp.text}")
        sys.exit(1)

    tokens = resp.json()
    refresh_token = tokens.get("refresh_token")
    access_token  = tokens.get("access_token")

    if not refresh_token:
        print(f"\n[ERROR] No refresh token in response: {tokens}")
        sys.exit(1)

    # Write to .env
    write_env_value("QBO_REFRESH_TOKEN", refresh_token)
    if realm_id:
        write_env_value("QBO_REALM_ID", realm_id)

    print("\n" + "=" * 60)
    print("  ✅ Success! Your .env has been updated.")
    print("=" * 60)
    print(f"\n  Refresh Token : {refresh_token[:12]}... (truncated for display)")
    if realm_id:
        print(f"  Realm ID      : {realm_id}")
    else:
        print(
            "\n  [NOTE] Realm ID was not returned. You may need to add\n"
            "  QBO_REALM_ID manually to your .env (see README for help)."
        )
    print(
        "\nYou can now run qb_job_number_search.py. "
        "This setup script does not need to be run again\n"
        "unless your token expires (100 days of inactivity) or is revoked.\n"
    )


if __name__ == "__main__":
    main()