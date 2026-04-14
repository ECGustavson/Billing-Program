"""
QuickBooks Online – Job Number Invoice Lookup
=============================================
Searches all invoices for a matching "Job Number" custom field value and
displays the invoice status, client name, total, and invoice date.

Setup
-----
1. Install dependencies:
       pip install requests python-dotenv

2. Create a .env file next to this script (or fill in the Settings tab in the GUI):
       QBO_CLIENT_ID=your_client_id
       QBO_CLIENT_SECRET=your_client_secret
       QBO_REFRESH_TOKEN=your_refresh_token
       QBO_REALM_ID=your_realm_id
       QBO_ENVIRONMENT=production          # or "sandbox"

   Obtain credentials at https://developer.intuit.com
   Required OAuth scope: com.intuit.quickbooks.accounting

Notes
-----
- QBO's SQL query language does not support filtering on CustomField values,
  so the script fetches all invoices (paginated) and filters client-side.
- Custom fields are returned when `include=enhancedAllCustomFields` is used
  with minorversion >= 70 (per Intuit docs).
- Field matching is case-insensitive on both the field name and value.
- A log file (qbo_lookup.log) is written next to this script automatically.
"""

import os
import re
import logging
import threading
import tkinter as tk
from tkinter import ttk, messagebox

import requests
from dotenv import load_dotenv

# Always load .env from the same directory as this script, regardless of
# where Python is launched from — important when running off a network drive.
ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "qbo_lookup.log")
load_dotenv(ENV_PATH)

# ── Logging setup ──────────────────────────────────────────────────────────────

logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("qbo_lookup")

# ── Constants ──────────────────────────────────────────────────────────────────

TOKEN_URL         = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
PAGE_SIZE         = 100
CUSTOM_FIELD_NAME = "Job Number"   # Change if your QBO field has a different name

COLUMNS = ("job_number", "status", "client", "total", "date", "invoice_no")
COL_LABELS = {
    "job_number": "Job Number",
    "status":     "Status",
    "client":     "Client",
    "total":      "Total",
    "date":       "Invoice Date",
    "invoice_no": "Invoice #",
}
COL_WIDTHS = {
    "job_number": 110,
    "status":      90,
    "client":      200,
    "total":        90,
    "date":        100,
    "invoice_no":   90,
}


# ── .env helpers ───────────────────────────────────────────────────────────────

def update_env_token(new_token: str) -> None:
    """
    Write the rotated refresh token back into the .env file in-place.
    Uses a regex replace so all other lines are preserved exactly.
    Safe to call from a background thread — writes atomically via a temp file.
    """
    try:
        if not os.path.exists(ENV_PATH):
            return

        with open(ENV_PATH, "r", encoding="utf-8") as f:
            contents = f.read()

        updated = re.sub(
            r"^(QBO_REFRESH_TOKEN\s*=\s*).*$",
            rf"\g<1>{new_token}",
            contents,
            flags=re.MULTILINE,
        )

        # Write to a temp file first, then replace — avoids corruption if the
        # process is killed mid-write (relevant on a network drive).
        tmp_path = ENV_PATH + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(updated)
        os.replace(tmp_path, ENV_PATH)
        log.info("Refresh token rotated and written to .env successfully.")

    except Exception as e:
        log.warning(f"Could not update .env with new refresh token: {e}")


# ── QBO API logic ──────────────────────────────────────────────────────────────

def _base_url(environment: str) -> str:
    if environment.lower() == "sandbox":
        return "https://sandbox-quickbooks.api.intuit.com"
    return "https://quickbooks.api.intuit.com"


def get_access_token(client_id: str, client_secret: str,
                     refresh_token: str) -> tuple:
    """Exchange a refresh token for a fresh access token.
    Returns (access_token, new_refresh_token).
    """
    log.info("Requesting access token.")
    resp = requests.post(
        TOKEN_URL,
        data={"grant_type": "refresh_token", "refresh_token": refresh_token},
        auth=(client_id, client_secret),
        headers={"Accept": "application/json"},
        timeout=15,
    )
    intuit_tid = resp.headers.get("intuit_tid", "N/A")
    if not resp.ok:
        log.error(
            f"Access token request failed | intuit_tid={intuit_tid} | "
            f"HTTP {resp.status_code} | {resp.text[:300]}"
        )
        raise requests.HTTPError(
            f"[intuit_tid: {intuit_tid}] HTTP {resp.status_code}: {resp.text}",
            response=resp,
        )
    log.info(f"Access token obtained successfully | intuit_tid={intuit_tid}")
    data = resp.json()
    new_refresh = data.get("refresh_token", refresh_token)
    return data["access_token"], new_refresh


def fetch_all_invoices(access_token: str, realm_id: str, environment: str,
                       progress_cb=None) -> list:
    """Paginate through every invoice in the company and return raw dicts."""
    base = _base_url(environment)
    url  = f"{base}/v3/company/{realm_id}/query"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept":        "application/json",
    }

    invoices = []
    start = 1
    log.info(f"Beginning invoice fetch | realm_id={realm_id} | environment={environment}")

    while True:
        sql = (
            f"SELECT * FROM Invoice "
            f"STARTPOSITION {start} MAXRESULTS {PAGE_SIZE}"
        )
        params = {
            "query":        sql,
            "minorversion": "70",
            "include":      "enhancedAllCustomFields",
        }

        resp = requests.get(url, headers=headers, params=params, timeout=20)
        intuit_tid = resp.headers.get("intuit_tid", "N/A")
        if not resp.ok:
            log.error(
                f"Invoice fetch failed | intuit_tid={intuit_tid} | "
                f"HTTP {resp.status_code} | start={start} | {resp.text[:300]}"
            )
            raise requests.HTTPError(
                f"[intuit_tid: {intuit_tid}] HTTP {resp.status_code}: {resp.text}",
                response=resp,
            )

        batch = resp.json().get("QueryResponse", {}).get("Invoice", [])
        invoices.extend(batch)
        log.info(f"Fetched page starting at {start} | {len(batch)} invoices | intuit_tid={intuit_tid}")

        if progress_cb:
            progress_cb(f"Fetched {len(invoices)} invoices…")

        if len(batch) < PAGE_SIZE:
            break
        start += PAGE_SIZE

    log.info(f"Invoice fetch complete | total={len(invoices)}")
    return invoices


def extract_job_number(invoice: dict):
    """Return the Job Number custom field value, or None if absent."""
    for cf in invoice.get("CustomField", []):
        if cf.get("Name", "").strip().lower() == CUSTOM_FIELD_NAME.lower():
            return cf.get("StringValue", "").strip() or None
    return None


def invoice_status(invoice: dict) -> str:
    balance = float(invoice.get("Balance", 0))
    total   = float(invoice.get("TotalAmt", 0))
    if balance == 0 and total > 0:
        return "Paid"
    if 0 < balance < total:
        return "Partial"
    return "Open"


def search_invoices(job_number: str, creds: dict, progress_cb=None) -> list:
    """Full search pipeline. Returns a list of result dicts."""
    log.info(f"Search started | job_number='{job_number}'")
    access_token, new_refresh = get_access_token(
        creds["client_id"], creds["client_secret"], creds["refresh_token"]
    )

    # If the token rotated, silently write it back to .env so all machines
    # stay in sync (since the script and .env live together on the network drive).
    if new_refresh != creds["refresh_token"]:
        update_env_token(new_refresh)
        creds["refresh_token"] = new_refresh

    all_invoices = fetch_all_invoices(
        access_token, creds["realm_id"], creds["environment"], progress_cb
    )

    if progress_cb:
        progress_cb(
            f"Scanning {len(all_invoices)} invoices for "
            f"Job Number '{job_number}'…"
        )

    results = []
    for inv in all_invoices:
        jn = extract_job_number(inv)
        if jn and jn.lower() == job_number.strip().lower():
            results.append({
                "job_number": jn,
                "status":     invoice_status(inv),
                "client":     inv.get("CustomerRef", {}).get("name", "—"),
                "total":      f"${float(inv.get('TotalAmt', 0)):,.2f}",
                "date":       inv.get("TxnDate", "—"),
                "invoice_no": inv.get("DocNumber", "—"),
            })

    log.info(f"Search complete | job_number='{job_number}' | matches={len(results)}")
    return results


# ── Tkinter GUI ────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QBO Job Number Lookup")
        self.resizable(True, True)
        self.minsize(740, 480)
        self._build_ui()
        self._load_env_into_fields()
        log.info("Application started.")

    # ── UI construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        self._build_search_tab()
        self._build_settings_tab()

    def _build_search_tab(self):
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text="  Search  ")

        # Search bar
        search_frame = ttk.LabelFrame(tab, text="Job Number", padding=8)
        search_frame.pack(fill="x")

        self.job_entry = ttk.Entry(search_frame, font=("", 12), width=28)
        self.job_entry.pack(side="left", padx=(0, 8), ipady=4)
        self.job_entry.bind("<Return>", lambda _: self._run_search())

        self.search_btn = ttk.Button(
            search_frame, text="Search", command=self._run_search
        )
        self.search_btn.pack(side="left")

        self.clear_btn = ttk.Button(
            search_frame, text="Clear", command=self._clear_results
        )
        self.clear_btn.pack(side="left", padx=(4, 0))

        # Status label
        self.status_var = tk.StringVar(value="Enter a Job Number and press Search.")
        ttk.Label(tab, textvariable=self.status_var, foreground="#555").pack(
            anchor="w", pady=(8, 2)
        )

        # Results table
        table_frame = ttk.Frame(tab)
        table_frame.pack(fill="both", expand=True, pady=(4, 0))

        self.tree = ttk.Treeview(
            table_frame, columns=COLUMNS, show="headings", selectmode="browse"
        )
        for col in COLUMNS:
            self.tree.heading(col, text=COL_LABELS[col])
            self.tree.column(col, width=COL_WIDTHS[col], anchor="w", minwidth=60)

        # Status-based row colours
        self.tree.tag_configure("Paid",    foreground="#1a7a1a")
        self.tree.tag_configure("Open",    foreground="#c0392b")
        self.tree.tag_configure("Partial", foreground="#d68910")

        vsb = ttk.Scrollbar(table_frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Progress bar
        self.progress = ttk.Progressbar(tab, mode="indeterminate")
        self.progress.pack(fill="x", pady=(8, 0))

    def _build_settings_tab(self):
        tab = ttk.Frame(self.notebook, padding=16)
        self.notebook.add(tab, text="  Settings  ")

        fields = [
            ("Client ID",          "client_id",     False),
            ("Client Secret",      "client_secret", True),
            ("Refresh Token",      "refresh_token", True),
            ("Realm / Company ID", "realm_id",      False),
        ]

        self._setting_vars = {}
        for row, (label, key, secret) in enumerate(fields):
            ttk.Label(tab, text=label + ":").grid(
                row=row, column=0, sticky="w", pady=5, padx=(0, 14)
            )
            var = tk.StringVar()
            entry = ttk.Entry(
                tab, textvariable=var, width=55,
                show="*" if secret else ""
            )
            entry.grid(row=row, column=1, sticky="ew", pady=5)
            self._setting_vars[key] = var

        # Environment toggle
        ttk.Label(tab, text="Environment:").grid(
            row=len(fields), column=0, sticky="w", pady=5
        )
        self._env_var = tk.StringVar(value="production")
        env_frame = ttk.Frame(tab)
        env_frame.grid(row=len(fields), column=1, sticky="w")
        ttk.Radiobutton(env_frame, text="Production",
                        variable=self._env_var, value="production").pack(side="left")
        ttk.Radiobutton(env_frame, text="Sandbox",
                        variable=self._env_var, value="sandbox").pack(
            side="left", padx=(14, 0)
        )

        tab.columnconfigure(1, weight=1)

        ttk.Button(tab, text="Save Settings", command=self._save_settings).grid(
            row=len(fields) + 1, column=1, sticky="e", pady=(18, 0)
        )
        ttk.Label(
            tab,
            text=(
                "Credentials are loaded from the .env file next to this script.\n"
                "The refresh token is automatically updated in .env when it rotates.\n"
                "Do not move the script and .env file apart."
            ),
            foreground="#777",
            justify="left",
        ).grid(row=len(fields) + 2, column=0, columnspan=2, sticky="w", pady=(14, 0))

    # ── Helpers ────────────────────────────────────────────────────────────────

    def _load_env_into_fields(self):
        """Pre-populate settings fields from environment / .env."""
        mapping = {
            "client_id":     "QBO_CLIENT_ID",
            "client_secret": "QBO_CLIENT_SECRET",
            "refresh_token": "QBO_REFRESH_TOKEN",
            "realm_id":      "QBO_REALM_ID",
        }
        for key, env_key in mapping.items():
            self._setting_vars[key].set(os.getenv(env_key, ""))

        env = os.getenv("QBO_ENVIRONMENT", "production").lower()
        self._env_var.set(env if env in ("production", "sandbox") else "production")

    def _get_creds(self):
        creds = {
            "client_id":     self._setting_vars["client_id"].get().strip(),
            "client_secret": self._setting_vars["client_secret"].get().strip(),
            "refresh_token": self._setting_vars["refresh_token"].get().strip(),
            "realm_id":      self._setting_vars["realm_id"].get().strip(),
            "environment":   self._env_var.get(),
        }
        missing = [k for k, v in creds.items() if not v]
        if missing:
            messagebox.showerror(
                "Missing Credentials",
                f"Please fill in the Settings tab:\n  {', '.join(missing)}"
            )
            return None
        return creds

    def _save_settings(self):
        """Write all settings fields back to the .env file."""
        mapping = {
            "QBO_CLIENT_ID":     self._setting_vars["client_id"].get().strip(),
            "QBO_CLIENT_SECRET": self._setting_vars["client_secret"].get().strip(),
            "QBO_REFRESH_TOKEN": self._setting_vars["refresh_token"].get().strip(),
            "QBO_REALM_ID":      self._setting_vars["realm_id"].get().strip(),
            "QBO_ENVIRONMENT":   self._env_var.get(),
        }

        try:
            if os.path.exists(ENV_PATH):
                with open(ENV_PATH, "r", encoding="utf-8") as f:
                    contents = f.read()
            else:
                contents = ""

            for key, value in mapping.items():
                if re.search(rf"^{key}\s*=", contents, flags=re.MULTILINE):
                    contents = re.sub(
                        rf"^({key}\s*=\s*).*$",
                        rf"\g<1>{value}",
                        contents,
                        flags=re.MULTILINE,
                    )
                else:
                    contents += f"\n{key}={value}"

            tmp_path = ENV_PATH + ".tmp"
            with open(tmp_path, "w", encoding="utf-8") as f:
                f.write(contents)
            os.replace(tmp_path, ENV_PATH)

            log.info("Settings saved to .env via GUI.")
            messagebox.showinfo("Settings Saved", f"Credentials saved to:\n{ENV_PATH}")

        except Exception as e:
            log.error(f"Failed to save settings to .env: {e}")
            messagebox.showerror("Save Failed", f"Could not write to .env:\n{e}")

    def _clear_results(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.job_entry.delete(0, "end")
        self.status_var.set("Enter a Job Number and press Search.")

    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        self.search_btn.config(state=state)
        self.job_entry.config(state=state)
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    # ── Search (runs in a background thread) ───────────────────────────────────

    def _run_search(self):
        job_number = self.job_entry.get().strip()
        if not job_number:
            messagebox.showwarning("No Input", "Please enter a Job Number.")
            return

        creds = self._get_creds()
        if creds is None:
            return

        # Clear old results
        for row in self.tree.get_children():
            self.tree.delete(row)

        self._set_busy(True)
        self.status_var.set("Connecting to QuickBooks Online…")

        threading.Thread(
            target=self._search_thread,
            args=(job_number, creds),
            daemon=True,
        ).start()

    def _search_thread(self, job_number: str, creds: dict):
        try:
            results = search_invoices(
                job_number, creds,
                progress_cb=lambda msg: self.after(0, self.status_var.set, msg),
            )
            self.after(0, self._display_results, results, job_number)
        except requests.HTTPError as e:
            msg = str(e)
            log.error(f"HTTPError during search | {msg}")
            self.after(0, self._on_error, msg)
        except requests.ConnectionError as e:
            log.error(f"ConnectionError during search | {e}")
            self.after(0, self._on_error,
                       "Connection failed.\nCheck your network and credentials.")
        except Exception as e:
            log.exception(f"Unexpected error during search: {e}")
            self.after(0, self._on_error, str(e))

    def _display_results(self, results: list, job_number: str):
        self._set_busy(False)

        if not results:
            self.status_var.set(
                f"No invoices found with Job Number '{job_number}'."
            )
            return

        for r in results:
            self.tree.insert(
                "", "end",
                values=(
                    r["job_number"],
                    r["status"],
                    r["client"],
                    r["total"],
                    r["date"],
                    r["invoice_no"],
                ),
                tags=(r["status"],),
            )

        n = len(results)
        self.status_var.set(
            f"Found {n} invoice{'s' if n != 1 else ''} "
            f"for Job Number '{job_number}'."
        )

    def _on_error(self, message: str):
        self._set_busy(False)
        self.status_var.set("Error – see dialog for details.")
        messagebox.showerror("QuickBooks API Error", message)


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()