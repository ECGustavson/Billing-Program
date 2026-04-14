"""
QuickBooks Online – Invoice Lookup
===================================
Search by one or more comma-separated values across three custom fields:
  • Order Number
  • PO Number
  • Quote Number

Any invoice matching any search value on any of the three fields is returned.
All three field values are always displayed for every matched invoice.

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
import sys
import subprocess
import re
import logging
import threading
import tkinter as tk
from tkinter import ttk, messagebox

import requests
from dotenv import load_dotenv

# Always load .env from the same directory as this script, regardless of
# where Python is launched from — important when running off a network drive.
# Resolve base directory — works both running from source and as a PyInstaller .exe
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ENV_PATH = os.path.join(BASE_DIR, ".env")
LOG_PATH = os.path.join(BASE_DIR, "qbo_lookup.log")
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

TOKEN_URL = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
PAGE_SIZE = 100

# Custom field names exactly as they appear in QBO (case-insensitive match)
FIELD_ORDER  = "Order Number"
FIELD_PO     = "PO Number"
FIELD_QUOTE  = "Quote Number"
SEARCH_FIELDS = (FIELD_ORDER, FIELD_PO, FIELD_QUOTE)

COLUMNS = ("order_no", "po_no", "quote_no", "status", "client", "total", "date", "invoice_no")
COL_LABELS = {
    "order_no":   "Order Number",
    "po_no":      "PO Number",
    "quote_no":   "Quote Number",
    "status":     "Status",
    "client":     "Client",
    "total":      "Total",
    "date":       "Invoice Date",
    "invoice_no": "Invoice #",
}
COL_WIDTHS = {
    "order_no":   120,
    "po_no":      110,
    "quote_no":   110,
    "status":      80,
    "client":     180,
    "total":       85,
    "date":        95,
    "invoice_no":  80,
}


# ── .env helpers ───────────────────────────────────────────────────────────────

def update_env_token(new_token: str) -> None:
    """Write the rotated refresh token back into the .env file in-place."""
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
        log.info(
            f"Fetched page starting at {start} | "
            f"{len(batch)} invoices | intuit_tid={intuit_tid}"
        )

        if progress_cb:
            progress_cb(f"Fetched {len(invoices)} invoices…")

        if len(batch) < PAGE_SIZE:
            break
        start += PAGE_SIZE

    log.info(f"Invoice fetch complete | total={len(invoices)}")
    return invoices


def extract_custom_fields(invoice: dict) -> dict:
    """
    Extract all three tracked custom field values from an invoice.
    Returns a dict with keys: order_no, po_no, quote_no.
    Any field not present on the invoice is returned as an empty string.
    """
    field_map = {f.lower(): f for f in SEARCH_FIELDS}
    values = {f: "" for f in SEARCH_FIELDS}

    for cf in invoice.get("CustomField", []):
        name = cf.get("Name", "").strip().lower()
        if name in field_map:
            values[field_map[name]] = cf.get("StringValue", "").strip()

    return {
        "order_no":  values[FIELD_ORDER],
        "po_no":     values[FIELD_PO],
        "quote_no":  values[FIELD_QUOTE],
    }


def invoice_status(invoice: dict) -> str:
    balance = float(invoice.get("Balance", 0))
    total   = float(invoice.get("TotalAmt", 0))
    if balance == 0 and total > 0:
        return "Paid"
    if 0 < balance < total:
        return "Partial"
    return "Open"


def search_invoices(search_input: str, creds: dict, progress_cb=None) -> list:
    """
    Full search pipeline.
    search_input: one or more comma-separated values to search for.
    Searches all three custom fields (Order Number, PO Number, Quote Number).
    Returns a list of result dicts — one per matching invoice.
    """
    # Parse comma-separated search terms, strip whitespace, drop blanks
    search_terms = [t.strip().lower() for t in search_input.split(",") if t.strip()]
    if not search_terms:
        return []

    log.info(f"Search started | terms={search_terms}")

    access_token, new_refresh = get_access_token(
        creds["client_id"], creds["client_secret"], creds["refresh_token"]
    )

    if new_refresh != creds["refresh_token"]:
        update_env_token(new_refresh)
        creds["refresh_token"] = new_refresh

    all_invoices = fetch_all_invoices(
        access_token, creds["realm_id"], creds["environment"], progress_cb
    )

    if progress_cb:
        progress_cb(
            f"Scanning {len(all_invoices)} invoices for "
            f"{len(search_terms)} term(s)…"
        )

    results = []
    seen_ids = set()  # Prevent duplicate rows if multiple fields match

    for inv in all_invoices:
        fields = extract_custom_fields(inv)

        # Check if any search term matches any of the three field values
        field_values = [
            fields["order_no"].lower(),
            fields["po_no"].lower(),
            fields["quote_no"].lower(),
        ]
        matched = any(
            term == val
            for term in search_terms
            for val in field_values
            if val  # skip empty fields
        )

        if matched:
            inv_id = inv.get("Id")
            if inv_id in seen_ids:
                continue
            seen_ids.add(inv_id)

            results.append({
                "order_no":   fields["order_no"],
                "po_no":      fields["po_no"],
                "quote_no":   fields["quote_no"],
                "status":     invoice_status(inv),
                "client":     inv.get("CustomerRef", {}).get("name", "—"),
                "total":      f"${float(inv.get('TotalAmt', 0)):,.2f}",
                "date":       inv.get("TxnDate", "—"),
                "invoice_no": inv.get("DocNumber", "—"),
            })

    log.info(
        f"Search complete | terms={search_terms} | matches={len(results)}"
    )
    return results


# ── Tkinter GUI ────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QBO Invoice Lookup")
        self.resizable(True, True)
        self.minsize(860, 480)
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
        search_frame = ttk.LabelFrame(
            tab, text="Search  (Order Number, PO Number, or Quote Number — comma-separate multiple)",
            padding=8
        )
        search_frame.pack(fill="x")

        self.job_entry = ttk.Entry(search_frame, font=("", 12), width=40)
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
        self.status_var = tk.StringVar(value="Enter a number and press Search.")
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

        btn_frame = ttk.Frame(tab)
        btn_frame.grid(row=len(fields) + 1, column=0, columnspan=2, sticky="ew", pady=(18, 0))
        ttk.Button(btn_frame, text="Save Settings", command=self._save_settings).pack(side="right")
        ttk.Button(
            btn_frame, text="Re-authorize QBO…", command=self._launch_get_refresh
        ).pack(side="right", padx=(0, 8))
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

    def _launch_get_refresh(self):
        """Launch GetRefresh.py (or GetRefresh.exe) to re-run the OAuth flow."""
        if getattr(sys, "frozen", False):
            # Running as PyInstaller bundle — look for GetRefresh.exe alongside the exe
            target = os.path.join(BASE_DIR, "GetRefresh.exe")
            if not os.path.exists(target):
                msg = "GetRefresh.exe not found next to IsBilled.exe.\n\nExpected at:\n" + target
                messagebox.showerror("Not Found", msg)
                return
            log.info("Launching GetRefresh.exe for re-authorization.")
            subprocess.Popen([target])
        else:
            # Running from source — launch GetRefresh.py with the same Python interpreter
            target = os.path.join(BASE_DIR, "GetRefresh.py")
            if not os.path.exists(target):
                msg = "GetRefresh.py not found next to IsBilled.py.\n\nExpected at:\n" + target
                messagebox.showerror("Not Found", msg)
                return
            log.info("Launching GetRefresh.py for re-authorization.")
            subprocess.Popen([sys.executable, target])

        messagebox.showinfo(
            "Re-authorization Launched",
            "GetRefresh has opened in a new window.\n\n"
            "Complete the QuickBooks login in your browser.\n"
            "Your .env will be updated automatically when done."
        )

    def _clear_results(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.job_entry.delete(0, "end")
        self.status_var.set("Enter a number and press Search.")

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
        search_input = self.job_entry.get().strip()
        if not search_input:
            messagebox.showwarning("No Input", "Please enter a number to search.")
            return

        creds = self._get_creds()
        if creds is None:
            return

        for row in self.tree.get_children():
            self.tree.delete(row)

        self._set_busy(True)
        self.status_var.set("Connecting to QuickBooks Online…")

        threading.Thread(
            target=self._search_thread,
            args=(search_input, creds),
            daemon=True,
        ).start()

    def _search_thread(self, search_input: str, creds: dict):
        try:
            results = search_invoices(
                search_input, creds,
                progress_cb=lambda msg: self.after(0, self.status_var.set, msg),
            )
            self.after(0, self._display_results, results, search_input)
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

    def _display_results(self, results: list, search_input: str):
        self._set_busy(False)

        terms = [t.strip() for t in search_input.split(",") if t.strip()]

        if not results:
            self.status_var.set(
                f"No invoices found matching: {', '.join(terms)}"
            )
            return

        for r in results:
            self.tree.insert(
                "", "end",
                values=(
                    r["order_no"],
                    r["po_no"],
                    r["quote_no"],
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
            f"matching: {', '.join(terms)}"
        )

    def _on_error(self, message: str):
        self._set_busy(False)
        self.status_var.set("Error – see dialog for details.")
        messagebox.showerror("QuickBooks API Error", message)


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()