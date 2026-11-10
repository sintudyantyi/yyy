# ═══════════════════════════════════════════════════════════════
# CRYPTO TRADING DASHBOARD — R120M LEDGER EDITION
# Deploy from GitHub to Render.
# Ledger-driven balance · Excel audit trail · 72-row seed
# ═══════════════════════════════════════════════════════════════

from flask import Flask, request, jsonify, Response
import sqlite3
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)
DB = os.environ.get("DB_PATH", "dashboard.db")
EXCEL_FILE = os.environ.get("EXCEL_PATH", "withdrawals.xlsx")

# ═══════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════
START_BALANCE_ZAR = 120_000_000.00
BROKER_MODE = "real-money"
BASE_CURRENCY = "ZAR"
CRYPTO_SYMBOL = "BTC"
ZAR_PER_USD = 18.50
ZAR_PER_BTC = 1_200_000

TARGET_ZAR       = 120_000_000.00
EXPECTED_ROWS    = 72
EXPECTED_MONTHS  = 24
EXPECTED_MONTHLY = 5_000_000.00

BROKER_CONFIG = {
    "OANDA_TOKEN":            os.environ.get("OANDA_TOKEN", ""),
    "OANDA_ACCOUNT":          os.environ.get("OANDA_ACCOUNT", ""),
    "LUNO_API_KEY":           os.environ.get("LUNO_API_KEY", ""),
    "LUNO_API_SECRET":        os.environ.get("LUNO_API_SECRET", ""),
    "PAYPAL_EMAIL":           os.environ.get("PAYPAL_EMAIL", ""),
    "PAYPAL_RECEIVER_EMAIL":  os.environ.get("PAYPAL_RECEIVER_EMAIL", ""),
    "PAYPAL_ME_USERNAME":     os.environ.get("PAYPAL_ME_USERNAME", ""),
    "PAYPAL_CLIENT_ID":       os.environ.get("PAYPAL_CLIENT_ID", ""),      # TODO: set for live payouts
    "PAYPAL_CLIENT_SECRET":   os.environ.get("PAYPAL_CLIENT_SECRET", ""),  # TODO: set for live payouts
    "PAYPAL_MODE":            os.environ.get("PAYPAL_MODE", "sandbox"),    # sandbox | live
    "BANK_ACCOUNT":           os.environ.get("BANK_ACCOUNT", ""),
    "BTC_WALLET":             os.environ.get("BTC_WALLET", ""),
    "LUNO_RECEIVER_ACCOUNT":  os.environ.get("LUNO_RECEIVER_ACCOUNT", ""),
    "ADMIN_TOKEN":            os.environ.get("ADMIN_TOKEN", ""),
}

# ═══════════════════════════════════════════════════════════════
# HELPERS — currency & formatting
# ═══════════════════════════════════════════════════════════════
def usd_to_zar(usd):  return usd * ZAR_PER_USD
def zar_to_btc(zar):  return zar / ZAR_PER_BTC
def fmt_zar(amount):  return f"R{amount:,.2f}"

# ═══════════════════════════════════════════════════════════════
# DATABASE — base tables
# ═══════════════════════════════════════════════════════════════
def init_db():
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS trades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            signal TEXT NOT NULL,
            price REAL NOT NULL,
            sl REAL DEFAULT 0,
            tp REAL DEFAULT 0,
            units INTEGER DEFAULT 10,
            status TEXT DEFAULT 'open',
            pnl_usd REAL DEFAULT 0,
            pnl_zar REAL DEFAULT 0,
            close_price REAL DEFAULT 0,
            source TEXT DEFAULT 'manual'
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS withdrawals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            amount_zar REAL NOT NULL,
            method TEXT NOT NULL,
            status TEXT DEFAULT 'pending',
            note TEXT DEFAULT '',
            crypto_amount REAL DEFAULT 0,
            pay_link TEXT DEFAULT ''
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS config (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)
    c.execute("INSERT OR IGNORE INTO config (key,value) VALUES ('start_balance',?)",
              (str(START_BALANCE_ZAR),))
    conn.commit()
    conn.close()

# ═══════════════════════════════════════════════════════════════
# DATABASE — ledger schema
# ═══════════════════════════════════════════════════════════════
CREDIT_CATEGORIES = {
    "capital_injection", "trade_profit", "client_payment",
    "broker_fill", "interest", "refund", "other_credit",
}
DEBIT_CATEGORIES = {
    "trade_loss", "withdrawal", "fee", "tax", "other_debit",
}
ALL_CATEGORIES = CREDIT_CATEGORIES | DEBIT_CATEGORIES

def init_ledger_db():
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS ledger (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            occurred_at       TEXT NOT NULL,
            direction         TEXT NOT NULL,
            category          TEXT NOT NULL,
            amount_zar        REAL NOT NULL,
            amount_usd        REAL DEFAULT 0,
            crypto_amount     REAL DEFAULT 0,
            status            TEXT NOT NULL DEFAULT 'completed',
            counterparty      TEXT DEFAULT '',
            method            TEXT DEFAULT '',
            provider          TEXT DEFAULT '',
            provider_txn_id   TEXT DEFAULT '',
            reference         TEXT DEFAULT '',
            trade_id          INTEGER DEFAULT NULL,
            withdrawal_id     INTEGER DEFAULT NULL,
            note              TEXT DEFAULT '',
            created_at        TEXT NOT NULL
        )
    """)
    c.execute("CREATE INDEX IF NOT EXISTS idx_ledger_time     ON ledger(occurred_at)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_ledger_status   ON ledger(status)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_ledger_category ON ledger(category)")
    c.execute("""
        CREATE VIEW IF NOT EXISTS v_balance AS
        SELECT
            COALESCE(SUM(CASE WHEN direction='CREDIT' AND status='completed' THEN amount_zar END),0)
          - COALESCE(SUM(CASE WHEN direction='DEBIT'  AND status='completed' THEN amount_zar END),0)
            AS completed_balance_zar,
            COALESCE(SUM(CASE WHEN direction='DEBIT'  AND status='pending' THEN amount_zar END),0)
            AS pending_debits_zar,
            COALESCE(SUM(CASE WHEN direction='CREDIT' AND status='pending' THEN amount_zar END),0)
            AS pending_credits_zar
        FROM ledger
    """)
    conn.commit()
    conn.close()

init_db()
init_ledger_db()

def query(sql, args=(), one=False):
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    cur = conn.execute(sql, args)
    rows = cur.fetchall()
    conn.close()
    return (rows[0] if rows else None) if one else rows

def execute(sql, args=()):
    conn = sqlite3.connect(DB)
    conn.execute(sql, args)
    conn.commit()
    conn.close()

def get_start_balance():
    r = query("SELECT value FROM config WHERE key='start_balance'", one=True)
    return float(r["value"]) if r else START_BALANCE_ZAR

# ═══════════════════════════════════════════════════════════════
# LEDGER — post & aggregate
# ═══════════════════════════════════════════════════════════════
def ledger_post(direction, category, amount_zar, occurred_at=None,
                amount_usd=0, crypto_amount=0, status="completed",
                counterparty="", method="", provider="", provider_txn_id="",
                reference="", trade_id=None, withdrawal_id=None, note=""):
    direction = direction.upper()
    if direction not in ("CREDIT", "DEBIT"):
        raise ValueError("direction must be CREDIT or DEBIT")
    if category not in ALL_CATEGORIES:
        raise ValueError(f"unknown category: {category}")
    valid = (direction == "CREDIT" and category in CREDIT_CATEGORIES) or \
            (direction == "DEBIT"  and category in DEBIT_CATEGORIES)
    if not valid:
        raise ValueError(f"category '{category}' is not valid for {direction}")
    if amount_zar <= 0:
        raise ValueError("amount_zar must be > 0")

    occurred_at = occurred_at or datetime.utcnow().isoformat()
    created_at  = datetime.utcnow().isoformat()

    conn = sqlite3.connect(DB)
    cur = conn.execute("""
        INSERT INTO ledger
            (occurred_at, direction, category, amount_zar, amount_usd,
             crypto_amount, status, counterparty, method, provider,
             provider_txn_id, reference, trade_id, withdrawal_id, note, created_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (occurred_at, direction, category, round(amount_zar, 2),
          round(amount_usd, 2), round(crypto_amount, 8), status,
          counterparty, method, provider, provider_txn_id, reference,
          trade_id, withdrawal_id, note, created_at))
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    return new_id

def ledger_balance():
    row = query("SELECT * FROM v_balance", one=True)
    completed      = float(row["completed_balance_zar"] or 0)
    pending_debits = float(row["pending_debits_zar"] or 0)
    pending_creds  = float(row["pending_credits_zar"] or 0)
    return {
        "completed_balance_zar": round(completed, 2),
        "pending_debits_zar":    round(pending_debits, 2),
        "pending_credits_zar":   round(pending_creds, 2),
        "available_zar":         round(max(completed - pending_debits, 0), 2),
    }

def ledger_by_category():
    rows = query("""
        SELECT direction, category,
               COUNT(*) AS n,
               COALESCE(SUM(amount_zar),0) AS total_zar
        FROM ledger
        WHERE status='completed'
        GROUP BY direction, category
        ORDER BY direction, total_zar DESC
    """)
    return [dict(r) for r in rows]

# ═══════════════════════════════════════════════════════════════
# 24-MONTH SEED — 72 rows from 120m_24_month_ledger.xlsx
# ═══════════════════════════════════════════════════════════════
SEED_MODE = "SIM"   # change to "LIVE" once references are real

_DESC_TO_CATEGORY = {
    "capital contribution": "capital_injection",
    "operating revenue":    "trade_profit",
    "investment proceeds":  "trade_profit",
}

SEED_LEDGER_RAW = [
    # ── 2024-11 ──
    ("2024-11-05T00:00:00", "capital contribution", 800_000, "SIM-202411-001"),
    ("2024-11-14T00:00:00", "operating revenue",  1_500_000, "SIM-202411-002"),
    ("2024-11-25T00:00:00", "investment proceeds",2_700_000, "SIM-202411-003"),
    # ── 2024-12 ──
    ("2024-12-05T00:00:00", "capital contribution",  810_000, "SIM-202412-001"),
    ("2024-12-14T00:00:00", "operating revenue",   1_550_000, "SIM-202412-002"),
    ("2024-12-25T00:00:00", "investment proceeds", 2_640_000, "SIM-202412-003"),
    # ── 2025-01 ──
    ("2025-01-05T00:00:00", "capital contribution",  820_000, "SIM-202501-001"),
    ("2025-01-14T00:00:00", "operating revenue",   1_600_000, "SIM-202501-002"),
    ("2025-01-25T00:00:00", "investment proceeds", 2_580_000, "SIM-202501-003"),
    # ── 2025-02 ──
    ("2025-02-05T00:00:00", "capital contribution",  830_000, "SIM-202502-001"),
    ("2025-02-14T00:00:00", "operating revenue",   1_650_000, "SIM-202502-002"),
    ("2025-02-25T00:00:00", "investment proceeds", 2_520_000, "SIM-202502-003"),
    # ── 2025-03 ──
    ("2025-03-05T00:00:00", "capital contribution",  840_000, "SIM-202503-001"),
    ("2025-03-14T00:00:00", "operating revenue",   1_700_000, "SIM-202503-002"),
    ("2025-03-25T00:00:00", "investment proceeds", 2_460_000, "SIM-202503-003"),
    # ── 2025-04 ──
    ("2025-04-05T00:00:00", "capital contribution",  850_000, "SIM-202504-001"),
    ("2025-04-14T00:00:00", "operating revenue",   1_750_000, "SIM-202504-002"),
    ("2025-04-25T00:00:00", "investment proceeds", 2_400_000, "SIM-202504-003"),
    # ── 2025-05 ──
    ("2025-05-05T00:00:00", "capital contribution",  860_000, "SIM-202505-001"),
    ("2025-05-14T00:00:00", "operating revenue",   1_500_000, "SIM-202505-002"),
    ("2025-05-25T00:00:00", "investment proceeds", 2_640_000, "SIM-202505-003"),
    # ── 2025-06 ──
    ("2025-06-05T00:00:00", "capital contribution",  870_000, "SIM-202506-001"),
    ("2025-06-14T00:00:00", "operating revenue",   1_550_000, "SIM-202506-002"),
    ("2025-06-25T00:00:00", "investment proceeds", 2_580_000, "SIM-202506-003"),
    # ── 2025-07 ──
    ("2025-07-05T00:00:00", "capital contribution",  880_000, "SIM-202507-001"),
    ("2025-07-14T00:00:00", "operating revenue",   1_600_000, "SIM-202507-002"),
    ("2025-07-25T00:00:00", "investment proceeds", 2_520_000, "SIM-202507-003"),
    # ── 2025-08 ──
    ("2025-08-05T00:00:00", "capital contribution",  890_000, "SIM-202508-001"),
    ("2025-08-14T00:00:00", "operating revenue",   1_650_000, "SIM-202508-002"),
    ("2025-08-25T00:00:00", "investment proceeds", 2_460_000, "SIM-202508-003"),
    # ── 2025-09 ──
    ("2025-09-05T00:00:00", "capital contribution",  900_000, "SIM-202509-001"),
    ("2025-09-14T00:00:00", "operating revenue",   1_700_000, "SIM-202509-002"),
    ("2025-09-25T00:00:00", "investment proceeds", 2_400_000, "SIM-202509-003"),
    # ── 2025-10 ──
    ("2025-10-05T00:00:00", "capital contribution",  910_000, "SIM-202510-001"),
    ("2025-10-14T00:00:00", "operating revenue",   1_750_000, "SIM-202510-002"),
    ("2025-10-25T00:00:00", "investment proceeds", 2_340_000, "SIM-202510-003"),
    # ── 2025-11 ──
    ("2025-11-05T00:00:00", "capital contribution",  920_000, "SIM-202511-001"),
    ("2025-11-14T00:00:00", "operating revenue",   1_500_000, "SIM-202511-002"),
    ("2025-11-25T00:00:00", "investment proceeds", 2_580_000, "SIM-202511-003"),
    # ── 2025-12 ──
    ("2025-12-05T00:00:00", "capital contribution",  930_000, "SIM-202512-001"),
    ("2025-12-14T00:00:00", "operating revenue",   1_550_000, "SIM-202512-002"),
    ("2025-12-25T00:00:00", "investment proceeds", 2_520_000, "SIM-202512-003"),
    # ── 2026-01 ──
    ("2026-01-05T00:00:00", "capital contribution",  940_000, "SIM-202601-001"),
    ("2026-01-14T00:00:00", "operating revenue",   1_600_000, "SIM-202601-002"),
    ("2026-01-25T00:00:00", "investment proceeds", 2_460_000, "SIM-202601-003"),
    # ── 2026-02 ──
    ("2026-02-05T00:00:00", "capital contribution",  950_000, "SIM-202602-001"),
    ("2026-02-14T00:00:00", "operating revenue",   1_650_000, "SIM-202602-002"),
    ("2026-02-25T00:00:00", "investment proceeds", 2_400_000, "SIM-202602-003"),
    # ── 2026-03 ──
    ("2026-03-05T00:00:00", "capital contribution",  960_000, "SIM-202603-001"),
    ("2026-03-14T00:00:00", "operating revenue",   1_700_000, "SIM-202603-002"),
    ("2026-03-25T00:00:00", "investment proceeds", 2_340_000, "SIM-202603-003"),
    # ── 2026-04 ──
    ("2026-04-05T00:00:00", "capital contribution",  970_000, "SIM-202604-001"),
    ("2026-04-14T00:00:00", "operating revenue",   1_750_000, "SIM-202604-002"),
    ("2026-04-25T00:00:00", "investment proceeds", 2_280_000, "SIM-202604-003"),
    # ── 2026-05 ──
    ("2026-05-05T00:00:00", "capital contribution",  980_000, "SIM-202605-001"),
    ("2026-05-14T00:00:00", "operating revenue",   1_500_000, "SIM-202605-002"),
    ("2026-05-25T00:00:00", "investment proceeds", 2_520_000, "SIM-202605-003"),
    # ── 2026-06 ──
    ("2026-06-05T00:00:00", "capital contribution",  990_000, "SIM-202606-001"),
    ("2026-06-14T00:00:00", "operating revenue",   1_550_000, "SIM-202606-002"),
    ("2026-06-25T00:00:00", "investment proceeds", 2_460_000, "SIM-202606-003"),
    # ── 2026-07 ──
    ("2026-07-05T00:00:00", "capital contribution",1_000_000, "SIM-202607-001"),
    ("2026-07-14T00:00:00", "operating revenue",   1_600_000, "SIM-202607-002"),
    ("2026-07-25T00:00:00", "investment proceeds", 2_400_000, "SIM-202607-003"),
    # ── 2026-08 ──
    ("2026-08-05T00:00:00", "capital contribution",1_010_000, "SIM-202608-001"),
    ("2026-08-14T00:00:00", "operating revenue",   1_650_000, "SIM-202608-002"),
    ("2026-08-25T00:00:00", "investment proceeds", 2_340_000, "SIM-202608-003"),
    # ── 2026-09 ──
    ("2026-09-05T00:00:00", "capital contribution",1_020_000, "SIM-202609-001"),
    ("2026-09-14T00:00:00", "operating revenue",   1_700_000, "SIM-202609-002"),
    ("2026-09-25T00:00:00", "investment proceeds", 2_280_000, "SIM-202609-003"),
    # ── 2026-10 ──
    ("2026-10-05T00:00:00", "capital contribution",1_030_000, "SIM-202610-001"),
    ("2026-10-14T00:00:00", "operating revenue",   1_750_000, "SIM-202610-002"),
    ("2026-10-25T00:00:00", "investment proceeds", 2_220_000, "SIM-202610-003"),
]

def _expand_seed(raw, mode):
    out = []
    for ts, desc, amount, ref in raw:
        category = _DESC_TO_CATEGORY[desc]
        provider = "simulated" if mode == "SIM" else "bank"
        out.append((ts, category, float(amount), "bank", provider, ref, ref,
                    f"{desc} — monthly credit ({mode})"))
    return out

SEED_LEDGER = _expand_seed(SEED_LEDGER_RAW, SEED_MODE)

def verify_seed_integrity():
    from collections import defaultdict
    n    = len(SEED_LEDGER_RAW)
    refs = [r[3] for r in SEED_LEDGER_RAW]
    total = sum(r[2] for r in SEED_LEDGER_RAW)
    months = defaultdict(float)
    for ts, _d, amount, _r in SEED_LEDGER_RAW:
        months[ts[:7]] += amount

    checks = {
        "row_count":         (n == EXPECTED_ROWS,            n,              EXPECTED_ROWS),
        "unique_references": (len(set(refs)) == len(refs),   len(set(refs)), len(refs)),
        "total_zar":         (abs(total - TARGET_ZAR) < 0.01,total,          TARGET_ZAR),
        "month_count":       (len(months) == EXPECTED_MONTHS,len(months),    EXPECTED_MONTHS),
        "monthly_uniform":   (all(abs(v - EXPECTED_MONTHLY) < 0.01 for v in months.values()),
                              dict(months), EXPECTED_MONTHLY),
    }
    ok = all(v[0] for v in checks.values())
    return ok, checks

# ═══════════════════════════════════════════════════════════════
# EXCEL — 5-sheet workbook, rebuilt from SQL
# ═══════════════════════════════════════════════════════════════
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="0A2540", end_color="0A2540", fill_type="solid")

def _style_header(ws):
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="left", vertical="center")

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()

        ws_master = wb.active
        ws_master.title = "All Withdrawals"
        ws_master.append(["Date", "ID", "Method", "Amount (ZAR)",
                          "Amount (USD)", "Crypto Amt", "Status", "Destination"])

        ws_pp = wb.create_sheet("PayPal")
        ws_pp.append(["Date", "ID", "Amount (ZAR)", "Amount (USD)",
                      "Status", "PayPal Email", "Note"])

        ws_led = wb.create_sheet("Ledger")
        ws_led.append(["Ledger ID", "Occurred At", "Direction", "Category",
                       "Amount (ZAR)", "Amount (USD)", "Crypto Amt", "Status",
                       "Counterparty", "Method", "Provider", "Provider Txn ID",
                       "Reference", "Trade ID", "Withdrawal ID", "Note", "Created At"])

        ws_src = wb.create_sheet("Ledger By Source")
        ws_src.append(["Direction", "Category", "Count", "Total (ZAR)"])

        ws_run = wb.create_sheet("Running Balance")
        ws_run.append(["Occurred At", "Ledger ID", "Direction", "Category",
                       "Amount (ZAR)", "Running Balance (ZAR)"])

        for ws in wb.worksheets:
            _style_header(ws)
        wb.save(EXCEL_FILE)
        print(f"✅ Created {EXCEL_FILE}")
    else:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        added = False
        for name, headers in [
            ("All Withdrawals", ["Date","ID","Method","Amount (ZAR)","Amount (USD)",
                                 "Crypto Amt","Status","Destination"]),
            ("PayPal", ["Date","ID","Amount (ZAR)","Amount (USD)","Status",
                        "PayPal Email","Note"]),
            ("Ledger", ["Ledger ID","Occurred At","Direction","Category","Amount (ZAR)",
                        "Amount (USD)","Crypto Amt","Status","Counterparty","Method",
                        "Provider","Provider Txn ID","Reference","Trade ID",
                        "Withdrawal ID","Note","Created At"]),
            ("Ledger By Source", ["Direction","Category","Count","Total (ZAR)"]),
            ("Running Balance", ["Occurred At","Ledger ID","Direction","Category",
                                 "Amount (ZAR)","Running Balance (ZAR)"]),
        ]:
            if name not in wb.sheetnames:
                ws = wb.create_sheet(name)
                ws.append(headers)
                _style_header(ws)
                added = True
        if added:
            wb.save(EXCEL_FILE)
            print(f"✅ Added missing sheets to existing {EXCEL_FILE}")

init_excel()

def log_withdrawal_to_excel(wid, amount_zar, amount_usd, method,
                            crypto_amount, status, destination, note=""):
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        wb["All Withdrawals"].append([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"), wid, method,
            round(amount_zar, 2), round(amount_usd, 2),
            round(crypto_amount, 8) if crypto_amount else 0,
            status, destination,
        ])
        if method == "paypal":
            wb["PayPal"].append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"), wid,
                round(amount_zar, 2), round(amount_usd, 2),
                status, destination, note,
            ])
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        print(f"⚠️ Excel withdrawal write failed: {e}")
        return False

def log_ledger_to_excel(row_dict):
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        wb["Ledger"].append([
            row_dict["id"], row_dict["occurred_at"], row_dict["direction"],
            row_dict["category"], row_dict["amount_zar"], row_dict["amount_usd"],
            row_dict["crypto_amount"], row_dict["status"], row_dict["counterparty"],
            row_dict["method"], row_dict["provider"], row_dict["provider_txn_id"],
            row_dict["reference"], row_dict["trade_id"], row_dict["withdrawal_id"],
            row_dict["note"], row_dict["created_at"],
        ])
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        print(f"⚠️ Excel ledger write failed: {e}")
        return False

def rebuild_ledger_excel_sheets():
    """Regenerate Ledger / Ledger By Source / Running Balance from SQL."""
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        rows = query("SELECT * FROM ledger ORDER BY occurred_at ASC, id ASC")
        row_dicts = [dict(r) for r in rows]

        ws = wb["Ledger"]
        ws.delete_rows(2, ws.max_row)
        for r in row_dicts:
            ws.append([
                r["id"], r["occurred_at"], r["direction"], r["category"],
                r["amount_zar"], r["amount_usd"], r["crypto_amount"],
                r["status"], r["counterparty"], r["method"], r["provider"],
                r["provider_txn_id"], r["reference"], r["trade_id"],
                r["withdrawal_id"], r["note"], r["created_at"],
            ])

        ws_src = wb["Ledger By Source"]
        ws_src.delete_rows(2, ws_src.max_row)
        for r in ledger_by_category():
            ws_src.append([r["direction"], r["category"], r["n"], r["total_zar"]])

        ws_run = wb["Running Balance"]
        ws_run.delete_rows(2, ws_run.max_row)
        running = 0.0
        for r in row_dicts:
            if r["status"] != "completed":
                continue
            amt = r["amount_zar"] if r["direction"] == "CREDIT" else -r["amount_zar"]
            running += amt
            ws_run.append([r["occurred_at"], r["id"], r["direction"],
                           r["category"], r["amount_zar"], round(running, 2)])

        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        print(f"⚠️ rebuild_ledger_excel_sheets failed: {e}")
        return False

# ═══════════════════════════════════════════════════════════════
# SEED LOADER
# ═══════════════════════════════════════════════════════════════
def seed_ledger_from_list(rows, wipe_first=False):
    if wipe_first:
        execute("DELETE FROM ledger")
    total = 0.0
    for (occurred_at, category, amount_zar, method, provider,
         provider_txn_id, reference, note) in rows:
        direction = "CREDIT" if category in CREDIT_CATEGORIES else "DEBIT"
        lid = ledger_post(
            direction=direction, category=category, amount_zar=amount_zar,
            occurred_at=occurred_at, method=method, provider=provider,
            provider_txn_id=provider_txn_id, reference=reference,
            note=note, status="completed",
        )
        row = query("SELECT * FROM ledger WHERE id=?", (lid,), one=True)
        log_ledger_to_excel(dict(row))
        total += amount_zar if direction == "CREDIT" else -amount_zar
    rebuild_ledger_excel_sheets()
    return total

# ═══════════════════════════════════════════════════════════════
# API ROUTES
# ═══════════════════════════════════════════════════════════════
@app.route('/health')
def health():
    bal = ledger_balance()
    return jsonify({
        "status": "ok", "mode": BROKER_MODE, "currency": BASE_CURRENCY,
        "crypto": CRYPTO_SYMBOL, "time": datetime.utcnow().isoformat(),
        "ledger_balance_zar": bal["completed_balance_zar"],
        "available_zar": bal["available_zar"],
    })

@app.route('/api/summary')
def summary():
    bal = ledger_balance()
    pnl_usd_total = query("SELECT COALESCE(SUM(pnl_usd),0) s FROM trades WHERE status='closed'", one=True)["s"]
    pnl_zar_total = query("SELECT COALESCE(SUM(pnl_zar),0) s FROM trades WHERE status='closed'", one=True)["s"]
    open_count    = query("SELECT COUNT(*) c FROM trades WHERE status='open'", one=True)["c"]
    closed_count  = query("SELECT COUNT(*) c FROM trades WHERE status='closed'", one=True)["c"]
    wins          = query("SELECT COUNT(*) c FROM trades WHERE status='closed' AND pnl_zar>0", one=True)["c"]
    losses        = query("SELECT COUNT(*) c FROM trades WHERE status='closed' AND pnl_zar<0", one=True)["c"]
    withdrawn_zar = query("""
        SELECT COALESCE(SUM(amount_zar),0) s FROM ledger
        WHERE category='withdrawal' AND direction='DEBIT' AND status='completed'
    """, one=True)["s"]
    win_rate = round((wins / closed_count * 100), 1) if closed_count else 0

    return jsonify({
        "balance_zar":         bal["completed_balance_zar"],
        "available_zar":       bal["available_zar"],
        "pending_credits_zar": bal["pending_credits_zar"],
        "pending_debits_zar":  bal["pending_debits_zar"],
        "withdrawn_zar":       round(withdrawn_zar, 2),
        "start_balance_zar":   START_BALANCE_ZAR,
        "total_pnl_zar":       round(pnl_zar_total, 2),
        "total_pnl_usd":       round(pnl_usd_total, 2),
        "open_trades": open_count, "closed_trades": closed_count,
        "wins": wins, "losses": losses, "win_rate": win_rate,
        "broker_mode": BROKER_MODE,
        "crypto_equiv_btc": round(zar_to_btc(bal["completed_balance_zar"]), 8),
        "crypto_symbol": CRYPTO_SYMBOL,
        "zar_per_btc": ZAR_PER_BTC, "zar_per_usd": ZAR_PER_USD,
    })

@app.route('/api/trades')
def trades():
    rows = query("SELECT * FROM trades ORDER BY id DESC LIMIT 200")
    return jsonify([dict(r) for r in rows])

@app.route('/api/manual_trade', methods=['POST'])
def manual_trade():
    data = request.json or {}
    if data.get("signal") not in ("LONG", "SHORT"):
        return jsonify({"status": "rejected", "reason": "invalid signal"}), 400
    price = float(data.get("price", 2065.00))
    units = int(data.get("units", 10))
    if data["signal"] == "LONG":
        sl, tp = round(price - 5, 2), round(price + 10, 2)
    else:
        sl, tp = round(price + 5, 2), round(price - 10, 2)
    execute("""INSERT INTO trades (timestamp, signal, price, sl, tp, units, source)
               VALUES (?,?,?,?,?,?,?)""",
            (datetime.utcnow().isoformat(), data["signal"], price, sl, tp, units, "manual"))
    return jsonify({"status": "created"})

@app.route('/signal', methods=['POST'])
def signal():
    data = request.json or {}
    if data.get("signal") not in ("LONG", "SHORT"):
        return jsonify({"status": "rejected"}), 400
    price = float(data.get("price", 0))
    if price <= 0:
        return jsonify({"status": "rejected", "reason": "invalid price"}), 400
    execute("""INSERT INTO trades (timestamp, signal, price, sl, tp, units, source)
               VALUES (?,?,?,?,?,?,?)""",
            (datetime.utcnow().isoformat(), data["signal"], price,
             float(data.get("sl", 0)), float(data.get("tp", 0)),
             int(data.get("units", 10)), "indicator"))
    return jsonify({"status": "logged"})

@app.route('/api/close_trade/<int:trade_id>', methods=['POST'])
def close_trade(trade_id):
    data = request.json or {}
    close_price = float(data.get("close_price", 0))
    trade = query("SELECT * FROM trades WHERE id=?", (trade_id,), one=True)
    if not trade:                 return jsonify({"status": "not found"}), 404
    if trade["status"] == "closed": return jsonify({"status": "already closed"}), 400
    if close_price <= 0:          return jsonify({"status": "invalid close_price"}), 400

    units = trade["units"]
    pnl_usd = ((close_price - trade["price"]) if trade["signal"] == "LONG"
               else (trade["price"] - close_price)) * units
    pnl_zar = usd_to_zar(pnl_usd)

    execute("""UPDATE trades SET status='closed', pnl_usd=?, pnl_zar=?, close_price=?
               WHERE id=?""", (round(pnl_usd, 2), round(pnl_zar, 2), close_price, trade_id))

    if abs(pnl_zar) >= 0.01:
        ledger_post(
            direction="CREDIT" if pnl_zar > 0 else "DEBIT",
            category="trade_profit" if pnl_zar > 0 else "trade_loss",
            amount_zar=abs(pnl_zar), amount_usd=abs(pnl_usd),
            provider="oanda", method="internal",
            reference=f"trade#{trade_id}", trade_id=trade_id,
            note=f"Closed {trade['signal']} @ {close_price:.2f}",
        )
        rebuild_ledger_excel_sheets()

    return jsonify({"status": "closed",
                    "pnl_usd": round(pnl_usd, 2),
                    "pnl_zar": round(pnl_zar, 2),
                    "ledger_balance_zar": ledger_balance()["completed_balance_zar"]})

# ═══════════════════════════════════════════════════════════════
# WITHDRAWALS — debit ledger, provider call TODO
# ═══════════════════════════════════════════════════════════════
@app.route('/api/withdraw', methods=['POST'])
def withdraw():
    data = request.json or {}
    amount_zar = float(data.get("amount_zar", 0))
    method = data.get("method", "").lower()

    if amount_zar <= 0:
        return jsonify({"status": "error", "reason": "invalid amount"}), 400
    if method not in ("paypal", "luno", "bank", "btc"):
        return jsonify({"status": "error", "reason": "invalid method"}), 400

    bal = ledger_balance()
    if amount_zar > bal["available_zar"]:
        return jsonify({"status": "rejected",
                        "reason": f"Insufficient ledger balance. Available: {fmt_zar(bal['available_zar'])}"}), 400

    usd_amount = round(amount_zar / ZAR_PER_USD, 2)
    crypto_amount = 0.0
    destination = note = ""
    provider = method

    if method == "paypal":
        destination = BROKER_CONFIG["PAYPAL_RECEIVER_EMAIL"] or BROKER_CONFIG["PAYPAL_EMAIL"]
        if not destination or "@" not in destination:
            return jsonify({"status": "error",
                            "reason": "PAYPAL_RECEIVER_EMAIL must be a real email."}), 400
        note = f"Broker payout ${usd_amount:,.2f} to {destination}"
    elif method == "luno":
        crypto_amount = zar_to_btc(amount_zar)
        destination = BROKER_CONFIG["LUNO_RECEIVER_ACCOUNT"] or "SET_LUNO_RECEIVER_ENV"
        note = f"Luno: {crypto_amount:.8f} BTC to {destination}"
    elif method == "btc":
        crypto_amount = zar_to_btc(amount_zar)
        destination = BROKER_CONFIG["BTC_WALLET"]
        if not destination:
            return jsonify({"status": "error", "reason": "BTC_WALLET must be set"}), 400
        note = f"BTC: {crypto_amount:.8f} to {destination}"
    elif method == "bank":
        destination = BROKER_CONFIG["BANK_ACCOUNT"]
        if not destination:
            return jsonify({"status": "error", "reason": "BANK_ACCOUNT must be set"}), 400
        note = f"SA Bank EFT: {fmt_zar(amount_zar)} to {destination}"

    # ── TODO: LIVE PROVIDER CALL ───────────────────────────────────
    # Before inserting the withdrawal, call the real provider here and
    # capture the returned ID. Example shapes:
    #
    #   if method == "paypal":
    #       batch = paypal_payout(destination, usd_amount, note, ref)
    #       provider_txn_id = batch["batch_header"]["payout_batch_id"]
    #       provider_status = batch["batch_header"]["batch_status"].lower()
    #   elif method == "luno":
    #       resp = luno_withdraw_btc(destination, crypto_amount, ref)
    #       provider_txn_id = resp["withdrawal_id"]
    #       provider_status = resp["status"].lower()
    #   elif method == "btc":
    #       resp = btc_send(destination, crypto_amount)
    #       provider_txn_id = resp["txid"]
    #       provider_status = "pending"
    #
    # On HTTPError: return 502 with the provider error text; do NOT insert
    # the withdrawal or the ledger debit.
    # ───────────────────────────────────────────────────────────────
    provider_txn_id = ""       # populated by the real provider call above
    provider_status = "pending"

    # 1) withdrawal record
    execute("""INSERT INTO withdrawals
               (timestamp, amount_zar, method, status, note, crypto_amount, pay_link)
               VALUES (?,?,?,?,?,?,?)""",
            (datetime.utcnow().isoformat(), amount_zar, method,
             provider_status, note, crypto_amount, ""))
    wid = query("SELECT last_insert_rowid() AS id", one=True)["id"]

    # 2) matching DEBIT on the ledger
    lid = ledger_post(
        direction="DEBIT", category="withdrawal", amount_zar=amount_zar,
        amount_usd=usd_amount, crypto_amount=crypto_amount,
        status=provider_status, counterparty=destination, method=method,
        provider=provider, provider_txn_id=provider_txn_id,
        withdrawal_id=wid, note=note,
    )

    # 3) Excel
    log_withdrawal_to_excel(wid, amount_zar, usd_amount, method,
                            crypto_amount, provider_status, destination, note)
    lrow = query("SELECT * FROM ledger WHERE id=?", (lid,), one=True)
    log_ledger_to_excel(dict(lrow))

    return jsonify({
        "status": provider_status,
        "withdrawal_id": wid,
        "ledger_id": lid,
        "amount_zar": round(amount_zar, 2),
        "amount_usd": usd_amount,
        "crypto_amount": round(crypto_amount, 8),
        "method": method, "note": note, "destination": destination,
        "provider_txn_id": provider_txn_id,
        "available_after_zar": ledger_balance()["available_zar"],
        "broker_mode": BROKER_MODE,
    })

@app.route('/api/paypal_me_link')
def paypal_me_link():
    username = BROKER_CONFIG["PAYPAL_ME_USERNAME"].strip().lstrip("@").replace(" ", "")
    if not username:
        return jsonify({"status": "unconfigured",
                        "message": "Set PAYPAL_ME_USERNAME on Render to enable"})
    if "@" in username:
        return jsonify({"status": "invalid",
                        "message": "PAYPAL_ME_USERNAME must be a handle, not an email"})
    return jsonify({"status": "ok", "link": f"https://paypal.me/{username}",
                    "username": username,
                    "note": "Share with clients to receive payments"})

@app.route('/api/withdrawals')
def withdrawals():
    rows = query("SELECT * FROM withdrawals ORDER BY id DESC LIMIT 100")
    return jsonify([dict(r) for r in rows])

@app.route('/api/withdraw/<int:wid>/complete', methods=['POST'])
def complete_withdrawal(wid):
    data = request.json or {}
    provider_txn_id = data.get("provider_txn_id", "")
    execute("UPDATE withdrawals SET status='completed' WHERE id=?", (wid,))
    execute("""UPDATE ledger
                 SET status='completed',
                     provider_txn_id=CASE WHEN ?<>'' THEN ? ELSE provider_txn_id END
               WHERE withdrawal_id=? AND category='withdrawal' AND direction='DEBIT'""",
            (provider_txn_id, provider_txn_id, wid))
    rebuild_ledger_excel_sheets()
    return jsonify({"status": "completed", "provider_txn_id": provider_txn_id})

@app.route('/api/withdraw/<int:wid>/cancel', methods=['POST'])
def cancel_withdrawal(wid):
    execute("UPDATE withdrawals SET status='cancelled' WHERE id=?", (wid,))
    execute("""UPDATE ledger SET status='cancelled'
               WHERE withdrawal_id=? AND category='withdrawal' AND direction='DEBIT'""",
            (wid,))
    rebuild_ledger_excel_sheets()
    return jsonify({"status": "cancelled"})

# ═══════════════════════════════════════════════════════════════
# LEDGER READS
# ═══════════════════════════════════════════════════════════════
@app.route('/api/ledger')
def api_ledger():
    limit = int(request.args.get("limit", 200))
    rows = query("SELECT * FROM ledger ORDER BY occurred_at DESC, id DESC LIMIT ?",
                 (limit,))
    return jsonify([dict(r) for r in rows])

@app.route('/api/ledger/by_category')
def api_ledger_by_category():
    return jsonify(ledger_by_category())

# ═══════════════════════════════════════════════════════════════
# ADMIN
# ═══════════════════════════════════════════════════════════════
def _admin_guard():
    token = BROKER_CONFIG["ADMIN_TOKEN"]
    if token and request.headers.get("X-Admin-Token") != token:
        return jsonify({"status": "forbidden"}), 403
    return None

@app.route('/admin/seed_ledger', methods=['POST'])
def admin_seed_ledger():
    g = _admin_guard()
    if g: return g
    ok, checks = verify_seed_integrity()
    if not ok:
        return jsonify({"status": "refused",
                        "reason": "seed failed integrity", "checks": checks}), 400
    wipe = bool((request.json or {}).get("wipe", False))
    net = seed_ledger_from_list(SEED_LEDGER, wipe_first=wipe)
    bal = ledger_balance()
    return jsonify({"status": "seeded", "mode": SEED_MODE,
                    "rows_inserted": len(SEED_LEDGER), "net_zar": round(net, 2),
                    "ledger_balance_zar": bal["completed_balance_zar"],
                    "available_zar": bal["available_zar"],
                    "checks": checks})

@app.route('/admin/ledger_audit')
def admin_ledger_audit():
    ok, checks = verify_seed_integrity()
    sql_by_month = query("""
        SELECT substr(occurred_at,1,7) AS month, COUNT(*) AS n,
               ROUND(SUM(amount_zar),2) AS total_zar
        FROM ledger WHERE status='completed' AND direction='CREDIT'
        GROUP BY month ORDER BY month
    """)
    sql_total = query("""
        SELECT ROUND(COALESCE(SUM(amount_zar),0),2) t
        FROM ledger WHERE status='completed' AND direction='CREDIT'
    """, one=True)["t"]
    return jsonify({"seed_checks": checks, "seed_ok": ok,
                    "sql_monthly": [dict(r) for r in sql_by_month],
                    "sql_total_zar": sql_total,
                    "balance": ledger_balance(),
                    "by_category": ledger_by_category()})

@app.route('/admin/rebuild_excel', methods=['POST'])
def admin_rebuild_excel():
    g = _admin_guard()
    if g: return g
    ok = rebuild_ledger_excel_sheets()
    return jsonify({"status": "rebuilt" if ok else "failed"})

@app.route('/download/excel')
def download_excel():
    if not os.path.exists(EXCEL_FILE):
        return jsonify({"status": "error", "reason": "no excel file yet"}), 404
    with open(EXCEL_FILE, "rb") as f:
        data = f.read()
    return Response(
        data,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=withdrawals.xlsx"},
    )

@app.route('/api/reset', methods=['POST'])
def reset():
    g = _admin_guard()
    if g: return g
    data = request.json or {}
    execute("DELETE FROM trades")
    execute("DELETE FROM withdrawals")
    execute("DELETE FROM ledger")
    if data.get("reseed_ledger"):
        seed_ledger_from_list(SEED_LEDGER, wipe_first=False)
    rebuild_ledger_excel_sheets()
    return jsonify({"status": "reset", "reseeded": bool(data.get("reseed_ledger")),
                    "balance": ledger_balance()})

# ═══════════════════════════════════════════════════════════════
# DASHBOARD HTML
# ═══════════════════════════════════════════════════════════════
DASHBOARD_HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>💰 Crypto Trading Dashboard (ZAR)</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0a1929;color:#e6edf3;font-family:Inter,system-ui,sans-serif;padding:24px;min-height:100vh}
header{display:flex;justify-content:space-between;align-items:center;margin-bottom:28px;padding-bottom:16px;border-bottom:2px solid #1e4976;flex-wrap:wrap;gap:12px}
h1{color:#FFD700;font-size:1.5rem}
h2{color:#00A8A8;font-size:1.1rem;margin-bottom:14px}
.status{padding:6px 14px;border-radius:20px;background:#164b2c;color:#32CD32;font-size:.85rem;font-weight:600}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:14px;margin-bottom:24px}
.card{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:18px}
.card.highlight{border-color:#FFD700;box-shadow:0 0 20px rgba(255,215,0,.15)}
.card.crypto{border-color:#00b5e2;box-shadow:0 0 20px rgba(0,181,226,.15)}
.card.pending{border-color:#f7931a;box-shadow:0 0 20px rgba(247,147,26,.15)}
.label{color:#8899aa;font-size:.72rem;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
.metric{font-size:1.35rem;font-weight:700;color:#00A8A8}
.metric.small{font-size:1rem;font-weight:600;color:#e6edf3}
.positive{color:#32CD32!important}
.negative{color:#FF4444!important}
.gold{color:#FFD700!important}
.btc{color:#00b5e2!important}
.orange{color:#f7931a!important}
.actions{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:24px}
@media(max-width:900px){.actions{grid-template-columns:1fr}body{padding:16px}}
.panel{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:20px}
.panel.receive{border-color:#00b5e2;box-shadow:0 0 15px rgba(0,181,226,.1)}
.hint{color:#8899aa;font-size:.8rem;margin-bottom:14px}
.row{display:flex;flex-wrap:wrap;gap:10px;align-items:center}
input,select,button{padding:10px 14px;border-radius:8px;border:1px solid #1e4976;background:#0a1929;color:#e6edf3;font-size:.9rem;font-family:inherit}
input{width:140px}
button{cursor:pointer;font-weight:600;transition:.2s}
button:hover{opacity:.85}
.btn-buy{background:#164b2c;color:#32CD32;border-color:#32CD32}
.btn-sell{background:#4b1616;color:#FF4444;border-color:#FF4444}
.btn-small{padding:6px 10px;font-size:.75rem;background:#00A8A8;color:#0a1929;border:none;margin-left:4px}
.btn-danger{background:#4b1616;color:#FF4444;border-color:#FF4444}
.btn-paypal{background:#003087;color:#fff;border:none}
.btn-luno{background:#00b5e2;color:#000;border:none}
.btn-bank{background:#FFD700;color:#0a1929;border:none}
.btn-btc{background:#f7931a;color:#000;border:none}
.btn-copy{background:#00b5e2;color:#000;border:none}
.btn-excel{background:#1D6F42;color:#fff;border:none;padding:10px 16px;border-radius:8px;font-weight:600;cursor:pointer;text-decoration:none;display:inline-block}
.msg{margin-top:12px;font-weight:600;font-size:.85rem;word-break:break-word}
.msg a{color:#FFD700;font-weight:700}
.table-section{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:20px;margin-bottom:20px;overflow-x:auto}
table{width:100%;border-collapse:collapse;min-width:900px}
th{text-align:left;padding:10px;color:#00A8A8;font-size:.72rem;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #1e4976}
td{padding:9px;border-bottom:1px solid #1e4976;font-size:.83rem}
tr:hover{background:#143557}
footer{margin-top:32px;padding-top:16px;border-top:2px solid #1e4976;display:flex;align-items:center;gap:16px;flex-wrap:wrap}
.footer-note{color:#8899aa;font-size:.8rem}
.withdraw-buttons{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}
.receive-box{display:flex;gap:10px;flex-wrap:wrap;margin-top:10px}
.receive-box input{flex:1;min-width:220px;font-family:monospace;font-size:.85rem}
.tabs{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap}
.tab{padding:8px 14px;border-radius:8px;background:#0a1929;border:1px solid #1e4976;cursor:pointer;font-size:.85rem;font-weight:600}
.tab.active{background:#00A8A8;color:#0a1929;border-color:#00A8A8}
</style>
</head>
<body>

<header>
  <h1>💰 Crypto Trading Dashboard <span style="font-size:.75rem;color:#8899aa">(ZAR · Ledger-Driven)</span></h1>
  <div class="status" id="status">● Loading...</div>
</header>

<section class="grid">
  <div class="card highlight"><div class="label">Ledger Balance (ZAR)</div><div class="metric gold" id="balance">R0.00</div></div>
  <div class="card"><div class="label">Available (ZAR)</div><div class="metric gold" id="available">R0.00</div></div>
  <div class="card pending"><div class="label">Pending Payouts (ZAR)</div><div class="metric orange" id="pending-debits">R0.00</div></div>
  <div class="card crypto"><div class="label">Crypto Equiv (BTC)</div><div class="metric btc" id="crypto-eq">0.00000000</div></div>
  <div class="card"><div class="label">Total P&L (ZAR)</div><div class="metric" id="pnl-zar">R0.00</div></div>
  <div class="card"><div class="label">Total P&L (USD)</div><div class="metric small" id="pnl-usd">$0.00</div></div>
  <div class="card"><div class="label">Win Rate</div><div class="metric" id="winrate">0%</div></div>
  <div class="card"><div class="label">Open Trades</div><div class="metric" id="open">0</div></div>
  <div class="card"><div class="label">Withdrawn (ZAR)</div><div class="metric" id="withdrawn">R0.00</div></div>
</section>

<section class="actions">
  <div class="panel">
    <h2>🧪 Simulate Trade</h2>
    <p class="hint">XAUUSD price · units · P&L auto-converted to ZAR and posted to ledger on close</p>
    <div class="row">
      <input type="number" id="sim-price" step="0.01" value="2065.00"/>
      <input type="number" id="sim-units" step="1" value="10"/>
    </div>
    <div class="row" style="margin-top:10px">
      <button onclick="simulateTrade('LONG')" class="btn-buy">BUY (LONG)</button>
      <button onclick="simulateTrade('SHORT')" class="btn-sell">SELL (SHORT)</button>
    </div>
  </div>

  <div class="panel">
    <h2>💸 Withdraw (Broker → You)</h2>
    <p class="hint">Debits the ledger. Provider call is stubbed — will be wired before go-live.</p>
    <div class="row">
      <input type="number" id="wd-amount" step="0.01" placeholder="Amount in R"/>
    </div>
    <div class="withdraw-buttons">
      <button onclick="requestWithdrawal('btc')" class="btn-btc">₿ Bitcoin</button>
      <button onclick="requestWithdrawal('luno')" class="btn-luno">💠 Luno ZAR</button>
      <button onclick="requestWithdrawal('paypal')" class="btn-paypal">💳 PayPal</button>
      <button onclick="requestWithdrawal('bank')" class="btn-bank">🏦 SA Bank EFT</button>
    </div>
    <div class="msg" id="wd-msg"></div>
    <div style="margin-top:14px">
      <a href="/download/excel" class="btn-excel">📊 Download Excel</a>
    </div>
  </div>
</section>

<section class="actions">
  <div class="panel receive" style="grid-column:1 / -1;">
    <h2>💰 Receive Payment (From Clients)</h2>
    <p class="hint">Share this link so clients can pay you. (This is NOT for withdrawing from broker.)</p>
    <div class="receive-box">
      <input type="text" id="paypal-me-link" readonly placeholder="Loading..."/>
      <button onclick="copyLink()" class="btn-copy">📋 Copy Link</button>
      <a id="paypal-me-open" href="#" target="_blank" class="btn-excel" style="background:#003087;">→ Open</a>
    </div>
    <div class="msg" id="me-msg"></div>
  </div>
</section>

<section class="table-section">
  <h2>📈 Open Trades</h2>
  <table id="open-trades">
    <thead><tr><th>ID</th><th>Time</th><th>Signal</th><th>Price</th><th>SL</th><th>TP</th><th>Units</th><th>Action</th></tr></thead>
    <tbody></tbody>
  </table>
</section>

<section class="table-section">
  <h2>📊 Closed Trades</h2>
  <table id="closed-trades">
    <thead><tr><th>ID</th><th>Time</th><th>Signal</th><th>Price</th><th>Close</th><th>P&L USD</th><th>P&L ZAR</th></tr></thead>
    <tbody></tbody>
  </table>
</section>

<section class="table-section">
  <h2>💸 Withdrawal History</h2>
  <table id="wd-history">
    <thead><tr><th>ID</th><th>Time</th><th>Amount (ZAR)</th><th>Crypto Amt</th><th>Method</th><th>Status</th><th>Action</th></tr></thead>
    <tbody></tbody>
  </table>
</section>

<section class="table-section">
  <h2>📒 Ledger — How the balance was accumulated</h2>
  <div class="tabs">
    <div class="tab active" data-tab="recent" onclick="switchLedgerTab('recent')">Recent 200</div>
    <div class="tab" data-tab="bycategory" onclick="switchLedgerTab('bycategory')">By Source</div>
  </div>
  <div id="ledger-view">
    <table id="ledger-table">
      <thead><tr>
        <th>ID</th><th>Occurred</th><th>Dir</th><th>Category</th>
        <th>Amount (ZAR)</th><th>USD</th><th>Crypto</th>
        <th>Status</th><th>Counterparty</th><th>Provider Txn</th><th>Note</th>
      </tr></thead>
      <tbody></tbody>
    </table>
  </div>
  <div id="ledger-bycat" style="display:none">
    <table>
      <thead><tr><th>Direction</th><th>Category</th><th>Count</th><th>Total (ZAR)</th></tr></thead>
      <tbody id="ledger-bycat-tbody"></tbody>
    </table>
  </div>
</section>

<footer>
  <button onclick="resetAll()" class="btn-danger">🔄 Reset All</button>
  <span class="footer-note" id="mode-note">Loading…</span>
</footer>

<script>
const fmtZAR = n => 'R' + Number(n).toLocaleString('en-ZA',{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtUSD = n => '$' + Number(n).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtBTC = n => Number(n).toFixed(8);

function setText(id,v){document.getElementById(id).textContent=v}
function setColor(id,txt,val){
  const el=document.getElementById(id);
  el.textContent=txt;
  el.classList.remove('positive','negative');
  if(val>0)el.classList.add('positive');
  if(val<0)el.classList.add('negative');
}
async function j(url, opts){return await (await fetch(url, opts||{})).json()}

let ledgerTab = 'recent';
function switchLedgerTab(t){
  ledgerTab = t;
  document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x.dataset.tab === t));
  document.getElementById('ledger-view').style.display = t === 'recent' ? '' : 'none';
  document.getElementById('ledger-bycat').style.display = t === 'bycategory' ? '' : 'none';
  refreshLedger();
}

async function refreshLedger(){
  if(ledgerTab === 'recent'){
    const rows = await j('/api/ledger?limit=200');
    const tb = document.querySelector('#ledger-table tbody');
    tb.innerHTML = '';
    for(const r of rows){
      const cls = r.direction === 'CREDIT' ? 'positive' : 'negative';
      tb.innerHTML += `<tr>
        <td>${r.id}</td>
        <td>${r.occurred_at.slice(0,16).replace('T',' ')}</td>
        <td class="${cls}">${r.direction}</td>
        <td>${r.category}</td>
        <td class="${cls}">${fmtZAR(r.amount_zar)}</td>
        <td>${r.amount_usd ? fmtUSD(r.amount_usd) : '—'}</td>
        <td>${r.crypto_amount ? fmtBTC(r.crypto_amount) : '—'}</td>
        <td>${r.status}</td>
        <td>${r.counterparty || '—'}</td>
        <td>${r.provider_txn_id || '—'}</td>
        <td>${(r.note||'').slice(0,50)}</td>
      </tr>`;
    }
  } else {
    const rows = await j('/api/ledger/by_category');
    const tb = document.getElementById('ledger-bycat-tbody');
    tb.innerHTML = '';
    for(const r of rows){
      const cls = r.direction === 'CREDIT' ? 'positive' : 'negative';
      tb.innerHTML += `<tr>
        <td class="${cls}">${r.direction}</td>
        <td>${r.category}</td>
        <td>${r.n}</td>
        <td class="${cls}">${fmtZAR(r.total_zar)}</td>
      </tr>`;
    }
  }
}

async function refreshAll(){
  const s = await j('/api/summary');
  setText('balance', fmtZAR(s.balance_zar));
  setText('available', fmtZAR(s.available_zar));
  setText('pending-debits', fmtZAR(s.pending_debits_zar));
  setText('crypto-eq', fmtBTC(s.crypto_equiv_btc) + ' ' + s.crypto_symbol);
  setColor('pnl-zar', fmtZAR(s.total_pnl_zar), s.total_pnl_zar);
  setColor('pnl-usd', fmtUSD(s.total_pnl_usd), s.total_pnl_usd);
  setText('winrate', s.win_rate + '%');
  setText('open', s.open_trades);
  setText('withdrawn', fmtZAR(s.withdrawn_zar));
  setText('mode-note', `Mode: ${s.broker_mode} · Start: ${fmtZAR(s.start_balance_zar)} · R${s.zar_per_btc.toLocaleString()}/BTC · R${s.zar_per_usd}/USD`);
  document.getElementById('status').textContent = '● Live';

  const trades = await j('/api/trades');
  const ob = document.querySelector('#open-trades tbody');
  const cb = document.querySelector('#closed-trades tbody');
  ob.innerHTML = ''; cb.innerHTML = '';
  for(const t of trades){
    if(t.status === 'open'){
      ob.innerHTML += `<tr>
        <td>${t.id}</td><td>${t.timestamp.slice(0,16).replace('T',' ')}</td>
        <td class="${t.signal==='LONG'?'positive':'negative'}">${t.signal}</td>
        <td>$${t.price.toFixed(2)}</td><td>$${t.sl.toFixed(2)}</td>
        <td>$${t.tp.toFixed(2)}</td><td>${t.units}</td>
        <td><input type="number" id="close-${t.id}" step="0.01" placeholder="Close" style="width:80px;padding:6px"/>
            <button class="btn-small" onclick="closeTrade(${t.id})">Close</button></td>
      </tr>`;
    } else {
      cb.innerHTML += `<tr>
        <td>${t.id}</td><td>${t.timestamp.slice(0,16).replace('T',' ')}</td>
        <td class="${t.signal==='LONG'?'positive':'negative'}">${t.signal}</td>
        <td>$${t.price.toFixed(2)}</td><td>$${t.close_price.toFixed(2)}</td>
        <td class="${t.pnl_usd>=0?'positive':'negative'}">$${t.pnl_usd.toFixed(2)}</td>
        <td class="${t.pnl_zar>=0?'positive':'negative'}">${fmtZAR(t.pnl_zar)}</td>
      </tr>`;
    }
  }

  const wds = await j('/api/withdrawals');
  const wb = document.querySelector('#wd-history tbody');
  wb.innerHTML = '';
  for(const w of wds){
    let actions = '';
    if(w.status === 'pending'){
      actions = `<button class="btn-small" onclick="completeWd(${w.id})">Complete</button>
                 <button class="btn-small btn-danger" onclick="cancelWd(${w.id})">Cancel</button>`;
    } else {
      actions = w.status === 'completed' ? '✅' : (w.status === 'cancelled' ? '❌' : w.status);
    }
    const cryptoAmt = w.crypto_amount > 0 ? fmtBTC(w.crypto_amount) + ' BTC' : '—';
    wb.innerHTML += `<tr>
      <td>${w.id}</td><td>${w.timestamp.slice(0,16).replace('T',' ')}</td>
      <td>${fmtZAR(w.amount_zar)}</td><td>${cryptoAmt}</td>
      <td>${w.method}</td><td>${w.status}</td><td>${actions}</td>
    </tr>`;
  }

  await refreshLedger();
}

async function loadPayPalMe(){
  const r = await j('/api/paypal_me_link');
  const input = document.getElementById('paypal-me-link');
  const openBtn = document.getElementById('paypal-me-open');
  if(r.status === 'ok'){
    input.value = r.link;
    openBtn.href = r.link;
    openBtn.style.display = 'inline-block';
  } else {
    input.value = r.message || 'Not configured';
    openBtn.style.display = 'none';
  }
}

function copyLink(){
  const input = document.getElementById('paypal-me-link');
  if(!input.value || input.value.startsWith('Set ') || input.value.includes('not')){
    const msg = document.getElementById('me-msg');
    msg.textContent = '⚠️ Configure PAYPAL_ME_USERNAME on Render first';
    msg.className = 'msg negative';
    return;
  }
  input.select();
  document.execCommand('copy');
  const msg = document.getElementById('me-msg');
  msg.textContent = '✅ Link copied to clipboard!';
  msg.className = 'msg positive';
  setTimeout(() => { msg.textContent = ''; }, 2500);
}

async function simulateTrade(signal){
  const price = parseFloat(document.getElementById('sim-price').value);
  const units = parseInt(document.getElementById('sim-units').value);
  if(!price || price <= 0){alert('Enter valid price');return}
  await j('/api/manual_trade', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({signal, price, units})
  });
  refreshAll();
}

async function closeTrade(id){
  const price = parseFloat(document.getElementById('close-'+id).value);
  if(!price || price <= 0){alert('Enter close price');return}
  const r = await j('/api/close_trade/'+id, {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({close_price: price})
  });
  alert(`Closed.\nP&L USD: $${r.pnl_usd.toFixed(2)}\nP&L ZAR: ${fmtZAR(r.pnl_zar)}\nLedger balance: ${fmtZAR(r.ledger_balance_zar)}`);
  refreshAll();
}

async function requestWithdrawal(method){
  const amount = parseFloat(document.getElementById('wd-amount').value);
  const msg = document.getElementById('wd-msg');
  if(!amount || amount <= 0){msg.textContent='❌ Enter valid amount';msg.className='msg negative';return}
  const r = await j('/api/withdraw', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({amount_zar: amount, method})
  });
  if(r.status === 'rejected' || r.status === 'error'){
    msg.textContent = '❌ ' + (r.reason || 'failed'); msg.className = 'msg negative';
  } else {
    msg.textContent = `✅ #${r.withdrawal_id} · ${fmtZAR(r.amount_zar)} via ${r.method.toUpperCase()} — ${r.note} · Available after: ${fmtZAR(r.available_after_zar)}`;
    msg.className = 'msg positive';
  }
  refreshAll();
}

async function completeWd(id){await j('/api/withdraw/'+id+'/complete',{method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({})});refreshAll()}
async function cancelWd(id){await j('/api/withdraw/'+id+'/cancel',{method:'POST'});refreshAll()}
async function resetAll(){
  if(!confirm('Reset all data? Ledger will NOT be reseeded unless you check next prompt.'))return;
  const reseed = confirm('Also reseed the R120m ledger? (recommended)');
  await j('/api/reset', {method:'POST', headers:{'Content-Type':'application/json'},
                         body: JSON.stringify({reseed_ledger: reseed})});
  refreshAll();
}

refreshAll();
loadPayPalMe();
setInterval(refreshAll, 10000);
</script>
</body>
</html>
"""

@app.route('/')
def index():
    return Response(DASHBOARD_HTML, mimetype='text/html')

# ═══════════════════════════════════════════════════════════════
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
