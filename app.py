# ═══════════════════════════════════════════════════════════════
# CRYPTO TRADING DASHBOARD - PHASE 1 + PAYPAL/EXCEL
# Deploy from GitHub to Render.
# Starting balance: R120,000,000 (simulated)
# ═══════════════════════════════════════════════════════════════

from flask import Flask, request, jsonify, Response
import sqlite3
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill

app = Flask(__name__)
DB = "dashboard.db"

# ═══════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════
START_BALANCE_ZAR = 120_000_000.00
BROKER_MODE = "SIMULATED"
BASE_CURRENCY = "ZAR"
CRYPTO_SYMBOL = "BTC"
ZAR_PER_USD = 18.50
ZAR_PER_BTC = 1_200_000

BROKER_CONFIG = {
    "OANDA_TOKEN": os.environ.get("OANDA_TOKEN", ""),
    "OANDA_ACCOUNT": os.environ.get("OANDA_ACCOUNT", ""),
    "LUNO_API_KEY": os.environ.get("LUNO_API_KEY", ""),
    "LUNO_API_SECRET": os.environ.get("LUNO_API_SECRET", ""),
    "PAYPAL_EMAIL": os.environ.get("PAYPAL_EMAIL", ""),        # for WITHDRAW (broker payout)
    "PAYPAL_ME_USERNAME": os.environ.get("PAYPAL_ME_USERNAME", ""),  # for RECEIVE (clients)
    "BANK_ACCOUNT": os.environ.get("BANK_ACCOUNT", ""),
    "BTC_WALLET": os.environ.get("BTC_WALLET", ""),
}

# ═══════════════════════════════════════════════════════════════
# EXCEL SETUP
# ═══════════════════════════════════════════════════════════════
EXCEL_FILE = "withdrawals.xlsx"

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()

        ws_master = wb.active
        ws_master.title = "All Withdrawals"
        master_headers = ["Date", "ID", "Method", "Amount (ZAR)",
                          "Amount (USD)", "Crypto Amt", "Status", "Destination"]
        ws_master.append(master_headers)

        ws_pp = wb.create_sheet("PayPal")
        pp_headers = ["Date", "ID", "Amount (ZAR)", "Amount (USD)",
                      "Status", "PayPal Email", "Note"]
        ws_pp.append(pp_headers)

        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="0A2540", end_color="0A2540", fill_type="solid")
        for ws in [ws_master, ws_pp]:
            for cell in ws[1]:
                cell.font = header_font
                cell.fill = header_fill

        wb.save(EXCEL_FILE)
        print(f"✅ Created {EXCEL_FILE}")
    else:
        print(f"📂 Existing {EXCEL_FILE} loaded")

init_excel()

def log_withdrawal_to_excel(wid, amount_zar, amount_usd, method,
                            crypto_amount, status, destination, note=""):
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)

        ws_master = wb["All Withdrawals"]
        ws_master.append([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            wid, method,
            round(amount_zar, 2), round(amount_usd, 2),
            round(crypto_amount, 8) if crypto_amount else 0,
            status, destination
        ])

        if method == "paypal":
            ws_pp = wb["PayPal"]
            ws_pp.append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                wid,
                round(amount_zar, 2),
                round(amount_usd, 2),
                status,
                destination,
                note
            ])

        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        print(f"⚠️ Excel write failed: {e}")
        return False

# ═══════════════════════════════════════════════════════════════
# DATABASE
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
    c.execute(
        "INSERT OR IGNORE INTO config (key,value) VALUES ('start_balance',?)",
        (str(START_BALANCE_ZAR),)
    )
    conn.commit()
    conn.close()

init_db()

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
# HELPERS
# ═══════════════════════════════════════════════════════════════
def usd_to_zar(usd):
    return usd * ZAR_PER_USD

def zar_to_btc(zar):
    return zar / ZAR_PER_BTC

def fmt_zar(amount):
    return f"R{amount:,.2f}"

# ═══════════════════════════════════════════════════════════════
# API ROUTES
# ═══════════════════════════════════════════════════════════════

@app.route('/health')
def health():
    return jsonify({
        "status": "ok",
        "mode": BROKER_MODE,
        "currency": BASE_CURRENCY,
        "crypto": CRYPTO_SYMBOL,
        "time": datetime.utcnow().isoformat()
    })

@app.route('/api/summary')
def summary():
    start_zar = get_start_balance()

    pnl_usd_total = query(
        "SELECT COALESCE(SUM(pnl_usd),0) AS s FROM trades WHERE status='closed'",
        one=True
    )["s"]

    pnl_zar_total = query(
        "SELECT COALESCE(SUM(pnl_zar),0) AS s FROM trades WHERE status='closed'",
        one=True
    )["s"]

    open_count = query(
        "SELECT COUNT(*) AS c FROM trades WHERE status='open'", one=True
    )["c"]

    closed_count = query(
        "SELECT COUNT(*) AS c FROM trades WHERE status='closed'", one=True
    )["c"]

    wins = query(
        "SELECT COUNT(*) AS c FROM trades WHERE status='closed' AND pnl_zar>0",
        one=True
    )["c"]

    losses = query(
        "SELECT COUNT(*) AS c FROM trades WHERE status='closed' AND pnl_zar<0",
        one=True
    )["c"]

    withdrawn_zar = query(
        "SELECT COALESCE(SUM(amount_zar),0) AS s FROM withdrawals WHERE status='completed'",
        one=True
    )["s"]

    pending_zar = query(
        "SELECT COALESCE(SUM(amount_zar),0) AS s FROM withdrawals WHERE status='pending'",
        one=True
    )["s"]

    win_rate = round((wins / closed_count * 100), 1) if closed_count else 0
    balance_zar = start_zar + pnl_zar_total
    available_zar = max(balance_zar - withdrawn_zar - pending_zar, 0)

    return jsonify({
        "start_balance_zar": round(start_zar, 2),
        "total_pnl_zar": round(pnl_zar_total, 2),
        "total_pnl_usd": round(pnl_usd_total, 2),
        "balance_zar": round(balance_zar, 2),
        "withdrawn_zar": round(withdrawn_zar, 2),
        "pending_zar": round(pending_zar, 2),
        "available_zar": round(available_zar, 2),
        "open_trades": open_count,
        "closed_trades": closed_count,
        "wins": wins,
        "losses": losses,
        "win_rate": win_rate,
        "broker_mode": BROKER_MODE,
        "crypto_equiv_btc": round(zar_to_btc(balance_zar), 8),
        "crypto_symbol": CRYPTO_SYMBOL,
        "zar_per_btc": ZAR_PER_BTC,
        "zar_per_usd": ZAR_PER_USD
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
        sl = round(price - 5, 2)
        tp = round(price + 10, 2)
    else:
        sl = round(price + 5, 2)
        tp = round(price - 10, 2)

    execute(
        """INSERT INTO trades
           (timestamp, signal, price, sl, tp, units, source)
           VALUES (?,?,?,?,?,?,?)""",
        (datetime.utcnow().isoformat(), data["signal"],
         price, sl, tp, units, "manual")
    )

    return jsonify({"status": "created"})

@app.route('/signal', methods=['POST'])
def signal():
    data = request.json or {}
    if data.get("signal") not in ("LONG", "SHORT"):
        return jsonify({"status": "rejected"}), 400

    price = float(data.get("price", 0))
    if price <= 0:
        return jsonify({"status": "rejected", "reason": "invalid price"}), 400

    execute(
        """INSERT INTO trades
           (timestamp, signal, price, sl, tp, units, source)
           VALUES (?,?,?,?,?,?,?)""",
        (datetime.utcnow().isoformat(), data["signal"],
         price, float(data.get("sl", 0)),
         float(data.get("tp", 0)),
         int(data.get("units", 10)), "indicator")
    )
    return jsonify({"status": "logged"})

@app.route('/api/close_trade/<int:trade_id>', methods=['POST'])
def close_trade(trade_id):
    data = request.json or {}
    close_price = float(data.get("close_price", 0))

    trade = query("SELECT * FROM trades WHERE id=?", (trade_id,), one=True)
    if not trade:
        return jsonify({"status": "not found"}), 404
    if trade["status"] == "closed":
        return jsonify({"status": "already closed"}), 400
    if close_price <= 0:
        return jsonify({"status": "invalid close_price"}), 400

    units = trade["units"]
    if trade["signal"] == "LONG":
        pnl_usd = (close_price - trade["price"]) * units
    else:
        pnl_usd = (trade["price"] - close_price) * units

    pnl_zar = usd_to_zar(pnl_usd)

    execute(
        """UPDATE trades
           SET status='closed', pnl_usd=?, pnl_zar=?, close_price=?
           WHERE id=?""",
        (round(pnl_usd, 2), round(pnl_zar, 2), close_price, trade_id)
    )

    return jsonify({
        "status": "closed",
        "pnl_usd": round(pnl_usd, 2),
        "pnl_zar": round(pnl_zar, 2)
    })

# ═══════════════════════════════════════════════════════════════
# WITHDRAWAL — BROKER PAYOUT (no paypal.me for self)
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

    s = summary().get_json()
    if amount_zar > s["available_zar"]:
        return jsonify({
            "status": "rejected",
            "reason": f"Insufficient funds. Available: {fmt_zar(s['available_zar'])}"
        }), 400

    crypto_amount = 0
    usd_amount = round(amount_zar / ZAR_PER_USD, 2)
    note = ""
    destination = ""
    pay_link = ""   # NOT USED for withdrawal anymore

    # ── PAYPAL — broker payout to YOUR email ──
    if method == "paypal":
        paypal_email = BROKER_CONFIG["PAYPAL_EMAIL"]
        if not paypal_email or "@" not in paypal_email:
            return jsonify({
                "status": "error",
                "reason": "PAYPAL_EMAIL must be your real email (e.g. you@email.com). Set it on Render."
            }), 400
        destination = paypal_email
        note = f"→ Broker will payout ${usd_amount:,.2f} to {paypal_email}"

    # ── LUNO ──
    elif method == "luno":
        btc = zar_to_btc(amount_zar)
        crypto_amount = btc
        destination = "Luno ZAR account"
        note = f"→ Luno: {btc:.8f} BTC"

    # ── BTC ──
    elif method == "btc":
        btc = zar_to_btc(amount_zar)
        crypto_amount = btc
        destination = BROKER_CONFIG["BTC_WALLET"] or "SET_BTC_WALLET_ENV"
        note = f"→ BTC: {btc:.8f} to {destination}"

    # ── BANK ──
    elif method == "bank":
        destination = BROKER_CONFIG["BANK_ACCOUNT"] or "SET_BANK_ACCOUNT_ENV"
        note = f"→ SA Bank EFT: {fmt_zar(amount_zar)} to {destination}"

    execute(
        """INSERT INTO withdrawals
           (timestamp, amount_zar, method, status, note, crypto_amount, pay_link)
           VALUES (?,?,?,?,?,?,?)""",
        (datetime.utcnow().isoformat(),
         amount_zar, method, "pending", note, crypto_amount, pay_link)
    )
    wid = query("SELECT last_insert_rowid() AS id", one=True)["id"]

    log_withdrawal_to_excel(
        wid=wid,
        amount_zar=amount_zar,
        amount_usd=usd_amount,
        method=method,
        crypto_amount=crypto_amount,
        status="pending",
        destination=destination,
        note=note
    )

    return jsonify({
        "status": "pending",
        "withdrawal_id": wid,
        "amount_zar": round(amount_zar, 2),
        "amount_usd": usd_amount,
        "crypto_amount": round(crypto_amount, 8),
        "method": method,
        "note": note,
        "destination": destination,
        "broker_mode": BROKER_MODE
    })

# ═══════════════════════════════════════════════════════════════
# PAYPAL.ME — RECEIVE PAYMENTS FROM CLIENTS (separate flow)
# ═══════════════════════════════════════════════════════════════
@app.route('/api/paypal_me_link')
def paypal_me_link():
    username = BROKER_CONFIG["PAYPAL_ME_USERNAME"].strip().lstrip("@").replace(" ", "")
    if not username:
        return jsonify({
            "status": "unconfigured",
            "message": "Set PAYPAL_ME_USERNAME on Render to enable"
        })
    if "@" in username:
        return jsonify({
            "status": "invalid",
            "message": "PAYPAL_ME_USERNAME must be a handle, not an email"
        })
    return jsonify({
        "status": "ok",
        "link": f"https://paypal.me/{username}",
        "username": username,
        "note": "Share with clients to receive payments"
    })

@app.route('/api/withdrawals')
def withdrawals():
    rows = query("SELECT * FROM withdrawals ORDER BY id DESC LIMIT 100")
    return jsonify([dict(r) for r in rows])

@app.route('/api/withdraw/<int:wid>/complete', methods=['POST'])
def complete_withdrawal(wid):
    execute("UPDATE withdrawals SET status='completed' WHERE id=?", (wid,))
    return jsonify({"status": "completed"})

@app.route('/api/withdraw/<int:wid>/cancel', methods=['POST'])
def cancel_withdrawal(wid):
    execute("UPDATE withdrawals SET status='cancelled' WHERE id=?", (wid,))
    return jsonify({"status": "cancelled"})

@app.route('/download/excel')
def download_excel():
    if not os.path.exists(EXCEL_FILE):
        return jsonify({"status": "error", "reason": "no excel file yet"}), 404
    with open(EXCEL_FILE, "rb") as f:
        data = f.read()
    return Response(
        data,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=withdrawals.xlsx"}
    )

@app.route('/api/reset', methods=['POST'])
def reset():
    execute("DELETE FROM trades")
    execute("DELETE FROM withdrawals")
    return jsonify({"status": "reset"})

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
.label{color:#8899aa;font-size:.72rem;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
.metric{font-size:1.35rem;font-weight:700;color:#00A8A8}
.metric.small{font-size:1rem;font-weight:600;color:#e6edf3}
.positive{color:#32CD32!important}
.negative{color:#FF4444!important}
.gold{color:#FFD700!important}
.btc{color:#00b5e2!important}
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
table{width:100%;border-collapse:collapse;min-width:700px}
th{text-align:left;padding:10px;color:#00A8A8;font-size:.72rem;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #1e4976}
td{padding:9px;border-bottom:1px solid #1e4976;font-size:.83rem}
tr:hover{background:#143557}
footer{margin-top:32px;padding-top:16px;border-top:2px solid #1e4976;display:flex;align-items:center;gap:16px;flex-wrap:wrap}
.footer-note{color:#8899aa;font-size:.8rem}
.withdraw-buttons{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}
.receive-box{display:flex;gap:10px;flex-wrap:wrap;margin-top:10px}
.receive-box input{flex:1;min-width:220px;font-family:monospace;font-size:.85rem}
</style>
</head>
<body>

<header>
  <h1>💰 Crypto Trading Dashboard <span style="font-size:.75rem;color:#8899aa">(ZAR)</span></h1>
  <div class="status" id="status">● Loading...</div>
</header>

<section class="grid">
  <div class="card highlight"><div class="label">Balance (ZAR)</div><div class="metric gold" id="balance">R0.00</div></div>
  <div class="card crypto"><div class="label">Crypto Equiv (BTC)</div><div class="metric btc" id="crypto-eq">0.00000000</div></div>
  <div class="card"><div class="label">Total P&L (ZAR)</div><div class="metric" id="pnl-zar">R0.00</div></div>
  <div class="card"><div class="label">Total P&L (USD)</div><div class="metric small" id="pnl-usd">$0.00</div></div>
  <div class="card"><div class="label">Win Rate</div><div class="metric" id="winrate">0%</div></div>
  <div class="card"><div class="label">Open</div><div class="metric" id="open">0</div></div>
  <div class="card"><div class="label">Withdrawn (ZAR)</div><div class="metric" id="withdrawn">R0.00</div></div>
  <div class="card highlight"><div class="label">Available (ZAR)</div><div class="metric gold" id="available">R0.00</div></div>
</section>

<section class="actions">
  <div class="panel">
    <h2>🧪 Simulate Trade</h2>
    <p class="hint">XAUUSD price · units · P&L auto-converted to ZAR</p>
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
    <p class="hint">Broker sends money to YOUR accounts. Logged to Excel.</p>
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

<footer>
  <button onclick="resetAll()" class="btn-danger">🔄 Reset All</button>
  <span class="footer-note" id="mode-note">Phase 1 · Simulated · No Broker</span>
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

async function refreshAll(){
  const s = await j('/api/summary');
  setText('balance', fmtZAR(s.balance_zar));
  setText('crypto-eq', fmtBTC(s.crypto_equiv_btc) + ' ' + s.crypto_symbol);
  setColor('pnl-zar', fmtZAR(s.total_pnl_zar), s.total_pnl_zar);
  setColor('pnl-usd', fmtUSD(s.total_pnl_usd), s.total_pnl_usd);
  setText('winrate', s.win_rate + '%');
  setText('open', s.open_trades);
  setText('withdrawn', fmtZAR(s.withdrawn_zar));
  setText('available', fmtZAR(s.available_zar));
  setText('mode-note', `Mode: ${s.broker_mode} · Start: ${fmtZAR(s.start_balance_zar)} · R${s.zar_per_btc.toLocaleString()}/BTC · R${s.zar_per_usd}/USD`);
  document.getElementById('status').textContent = '● Live';

  const trades = await j('/api/trades');
  const ob = document.querySelector('#open-trades tbody');
  const cb = document.querySelector('#closed-trades tbody');
  ob.innerHTML = ''; cb.innerHTML = '';

  for(const t of trades){
    if(t.status==='open'){
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
    if(w.status==='pending'){
      actions = `<button class="btn-small" onclick="completeWd(${w.id})">Complete</button>
                 <button class="btn-small btn-danger" onclick="cancelWd(${w.id})">Cancel</button>`;
    } else {
      actions = w.status==='completed' ? '✅' : '❌';
    }
    const cryptoAmt = w.crypto_amount > 0 ? fmtBTC(w.crypto_amount) + ' BTC' : '—';
    wb.innerHTML += `<tr>
      <td>${w.id}</td><td>${w.timestamp.slice(0,16).replace('T',' ')}</td>
      <td>${fmtZAR(w.amount_zar)}</td><td>${cryptoAmt}</td>
      <td>${w.method}</td><td>${w.status}</td><td>${actions}</td>
    </tr>`;
  }
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
  if(!price||price<=0){alert('Enter valid price');return}
  await j('/api/manual_trade', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({signal, price, units})
  });
  refreshAll();
}

async function closeTrade(id){
  const price = parseFloat(document.getElementById('close-'+id).value);
  if(!price||price<=0){alert('Enter close price');return}
  const r = await j('/api/close_trade/'+id, {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({close_price: price})
  });
  alert(`Closed.\nP&L USD: $${r.pnl_usd.toFixed(2)}\nP&L ZAR: ${fmtZAR(r.pnl_zar)}`);
  refreshAll();
}

async function requestWithdrawal(method){
  const amount = parseFloat(document.getElementById('wd-amount').value);
  const msg = document.getElementById('wd-msg');
  if(!amount||amount<=0){msg.textContent='❌ Enter valid amount';msg.className='msg negative';return}
  const r = await j('/api/withdraw', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({amount_zar: amount, method})
  });
  if(r.status==='rejected' || r.status==='error'){
    msg.textContent='❌ '+(r.reason||'failed'); msg.className='msg negative';
  } else {
    msg.textContent=`✅ #${r.withdrawal_id} · ${fmtZAR(r.amount_zar)} via ${r.method.toUpperCase()} — ${r.note}`;
    msg.className='msg positive';
  }
  refreshAll();
}

async function completeWd(id){await j('/api/withdraw/'+id+'/complete',{method:'POST'});refreshAll()}
async function cancelWd(id){await j('/api/withdraw/'+id+'/cancel',{method:'POST'});refreshAll()}
async function resetAll(){if(!confirm('Reset all data?'))return;await j('/api/reset',{method:'POST'});refreshAll()}

// Init
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
