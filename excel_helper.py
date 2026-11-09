import openpyxl
from openpyxl.styles import Font, PatternFill
import os
from datetime import datetime

EXCEL_FILE = "withdrawals.xlsx"

def init_excel():
    """Create the Excel file with headers if it doesn't exist."""
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PayPal Withdrawals"
        
        # Headers
        headers = ["Date", "Withdrawal ID", "Amount (ZAR)", "Amount (USD)", "Status", "PayPal Email"]
        ws.append(headers)
        
        # Style headers
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="003087", end_color="003087", fill_type="solid") # PayPal Blue
        
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            
        wb.save(EXCEL_FILE)
        print(f"✅ Created {EXCEL_FILE}")

def log_withdrawal_to_excel(wid, amount_zar, amount_usd, method, paypal_email):
    """Append a new withdrawal row to the Excel file."""
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    
    ws.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        wid,
        amount_zar,
        amount_usd,
        "PENDING",
        paypal_email
    ])
    
    wb.save(EXCEL_FILE)
    return True