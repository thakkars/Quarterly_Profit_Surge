"""
R06_Quarterly_Archive_And_Reset.py

Automated Quarterly Record Keeping & Reset:
1. Archives current quarter results into Records/Results_<Quarter>.xlsx (e.g., Results_Mar26.xlsx)
2. Segregates stocks into 3 styled tabs:
   - Accepted
   - Rejected
   - No action
3. Saves a full raw JSON backup in Records/rawdata_<Quarter>.json
4. Refreshes/resets rawdata.json to [] for the upcoming quarter
5. Sends Telegram notification with archive summary (if configured)
"""

import os
import sys

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import json
import re
import argparse
import urllib.request
import urllib.error
from datetime import datetime
import pandas as pd

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RECORDS_DIR = os.path.join(BASE_DIR, "Records")
RAW_JSON_PATH = os.path.join(BASE_DIR, "rawdata.json")
DECISIONS_JSON_PATH = os.path.join(BASE_DIR, "decisions.json")

MONTH_MAP = {
    1: ("Dec", -1),   # Jan EOD -> Q3 ending Dec 31 of previous year
    4: ("Mar", 0),    # Apr EOD -> Q4 ending Mar 31 of current year
    7: ("Jun", 0),    # Jul EOD -> Q1 ending Jun 30 of current year
    10: ("Sep", 0),   # Oct EOD -> Q2 ending Sep 30 of current year
}

MONTH_NAMES = {
    "jan": "Jan", "january": "Jan",
    "feb": "Feb", "february": "Feb",
    "mar": "Mar", "march": "Mar",
    "apr": "Apr", "april": "Apr",
    "may": "May",
    "jun": "Jun", "june": "Jun",
    "jul": "Jul", "july": "Jul",
    "aug": "Aug", "august": "Aug",
    "sep": "Sep", "september": "Sep",
    "oct": "Oct", "october": "Oct",
    "nov": "Nov", "november": "Nov",
    "dec": "Dec", "december": "Dec"
}

def normalize_qtr_key(val):
    """
    Normalizes quarter strings like 'Mar 2026', 'Mar26', 'March 2026' to ('Mar', '26')
    """
    if not val:
        return None
    s = str(val).strip().lower()
    match = re.search(r'([a-z]+)\s*[-_]?\s*(\d{2,4})', s)
    if match:
        m_str, y_str = match.groups()
        mon = MONTH_NAMES.get(m_str, m_str.capitalize()[:3])
        yy = y_str[-2:] if len(y_str) >= 2 else y_str
        return (mon, yy)
    return None

def detect_quarter_from_date():
    """
    Determines quarter based on execution date:
    - 3 Apr (or April run): Mar<YY>
    - 3 Jul (or July run): Jun<YY>
    - 3 Oct (or October run): Sep<YY>
    - 3 Jan (or January run): Dec<YY-1>
    """
    now = datetime.now()
    month = now.month
    year = now.year

    if month in MONTH_MAP:
        mon_prefix, yr_offset = MONTH_MAP[month]
        target_year = year + yr_offset
        return f"{mon_prefix}{str(target_year)[-2:]}"

    # Fallback if executed on non-standard month: inspect rawdata.json columns
    if os.path.exists(RAW_JSON_PATH):
        try:
            with open(RAW_JSON_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if data and isinstance(data, list) and len(data) > 0:
                first_row = data[0]
                q_cols = []
                for k in first_row.keys():
                    try:
                        dt = pd.to_datetime(k, format="%b %Y")
                        q_cols.append((dt, k))
                    except Exception:
                        pass
                if q_cols:
                    q_cols.sort(key=lambda x: x[0], reverse=True)
                    latest_q = q_cols[0][1]
                    norm = normalize_qtr_key(latest_q)
                    if norm:
                        return f"{norm[0]}{norm[1]}"
        except Exception as e:
            print(f"[WARN] Failed to auto-detect quarter from rawdata.json: {e}")

    # Final fallback: most recent ended quarter
    if month in [1, 2, 3]:
        return f"Dec{str(year - 1)[-2:]}"
    elif month in [4, 5, 6]:
        return f"Mar{str(year)[-2:]}"
    elif month in [7, 8, 9]:
        return f"Jun{str(year)[-2:]}"
    else:
        return f"Sep{str(year)[-2:]}"

def load_quarter_decisions(target_qtr_norm):
    """
    Loads decisions for target quarter from decisions.json
    """
    if not os.path.exists(DECISIONS_JSON_PATH):
        print(f"[WARN] {DECISIONS_JSON_PATH} not found.")
        return {}

    try:
        with open(DECISIONS_JSON_PATH, "r", encoding="utf-8") as f:
            all_decisions = json.load(f)
    except Exception as e:
        print(f"[ERROR] Failed to read decisions.json: {e}")
        return {}

    # Match normalized quarter key
    for qtr_key, decisions in all_decisions.items():
        norm = normalize_qtr_key(qtr_key)
        if norm and norm == target_qtr_norm:
            print(f"[OK] Found decisions matching quarter '{qtr_key}' ({len(decisions)} evaluations)")
            return decisions

    print(f"[WARN] No matching quarter found in decisions.json for {target_qtr_norm}. Available: {list(all_decisions.keys())}")
    return {}

def get_row_decision(row, qtr_decisions):
    """
    Determines decision ('accepted', 'rejected', or 'no action') for a raw data row
    """
    if not qtr_decisions:
        return "no action"

    candidates = []
    
    nse = str(row.get("NSE Code Exact", "") or "").strip()
    if nse and nse not in ["0", "None", "nan"]:
        candidates.append(nse)

    scr = str(row.get("Screener Name Exact", "") or "").strip()
    if scr and scr not in ["0", "None", "nan"]:
        candidates.append(scr)

    comp = str(row.get("Company (Page)", "") or "").strip()
    if comp and comp not in ["0", "None", "nan"]:
        candidates.append(comp)

    name = str(row.get("Name", "") or "").strip()
    if name and name not in ["0", "None", "nan"]:
        candidates.append(name)

    for cand in candidates:
        if cand in qtr_decisions:
            val = str(qtr_decisions[cand]).strip().lower()
            if val in ["accepted", "rejected"]:
                return val

    return "no action"

def style_excel_workbook(writer, sheet_configs):
    """
    Styles openpyxl workbook sheets with frozen headers, styled table headers, and tab colors
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    border_side = Side(border_style="thin", color="E0E0E0")
    cell_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    
    for cfg in sheet_configs:
        sheet_name = cfg["name"]
        tab_color = cfg["tab_color"]
        header_fill_color = cfg["header_color"]
        
        ws = writer.sheets.get(sheet_name)
        if ws is None:
            continue
            
        ws.sheet_properties.tabColor = tab_color
        ws.freeze_panes = "A2"
        ws.views.sheetView[0].showGridLines = True
        
        header_fill = PatternFill(start_color=header_fill_color, end_color=header_fill_color, fill_type="solid")
        
        # Style headers
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
        # Apply borders and auto column width
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                cell.border = cell_border
                val_str = str(cell.value or "")
                max_len = max(max_len, len(val_str))
            
            # Set proportional width
            ws.column_dimensions[col_letter].width = max(12, min(max_len + 3, 40))

def send_telegram_alert(qtr_code, accepted_count, rejected_count, no_action_count, total_count, excel_filename):
    """
    Sends summary notification to Telegram if secrets/env vars are configured
    """
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    raw_chat_ids = os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not raw_chat_ids:
        print("[INFO] Telegram notification skipped (no TELEGRAM_BOT_TOKEN or TELEGRAM_CHAT_ID).")
        return

    chat_id_str = raw_chat_ids.replace(";", ",").replace("\n", ",")
    chat_ids = [c.strip() for c in chat_id_str.split(",") if c.strip()]
    if not chat_ids:
        return

    now_str = datetime.now().strftime("%Y-%m-%d %I:%M %p")
    msg = (
        f"📦 <b>Quarterly Archive & Refresh Complete</b>\n"
        f"📅 Date: <code>{now_str}</code>\n"
        f"🏷️ Quarter: <b>{qtr_code}</b>\n\n"
        f"📊 <b>Quarter Results Summary:</b>\n"
        f"• Total Universe: <code>{total_count}</code> stocks\n"
        f"• ✅ <b>Accepted:</b> <code>{accepted_count}</code> stocks\n"
        f"• ❌ <b>Rejected:</b> <code>{rejected_count}</code> stocks\n"
        f"• ⏳ <b>No Action:</b> <code>{no_action_count}</code> stocks\n\n"
        f"📁 <b>Excel Saved:</b> <code>Records/{excel_filename}</code>\n"
        f"🔄 <b>Status:</b> <code>rawdata.json</code> refreshed for new quarter results."
    )

    url = f"https://api.telegram.org/bot{token}/sendMessage"
    for cid in chat_ids:
        try:
            payload = {
                "chat_id": cid,
                "text": msg,
                "parse_mode": "HTML",
                "disable_web_page_preview": True
            }
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode("utf-8"),
                headers={"Content-Type": "application/json"},
                method="POST"
            )
            with urllib.request.urlopen(req, timeout=15) as res:
                print(f"[OK] Telegram notification sent to chat_id: {cid}")
        except Exception as e:
            print(f"[WARN] Failed to send Telegram alert to chat_id {cid}: {e}")

def main():
    parser = argparse.ArgumentParser(description="Quarterly Archive & Reset for Stock Screener")
    parser.add_argument("--quarter", default=os.environ.get("QUARTER", "auto"),
                        help="Target quarter to archive (e.g. Mar26, Jun26, Sep26, Dec26, or 'auto')")
    parser.add_argument("--no-reset", action="store_true",
                        default=(os.environ.get("RESET_RAWDATA", "true").lower() == "false"),
                        help="Skip resetting rawdata.json")
    args = parser.parse_args()

    # 1. Determine quarter code
    if args.quarter.lower() == "auto":
        qtr_code = detect_quarter_from_date()
    else:
        qtr_code = args.quarter.strip()

    qtr_norm = normalize_qtr_key(qtr_code)
    if not qtr_norm:
        qtr_norm = ("Qtr", "26")
    formatted_qtr = f"{qtr_norm[0]}{qtr_norm[1]}"

    print(f"==================================================")
    print(f"📦 QUARTERLY ARCHIVE & REFRESH: {formatted_qtr}")
    print(f"==================================================")

    # 2. Load rawdata.json
    if not os.path.exists(RAW_JSON_PATH):
        print(f"[ERROR] {RAW_JSON_PATH} not found.")
        sys.exit(1)

    with open(RAW_JSON_PATH, "r", encoding="utf-8") as f:
        raw_data = json.load(f)

    if not isinstance(raw_data, list):
        print(f"[ERROR] rawdata.json does not contain a list of records.")
        sys.exit(1)

    total_records = len(raw_data)
    print(f"[OK] Loaded {total_records} stocks from rawdata.json")

    # 3. Load decisions
    qtr_decisions = load_quarter_decisions(qtr_norm)

    # 4. Classify into 3 groups
    accepted_rows = []
    rejected_rows = []
    no_action_rows = []

    for row in raw_data:
        decision = get_row_decision(row, qtr_decisions)
        if decision == "accepted":
            accepted_rows.append(row)
        elif decision == "rejected":
            rejected_rows.append(row)
        else:
            no_action_rows.append(row)

    print(f"[SUMMARY] Classification Results:")
    print(f"  • Accepted:  {len(accepted_rows)}")
    print(f"  • Rejected:  {len(rejected_rows)}")
    print(f"  • No action: {len(no_action_rows)}")

    # 5. Build DataFrames
    all_columns = list(raw_data[0].keys()) if total_records > 0 else []

    df_accepted = pd.DataFrame(accepted_rows) if accepted_rows else pd.DataFrame(columns=all_columns)
    df_rejected = pd.DataFrame(rejected_rows) if rejected_rows else pd.DataFrame(columns=all_columns)
    df_no_action = pd.DataFrame(no_action_rows) if no_action_rows else pd.DataFrame(columns=all_columns)

    # Ensure consistent column ordering across tabs
    for df in [df_accepted, df_rejected, df_no_action]:
        for c in all_columns:
            if c not in df.columns:
                df[c] = ""
        if all_columns:
            df = df[[c for c in all_columns if c in df.columns]]

    # 6. Ensure Records directory exists
    os.makedirs(RECORDS_DIR, exist_ok=True)

    # 7. Write Excel file
    excel_filename = f"Results_{formatted_qtr}.xlsx"
    excel_filepath = os.path.join(RECORDS_DIR, excel_filename)

    sheet_configs = [
        {"name": "Accepted", "tab_color": "28A745", "header_color": "1E7E34"},
        {"name": "Rejected", "tab_color": "DC3545", "header_color": "BD2130"},
        {"name": "No action", "tab_color": "6C757D", "header_color": "495057"},
    ]

    print(f"[WRITE] Writing 3 tabs to {excel_filepath}...")
    with pd.ExcelWriter(excel_filepath, engine="openpyxl") as writer:
        df_accepted.to_excel(writer, sheet_name="Accepted", index=False)
        df_rejected.to_excel(writer, sheet_name="Rejected", index=False)
        df_no_action.to_excel(writer, sheet_name="No action", index=False)
        style_excel_workbook(writer, sheet_configs)

    print(f"[OK] Successfully saved Excel archive: {excel_filepath}")

    # 8. Archive rawdata JSON backup
    raw_backup_filename = f"rawdata_{formatted_qtr}.json"
    raw_backup_path = os.path.join(RECORDS_DIR, raw_backup_filename)
    with open(raw_backup_path, "w", encoding="utf-8") as f:
        json.dump(raw_data, f, indent=2)
    print(f"[OK] Successfully saved raw JSON backup: {raw_backup_path}")

    # 9. Refresh / Reset rawdata.json
    if not args.no_reset:
        print("[RESET] Refreshing rawdata.json for incoming quarter...")
        with open(RAW_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump([], f)
        print("[OK] rawdata.json has been reset to []")
    else:
        print("[INFO] --no-reset specified; rawdata.json was preserved.")

    # 10. Send Telegram Alert
    send_telegram_alert(
        formatted_qtr,
        len(accepted_rows),
        len(rejected_rows),
        len(no_action_rows),
        total_records,
        excel_filename
    )

    print("[DONE] Quarterly archive completed successfully.")

if __name__ == "__main__":
    main()
