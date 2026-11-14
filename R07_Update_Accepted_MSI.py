"""
R07_Update_Accepted_MSI.py
--------------------------------------------------------------------------------
Fetches current MarketSmith India (MSI) Price Strength and related metrics
for ONLY stocks marked as "accepted" in decisions.json for the active quarter.
Updates rawdata.json, Raw Data.xlsx, and Master File.xlsx.

Can be run:
  - Locally: python R07_Update_Accepted_MSI.py [--quarter "Sep 2026"]
  - On GitHub Actions via update_msi.yml workflow_dispatch
--------------------------------------------------------------------------------
"""

import os
import sys
import time
import json
import re
import argparse
from datetime import datetime

# Windows console UTF-8 fix
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# ─────────────────────────────────────────────────────────────
# ENVIRONMENT & PATHS
# ─────────────────────────────────────────────────────────────
IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"
LOCAL_REPO_DIR    = os.path.dirname(os.path.abspath(__file__))
RAW_JSON_PATH     = os.path.join(LOCAL_REPO_DIR, "rawdata.json")
RAW_EXCEL_PATH    = os.path.join(LOCAL_REPO_DIR, "Raw Data.xlsx")
MASTER_EXCEL_PATH = os.path.join(LOCAL_REPO_DIR, "Master File.xlsx")
DECISIONS_PATH    = os.path.join(LOCAL_REPO_DIR, "decisions.json")
PROGRESS_JSON_PATH= os.path.join(LOCAL_REPO_DIR, "progress.json")

MSI_BASE_URL = "https://marketsmithindia.com/mstool/eval"
MSI_COLS = [
    "MSI EPS Strength",
    "MSI Price Strength",
    "MSI Buyer Demand",
    "MSI Group Rank",
    "MSI Master Score",
]

if not IS_GITHUB_ACTIONS:
    try:
        from webdriver_manager.chrome import ChromeDriverManager
    except ImportError:
        ChromeDriverManager = None

# ─────────────────────────────────────────────────────────────
# PROGRESS TELEMETRY
# ─────────────────────────────────────────────────────────────
def push_progress(phase, completed, total, current_company="", status="running", elapsed_sec=0):
    pct = round((completed / total * 100)) if total > 0 else 0
    eta_sec = 0
    if completed > 0 and total > completed and elapsed_sec > 0:
        rate = elapsed_sec / completed
        eta_sec = int(rate * (total - completed))

    payload = {
        "phase":           phase,
        "status":          status,
        "completed":       completed,
        "total":           total,
        "pct":             pct,
        "current_company": current_company,
        "elapsed_sec":     int(elapsed_sec),
        "eta_sec":         eta_sec,
        "ts":              datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ"),
    }
    try:
        with open(PROGRESS_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
    except Exception as e:
        print(f"[WARN] Progress push failed: {e}")

# ─────────────────────────────────────────────────────────────
# TELEGRAM NOTIFICATION (DISABLED FOR MSI UPDATES)
# ─────────────────────────────────────────────────────────────
def send_telegram_notification(msg):
    # Telegram notifications disabled for MSI updates per user preference
    pass

# ─────────────────────────────────────────────────────────────
# CHROME DRIVER INIT
# ─────────────────────────────────────────────────────────────
def init_driver(max_retries=3):
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1600,900")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    for attempt in range(1, max_retries + 1):
        try:
            if IS_GITHUB_ACTIONS or not ChromeDriverManager:
                d = webdriver.Chrome(options=opts)
            else:
                d = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
            d.set_page_load_timeout(35)
            return d
        except Exception as e:
            print(f"[WARN] Chrome driver attempt {attempt}/{max_retries} failed: {e}")
            time.sleep(2)
            if attempt == max_retries:
                raise RuntimeError(f"Could not launch Chrome driver: {e}")

# ─────────────────────────────────────────────────────────────
# MSI SCRAPER FOR A SINGLE SYMBOL
# ─────────────────────────────────────────────────────────────
def msi_scrape_symbol(driver, symbol):
    result = {
        "MSI EPS Strength":   "",
        "MSI Price Strength": "",
        "MSI Buyer Demand":   "",
        "MSI Group Rank":     "",
        "MSI Master Score":   "",
    }
    if not symbol or not driver:
        return result

    url = f"{MSI_BASE_URL}/{symbol.lower()}/evaluation.jsp"
    try:
        driver.get(url)
        time.sleep(1.2)

        h4_list = driver.find_elements(By.TAG_NAME, "h4")
        for h4 in h4_list:
            try:
                htext = h4.text.strip()
                if not any(k in htext for k in ["EPS Strength", "Price Strength", "Buyer Demand", "Group Rank"]):
                    continue
                h3 = h4.find_element(By.XPATH, "following::h3[1]")
                spans = h3.find_elements(By.TAG_NAME, "span")
                val = ""
                if len(spans) >= 2:
                    val = spans[1].text.strip()
                elif spans:
                    val = spans[0].text.split(":")[-1].strip()
                else:
                    val = h3.text.split(":")[-1].strip()

                # Clean numeric values if possible
                num_val = val
                clean_num = re.sub(r"[^\d.]", "", val)
                if clean_num:
                    try:
                        num_val = float(clean_num) if "." in clean_num else int(clean_num)
                    except ValueError:
                        pass

                if "EPS Strength" in htext:
                    result["MSI EPS Strength"] = num_val
                elif "Price Strength" in htext:
                    result["MSI Price Strength"] = num_val
                elif "Buyer Demand" in htext:
                    result["MSI Buyer Demand"] = val
                elif "Group Rank" in htext:
                    result["MSI Group Rank"] = num_val
            except Exception:
                continue

        # Master Score
        try:
            ms = driver.find_elements(By.CSS_SELECTOR, "#masterScoreIndicator .masters b")
            if ms and ms[0].text.strip():
                result["MSI Master Score"] = ms[0].text.strip()
            else:
                h2 = driver.find_elements(By.CSS_SELECTOR, "h2.contentMasters b")
                if h2:
                    parts = h2[0].text.strip().split()
                    if len(parts) >= 3:
                        result["MSI Master Score"] = parts[2].rstrip(":")
        except Exception:
            pass

    except Exception as e:
        print(f"  [WARN] MSI scrape error for {symbol}: {e}")

    return result

# ─────────────────────────────────────────────────────────────
# IDENTIFIER UTILITIES
# ─────────────────────────────────────────────────────────────
def get_row_identifiers(row):
    nse_raw = str(row.get("NSE Code Exact", "") or "").strip()
    screener_raw = str(row.get("Screener Name Exact", "") or "").strip()
    nse = "" if (nse_raw.isdigit() or nse_raw == "0" or nse_raw == "") else nse_raw
    bse = screener_raw if screener_raw.isdigit() else ""
    comp_name = str(row.get("Company (Page)", "") or row.get("Name", "") or "").strip()
    ticker = (nse or bse or comp_name).strip()
    return {
        "ticker": ticker.upper(),
        "nse": nse.upper(),
        "bse": bse.upper(),
        "screener": screener_raw.upper(),
        "company": comp_name.upper(),
        "name": str(row.get("Name", "") or "").strip().upper(),
        "raw_sym": (nse or bse or screener_raw).strip()
    }

def find_excel_sheet(wb, quarter_str):
    if not quarter_str:
        return wb.active
    if quarter_str in wb.sheetnames:
        return wb[quarter_str]
    # Check short form "Sep 2026" -> "Sep 26"
    parts = quarter_str.split()
    if len(parts) == 2 and len(parts[1]) == 4:
        short_q = f"{parts[0]} {parts[1][2:]}"
        if short_q in wb.sheetnames:
            return wb[short_q]
    # Check lowercase / whitespace insensitive
    for s in wb.sheetnames:
        if s.strip().lower() == quarter_str.strip().lower():
            return wb[s]
    # Fallback to last sheet
    return wb[wb.sheetnames[-1]]

# ─────────────────────────────────────────────────────────────
# MAIN EXECUTION
# ─────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Update MSI Price Strength for accepted stocks")
    parser.add_argument("--quarter", "-q", default=os.environ.get("TARGET_QUARTER", ""), help="Target quarter e.g. 'Sep 2026'")
    args = parser.parse_args()

    start_time = time.time()
    print("=" * 60)
    print("📈 MARKETSMITH INDIA (MSI) PRICE STRENGTH UPDATER")
    print("=" * 60)

    # 1. Load decisions
    if not os.path.exists(DECISIONS_PATH):
        print(f"❌ decisions.json not found at {DECISIONS_PATH}. Aborting.")
        return 1

    try:
        with open(DECISIONS_PATH, "r", encoding="utf-8") as f:
            decisions = json.load(f)
    except Exception as e:
        print(f"❌ Failed to parse decisions.json: {e}")
        return 1

    # 2. Load rawdata.json
    if not os.path.exists(RAW_JSON_PATH):
        print(f"❌ rawdata.json not found at {RAW_JSON_PATH}. Aborting.")
        return 1

    try:
        with open(RAW_JSON_PATH, "r", encoding="utf-8") as f:
            raw_data = json.load(f)
    except Exception as e:
        print(f"❌ Failed to parse rawdata.json: {e}")
        return 1

    # 3. Determine active quarter
    target_quarter = (args.quarter or "").strip()
    if not target_quarter and raw_data:
        # Detect first quarter column from first row
        for col in raw_data[0].keys():
            try:
                pd.to_datetime(col, format="%b %Y")
                target_quarter = col
                break
            except Exception:
                continue

    if not target_quarter:
        # Fallback to latest key in decisions.json
        if decisions:
            target_quarter = list(decisions.keys())[-1]

    print(f"🎯 Target Quarter: '{target_quarter}'")

    # 4. Extract accepted set for this quarter
    q_decisions = decisions.get(target_quarter, {})
    accepted_keys = {
        str(k).strip().upper()
        for k, v in q_decisions.items()
        if str(v).lower() == "accepted"
    }

    # If target_quarter had no accepted entries, check if quarter was stored in short format
    if not accepted_keys:
        parts = target_quarter.split()
        if len(parts) == 2 and len(parts[1]) == 4:
            short_q = f"{parts[0]} {parts[1][2:]}"
            if short_q in decisions:
                accepted_keys = {
                    str(k).strip().upper()
                    for k, v in decisions[short_q].items()
                    if str(v).lower() == "accepted"
                }

    print(f"📋 Found {len(accepted_keys)} accepted key(s) in decisions for quarter '{target_quarter}':")
    for k in sorted(accepted_keys):
        print(f"   - {k}")

    if not accepted_keys:
        print(f"ℹ️ No accepted stocks found for quarter '{target_quarter}'. Nothing to update.")
        push_progress("complete", 0, 0, "No accepted stocks found", "complete", time.time() - start_time)
        return 0

    # 5. Identify matching rows in raw_data
    accepted_rows = []
    for idx, row in enumerate(raw_data):
        ids = get_row_identifiers(row)
        matched = (
            ids["ticker"] in accepted_keys or
            ids["nse"] in accepted_keys or
            ids["bse"] in accepted_keys or
            ids["screener"] in accepted_keys or
            ids["company"] in accepted_keys or
            ids["name"] in accepted_keys
        )
        if matched:
            accepted_rows.append((idx, row, ids))

    print(f"\n🔍 Matched {len(accepted_rows)} row(s) in rawdata.json marked as Accepted:")
    for idx, row, ids in accepted_rows:
        print(f"   • Row #{idx+1}: {ids['company']} (Symbol: {ids['raw_sym']})")

    if not accepted_rows:
        print("⚠️ Accepted keys found in decisions.json, but none matched any row in rawdata.json.")
        return 0

    # 6. Initialize headless Chrome driver
    print("\n🚀 Initializing Headless Chrome...")
    push_progress("initialising", 0, len(accepted_rows), "Launching browser engine...", "running", time.time() - start_time)
    driver = init_driver()

    updated_count = 0
    scraped_cache = {}
    scrape_results = []

    try:
        total = len(accepted_rows)
        for i, (idx, row, ids) in enumerate(accepted_rows):
            sym = ids["raw_sym"] or ids["nse"] or ids["bse"] or ids["ticker"]
            comp_name = row.get("Company (Page)") or row.get("Name") or sym
            elapsed = time.time() - start_time

            print(f"\n[{i+1}/{total}] Scraping MarketSmith India for: {comp_name} (Symbol: {sym})...")
            push_progress("scraping", i, total, f"Scraping MSI: {sym}", "running", elapsed)

            if not sym or sym == "0":
                print(f"   ⚠️ No valid exchange symbol for {comp_name} - skipping MSI scrape.")
                continue

            if sym.upper() in scraped_cache:
                msi_data = scraped_cache[sym.upper()]
                print(f"   ⚡ Using cached MSI data for {sym}")
            else:
                msi_data = msi_scrape_symbol(driver, sym)
                scraped_cache[sym.upper()] = msi_data

            # Update row in raw_data
            for col in MSI_COLS:
                val = msi_data.get(col, "")
                row[col] = val

            row["MSI Updated At"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            updated_count += 1
            scrape_results.append({
                "symbol": sym,
                "name": comp_name,
                "price_strength": msi_data.get("MSI Price Strength", "N/A"),
                "eps_strength": msi_data.get("MSI EPS Strength", "N/A"),
                "group_rank": msi_data.get("MSI Group Rank", "N/A"),
                "buyer_demand": msi_data.get("MSI Buyer Demand", "N/A"),
                "master_score": msi_data.get("MSI Master Score", "N/A"),
            })

            print(f"   ✅ Price Strength: {msi_data.get('MSI Price Strength')} | EPS: {msi_data.get('MSI EPS Strength')} | Master Score: {msi_data.get('MSI Master Score')}")
            time.sleep(1.0)

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    # 7. Save rawdata.json
    print(f"\n💾 Saving updated rawdata.json ({len(raw_data)} total records)...")
    try:
        with open(RAW_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(raw_data, f, indent=2)
        print("   ✅ rawdata.json saved successfully.")
    except Exception as e:
        print(f"   ❌ Failed to save rawdata.json: {e}")

    # 8. Update Raw Data.xlsx
    if os.path.exists(RAW_EXCEL_PATH):
        print(f"📊 Updating '{RAW_EXCEL_PATH}'...")
        try:
            wb = openpyxl.load_workbook(RAW_EXCEL_PATH)
            ws = find_excel_sheet(wb, target_quarter)
            print(f"   Target sheet: '{ws.title}'")

            # Map header names to column index
            header_col_map = {}
            for col_idx in range(1, ws.max_column + 1):
                val = ws.cell(row=1, column=col_idx).value
                if val:
                    header_col_map[str(val).strip()] = col_idx

            # Ensure all MSI cols exist in header; append if missing
            for mcol in MSI_COLS:
                if mcol not in header_col_map:
                    new_col_idx = ws.max_column + 1
                    ws.cell(row=1, column=new_col_idx).value = mcol
                    header_col_map[mcol] = new_col_idx

            # Update rows in Excel
            for row_idx in range(2, ws.max_row + 1):
                cell_name = str(ws.cell(row=row_idx, column=header_col_map.get("Name", 3)).value or "").strip().upper()
                cell_comp = str(ws.cell(row=row_idx, column=header_col_map.get("Company (Page)", 4)).value or "").strip().upper()
                cell_nse  = str(ws.cell(row=row_idx, column=header_col_map.get("NSE Code Exact", 7)).value or "").strip().upper()
                cell_scr  = str(ws.cell(row=row_idx, column=header_col_map.get("Screener Name Exact", 8)).value or "").strip().upper()

                # Check if this Excel row matches any scraped accepted row
                for idx, r, ids in accepted_rows:
                    if (
                        (cell_nse and cell_nse == ids["nse"]) or
                        (cell_scr and cell_scr == ids["screener"]) or
                        (cell_comp and cell_comp == ids["company"]) or
                        (cell_name and cell_name == ids["name"])
                    ):
                        for mcol in MSI_COLS:
                            col_idx = header_col_map.get(mcol)
                            if col_idx:
                                ws.cell(row=row_idx, column=col_idx).value = r.get(mcol, "")
                        break

            wb.save(RAW_EXCEL_PATH)
            wb.close()
            print("   ✅ Raw Data.xlsx updated and saved successfully.")
        except Exception as e:
            print(f"   [WARN] Failed to update Raw Data.xlsx: {e}")

    # 9. Update Master File.xlsx if sheet or matching column exists
    if os.path.exists(MASTER_EXCEL_PATH):
        try:
            wb_m = openpyxl.load_workbook(MASTER_EXCEL_PATH)
            ws_m = find_excel_sheet(wb_m, target_quarter)
            header_map_m = {str(ws_m.cell(row=1, column=c).value).strip(): c for c in range(1, ws_m.max_column + 1) if ws_m.cell(row=1, column=c).value}
            
            # If MSI Price Strength is a column in master
            if "MSI Price Strength" in header_map_m:
                for row_idx in range(2, ws_m.max_row + 1):
                    c_name = str(ws_m.cell(row=row_idx, column=header_map_m.get("Company (Page)", 3)).value or "").strip().upper()
                    for idx, r, ids in accepted_rows:
                        if c_name and c_name == ids["company"]:
                            for mcol in MSI_COLS:
                                if mcol in header_map_m:
                                    ws_m.cell(row=row_idx, column=header_map_m[mcol]).value = r.get(mcol, "")
                            break
                wb_m.save(MASTER_EXCEL_PATH)
                wb_m.close()
                print("   ✅ Master File.xlsx updated.")
        except Exception as e:
            print(f"   [NOTE] Master File.xlsx skipped: {e}")

    total_elapsed = round(time.time() - start_time)
    print("\n" + "=" * 60)
    print(f"🎉 COMPLETED: Successfully refreshed MSI metrics for {updated_count} accepted stock(s) in {total_elapsed}s.")
    print("=" * 60)

    # 10. Update progress.json
    push_progress("complete", updated_count, len(accepted_rows), "MSI Refresh Complete", "complete", total_elapsed)

    return 0

if __name__ == "__main__":
    sys.exit(main())
