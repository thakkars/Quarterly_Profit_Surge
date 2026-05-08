import os
import json
import time
import traceback
import requests
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from openpyxl.utils import get_column_letter

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

if os.environ.get("GITHUB_ACTIONS", "false").lower() != "true":
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.service import Service

# ═══════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════
IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"

if IS_GITHUB_ACTIONS:
    FOLDER_PATH = "/tmp/screener_data"
else:
    FOLDER_PATH = r"D:\Investment Research\Python Projects\Operating Profit more than 70% jump"

os.makedirs(FOLDER_PATH, exist_ok=True)

CALENDAR_EXCEL_PATH = os.path.join(FOLDER_PATH, "NSE_Results_Calendar.xlsx")
CALENDAR_JSON_PATH  = os.path.join(FOLDER_PATH, "nse_results_calendar.json")

R2_BUCKET     = os.environ.get("R2_BUCKET_NAME",       "quarterly-profit-surge")
R2_ACCESS_KEY = os.environ.get("R2_ACCESS_KEY_ID",     "")
R2_SECRET_KEY = os.environ.get("R2_SECRET_ACCESS_KEY", "")
R2_ENDPOINT   = os.environ.get("R2_ENDPOINT_URL",      "")

R2_CALENDAR_EXCEL_KEY = "NSE_Results_Calendar.xlsx"
R2_CALENDAR_JSON_KEY  = "nse_results_calendar.json"

NSE_BASE_URL  = "https://www.nseindia.com"
NSE_XBRL_PAGE = "https://www.nseindia.com/companies-listing/corporate-filings-announcements-xbrl"

COLS_ORDER = [
    "Sr No", "Symbol", "Company Name", "ISIN",
    "Broadcast Date", "Broadcast Time", "Subject",
    "Description", "Series", "Exchange", "XBRL Link", "Scraped At",
]

# API endpoints to try (in priority order)
API_ENDPOINTS = [
    "https://www.nseindia.com/api/corporate-announcements?index=equities&type=Board+Meeting&category=Result",
    "https://www.nseindia.com/api/corporate-announcements?index=equities&category=Result",
    "https://www.nseindia.com/api/home-corporate-announcements?index=equities&category=Result",
    "https://www.nseindia.com/api/corporates-pit?index=equities&type=result",
]

# ═══════════════════════════════════════════════════════════
# R2 HELPERS
# ═══════════════════════════════════════════════════════════
def get_r2_client():
    import boto3
    return boto3.client(
        "s3",
        endpoint_url=R2_ENDPOINT,
        aws_access_key_id=R2_ACCESS_KEY,
        aws_secret_access_key=R2_SECRET_KEY,
        region_name="auto",
    )

def r2_download(r2_key, local_path):
    try:
        get_r2_client().download_file(R2_BUCKET, r2_key, local_path)
        print(f"⬇️  R2 download: {r2_key}")
        return True
    except Exception as e:
        print(f"⚠️  R2 download failed ({r2_key}): {e}")
        return False

def r2_upload(local_path, r2_key):
    try:
        if not os.path.exists(local_path):
            print(f"⚠️  R2 upload skipped — file missing: {local_path}")
            return False
        extra = {}
        if r2_key.endswith(".json"):
            extra = {"ContentType": "application/json", "CacheControl": "no-cache, no-store, must-revalidate"}
        get_r2_client().upload_file(local_path, R2_BUCKET, r2_key, ExtraArgs=extra or None)
        print(f"⬆️  R2 upload: {r2_key}")
        return True
    except Exception as e:
        print(f"⚠️  R2 upload failed ({r2_key}): {e}")
        return False

# ═══════════════════════════════════════════════════════════
# SELENIUM DRIVER
# ═══════════════════════════════════════════════════════════
def init_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1600,900")
    opts.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
    if IS_GITHUB_ACTIONS:
        return webdriver.Chrome(options=opts)
    else:
        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=opts
        )

# ═══════════════════════════════════════════════════════════
# HARVEST NSE COOKIES VIA SELENIUM
# NSE injects `nsit` + `nseappid` cookies via JS — we need
# a real browser to land on the page and collect them.
# We then hand these cookies to a requests.Session so the
# API calls are authenticated properly.
# ═══════════════════════════════════════════════════════════
def harvest_nse_cookies():
    """
    1. Launch headless Chrome
    2. Visit NSE home + XBRL page (lets JS set nsit/nseappid cookies)
    3. Extract all cookies from the browser
    4. Return a requests.Session pre-loaded with those cookies
    """
    driver = None
    try:
        print("  Launching headless Chrome to harvest NSE cookies...")
        driver = init_driver()

        # Step 1 — hit home page first (sets initial cookies)
        driver.get(NSE_BASE_URL)
        time.sleep(3)
        print(f"  Visited NSE home → title: {driver.title[:60]}")

        # Step 2 — hit the XBRL announcements page (sets nsit cookie)
        driver.get(NSE_XBRL_PAGE)
        time.sleep(4)
        print(f"  Visited XBRL page → title: {driver.title[:60]}")

        # Step 3 — collect all cookies
        selenium_cookies = driver.get_cookies()
        print(f"  Harvested {len(selenium_cookies)} cookies from NSE")

        # Step 4 — build requests.Session with these cookies + correct headers
        session = requests.Session()
        session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept":           "application/json, text/plain, */*",
            "Accept-Language":  "en-US,en;q=0.9",
            "Accept-Encoding":  "gzip, deflate, br",
            "Referer":          NSE_XBRL_PAGE,
            "X-Requested-With": "XMLHttpRequest",
            "Connection":       "keep-alive",
        })
        for c in selenium_cookies:
            session.cookies.set(c["name"], c["value"], domain=c.get("domain", ".nseindia.com"))

        # Log key cookie names (not values) for debugging
        cookie_names = [c["name"] for c in selenium_cookies]
        print(f"  Cookie names: {cookie_names}")

        return session

    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass


# ═══════════════════════════════════════════════════════════
# FETCH ANNOUNCEMENTS
# ═══════════════════════════════════════════════════════════
def fetch_nse_announcements(session):
    all_records = []
    for endpoint in API_ENDPOINTS:
        try:
            print(f"  Trying: {endpoint}")
            resp = session.get(endpoint, timeout=20)
            print(f"  Status: {resp.status_code} | Content-Type: {resp.headers.get('Content-Type','?')[:50]}")

            if resp.status_code != 200:
                print(f"  ⚠️  Non-200 response, skipping.")
                time.sleep(1)
                continue

            # Guard: make sure response is actually JSON
            ct = resp.headers.get("Content-Type", "")
            if "json" not in ct and not resp.text.strip().startswith(("[", "{")):
                print(f"  ⚠️  Response is not JSON (got HTML/text). NSE may still be blocking.")
                print(f"  Response preview: {resp.text[:200]}")
                time.sleep(1)
                continue

            data = resp.json()

            if isinstance(data, list):
                all_records = data
            elif isinstance(data, dict):
                for key in ("data", "announcements", "results", "body", "table"):
                    if key in data and isinstance(data[key], list):
                        all_records = data[key]
                        break

            if all_records:
                print(f"  ✅ {len(all_records)} records fetched from {endpoint}")
                break
            else:
                print(f"  ⚠️  Empty result from {endpoint}")

        except json.JSONDecodeError as e:
            print(f"  ❌ JSON decode error on {endpoint}: {e}")
            print(f"  Response preview: {resp.text[:200] if 'resp' in dir() else 'n/a'}")
        except Exception as e:
            print(f"  ❌ Request error on {endpoint}: {e}")
        time.sleep(1.5)

    return all_records


# ═══════════════════════════════════════════════════════════
# PARSE RECORD
# ═══════════════════════════════════════════════════════════
def parse_record(raw, sr_no):
    def g(*keys):
        raw_lower = {k.lower(): v for k, v in raw.items()}
        for k in keys:
            val = raw_lower.get(k.lower(), "")
            if val and str(val).strip() not in ("", "null", "None", "-"):
                return str(val).strip()
        return ""

    bcast_raw  = g("broadcastdatetime", "broadcast_date_time", "an_dt", "date", "an_date")
    bcast_date, bcast_time = "", ""
    if bcast_raw:
        for fmt in ("%d-%b-%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
            try:
                dt_obj     = datetime.strptime(bcast_raw[:19], fmt)
                bcast_date = dt_obj.strftime("%d-%b-%Y")
                bcast_time = dt_obj.strftime("%H:%M:%S")
                break
            except:
                pass
        if not bcast_date:
            bcast_date = bcast_raw[:10]
            bcast_time = bcast_raw[11:19] if len(bcast_raw) > 10 else ""

    xbrl_link = g("attchmntfile", "attachment_file", "xbrl", "attachmentfile", "attchmnt_file")
    if xbrl_link and not xbrl_link.startswith("http"):
        xbrl_link = f"{NSE_BASE_URL}{xbrl_link}" if xbrl_link.startswith("/") else xbrl_link

    return {
        "Sr No":          sr_no,
        "Symbol":         g("symbol", "scrip_cd", "nsecode"),
        "Company Name":   g("comp", "company", "company_name", "name", "longname"),
        "ISIN":           g("isin", "isin_code"),
        "Broadcast Date": bcast_date,
        "Broadcast Time": bcast_time,
        "Subject":        g("subject", "an_subject", "desc", "reason"),
        "Description":    g("details", "an_det", "description", "body"),
        "Series":         g("series"),
        "Exchange":       g("exchange", "exch") or "NSE",
        "XBRL Link":      xbrl_link,
        "Scraped At":     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


# ═══════════════════════════════════════════════════════════
# EXCEL WRITER
# ═══════════════════════════════════════════════════════════
def save_to_excel(df, path):
    sheet_name = "NSE Results Calendar"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    wb = load_workbook(path)
    ws = wb[sheet_name]

    header_fill = PatternFill("solid", fgColor="1A365D")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    light_fill  = PatternFill("solid", fgColor="EBF8FF")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.border    = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=False)
            if row_idx % 2 == 0:
                cell.fill = light_fill

    col_widths = {
        "Sr No": 7, "Symbol": 12, "Company Name": 35, "ISIN": 16,
        "Broadcast Date": 16, "Broadcast Time": 14, "Subject": 40,
        "Description": 50, "Series": 8, "Exchange": 10,
        "XBRL Link": 20, "Scraped At": 20,
    }
    for i, col_name in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = col_widths.get(col_name, 15)

    ws.freeze_panes    = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)
    print(f"✅ Saved Excel: {path}  ({len(df)} rows)")


# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════
def main():
    print("\n" + "="*60)
    print("  NSE Results Calendar Scraper")
    print(f"  Run time (IST): {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)

    if IS_GITHUB_ACTIONS:
        r2_download(R2_CALENDAR_EXCEL_KEY, CALENDAR_EXCEL_PATH)

    # Step 1 — Use Selenium to get real NSE cookies
    print("\n[1/4] Harvesting NSE cookies via Selenium...")
    session = harvest_nse_cookies()

    # Step 2 — Fetch announcements with cookie-authenticated session
    print("\n[2/4] Fetching NSE announcements...")
    raw_records = fetch_nse_announcements(session)

    if not raw_records:
        print("❌ No records fetched. The NSE API may have changed or is temporarily unavailable.")
        return

    # Step 3 — Parse
    print(f"\n[3/4] Parsing {len(raw_records)} records...")
    parsed_rows = [parse_record(r, i + 1) for i, r in enumerate(raw_records)]
    df_new = pd.DataFrame(parsed_rows, columns=COLS_ORDER)

    # Merge with history
    if os.path.exists(CALENDAR_EXCEL_PATH):
        try:
            df_existing = pd.read_excel(CALENDAR_EXCEL_PATH, sheet_name="NSE Results Calendar")
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_combined.drop_duplicates(
                subset=["Symbol", "Broadcast Date", "Broadcast Time", "Subject"],
                keep="last", inplace=True,
            )
            try:
                df_combined["_sort_dt"] = pd.to_datetime(
                    df_combined["Broadcast Date"] + " " + df_combined["Broadcast Time"],
                    format="%d-%b-%Y %H:%M:%S", errors="coerce",
                )
                df_combined.sort_values("_sort_dt", ascending=False, inplace=True)
                df_combined.drop(columns=["_sort_dt"], inplace=True)
            except:
                pass
            df_combined["Sr No"] = range(1, len(df_combined) + 1)
            df_final = df_combined[COLS_ORDER]
            print(f"  Merged: {len(df_existing)} old + {len(df_new)} new = {len(df_final)} total")
        except Exception as e:
            print(f"  ⚠️  Merge failed ({e}), using fresh data.")
            df_final = df_new
    else:
        df_final = df_new

    # Step 4 — Save & upload
    print("\n[4/4] Saving files...")
    save_to_excel(df_final, CALENDAR_EXCEL_PATH)
    df_final.fillna("").to_json(CALENDAR_JSON_PATH, orient="records")
    print(f"✅ Saved JSON: {CALENDAR_JSON_PATH}")

    if IS_GITHUB_ACTIONS:
        r2_upload(CALENDAR_EXCEL_PATH, R2_CALENDAR_EXCEL_KEY)
        r2_upload(CALENDAR_JSON_PATH,  R2_CALENDAR_JSON_KEY)

    print("\n✅ NSE Results Calendar scraper complete!")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
        traceback.print_exc()
