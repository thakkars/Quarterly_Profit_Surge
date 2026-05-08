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

# R2 config
R2_BUCKET     = os.environ.get("R2_BUCKET_NAME",       "quarterly-profit-surge")
R2_ACCESS_KEY = os.environ.get("R2_ACCESS_KEY_ID",     "")
R2_SECRET_KEY = os.environ.get("R2_SECRET_ACCESS_KEY", "")
R2_ENDPOINT   = os.environ.get("R2_ENDPOINT_URL",      "")

R2_CALENDAR_EXCEL_KEY = "NSE_Results_Calendar.xlsx"
R2_CALENDAR_JSON_KEY  = "nse_results_calendar.json"

# NSE API endpoints
NSE_BASE_URL    = "https://www.nseindia.com"
NSE_XBRL_API   = "https://www.nseindia.com/api/corporate-announcements?index=equities&type=Board+Meeting&category=Result"
NSE_QUOTE_API  = "https://www.nseindia.com/api/quote-equity?symbol={symbol}"

# Columns to capture from NSE API
COLS_ORDER = [
    "Sr No",
    "Symbol",
    "Company Name",
    "ISIN",
    "Broadcast Date",
    "Broadcast Time",
    "Subject",
    "Description",
    "Series",
    "Exchange",
    "XBRL Link",
    "Scraped At",
]

# ═══════════════════════════════════════════════════════════
# R2 HELPERS  (same pattern as R01)
# ═══════════════════════════════════════════════════════════
def get_r2_client():
    import boto3
    return boto3.client(
        "s3",
        endpoint_url          = R2_ENDPOINT,
        aws_access_key_id     = R2_ACCESS_KEY,
        aws_secret_access_key = R2_SECRET_KEY,
        region_name           = "auto",
    )

def r2_download(r2_key, local_path):
    try:
        client = get_r2_client()
        client.download_file(R2_BUCKET, r2_key, local_path)
        print(f"⬇️  Downloaded from R2: {r2_key} → {local_path}")
        return True
    except Exception as e:
        print(f"⚠️  R2 download failed for {r2_key}: {e}")
        return False

def r2_upload(local_path, r2_key, content_type=None):
    try:
        if not os.path.exists(local_path):
            print(f"⚠️  Skipping R2 upload — file not found: {local_path}")
            return False
        client = get_r2_client()
        extra  = {}
        if content_type:
            extra["ContentType"] = content_type
        if r2_key.endswith(".json"):
            extra["ContentType"]  = "application/json"
            extra["CacheControl"] = "no-cache, no-store, must-revalidate"
        client.upload_file(local_path, R2_BUCKET, r2_key, ExtraArgs=extra if extra else None)
        print(f"⬆️  Uploaded to R2: {local_path} → {r2_key}")
        return True
    except Exception as e:
        print(f"⚠️  R2 upload failed for {r2_key}: {e}")
        return False

# ═══════════════════════════════════════════════════════════
# NSE SESSION HELPER
# NSE blocks plain requests — we need to first hit the home
# page to get cookies, then call the API.
# ═══════════════════════════════════════════════════════════
def build_nse_session():
    session = requests.Session()
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "*/*",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Referer": "https://www.nseindia.com/",
        "Connection": "keep-alive",
    }
    session.headers.update(headers)

    # Warm up — get cookies from main page and XBRL page
    for warm_url in [
        NSE_BASE_URL,
        f"{NSE_BASE_URL}/companies-listing/corporate-filings-announcements-xbrl",
    ]:
        try:
            resp = session.get(warm_url, timeout=15)
            print(f"  Warmed up NSE session via {warm_url} → status {resp.status_code}")
            time.sleep(2)
        except Exception as e:
            print(f"  ⚠️  Warm-up failed for {warm_url}: {e}")

    return session


def fetch_nse_announcements(session, pages=5):
    """
    Fetches NSE XBRL board meeting / result announcements.
    NSE paginates via `from` and `to` date params or just returns
    a flat list — we grab as many as available.
    Returns a list of raw dicts.
    """
    all_records = []

    # Try the primary XBRL result endpoint
    endpoints = [
        "https://www.nseindia.com/api/corporate-announcements?index=equities&type=Board+Meeting&category=Result",
        "https://www.nseindia.com/api/corporate-announcements?index=equities&category=Result",
        "https://www.nseindia.com/api/home-corporate-announcements?index=equities&category=Result",
    ]

    for endpoint in endpoints:
        try:
            print(f"  Trying endpoint: {endpoint}")
            resp = session.get(endpoint, timeout=20)
            if resp.status_code == 200:
                data = resp.json()
                # NSE can return list directly or wrapped in a dict
                if isinstance(data, list):
                    all_records = data
                elif isinstance(data, dict):
                    for key in ("data", "announcements", "results", "body"):
                        if key in data and isinstance(data[key], list):
                            all_records = data[key]
                            break
                    if not all_records and data:
                        # try to use dict values if it's a single record
                        all_records = [data]
                if all_records:
                    print(f"  ✅ Got {len(all_records)} records from {endpoint}")
                    break
            else:
                print(f"  ⚠️  Status {resp.status_code} from {endpoint}")
        except Exception as e:
            print(f"  ⚠️  Error fetching {endpoint}: {e}")
        time.sleep(1.5)

    return all_records


def parse_record(raw, sr_no):
    """
    Normalise a single NSE announcement record into our column schema.
    NSE field names vary slightly — we handle multiple known variants.
    """
    def g(*keys):
        """Get first matching key from raw dict, case-insensitive."""
        raw_lower = {k.lower(): v for k, v in raw.items()}
        for k in keys:
            val = raw_lower.get(k.lower(), "")
            if val and str(val).strip() not in ("", "null", "None", "-"):
                return str(val).strip()
        return ""

    # Broadcast date + time
    bcast_raw  = g("broadcastdatetime", "broadcast_date_time", "an_dt", "date", "an_date")
    bcast_date = ""
    bcast_time = ""
    if bcast_raw:
        try:
            # NSE format: "08-May-2026 18:30:00" or "2026-05-08T18:30:00"
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
        except:
            bcast_date = bcast_raw

    xbrl_link = g("attchmntfile", "attachment_file", "xbrl", "attachmentfile", "attchmnt_file")
    if xbrl_link and not xbrl_link.startswith("http"):
        xbrl_link = f"https://www.nseindia.com{xbrl_link}" if xbrl_link.startswith("/") else xbrl_link

    return {
        "Sr No":         sr_no,
        "Symbol":        g("symbol", "scrip_cd", "nsecode"),
        "Company Name":  g("comp", "company", "company_name", "name", "longname"),
        "ISIN":          g("isin", "isin_code"),
        "Broadcast Date": bcast_date,
        "Broadcast Time": bcast_time,
        "Subject":       g("subject", "an_subject", "desc", "reason"),
        "Description":   g("details", "an_det", "description", "body"),
        "Series":        g("series"),
        "Exchange":      g("exchange", "exch") or "NSE",
        "XBRL Link":     xbrl_link,
        "Scraped At":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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

    # Header style
    header_fill = PatternFill("solid", fgColor="1A365D")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Row zebra striping + borders
    light_fill = PatternFill("solid", fgColor="EBF8FF")
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

    # Column widths
    col_widths = {
        "Sr No": 7, "Symbol": 12, "Company Name": 35, "ISIN": 16,
        "Broadcast Date": 16, "Broadcast Time": 14, "Subject": 40,
        "Description": 50, "Series": 8, "Exchange": 10,
        "XBRL Link": 20, "Scraped At": 20,
    }
    for col_idx, col_name in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(col_name, 15)

    ws.freeze_panes = "A2"
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

    # Download existing file from R2 to preserve history
    if IS_GITHUB_ACTIONS:
        r2_download(R2_CALENDAR_EXCEL_KEY, CALENDAR_EXCEL_PATH)

    # Build NSE session with cookie warm-up
    print("\n[1/4] Building NSE session...")
    session = build_nse_session()

    # Fetch announcements
    print("\n[2/4] Fetching NSE announcements...")
    raw_records = fetch_nse_announcements(session)

    if not raw_records:
        print("❌ No records fetched. Exiting.")
        return

    # Parse into clean rows
    print(f"\n[3/4] Parsing {len(raw_records)} records...")
    parsed_rows = [parse_record(r, i + 1) for i, r in enumerate(raw_records)]
    df_new = pd.DataFrame(parsed_rows, columns=COLS_ORDER)

    # Merge with existing data to preserve history (deduplicate by Symbol+BroadcastDate+BroadcastTime)
    if os.path.exists(CALENDAR_EXCEL_PATH):
        try:
            df_existing = pd.read_excel(CALENDAR_EXCEL_PATH, sheet_name="NSE Results Calendar")
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_combined.drop_duplicates(
                subset=["Symbol", "Broadcast Date", "Broadcast Time", "Subject"],
                keep="last",
                inplace=True,
            )
            # Sort by Broadcast Date desc, then Time desc
            try:
                df_combined["_sort_dt"] = pd.to_datetime(
                    df_combined["Broadcast Date"] + " " + df_combined["Broadcast Time"],
                    format="%d-%b-%Y %H:%M:%S",
                    errors="coerce",
                )
                df_combined.sort_values("_sort_dt", ascending=False, inplace=True)
                df_combined.drop(columns=["_sort_dt"], inplace=True)
            except:
                pass
            df_combined["Sr No"] = range(1, len(df_combined) + 1)
            df_final = df_combined[COLS_ORDER]
            print(f"  Merged: {len(df_existing)} existing + {len(df_new)} new = {len(df_final)} total (after dedup)")
        except Exception as e:
            print(f"  ⚠️  Merge failed ({e}), using fresh data only.")
            df_final = df_new
    else:
        df_final = df_new

    # Save Excel
    print("\n[4/4] Saving files...")
    save_to_excel(df_final, CALENDAR_EXCEL_PATH)

    # Save JSON (for the HTML viewer)
    df_final.fillna("").to_json(CALENDAR_JSON_PATH, orient="records")
    print(f"✅ Saved JSON: {CALENDAR_JSON_PATH}")

    # Upload to R2
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
