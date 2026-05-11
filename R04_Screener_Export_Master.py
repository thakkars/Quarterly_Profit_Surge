import os
import re
import time
import shutil
import traceback
from datetime import datetime
import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ═══════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════
IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"

USERNAME   = os.environ.get("SCREENER_USERNAME", "")
PASSWORD   = os.environ.get("SCREENER_PASSWORD", "")

SCREEN_ID  = "3650931"
SCREEN_URL = f"https://www.screener.in/screens/{SCREEN_ID}/all-nse-bse-stocks/"
EXPORT_URL = f"https://www.screener.in/screens/{SCREEN_ID}/export/"

TMP_DIR      = "/tmp/screener_export" if IS_GITHUB_ACTIONS else "."
os.makedirs(TMP_DIR, exist_ok=True)
DOWNLOAD_DIR = os.path.join(TMP_DIR, "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

TODAY           = datetime.utcnow().strftime("%Y-%m-%d")
RAW_CSV_KEY     = f"raw/query-results-{TODAY}.csv"
MASTER_CSV_KEY  = "master/exchange_codes_fullnames.csv"
MASTER_XLSX_KEY = "master/exchange_codes_fullnames.xlsx"

MASTER_CSV_PATH  = os.path.join(TMP_DIR, "exchange_codes_fullnames.csv")
MASTER_XLSX_PATH = os.path.join(TMP_DIR, "exchange_codes_fullnames.xlsx")

R2_BUCKET     = os.environ.get("R2_BUCKET_NAME",       "") or "quarterly-profit-surge"
R2_ACCESS_KEY = os.environ.get("R2_ACCESS_KEY_ID",     "")
R2_SECRET_KEY = os.environ.get("R2_SECRET_ACCESS_KEY", "")
R2_ENDPOINT   = os.environ.get("R2_ENDPOINT_URL",      "")

HEADLESS  = IS_GITHUB_ACTIONS
PAGE_DELAY = 8.0
NAME_DELAY = 1.2

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
        print(f"⬇️  Downloaded R2:{r2_key} → {local_path}")
        return True
    except Exception as e:
        print(f"ℹ️  R2 download skipped ({r2_key}): {e}")
        return False

def r2_upload(local_path, r2_key, content_type="text/csv"):
    try:
        print(f"ℹ️  Uploading → bucket='{R2_BUCKET}'  key='{r2_key}'")
        get_r2_client().upload_file(
            local_path, R2_BUCKET, r2_key,
            ExtraArgs={"ContentType": content_type,
                       "CacheControl": "no-cache, no-store, must-revalidate"},
        )
        print(f"⬆️  Uploaded {local_path} → R2:{r2_key}")
        return True
    except Exception as e:
        print(f"⚠️  R2 upload failed ({r2_key}): {e}")
        return False

# ═══════════════════════════════════════════════════════════
# SELENIUM DRIVER
# ═══════════════════════════════════════════════════════════
def init_driver():
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1600,900")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    opts.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=opts)

# ═══════════════════════════════════════════════════════════
# LOGIN  —  returns session cookies for requests-based fallback
# ═══════════════════════════════════════════════════════════
def login(driver, wait):
    driver.get("https://www.screener.in/login/")
    wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(USERNAME)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    driver.find_element(By.NAME, "password").submit()
    time.sleep(PAGE_DELAY)
    if "login" in driver.current_url.lower():
        raise RuntimeError("❌ Login failed — check SCREENER_USERNAME / SCREENER_PASSWORD secrets.")
    print("✅ Logged in to Screener.in (Selenium)")

def get_requests_session(driver):
    """
    Build a requests.Session using cookies from the logged-in Selenium browser.
    This lets us download the CSV directly without involving Chrome's download manager.
    """
    session = requests.Session()
    for cookie in driver.get_cookies():
        session.cookies.set(cookie["name"], cookie["value"],
                            domain=cookie.get("domain", ".screener.in"))
    session.headers.update({
        "User-Agent": driver.execute_script("return navigator.userAgent;"),
        "Referer": SCREEN_URL,
    })
    return session

# ═══════════════════════════════════════════════════════════
# METHOD 1  —  requests-based direct download (most reliable)
# ═══════════════════════════════════════════════════════════
def export_via_requests(session):
    """
    Use the logged-in session to hit the export endpoint directly.
    Screener returns the CSV as a file attachment when authenticated.
    """
    print(f"  📥 Method 1 (requests): GET {EXPORT_URL}")
    try:
        resp = session.get(EXPORT_URL, timeout=60, allow_redirects=True)
        print(f"     Status: {resp.status_code}  Content-Type: {resp.headers.get('Content-Type', '')}")
        ct = resp.headers.get("Content-Type", "")
        # Accept if response is CSV or octet-stream and has real content
        if resp.status_code == 200 and ("csv" in ct or "octet" in ct or "text/plain" in ct):
            out = os.path.join(DOWNLOAD_DIR, f"screener_export_{TODAY}.csv")
            with open(out, "wb") as f:
                f.write(resp.content)
            print(f"  ✅ Method 1 success: {out} ({len(resp.content)} bytes)")
            return out
        # Sometimes Content-Type is text/html but body IS a CSV (starts with column headers)
        if resp.status_code == 200 and resp.text.strip().startswith(("Name", "name", "S.No", "Sr")):
            out = os.path.join(DOWNLOAD_DIR, f"screener_export_{TODAY}.csv")
            with open(out, "w", encoding="utf-8") as f:
                f.write(resp.text)
            print(f"  ✅ Method 1 success (text fallback): {out}")
            return out
        print(f"  ⚠️  Method 1 response is not a CSV (content-type={ct}, len={len(resp.content)})")
        print(f"     First 300 chars: {resp.text[:300]}")
    except Exception as e:
        print(f"  ⚠️  Method 1 exception: {e}")
    return None

# ═══════════════════════════════════════════════════════════
# METHOD 2  —  Selenium button click (waits for JS render)
# ═══════════════════════════════════════════════════════════
def export_via_button(driver, wait, session):
    print(f"  🖱️  Method 2 (button): navigating to {SCREEN_URL}")
    driver.get(SCREEN_URL)

    # Wait for the results table to appear (proves JS has finished rendering)
    print("     Waiting for results table...")
    try:
        wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "table.data-table, #table-wrapper table, .results-table")
        ))
        print("     Results table found ✅")
    except Exception:
        print("     ⚠️  Table not found; continuing anyway...")

    time.sleep(3)  # extra buffer for export button to render

    # Print all <a> tags on the page to diagnose what's available
    try:
        links = driver.find_elements(By.TAG_NAME, "a")
        export_links = [(a.get_attribute("href") or "", a.text.strip()) for a in links
                        if "export" in (a.get_attribute("href") or "").lower()
                        or "export" in a.text.lower()]
        print(f"     Export-related links found: {export_links}")
    except Exception:
        pass

    # Ordered selectors — Screener's actual DOM
    SELECTORS = [
        (By.CSS_SELECTOR,  "p.export a"),
        (By.CSS_SELECTOR,  ".export a"),
        (By.CSS_SELECTOR,  "a.button-green"),
        (By.XPATH,         "//a[normalize-space(text())='Export']"),
        (By.XPATH,         "//a[normalize-space(text())='Export CSV']"),
        (By.XPATH,         "//button[normalize-space(text())='Export']"),
        (By.XPATH,         "//*[contains(@class,'export')]//a"),
        (By.XPATH,         "//*[contains(text(),'Export') and (self::a or self::button)]"),
    ]

    for by, selector in SELECTORS:
        try:
            btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((by, selector)))
            href = btn.get_attribute("href") or ""
            print(f"     Found button via '{selector}', href='{href}'")

            # If there's an href with /export/ in it, use requests to download directly
            if href and "export" in href.lower():
                print(f"     📥 Downloading export href directly via requests: {href}")
                try:
                    resp = session.get(href, timeout=60, allow_redirects=True)
                    ct = resp.headers.get("Content-Type", "")
                    if resp.status_code == 200 and len(resp.content) > 100:
                        out = os.path.join(DOWNLOAD_DIR, f"screener_export_{TODAY}.csv")
                        with open(out, "wb") as f:
                            f.write(resp.content)
                        print(f"  ✅ Method 2 href download success: {out}")
                        return out
                except Exception as he:
                    print(f"     href download failed: {he}")

            # Otherwise click and wait for file download
            driver.execute_script("arguments[0].scrollIntoView(true);", btn)
            time.sleep(0.4)
            driver.execute_script("arguments[0].click();", btn)
            print(f"     Clicked. Waiting for download...")
            for _ in range(30):
                files = [f for f in os.listdir(DOWNLOAD_DIR)
                         if f.endswith(".csv") and not f.endswith(".crdownload")]
                if files:
                    path = os.path.join(DOWNLOAD_DIR, sorted(files)[-1])
                    print(f"  ✅ Method 2 click download success: {path}")
                    return path
                time.sleep(1)
            print("     ⚠️  Download timed out after 30s.")
        except Exception:
            continue

    return None

# ═══════════════════════════════════════════════════════════
# EXPORT ORCHESTRATOR
# ═══════════════════════════════════════════════════════════
def export_csv(driver, wait, session):
    for f in os.listdir(DOWNLOAD_DIR):
        if f.endswith(".csv"):
            os.remove(os.path.join(DOWNLOAD_DIR, f))

    path = export_via_requests(session)
    if path:
        return path

    path = export_via_button(driver, wait, session)
    if path:
        return path

    # Full page source dump for diagnosis
    try:
        print("  🔍 Full page source (first 5000 chars):")
        print(driver.page_source[:5000])
    except Exception:
        pass

    raise RuntimeError(
        "❌ All export methods failed.\n"
        "   Check the page source above for clues."
    )

# ═══════════════════════════════════════════════════════════
# FETCH FULL COMPANY NAME
# ═══════════════════════════════════════════════════════════
def fetch_full_name(driver, wait, nse_code, bse_code):
    slug = (nse_code or bse_code or "").strip()
    if not slug:
        return ""
    url = f"https://www.screener.in/company/{slug}/"
    try:
        driver.get(url)
        h1 = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "h1.h2, h1, .company-name h1, #overview h1")
        ))
        name = h1.text.strip()
        name = re.sub(r"\s*(BSE|NSE|\|).*$", "", name).strip()
        time.sleep(NAME_DELAY)
        return name
    except Exception as e:
        print(f"    ⚠️  Could not fetch name for {slug}: {e}")
        return ""

# ═══════════════════════════════════════════════════════════
# LOAD / STANDARDISE EXPORTED CSV
# ═══════════════════════════════════════════════════════════
EXPECTED_COLS = {
    "name":      ["name", "company name", "company"],
    "bse_code":  ["bse code", "bse"],
    "nse_code":  ["nse code", "nse"],
    "isin_code": ["isin code", "isin"],
}

def find_col(df_cols_lower, aliases):
    for alias in aliases:
        for col in df_cols_lower:
            if alias in col:
                return col
    return None

def load_export(csv_path):
    df = pd.read_csv(csv_path, dtype=str).fillna("")
    df.columns = df.columns.str.strip()
    lc = {c.lower(): c for c in df.columns}
    col_name = find_col(lc, EXPECTED_COLS["name"])
    col_bse  = find_col(lc, EXPECTED_COLS["bse_code"])
    col_nse  = find_col(lc, EXPECTED_COLS["nse_code"])
    col_isin = find_col(lc, EXPECTED_COLS["isin_code"])
    print(f"  Columns → name:'{col_name}' bse:'{col_bse}' nse:'{col_nse}' isin:'{col_isin}'")
    out = pd.DataFrame()
    out["Screener_Name"] = df[col_name].str.strip() if col_name else ""
    out["BSE_Code"]      = df[col_bse].str.strip()  if col_bse  else ""
    out["NSE_Code"]      = df[col_nse].str.strip()  if col_nse  else ""
    out["ISIN_Code"]     = df[col_isin].str.strip()  if col_isin else ""
    def clean(v):
        v = str(v).strip()
        return "" if v in ("-", "—", "N/A", "NA", "nan") else v
    for col in ["BSE_Code", "NSE_Code", "ISIN_Code"]:
        out[col] = out[col].apply(clean)
    return out

# ═══════════════════════════════════════════════════════════
# DIFF
# ═══════════════════════════════════════════════════════════
def make_key(row):
    if row.get("ISIN_Code", ""):
        return row["ISIN_Code"]
    parts = [p for p in [row.get("NSE_Code", ""), row.get("BSE_Code", "")] if p]
    return "|".join(parts) if parts else str(row.get("Screener_Name", "")).lower()

def diff_new_rows(df_today, df_master):
    master_keys = set(df_master["Key"].tolist()) if "Key" in df_master.columns else set()
    df_today["Key"] = df_today.apply(make_key, axis=1)
    new_rows = df_today[~df_today["Key"].isin(master_keys)].copy()
    print(f"  Today: {len(df_today)}  Master: {len(df_master)}  New: {len(new_rows)}")
    return new_rows

# ═══════════════════════════════════════════════════════════
# SAVE MASTER
# ═══════════════════════════════════════════════════════════
def save_master(df):
    display_cols = ["Sr_No", "Full_Company_Name", "Screener_Name",
                    "NSE_Code", "BSE_Code", "ISIN_Code",
                    "First_Seen", "Last_Seen", "Source_URL"]
    for col in display_cols:
        if col not in df.columns:
            df[col] = ""
    df = df.sort_values("Full_Company_Name", ignore_index=True)
    df["Sr_No"] = range(1, len(df) + 1)
    df = df[display_cols]
    df.to_csv(MASTER_CSV_PATH, index=False)
    print(f"  💾 CSV: {MASTER_CSV_PATH} ({len(df)} rows)")
    with pd.ExcelWriter(MASTER_XLSX_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Exchange Codes")
    print(f"  💾 XLSX: {MASTER_XLSX_PATH} ({len(df)} rows)")
    return df

# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════
print("\n🚀 R04_Screener_Export_Master.py starting...")
print(f"   Screen URL : {SCREEN_URL}")
print(f"   Export URL : {EXPORT_URL}")
print(f"   R2 Bucket  : '{R2_BUCKET}'")
print(f"   Date (UTC) : {TODAY}")
print(f"   Download   : {DOWNLOAD_DIR}")

try:
    if IS_GITHUB_ACTIONS:
        r2_download(MASTER_CSV_KEY, MASTER_CSV_PATH)

    df_master = pd.DataFrame(columns=["Key", "Full_Company_Name", "Screener_Name",
                                       "NSE_Code", "BSE_Code", "ISIN_Code",
                                       "First_Seen", "Last_Seen", "Source_URL"])
    if os.path.exists(MASTER_CSV_PATH):
        df_master = pd.read_csv(MASTER_CSV_PATH, dtype=str).fillna("")
        if "Key" not in df_master.columns:
            df_master["Key"] = df_master.apply(make_key, axis=1)
        print(f"ℹ️  Loaded existing master: {len(df_master)} rows")
    else:
        print("ℹ️  No existing master — starting fresh")

    driver = init_driver()
    wait   = WebDriverWait(driver, 30)
    login(driver, wait)

    # Build requests session from Selenium cookies
    session = get_requests_session(driver)

    print("\n📥 Exporting CSV...")
    downloaded_csv = export_csv(driver, wait, session)

    df_today = load_export(downloaded_csv)
    df_today["Key"]       = df_today.apply(make_key, axis=1)
    df_today["Last_Seen"] = TODAY

    new_rows = diff_new_rows(df_today, df_master)

    if not new_rows.empty:
        print(f"\n🔍 Fetching names for {len(new_rows)} new stocks...")
        full_names = []
        for i, (_, row) in enumerate(new_rows.iterrows(), 1):
            name = fetch_full_name(driver, wait, row["NSE_Code"], row["BSE_Code"])
            full_names.append(name or row["Screener_Name"])
            print(f"  [{i}/{len(new_rows)}] {row['Screener_Name']} → '{name}'")
        new_rows = new_rows.copy()
        new_rows["Full_Company_Name"] = full_names
        new_rows["First_Seen"]        = TODAY
        nse = new_rows["NSE_Code"].str.strip()
        bse = new_rows["BSE_Code"].str.strip()
        new_rows["Source_URL"] = [
            f"https://www.screener.in/company/{(n or b)}/"
            for n, b in zip(nse, bse)
        ]
    else:
        print("\n✅ No new stocks today.")

    driver.quit()

    if "Last_Seen" in df_master.columns:
        today_keys = set(df_today["Key"].tolist())
        df_master.loc[df_master["Key"].isin(today_keys), "Last_Seen"] = TODAY

    if not new_rows.empty:
        df_master = pd.concat([df_master, new_rows], ignore_index=True)

    df_final = save_master(df_master)
    print(df_final.head(10).to_string())

    if IS_GITHUB_ACTIONS:
        dated_raw = os.path.join(DOWNLOAD_DIR, f"query-results-{TODAY}.csv")
        shutil.copy(downloaded_csv, dated_raw)
        r2_upload(dated_raw,        RAW_CSV_KEY,     "text/csv")
        r2_upload(MASTER_CSV_PATH,  MASTER_CSV_KEY,  "text/csv")
        r2_upload(MASTER_XLSX_PATH, MASTER_XLSX_KEY,
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        print("ℹ️  Not in GitHub Actions — skipping R2 upload.")

    print(f"\n✅ Done. Total master rows: {len(df_final)}")

except Exception as e:
    print(f"\n❌ Fatal error: {e}")
    traceback.print_exc()
    try:
        driver.quit()
    except Exception:
        pass
    exit(1)
