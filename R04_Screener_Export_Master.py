import os
import re
import time
import traceback
from datetime import datetime
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
# Direct export URL — most reliable; works after login cookie is set
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

# R2 CONFIG
R2_BUCKET     = os.environ.get("R2_BUCKET_NAME",       "") or "quarterly-profit-surge"
R2_ACCESS_KEY = os.environ.get("R2_ACCESS_KEY_ID",     "")
R2_SECRET_KEY = os.environ.get("R2_SECRET_ACCESS_KEY", "")
R2_ENDPOINT   = os.environ.get("R2_ENDPOINT_URL",      "")

HEADLESS   = IS_GITHUB_ACTIONS
PAGE_DELAY = 5.0
NAME_DELAY = 1.2

# ═══════════════════════════════════════════════════════════
# R2 HELPERS
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
            ExtraArgs={
                "ContentType": content_type,
                "CacheControl": "no-cache, no-store, must-revalidate",
            },
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
# LOGIN
# ═══════════════════════════════════════════════════════════
def login(driver, wait):
    driver.get("https://www.screener.in/login/")
    wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(USERNAME)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    driver.find_element(By.NAME, "password").submit()
    time.sleep(PAGE_DELAY)
    if "login" in driver.current_url.lower():
        raise RuntimeError("❌ Login failed — check SCREENER_USERNAME / SCREENER_PASSWORD secrets.")
    print("✅ Logged in to Screener.in")

# ═══════════════════════════════════════════════════════════
# EXPORT CSV — METHOD 1: DIRECT URL (most reliable)
# ═══════════════════════════════════════════════════════════
def export_via_direct_url(driver):
    """
    Navigate directly to the export endpoint.
    Screener serves the CSV as a file download when the user is logged in.
    """
    print(f"  📥 Method 1: Direct export URL → {EXPORT_URL}")
    driver.get(EXPORT_URL)
    time.sleep(6)   # give browser time to trigger the download

    files = [f for f in os.listdir(DOWNLOAD_DIR)
             if f.endswith(".csv") and not f.endswith(".crdownload")]
    if files:
        path = os.path.join(DOWNLOAD_DIR, sorted(files)[-1])
        print(f"  ✅ Downloaded via direct URL: {path}")
        return path
    return None

# ═══════════════════════════════════════════════════════════
# EXPORT CSV — METHOD 2: CLICK EXPORT BUTTON ON PAGE
# ═══════════════════════════════════════════════════════════
def export_via_button(driver, wait):
    """
    Navigate to the screen page, wait for results table,
    then click the Export / Export CSV button.
    """
    print(f"  🖱️  Method 2: Click Export button on screen page → {SCREEN_URL}")
    driver.get(SCREEN_URL)
    time.sleep(PAGE_DELAY)

    # Scroll down then back up to trigger lazy-render
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1.5)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1.5)

    # Screener.in known selectors (ordered by specificity)
    SELECTORS = [
        # Screener's <p class="export"><a ...>Export</a></p> pattern
        (By.CSS_SELECTOR,  "p.export a"),
        (By.CSS_SELECTOR,  ".export a"),
        (By.CSS_SELECTOR,  "a.button-green"),
        # Generic text-match fallbacks
        (By.XPATH,         "//a[normalize-space(text())='Export']"),
        (By.XPATH,         "//a[normalize-space(text())='Export CSV']"),
        (By.XPATH,         "//button[normalize-space(text())='Export']"),
        (By.XPATH,         "//*[contains(@class,'export')]//a"),
        (By.XPATH,         "//*[contains(text(),'Export') and (self::a or self::button)]"),
    ]

    for by, selector in SELECTORS:
        try:
            btn = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((by, selector)))
            driver.execute_script("arguments[0].scrollIntoView(true);", btn)
            time.sleep(0.4)
            driver.execute_script("arguments[0].click();", btn)
            print(f"  ✅ Clicked Export via: {selector}")

            # Wait up to 30s for download
            for _ in range(30):
                files = [f for f in os.listdir(DOWNLOAD_DIR)
                         if f.endswith(".csv") and not f.endswith(".crdownload")]
                if files:
                    path = os.path.join(DOWNLOAD_DIR, sorted(files)[-1])
                    print(f"  ✅ Downloaded via button click: {path}")
                    return path
                time.sleep(1)
            print("  ⚠️  Button clicked but download timed out — trying next selector.")
        except Exception:
            continue

    return None

# ═══════════════════════════════════════════════════════════
# EXPORT CSV — ORCHESTRATOR (tries both methods)
# ═══════════════════════════════════════════════════════════
def export_csv(driver, wait):
    # Clear any stale CSV files from a previous run
    for f in os.listdir(DOWNLOAD_DIR):
        if f.endswith(".csv"):
            os.remove(os.path.join(DOWNLOAD_DIR, f))

    # Try Method 1: direct URL first (fastest, no DOM dependency)
    path = export_via_direct_url(driver)
    if path:
        return path

    # Try Method 2: button click
    path = export_via_button(driver, wait)
    if path:
        return path

    # Debug: dump page source excerpt to help diagnose
    try:
        src = driver.page_source[:3000]
        print("  🔍 Page source excerpt for debugging:")
        print(src)
    except Exception:
        pass

    raise RuntimeError(
        "❌ Could not export CSV via direct URL or button click.\n"
        "   Check: (1) Login succeeded, (2) Screen ID is correct, "
        "(3) Your Screener plan allows exports."
    )

# ═══════════════════════════════════════════════════════════
# FETCH FULL COMPANY NAME FROM SCREENER PAGE
# ═══════════════════════════════════════════════════════════
def screener_slug(nse_code, bse_code):
    code = (nse_code or bse_code or "").strip()
    return code if code else None

def fetch_full_name(driver, wait, nse_code, bse_code):
    slug = screener_slug(nse_code, bse_code)
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

    print(f"  Export columns detected → name:'{col_name}' bse:'{col_bse}' nse:'{col_nse}' isin:'{col_isin}'")

    out = pd.DataFrame()
    out["Screener_Name"] = df[col_name].str.strip()  if col_name else ""
    out["BSE_Code"]      = df[col_bse].str.strip()   if col_bse  else ""
    out["NSE_Code"]      = df[col_nse].str.strip()   if col_nse  else ""
    out["ISIN_Code"]     = df[col_isin].str.strip()  if col_isin else ""

    def clean(v):
        v = str(v).strip()
        return "" if v in ("-", "—", "N/A", "NA", "nan") else v

    for col in ["BSE_Code", "NSE_Code", "ISIN_Code"]:
        out[col] = out[col].apply(clean)

    return out

# ═══════════════════════════════════════════════════════════
# DIFF: FIND NEW ROWS NOT IN MASTER
# ═══════════════════════════════════════════════════════════
def make_key(row):
    """Primary key = ISIN; fallback = NSE|BSE."""
    if row.get("ISIN_Code", ""):
        return row["ISIN_Code"]
    parts = [p for p in [row.get("NSE_Code", ""), row.get("BSE_Code", "")] if p]
    return "|".join(parts) if parts else str(row.get("Screener_Name", "")).lower()

def diff_new_rows(df_today, df_master):
    master_keys = set(df_master["Key"].tolist()) if "Key" in df_master.columns else set()
    df_today["Key"] = df_today.apply(make_key, axis=1)
    new_mask = ~df_today["Key"].isin(master_keys)
    new_rows = df_today[new_mask].copy()
    print(f"  Today total: {len(df_today)}  Master existing: {len(df_master)}  New: {len(new_rows)}")
    return new_rows

# ═══════════════════════════════════════════════════════════
# SAVE MASTER (CSV + XLSX)
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
    print(f"  💾 CSV saved: {MASTER_CSV_PATH} ({len(df)} rows)")

    with pd.ExcelWriter(MASTER_XLSX_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Exchange Codes")
    print(f"  💾 XLSX saved: {MASTER_XLSX_PATH} ({len(df)} rows)")
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
    # 1. Download existing master from R2
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

    # 2. Login + export CSV
    driver = init_driver()
    wait   = WebDriverWait(driver, 25)
    login(driver, wait)

    print(f"\n📥 Exporting CSV from screen...")
    downloaded_csv = export_csv(driver, wait)

    # 3. Load + standardise today's export
    df_today = load_export(downloaded_csv)
    df_today["Key"]       = df_today.apply(make_key, axis=1)
    df_today["Last_Seen"] = TODAY

    # 4. Find new rows
    new_rows = diff_new_rows(df_today, df_master)

    # 5. For each new row: fetch full company name from Screener
    if not new_rows.empty:
        print(f"\n🔍 Fetching full company names for {len(new_rows)} new stocks...")
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
        print("\n✅ No new stocks today — master is up to date.")

    driver.quit()

    # 6. Update Last_Seen for all existing rows still in today's export
    if "Last_Seen" in df_master.columns:
        today_keys = set(df_today["Key"].tolist())
        df_master.loc[df_master["Key"].isin(today_keys), "Last_Seen"] = TODAY

    # 7. Append new rows to master
    if not new_rows.empty:
        df_master = pd.concat([df_master, new_rows], ignore_index=True)

    # 8. Save master CSV + XLSX
    df_final = save_master(df_master)
    print(df_final.head(10).to_string())

    # 9. Upload to R2
    if IS_GITHUB_ACTIONS:
        import shutil
        dated_raw = os.path.join(DOWNLOAD_DIR, f"query-results-{TODAY}.csv")
        shutil.copy(downloaded_csv, dated_raw)

        r2_upload(dated_raw,        RAW_CSV_KEY,     "text/csv")
        r2_upload(MASTER_CSV_PATH,  MASTER_CSV_KEY,  "text/csv")
        r2_upload(MASTER_XLSX_PATH, MASTER_XLSX_KEY,
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        print("ℹ️  Not in GitHub Actions — skipping R2 upload.")

    print(f"\n✅ R04_Screener_Export_Master.py completed. Total master rows: {len(df_final)}")

except Exception as e:
    print(f"\n❌ Fatal error: {e}")
    traceback.print_exc()
    try:
        driver.quit()
    except Exception:
        pass
    exit(1)
