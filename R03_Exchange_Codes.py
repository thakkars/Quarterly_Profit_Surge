import time
import traceback
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

# ═══════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════
IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"

USERNAME   = os.environ.get("SCREENER_USERNAME", "")
PASSWORD   = os.environ.get("SCREENER_PASSWORD", "")
SCREEN_URL = "https://www.screener.in/screens/3650931/all-nse-bse-stocks/"

FOLDER_PATH    = "/tmp/exchange_codes" if IS_GITHUB_ACTIONS else "."
os.makedirs(FOLDER_PATH, exist_ok=True)

CSV_LOCAL_PATH = os.path.join(FOLDER_PATH, "exchange_codes.csv")

# R2 CONFIG
R2_BUCKET     = os.environ.get("R2_BUCKET_NAME",       "quarterly-profit-surge")
R2_ACCESS_KEY = os.environ.get("R2_ACCESS_KEY_ID",     "")
R2_SECRET_KEY = os.environ.get("R2_SECRET_ACCESS_KEY", "")
R2_ENDPOINT   = os.environ.get("R2_ENDPOINT_URL",      "")
R2_CSV_KEY    = "exchange_codes.csv"

HEADLESS   = IS_GITHUB_ACTIONS
PAGE_DELAY = 3.0   # slightly longer for JS render

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

def r2_download_existing():
    try:
        client = get_r2_client()
        client.download_file(R2_BUCKET, R2_CSV_KEY, CSV_LOCAL_PATH)
        print(f"⬇️  Downloaded existing exchange_codes.csv from R2")
        return True
    except Exception as e:
        print(f"ℹ️  No existing CSV on R2 (will create fresh): {e}")
        return False

def r2_upload(local_path, r2_key):
    try:
        client = get_r2_client()
        client.upload_file(
            local_path, R2_BUCKET, r2_key,
            ExtraArgs={
                'ContentType': 'text/csv',
                'CacheControl': 'no-cache, no-store, must-revalidate'
            }
        )
        print(f"⬆️  Uploaded {local_path} → R2:{r2_key}")
        return True
    except Exception as e:
        print(f"⚠️  R2 upload failed: {e}")
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
    return webdriver.Chrome(options=opts)

# ═══════════════════════════════════════════════════════════
# SCREENER LOGIN
# ═══════════════════════════════════════════════════════════
def login(driver, wait):
    driver.get("https://www.screener.in/login/")
    try:
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.NAME, "password").submit()
        time.sleep(4)
        print("✅ Logged in to Screener.in")
    except Exception as e:
        print(f"⚠️  Login issue: {e}")

# ═══════════════════════════════════════════════════════════
# FIND TABLE — Screener puts <th> inside <tbody> row 1
# ═══════════════════════════════════════════════════════════
def get_table_and_headers(driver, wait):
    """
    Screener.in screen tables do NOT use <thead>.
    The first <tr> inside <tbody> contains <th> elements as headers.
    Returns (table_element, headers_list, col_map_dict) or (None,[],{})
    """
    # Wait for ANY table row with th to appear
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr")))
    except Exception as e:
        print(f"  ⚠️  Timed out waiting for table rows: {e}")
        return None, [], {}

    time.sleep(1)  # extra settle time

    # Try data-table first, fallback to first table
    for selector in ["table.data-table", "table"]:
        try:
            table = driver.find_element(By.CSS_SELECTOR, selector)

            # --- Strategy 1: th inside tbody (Screener default) ---
            header_els = table.find_elements(By.CSS_SELECTOR, "tbody tr:first-child th")
            if not header_els:
                # --- Strategy 2: th inside thead ---
                header_els = table.find_elements(By.CSS_SELECTOR, "thead th")
            if not header_els:
                # --- Strategy 3: any th anywhere in table ---
                header_els = table.find_elements(By.TAG_NAME, "th")

            headers = [el.text.strip() for el in header_els]
            headers = [h for h in headers if h]   # drop blanks

            if headers:
                col_map = {h.lower(): i for i, h in enumerate(headers)}
                print(f"  Headers detected ({len(headers)}): {headers}")
                return table, headers, col_map
        except Exception:
            continue

    # Last resort: dump page source snippet for debugging
    print("  ⚠️  No headers found. Dumping page source snippet for debug:")
    try:
        src = driver.page_source
        start = src.find("<table")
        print(src[start:start+2000])
    except:
        pass
    return None, [], {}

# ═══════════════════════════════════════════════════════════
# SCRAPE SCREEN TABLE (all pages)
# ═══════════════════════════════════════════════════════════
def scrape_all_pages(driver, wait):
    driver.get(SCREEN_URL)
    print(f"  Waiting for page JS to render...")
    time.sleep(PAGE_DELAY)

    # Scroll down to trigger lazy-load
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1.5)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)

    all_rows = []
    page_num = 1

    while True:
        print(f"\n  📄 Scraping page {page_num}...")

        table, headers, col_map = get_table_and_headers(driver, wait)

        if not headers or table is None:
            print("  ❌ Could not read headers — aborting pagination.")
            break

        # Locate column indices
        name_idx = next((col_map[k] for k in col_map if "name" in k), None)
        bse_idx  = next((col_map[k] for k in col_map if "bse" in k), None)
        nse_idx  = next((col_map[k] for k in col_map if "nse" in k), None)

        print(f"  Column indices → name:{name_idx}  bse:{bse_idx}  nse:{nse_idx}")

        if name_idx is None:
            print("  ⚠️  'Name' column not found in headers.")
            break

        # Data rows: skip the first tr if it contains th (header row inside tbody)
        all_tr = table.find_elements(By.CSS_SELECTOR, "tbody tr")
        data_rows = []
        for tr in all_tr:
            if tr.find_elements(By.TAG_NAME, "th"):
                continue   # skip header row
            data_rows.append(tr)

        if not data_rows:
            print("  ℹ️  No data rows found on this page.")
            break

        page_count = 0
        for tr in data_rows:
            cells = tr.find_elements(By.TAG_NAME, "td")
            if not cells or len(cells) <= name_idx:
                continue

            # Full company name from anchor if present
            name_cell = cells[name_idx]
            anchors   = name_cell.find_elements(By.TAG_NAME, "a")
            company   = anchors[0].text.strip() if anchors else name_cell.text.strip()

            bse_code = cells[bse_idx].text.strip() if bse_idx is not None and len(cells) > bse_idx else ""
            nse_code = cells[nse_idx].text.strip() if nse_idx is not None and len(cells) > nse_idx else ""

            # Normalise dash/empty codes
            def clean(v):
                v = v.strip()
                if v in ("-", "—", "–", "N/A", "NA", ""):
                    return ""
                if v.endswith(".0"):
                    v = v[:-2]
                return v

            bse_code = clean(bse_code)
            nse_code = clean(nse_code)

            if not company:
                continue

            all_rows.append({
                "Company":  company,
                "NSE_Code": nse_code,
                "BSE_Code": bse_code,
            })
            page_count += 1

        print(f"  ✅ Page {page_num}: {page_count} rows  (running total: {len(all_rows)})")

        # ── Next page ──
        try:
            # Screener uses li.next > a  OR  a.next
            next_btn = None
            for xpath in [
                "//li[contains(@class,'next')]/a",
                "//a[contains(@class,'next')]",
                "//a[normalize-space()='Next']",
                "//a[normalize-space()='›']",
            ]:
                els = driver.find_elements(By.XPATH, xpath)
                if els:
                    next_btn = els[0]
                    break

            if next_btn is None:
                print("  ⏹️  No 'Next' button found — done.")
                break

            classes  = (next_btn.get_attribute("class") or "").lower()
            parent   = next_btn.find_element(By.XPATH, "..") if next_btn else None
            p_class  = (parent.get_attribute("class") or "").lower() if parent else ""

            if "disabled" in classes or "disabled" in p_class:
                print("  ⏹️  Last page reached.")
                break

            driver.execute_script("arguments[0].click();", next_btn)
            page_num += 1
            time.sleep(PAGE_DELAY)

        except Exception as ex:
            print(f"  ⏹️  Pagination ended: {ex}")
            break

    return all_rows

# ═══════════════════════════════════════════════════════════
# MERGE WITH EXISTING CSV (preserve manual edits)
# ═══════════════════════════════════════════════════════════
def merge_with_existing(new_rows):
    df_new = pd.DataFrame(new_rows, columns=["Company", "NSE_Code", "BSE_Code"])

    if os.path.exists(CSV_LOCAL_PATH):
        try:
            df_old = pd.read_csv(CSV_LOCAL_PATH, dtype=str).fillna("")
            df_old.columns = [c.strip() for c in df_old.columns]

            old_lookup = {str(r.get("Company", "")).strip().lower(): dict(r)
                          for _, r in df_old.iterrows()}
            new_lookup = {str(r["Company"]).strip().lower(): dict(r)
                          for _, r in df_new.iterrows()}

            merged, seen = [], set()
            for key, row in new_lookup.items():
                merged.append(row)
                seen.add(key)
            for key, row in old_lookup.items():
                if key not in seen:
                    merged.append(row)   # keep manual-only rows

            df_final = pd.DataFrame(merged)
            for col in ["Company", "NSE_Code", "BSE_Code"]:
                if col not in df_final.columns:
                    df_final[col] = ""
            print(f"✅ Merged: {len(df_final)} rows total")
            return df_final
        except Exception as e:
            print(f"⚠️  Merge failed, using fresh data: {e}")

    print(f"✅ Fresh CSV: {len(df_new)} rows")
    return df_new

# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════
print("\n🚀 R03_Exchange_Codes.py starting...")
print(f"   Screen URL : {SCREEN_URL}")
print(f"   Target R2  : {R2_BUCKET}/{R2_CSV_KEY}")
print(f"   Timestamp  : {datetime.utcnow().isoformat()}Z")

try:
    if IS_GITHUB_ACTIONS:
        r2_download_existing()

    driver = init_driver()
    wait   = WebDriverWait(driver, 25)

    login(driver, wait)

    print(f"\n🌐 Navigating to screen: {SCREEN_URL}")
    scraped_rows = scrape_all_pages(driver, wait)
    print(f"\n📊 Total rows scraped: {len(scraped_rows)}")

    driver.quit()

    if not scraped_rows:
        print("❌ No rows scraped — aborting upload.")
        exit(1)

    df_final = merge_with_existing(scraped_rows)
    df_final = df_final.sort_values("Company", ignore_index=True)
    df_final.insert(0, "Sr_No", range(1, len(df_final) + 1))

    df_final.to_csv(CSV_LOCAL_PATH, index=False)
    print(f"\n💾 Saved: {CSV_LOCAL_PATH}  ({len(df_final)} rows)")
    print(df_final.head(10).to_string())

    if IS_GITHUB_ACTIONS:
        r2_upload(CSV_LOCAL_PATH, R2_CSV_KEY)

    print("\n✅ R03_Exchange_Codes.py completed successfully.")

except Exception as e:
    print(f"\n❌ Fatal error: {e}")
    traceback.print_exc()
    try:
        driver.quit()
    except:
        pass
    exit(1)
