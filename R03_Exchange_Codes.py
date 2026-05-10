import time
import traceback
import pandas as pd
import os
import json
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

HEADLESS  = IS_GITHUB_ACTIONS
PAGE_DELAY = 1.5

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
    """Download existing CSV from R2 so we can merge/preserve manual edits."""
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
        time.sleep(3)
        print("✅ Logged in to Screener.in")
    except Exception as e:
        print(f"⚠️  Login issue: {e}")

# ═══════════════════════════════════════════════════════════
# SCRAPE SCREEN TABLE (all pages)
# ═══════════════════════════════════════════════════════════
def scrape_all_pages(driver, wait):
    """
    Scrapes every page of the Screener screen at SCREEN_URL.
    Expects table columns: S.No., Name, BSE Code, NSE Code, ...
    Returns list of dicts with keys: Sr_No, Company, NSE_Code, BSE_Code
    """
    driver.get(SCREEN_URL)
    time.sleep(PAGE_DELAY)

    # Wait for the results table to load
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.data-table")))
    except:
        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
        except Exception as e:
            print(f"❌ Could not find results table: {e}")
            return []

    all_rows = []
    page_num = 1

    while True:
        print(f"  📄 Scraping page {page_num}...")
        time.sleep(PAGE_DELAY)

        # Parse table headers
        try:
            table = driver.find_element(By.CSS_SELECTOR, "table.data-table")
        except:
            try:
                table = driver.find_element(By.CSS_SELECTOR, "table")
            except:
                print("  ⚠️  Table not found on this page, stopping.")
                break

        headers = [th.text.strip() for th in table.find_elements(By.CSS_SELECTOR, "thead th")]
        print(f"  Headers found: {headers}")

        # Map column indices (case-insensitive)
        col_map = {h.lower(): i for i, h in enumerate(headers)}

        # Find relevant column indices
        name_idx = next((col_map[k] for k in col_map if "name" in k), None)
        bse_idx  = next((col_map[k] for k in col_map if "bse" in k), None)
        nse_idx  = next((col_map[k] for k in col_map if "nse" in k), None)

        if name_idx is None:
            print("  ⚠️  Could not find 'Name' column — skipping page.")
            break

        # Parse each data row
        rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
        if not rows:
            print("  ℹ️  No rows on this page.")
            break

        for tr in rows:
            cells = tr.find_elements(By.TAG_NAME, "td")
            if not cells or len(cells) <= name_idx:
                continue

            # Full company name — get from anchor tag inside name cell if present
            name_cell = cells[name_idx]
            anchors   = name_cell.find_elements(By.TAG_NAME, "a")
            company   = anchors[0].text.strip() if anchors else name_cell.text.strip()

            bse_code = cells[bse_idx].text.strip()  if bse_idx  is not None and len(cells) > bse_idx  else ""
            nse_code = cells[nse_idx].text.strip()  if nse_idx  is not None and len(cells) > nse_idx  else ""

            # Clean up codes
            bse_code = bse_code.replace(".0", "") if bse_code.endswith(".0") else bse_code
            nse_code = "" if nse_code in ("-", "—", "–", "N/A", "NA") else nse_code
            bse_code = "" if bse_code in ("-", "—", "–", "N/A", "NA") else bse_code

            if not company:
                continue

            all_rows.append({
                "Company":  company,
                "NSE_Code": nse_code,
                "BSE_Code": bse_code,
            })

        print(f"  ✅ Page {page_num}: {len(rows)} rows collected (total so far: {len(all_rows)})")

        # Navigate to next page
        try:
            next_btn = driver.find_element(
                By.XPATH,
                "//a[contains(@class,'next') or contains(text(),'Next') or contains(text(),'›')]"
            )
            classes = next_btn.get_attribute("class") or ""
            if "disabled" in classes.lower():
                print("  ⏹️  Last page reached.")
                break
            driver.execute_script("arguments[0].click();", next_btn)
            page_num += 1
            time.sleep(PAGE_DELAY)
        except Exception:
            print("  ⏹️  No next page button found — done.")
            break

    return all_rows

# ═══════════════════════════════════════════════════════════
# MERGE WITH EXISTING CSV (preserve manual edits)
# ═══════════════════════════════════════════════════════════
def merge_with_existing(new_rows):
    """
    If an existing exchange_codes.csv exists (downloaded from R2),
    preserve any manually added/edited rows, and update/add scraped rows.
    Key = Company name (case-insensitive).
    """
    df_new = pd.DataFrame(new_rows, columns=["Company", "NSE_Code", "BSE_Code"])

    if os.path.exists(CSV_LOCAL_PATH):
        try:
            df_old = pd.read_csv(CSV_LOCAL_PATH, dtype=str).fillna("")
            # Normalise column names
            df_old.columns = [c.strip() for c in df_old.columns]

            # Build a lookup from old data keyed by lowercase company name
            old_lookup = {str(r.get("Company", "")).strip().lower(): r
                          for _, r in df_old.iterrows()}
            new_lookup = {str(r["Company"]).strip().lower(): r
                          for _, r in df_new.iterrows()}

            # Merge: scraped data wins for NSE/BSE codes, but keep extra manual columns
            merged = []
            seen   = set()
            for key, new_row in new_lookup.items():
                merged.append(new_row)
                seen.add(key)
            # Append old rows not present in new scrape (manual-only entries)
            for key, old_row in old_lookup.items():
                if key not in seen:
                    merged.append(old_row)

            df_final = pd.DataFrame(merged)
            # Ensure core columns exist
            for col in ["Company", "NSE_Code", "BSE_Code"]:
                if col not in df_final.columns:
                    df_final[col] = ""
            print(f"✅ Merged: {len(df_final)} total rows (scraped + preserved manual entries)")
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
    # 1. Download existing CSV from R2 (to preserve manual edits)
    if IS_GITHUB_ACTIONS:
        r2_download_existing()

    # 2. Scrape Screener
    driver = init_driver()
    wait   = WebDriverWait(driver, 20)

    login(driver, wait)

    print(f"\n🌐 Navigating to screen: {SCREEN_URL}")
    scraped_rows = scrape_all_pages(driver, wait)
    print(f"\n📊 Total rows scraped: {len(scraped_rows)}")

    driver.quit()

    if not scraped_rows:
        print("❌ No rows scraped — aborting upload.")
        exit(1)

    # 3. Merge with existing
    df_final = merge_with_existing(scraped_rows)

    # 4. Sort by Company name
    df_final = df_final.sort_values("Company", ignore_index=True)
    df_final.insert(0, "Sr_No", range(1, len(df_final) + 1))

    # 5. Save CSV
    df_final.to_csv(CSV_LOCAL_PATH, index=False)
    print(f"\n💾 Saved CSV: {CSV_LOCAL_PATH} ({len(df_final)} rows)")
    print(df_final.head(10).to_string())

    # 6. Upload to R2
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
