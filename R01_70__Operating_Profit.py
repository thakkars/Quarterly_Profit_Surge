import time
import random
import traceback
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, InvalidSessionIdException
from datetime import datetime
import os
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Side

# webdriver_manager only needed locally — GitHub Actions uses setup-chrome action
if os.environ.get("GITHUB_ACTIONS", "false").lower() != "true":
    from webdriver_manager.chrome import ChromeDriverManager

# ═══════════════════════════════════════════════════════════
# ENVIRONMENT DETECTION
# ═══════════════════════════════════════════════════════════
IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"

# ═══════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════
USERNAME   = os.environ.get("SCREENER_USERNAME", "deepak.khatwani2712@gmail.com")
PASSWORD   = os.environ.get("SCREENER_PASSWORD", "gaurav23")
SCREEN_URL = "https://www.screener.in/screens/3391124/test/"

if IS_GITHUB_ACTIONS:
    FOLDER_PATH = "/tmp/screener_data"
else:
    FOLDER_PATH = r"D:\Investment Research\Python Projects\Operating Profit more than 70% jump"

MASTER_FILE_PATH    = os.path.join(FOLDER_PATH, "Master File.xlsx")
RAW_FILE_PATH       = os.path.join(FOLDER_PATH, "Raw Data.xlsx")
QUALIFIED_FILE_PATH = os.path.join(FOLDER_PATH, "Qualified_Stocks.xlsx")
INDUSTRY_FILE_PATH  = os.path.join(FOLDER_PATH, "Screener_Industry_Data.xlsx")

os.makedirs(FOLDER_PATH, exist_ok=True)

HEADLESS   = IS_GITHUB_ACTIONS
PAGE_DELAY = 1.0

# ─── R2 CONFIG ─────────────────────────────────────────────
R2_BUCKET      = os.environ.get("R2_BUCKET_NAME",        "quarterly-profit-surge")
R2_ACCESS_KEY  = os.environ.get("R2_ACCESS_KEY_ID",      "")
R2_SECRET_KEY  = os.environ.get("R2_SECRET_ACCESS_KEY",  "")
R2_ENDPOINT    = os.environ.get("R2_ENDPOINT_URL",       "")

R2_MASTER_KEY    = "Master File.xlsx"
R2_RAW_KEY       = "Raw Data.xlsx"
R2_QUALIFIED_KEY = "Qualified_Stocks.xlsx"
R2_INDUSTRY_KEY  = "Screener_Industry_Data.xlsx"

print(f"🌍 Environment : {'GitHub Actions (cloud)' if IS_GITHUB_ACTIONS else 'Local Windows'}")
print(f"📁 Working dir : {FOLDER_PATH}")


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
        client = get_r2_client()
        client.download_file(R2_BUCKET, r2_key, local_path)
        print(f"⬇️  Downloaded from R2: {r2_key} → {local_path}")
        return True
    except Exception as e:
        if "NoSuchKey" in str(e) or "404" in str(e):
            print(f"ℹ️  {r2_key} not found in R2 (first run — will create fresh).")
        else:
            print(f"⚠️  R2 download failed for {r2_key}: {e}")
        return False

def r2_upload(local_path, r2_key):
    try:
        if not os.path.exists(local_path):
            print(f"⚠️  Skipping R2 upload — file not found: {local_path}")
            return False
        client = get_r2_client()
        client.upload_file(local_path, R2_BUCKET, r2_key)
        print(f"⬆️  Uploaded to R2 : {local_path} → {r2_key}")
        return True
    except Exception as e:
        print(f"⚠️  R2 upload failed for {r2_key}: {e}")
        return False


# ═══════════════════════════════════════════════════════════
# STEP 0 — DOWNLOAD FILES FROM R2 (GitHub Actions only)
# ═══════════════════════════════════════════════════════════
if IS_GITHUB_ACTIONS:
    print("\n📥 Downloading existing files from Cloudflare R2...")
    r2_download(R2_MASTER_KEY,    MASTER_FILE_PATH)
    r2_download(R2_RAW_KEY,       RAW_FILE_PATH)
    r2_download(R2_QUALIFIED_KEY, QUALIFIED_FILE_PATH)
    r2_download(R2_INDUSTRY_KEY,  INDUSTRY_FILE_PATH)
    print("📥 Download step complete.\n")


# ═══════════════════════════════════════════════════════════
# INDUSTRY LOOKUP  (loaded once at startup)
# ═══════════════════════════════════════════════════════════
def load_industry_maps():
    if not os.path.exists(INDUSTRY_FILE_PATH):
        print("⚠️  Screener_Industry_Data.xlsx not found — industry columns will be blank.")
        return {}, {}
    try:
        df = pd.read_excel(INDUSTRY_FILE_PATH, sheet_name="Screener Industry")
        df["Name_Clean"] = df["Full Name"].astype(str).str.strip().str.lower()
        industry_map = df.set_index("Name_Clean")["Industry"].to_dict()
        group_map    = df.set_index("Name_Clean")["Industry Group"].to_dict()
        print(f"✅ Industry lookup loaded — {len(industry_map):,} companies.")
        return industry_map, group_map
    except Exception as e:
        print(f"⚠️  Failed to load industry data: {e}")
        return {}, {}


def enrich_with_industry(row_dict, industry_map, group_map):
    name_clean = str(row_dict.get("Company (Page)", "")).strip().lower()
    row_dict["Industry"]       = industry_map.get(name_clean, "Not Found")
    row_dict["Industry Group"] = group_map.get(name_clean, "Not Found")
    return row_dict


# ═══════════════════════════════════════════════════════════
# DRIVER SETUP
# ═══════════════════════════════════════════════════════════
def init_driver():
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1600,900")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    if IS_GITHUB_ACTIONS:
        return webdriver.Chrome(options=opts)
    else:
        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()), options=opts
        )

driver = init_driver()
wait   = WebDriverWait(driver, 15)


# ═══════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════
def login():
    driver.get("https://www.screener.in/login/")
    try:
        wait.until(EC.presence_of_element_located((By.NAME, "username")))
        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.NAME, "password").submit()
        time.sleep(3)
    except:
        pass

def safe_text(el):
    try:
        return el.text.strip()
    except:
        return ""

def safe_get(url, retries=1):
    global driver, wait
    try:
        driver.get(url)
        return True
    except InvalidSessionIdException:
        print("⚠️  Session died. Restarting Chrome...")
        try:
            driver.quit()
        except:
            pass
        driver = init_driver()
        wait   = WebDriverWait(driver, 15)
        login()
        if retries > 0:
            return safe_get(url, retries - 1)
        return False
    except Exception as e:
        print(f"⚠️  Unexpected error in safe_get: {e}")
        return False

def get_company_full_name():
    try:
        h1 = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.margin-0")))
        return h1.text.strip()
    except:
        return "Unknown"

def select_consolidated_if_exists():
    try:
        main_window = driver.current_window_handle
        tab = driver.find_element(By.LINK_TEXT, "Consolidated")
        tab.click()
        time.sleep(1.5)
        for w in driver.window_handles:
            if w != main_window:
                driver.switch_to.window(w)
                return main_window
        return main_window
    except:
        return driver.current_window_handle


# ═══════════════════════════════════════════════════════════
# METRIC EXTRACTION
# ═══════════════════════════════════════════════════════════
def get_top_dashboard_metrics():
    metrics = {"Promoter Holding": 0.0, "Pledged %": 0.0, "WC Days": 0.0}
    try:
        ratio_items = driver.find_elements(By.CSS_SELECTOR, "ul#top-ratios li")
        for li in ratio_items:
            try:
                name_el = li.find_element(By.CSS_SELECTOR, "span.name")
                val_el  = li.find_element(By.CSS_SELECTOR, "span.number")
                label   = name_el.text.lower().strip()
                value   = float(val_el.text.strip().replace(",", "")) if val_el.text.strip() else 0.0
                if "promoter holding" in label:
                    metrics["Promoter Holding"] = value
                elif "pledged percentage" in label or "pledged" in label:
                    metrics["Pledged %"] = value
                elif "wc days" in label or "working capital days" in label:
                    metrics["WC Days"] = value
            except:
                continue
    except:
        pass
    return metrics

def extract_operating_profit():
    try:
        table = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//section[@id='quarters']//table")))
    except TimeoutException:
        return {}

    headers    = [th.text.strip() for th in table.find_elements(By.TAG_NAME, "th")]
    rows       = table.find_elements(By.TAG_NAME, "tr")
    target_row = None

    for r in rows:
        if "operating profit" in r.text.lower():
            target_row = r
            break
    if not target_row:
        for r in rows:
            if "financing profit" in r.text.lower():
                target_row = r
                break
    if not target_row:
        return {}

    values = [safe_text(c) for c in target_row.find_elements(By.TAG_NAME, "td")]
    return {h: v for h, v in zip(headers[1:], values[1:])}

def collect_all_company_data():
    if not safe_get(SCREEN_URL):
        return []
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[href*='/company/']")))
    company_list = []
    seen_urls    = set()

    while True:
        rows = driver.find_elements(By.CSS_SELECTOR, "a[href*='/company/']")
        for a in rows:
            href       = a.get_attribute("href")
            short_name = a.text.strip()
            if href and "/company/" in href and href not in seen_urls:
                seen_urls.add(href)
                company_list.append({"short_name": short_name, "url": href})
        try:
            next_btn = driver.find_element(By.XPATH, "//a[contains(text(), 'Next')]")
            if "disabled" in (next_btn.get_attribute("class") or "").lower():
                break
            driver.execute_script("arguments[0].click();", next_btn)
            wait.until(EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "a[href*='/company/']")))
            time.sleep(1.5)
        except:
            break

    print(f"✅ Collected {len(company_list)} companies from screen.")
    return company_list

def scrape_company(company_info):
    url        = company_info['url']
    short_name = company_info['short_name']

    if not safe_get(url):
        return {"Name": short_name, "Company (Page)": "Load Failed", "date": datetime.now().strftime("%Y-%m-%d")}

    main_window = driver.current_window_handle
    try:
        _           = select_consolidated_if_exists()
        full_name   = get_company_full_name()
        op_data     = extract_operating_profit()
        top_metrics = get_top_dashboard_metrics()

        for w in driver.window_handles:
            if w != main_window:
                driver.close()
        driver.switch_to.window(main_window)

        return {
            "Name"           : short_name,
            "Company (Page)" : full_name,
            "Industry Group" : "",    
            "Industry"       : "",    
            **op_data,
            **top_metrics,
            "date"           : datetime.now().strftime("%Y-%m-%d")
        }
    except Exception as e:
        print(f"⚠️  Error scraping {short_name}: {e}")
        for w in driver.window_handles:
            if w != main_window:
                driver.close()
        driver.switch_to.window(main_window)
        return {"Name": short_name, "Company (Page)": "Error", "date": datetime.now().strftime("%Y-%m-%d")}


# ═══════════════════════════════════════════════════════════
# FORMATTING HELPERS
# ═══════════════════════════════════════════════════════════
THIN_BORDER  = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'),  bottom=Side(style='thin'),
)
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')

def format_worksheet(ws, header_row=1):
    for row in ws.iter_rows():
        for cell in row:
            cell.border = THIN_BORDER
            if cell.row == header_row or cell.col_idx == 1:
                cell.alignment = CENTER_ALIGN


# ═══════════════════════════════════════════════════════════
# DATA TRANSFORMATION & SAVING HELPERS
# ═══════════════════════════════════════════════════════════
def detect_latest_quarter(row_dict):
    non_quarter_keys = {
        "Name", "Company (Page)", "Industry Group", "Industry",
        "Promoter Holding", "Pledged %", "WC Days", "date"
    }
    quarter_candidates = []
    for k in row_dict:
        if k in non_quarter_keys or "(>70%)" in k:
            continue
        try:
            dt = pd.to_datetime(k, format="%b %Y")
            quarter_candidates.append((dt, k))
        except Exception:
            pass

    if not quarter_candidates:
        return None

    quarter_candidates.sort(key=lambda x: x[0], reverse=True)
    return quarter_candidates[0][0].strftime("%b %y")

def get_existing_companies_in_sheet(sheet_name):
    if not os.path.exists(RAW_FILE_PATH):
        return set()
    try:
        xl = pd.ExcelFile(RAW_FILE_PATH)
        if sheet_name not in xl.sheet_names:
            return set()
        df_existing = pd.read_excel(RAW_FILE_PATH, sheet_name=sheet_name)
        if "Name" in df_existing.columns:
            return set(df_existing["Name"].dropna().astype(str).str.strip())
        return set()
    except Exception as e:
        print(f"⚠️  Could not read sheet '{sheet_name}': {e}")
        return set()

def transform_and_order_data(df):
    """
    Reorders columns and creates dynamic YoY 70% growth matrix columns.
    Expected order: Base Cols -> Sorted Quarters -> Metrics -> Matrix
    """
    base_cols = ["date", "Name", "Company (Page)", "Industry", "Industry Group"]
    metric_cols = ["Pledged %", "Promoter Holding", "WC Days"]
    
    for col in base_cols + metric_cols:
        if col not in df.columns:
            df[col] = ""

    quarter_cols_raw = [c for c in df.columns if c not in base_cols and c not in metric_cols and "(>70%)" not in c]

    parsed_dates = []
    for c in quarter_cols_raw:
        try:
            dt = pd.to_datetime(c, format="%b %Y")
            parsed_dates.append((dt, c))
        except Exception:
            pass 
    
    parsed_dates.sort(key=lambda x: x[0], reverse=True)
    sorted_quarter_cols = [x[1] for x in parsed_dates]

    temp_df = pd.DataFrame()
    for c in sorted_quarter_cols:
        temp_df[c] = pd.to_numeric(df[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    matrix_cols = []
    for dt, col_current in parsed_dates:
        try:
            dt_prev_year = dt.replace(year=dt.year - 1)
            col_prev = dt_prev_year.strftime("%b %Y")
            
            if col_prev in temp_df.columns:
                jump_col = f"{col_current} vs {col_prev} (>70%)"
                matrix_cols.append(jump_col)
                
                condition = (temp_df[col_current] >= 1.7 * temp_df[col_prev]) & (temp_df[col_prev] > 0)
                df[jump_col] = np.where(condition, "Yes", "")
        except Exception:
            pass

    final_cols = base_cols + sorted_quarter_cols + metric_cols + matrix_cols
    return df[final_cols]

BATCH_SIZE = 10

def append_rows_to_raw_excel(rows, sheet_name):
    """Append row-dicts to the quarter sheet with dynamic transformations."""
    if not rows:
        return

    df_new = transform_and_order_data(pd.DataFrame(rows))

    if os.path.exists(RAW_FILE_PATH):
        try:
            wb = load_workbook(RAW_FILE_PATH)
            if sheet_name in wb.sheetnames:
                df_existing = pd.read_excel(RAW_FILE_PATH, sheet_name=sheet_name)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                df_final = transform_and_order_data(df_combined)
            else:
                df_final = df_new
            wb.close()
            
            with pd.ExcelWriter(RAW_FILE_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_final.to_excel(writer, sheet_name=sheet_name, index=False)
                
        except Exception as e:
            print(f"⚠️  Error appending to existing Excel: {e}")
            return
    else:
        with pd.ExcelWriter(RAW_FILE_PATH, engine='openpyxl') as writer:
            df_new.to_excel(writer, sheet_name=sheet_name, index=False)

    wb = load_workbook(RAW_FILE_PATH)
    ws = wb[sheet_name]
    format_worksheet(ws)
    
    # Auto-adjust column widths for better readability
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[column_letter].width = max_length + 2

    wb.save(RAW_FILE_PATH)
    print(f"  ↳ Saved and formatted {len(rows)} rows to sheet '{sheet_name}'.")


# ═══════════════════════════════════════════════════════════
# QUALIFIED STOCKS FILTERING
# ═══════════════════════════════════════════════════════════
def generate_qualified_stocks(sheet_name):
    """Reads Raw Data, applies strict criteria, and creates/updates Qualified_Stocks.xlsx"""
    if not os.path.exists(RAW_FILE_PATH):
        print("⚠️  Raw Data file missing. Cannot generate Qualified Stocks.")
        return

    print(f"\n⚙️  Applying filters to generate Qualified_Stocks.xlsx for '{sheet_name}'...")
    
    try:
        # READ ONLY: Load the raw data
        df = pd.read_excel(RAW_FILE_PATH, sheet_name=sheet_name)
        
        # 1. Ensure the metrics exist before filtering
        metrics = ["Promoter Holding", "Pledged %", "WC Days"]
        for col in metrics:
            if col not in df.columns:
                print(f"⚠️  Missing column '{col}' in raw data. Cannot filter.")
                return
            
            # Convert to numeric. For WC Days, fill missing with 999 so it safely fails the <= 150 filter.
            if col == "WC Days":
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(999) 
            else:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # 2. Apply strict baseline filtering criteria
        qualified_df = df[
            (df["Promoter Holding"] >= 50.0) &
            (df["Pledged %"] <= 10.0) &
            (df["WC Days"] <= 150.0)
        ].copy()

        # 3. Dynamic 70% Growth Matrix Filter (Latest Quarter Only)
        # Find the latest quarter from the available columns
        quarter_candidates = []
        for col in qualified_df.columns:
            if col not in ["date", "Name", "Company (Page)", "Industry", "Industry Group", "Promoter Holding", "Pledged %", "WC Days"] and "(>70%)" not in col:
                try:
                    quarter_candidates.append((pd.to_datetime(col, format="%b %Y"), col))
                except: pass
        
        if quarter_candidates:
            # Sort to find the latest quarter date
            quarter_candidates.sort(key=lambda x: x[0], reverse=True)
            latest_dt = quarter_candidates[0][0]
            latest_q_col = quarter_candidates[0][1]
            
            # Calculate exactly 1 year ago to construct the target matrix column name
            prev_year_col = latest_dt.replace(year=latest_dt.year - 1).strftime("%b %Y")
            target_matrix_col = f"{latest_q_col} vs {prev_year_col} (>70%)"
            
            # Check if this matrix column was generated and filter where it is "Yes"
            if target_matrix_col in qualified_df.columns:
                qualified_df = qualified_df[qualified_df[target_matrix_col] == "Yes"]
                print(f"🔍  Applied 70% Growth check for latest quarter: '{target_matrix_col}'")
            else:
                print(f"⚠️  Growth matrix column '{target_matrix_col}' not found. Cannot apply 70% jump filter.")

        if qualified_df.empty:
            print("ℹ️  No stocks met the Qualified Criteria in this batch.")
            return
            
        print(f"🎯  Found {len(qualified_df)} Qualified Stocks!")

        # 4. Save to Qualified_Stocks.xlsx
        if os.path.exists(QUALIFIED_FILE_PATH):
            with pd.ExcelWriter(QUALIFIED_FILE_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                qualified_df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(QUALIFIED_FILE_PATH, engine='openpyxl') as writer:
                qualified_df.to_excel(writer, sheet_name=sheet_name, index=False)

        # 5. Apply standard formatting
        wb = load_workbook(QUALIFIED_FILE_PATH)
        ws = wb[sheet_name]
        format_worksheet(ws)
        
        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except: pass
            ws.column_dimensions[column_letter].width = max_length + 2

        wb.save(QUALIFIED_FILE_PATH)
        print(f"✅  Saved and formatted Qualified_Stocks.xlsx")

    except Exception as e:
        print(f"⚠️  Error generating Qualified Stocks: {e}")


# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════
try:
    INDUSTRY_MAP, GROUP_MAP = load_industry_maps()

    login()
    company_data_list = collect_all_company_data()

    if not company_data_list:
        print("⚠️  No companies found on the screen. Exiting.")
    else:
        total = len(company_data_list)

        print(f"\n🔍 Scraping company #1 to detect latest quarter...")
        first_row = scrape_company(company_data_list[0])
        first_row = enrich_with_industry(first_row, INDUSTRY_MAP, GROUP_MAP)
        print(f"✅ [1/{total}] {company_data_list[0]['short_name']}  "
              f"| Industry: {first_row.get('Industry', '—')}")
        time.sleep(PAGE_DELAY + random.uniform(0.3, 0.7))

        target_sheet = detect_latest_quarter(first_row)
        if not target_sheet:
            target_sheet = datetime.now().strftime("%b %y")
            print(f"⚠️  Could not detect quarter — using today: '{target_sheet}'")
        else:
            print(f"\n📅 Detected latest quarter : '{target_sheet}'")

        existing_names = get_existing_companies_in_sheet(target_sheet)
        if existing_names:
            print(f"📋 Sheet '{target_sheet}' already has {len(existing_names)} "
                  f"companies — those will be skipped.")
        else:
            print(f"📋 Sheet '{target_sheet}' is empty / new — scraping all companies.")

        first_name   = str(first_row.get("Name", "")).strip()
        pending_rows = []

        if first_name in existing_names:
            print(f"ℹ️  '{first_name}' already in sheet — not saving again.")
        else:
            pending_rows.append(first_row)

        remaining = [
            c for c in company_data_list[1:]
            if str(c['short_name']).strip() not in existing_names
        ]
        print(f"\n🔄 Scraping {len(remaining)} new companies "
              f"(skipping {total - 1 - len(remaining)} already present)...\n")

        for i, info in enumerate(remaining, start=2):
            row = scrape_company(info)
            row = enrich_with_industry(row, INDUSTRY_MAP, GROUP_MAP)
            pending_rows.append(row)
            print(f"✅ [{i}/{total}] {info['short_name']}  "
                  f"| Industry: {row.get('Industry', '—')}")
            time.sleep(PAGE_DELAY + random.uniform(0.3, 0.7))

            if len(pending_rows) >= BATCH_SIZE:
                append_rows_to_raw_excel(pending_rows, target_sheet)
                pending_rows = []

        if pending_rows:
            append_rows_to_raw_excel(pending_rows, target_sheet)

        print(f"\n✅ Raw Data.xlsx → sheet '{target_sheet}' is up to date.")

        # ── NEW: GENERATE QUALIFIED STOCKS ───────────────────────────────
        generate_qualified_stocks(target_sheet)

        if IS_GITHUB_ACTIONS:
            print("\n⬆️  Uploading updated files to Cloudflare R2...")
            r2_upload(RAW_FILE_PATH,       R2_RAW_KEY)
            r2_upload(QUALIFIED_FILE_PATH, R2_QUALIFIED_KEY)
            r2_upload(MASTER_FILE_PATH,    R2_MASTER_KEY)

except Exception as e:
    print(f"\n❌ Fatal error: {e}")
    traceback.print_exc()
finally:
    try:
        driver.quit()
    except:
        pass
    print("\n🏁 Script finished.")
