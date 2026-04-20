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
import json
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Side

if os.environ.get("GITHUB_ACTIONS", "false").lower() != "true":
    from webdriver_manager.chrome import ChromeDriverManager

# ═══════════════════════════════════════════════════════════
# ENVIRONMENT DETECTION & CONFIG
# ═══════════════════════════════════════════════════════════
IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"

USERNAME   = os.environ.get("SCREENER_USERNAME", "deepak.khatwani2712@gmail.com")
PASSWORD   = os.environ.get("SCREENER_PASSWORD", "gaurav23")
SCREEN_URL = "https://www.screener.in/screens/742984/top-growers/"

if IS_GITHUB_ACTIONS:
    FOLDER_PATH = "/tmp/screener_data"
else:
    FOLDER_PATH = r"D:\Investment Research\Python Projects\Operating Profit more than 70% jump"

MASTER_FILE_PATH    = os.path.join(FOLDER_PATH, "Master File.xlsx")
RAW_FILE_PATH       = os.path.join(FOLDER_PATH, "Raw Data.xlsx")
RAW_JSON_PATH       = os.path.join(FOLDER_PATH, "rawdata.json")
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
R2_RAW_JSON_KEY  = "rawdata.json"
R2_INDUSTRY_KEY  = "Screener_Industry_Data.xlsx"

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
        print(f"⚠️  R2 download failed for {r2_key}: {e}")
        return False

def r2_upload(local_path, r2_key):
    try:
        if not os.path.exists(local_path):
            print(f"⚠️  Skipping R2 upload — file not found: {local_path}")
            return False
        client = get_r2_client()
        
        extra_args = {}
        if r2_key.endswith('.json'):
            extra_args = {
                'ContentType': 'application/json',
                'CacheControl': 'no-cache, no-store, must-revalidate'
            }
            
        client.upload_file(local_path, R2_BUCKET, r2_key, ExtraArgs=extra_args)
        print(f"⬆️  Uploaded to R2 : {local_path} → {r2_key}")
        return True
    except Exception as e:
        print(f"⚠️  R2 upload failed for {r2_key}: {e}")
        return False

if IS_GITHUB_ACTIONS:
    print("\n📥 Downloading existing files from Cloudflare R2...")
    r2_download(R2_MASTER_KEY,    MASTER_FILE_PATH)
    r2_download(R2_RAW_KEY,       RAW_FILE_PATH)
    r2_download(R2_INDUSTRY_KEY,  INDUSTRY_FILE_PATH)

# ═══════════════════════════════════════════════════════════
# MAPPING & ENRICHMENT
# ═══════════════════════════════════════════════════════════
def load_industry_maps():
    if not os.path.exists(INDUSTRY_FILE_PATH):
        return {}, {}
    try:
        df = pd.read_excel(INDUSTRY_FILE_PATH, sheet_name="Screener Industry")
        df["Name_Clean"] = df["Full Name"].astype(str).str.strip().str.lower()
        return df.set_index("Name_Clean")["Industry"].to_dict(), df.set_index("Name_Clean")["Industry Group"].to_dict()
    except:
        return {}, {}

def load_symbol_maps():
    """Loads the exact NSE Code and Screener Name from the master file."""
    if not os.path.exists(MASTER_FILE_PATH):
        return {}
    try:
        df = pd.read_excel(MASTER_FILE_PATH, sheet_name="screener_all_stocks")
        # Clean name for matching (removing spaces and making lower case helps catch slight variations)
        df["Match_Name"] = df["Name"].astype(str).str.strip().str.lower()
        
        mapping = {}
        for _, row in df.iterrows():
            mapping[row["Match_Name"]] = {
                "NSE_Code": str(row["NSE Code"]).strip() if pd.notna(row["NSE Code"]) else "",
                "Screener_Name": str(row["Name"]).strip()
            }
        return mapping
    except Exception as e:
        print(f"⚠️ Could not load symbol maps: {e}")
        return {}

def enrich_with_data(row_dict, industry_map, group_map, symbol_map):
    name_clean = str(row_dict.get("Company (Page)", "")).strip().lower()
    short_name_clean = str(row_dict.get("Name", "")).strip().lower()
    
    row_dict["Industry"]       = industry_map.get(name_clean, "Not Found")
    row_dict["Industry Group"] = group_map.get(name_clean, "Not Found")
    
    # Try to find the exact codes using either the short name or the full company page name
    mapped_data = symbol_map.get(short_name_clean) or symbol_map.get(name_clean) or {}
    
    row_dict["NSE Code Exact"] = mapped_data.get("NSE_Code", "")
    row_dict["Screener Name Exact"] = mapped_data.get("Screener_Name", str(row_dict.get("Name", "")))
    
    return row_dict

# ═══════════════════════════════════════════════════════════
# SCRAPER LOGIC
# ═══════════════════════════════════════════════════════════
def init_driver():
    opts = Options()
    if HEADLESS: opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1600,900")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    if IS_GITHUB_ACTIONS: return webdriver.Chrome(options=opts)
    else: return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

driver = init_driver()
wait   = WebDriverWait(driver, 15)

def login():
    driver.get("https://www.screener.in/login/")
    try:
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.NAME, "password").submit()
        time.sleep(3)
    except: pass

def safe_get(url):
    global driver, wait
    try:
        driver.get(url)
        return True
    except InvalidSessionIdException:
        try: driver.quit()
        except: pass
        driver = init_driver()
        wait   = WebDriverWait(driver, 15)
        login()
        return safe_get(url)
    except: return False

def get_company_full_name():
    try: return wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.margin-0"))).text.strip()
    except: return "Unknown"

def get_top_dashboard_metrics():
    metrics = {"Promoter Holding": 0.0, "Pledged %": 0.0, "WC Days": 0.0}
    try:
        ratio_items = driver.find_elements(By.CSS_SELECTOR, "ul#top-ratios li")
        for li in ratio_items:
            try:
                label = li.find_element(By.CSS_SELECTOR, "span.name").text.lower().strip()
                val_el = li.find_element(By.CSS_SELECTOR, "span.number").text.strip().replace(",", "")
                value = float(val_el) if val_el else 0.0
                if "promoter holding" in label: metrics["Promoter Holding"] = value
                elif "pledged" in label: metrics["Pledged %"] = value
                elif "wc days" in label: metrics["WC Days"] = value
            except: continue
    except: pass
    return metrics

def extract_operating_profit():
    try: table = wait.until(EC.presence_of_element_located((By.XPATH, "//section[@id='quarters']//table")))
    except: return {}
    headers = [th.text.strip() for th in table.find_elements(By.TAG_NAME, "th")]
    rows = table.find_elements(By.TAG_NAME, "tr")
    target_row = next((r for r in rows if "operating profit" in r.text.lower()), None)
    if not target_row: target_row = next((r for r in rows if "financing profit" in r.text.lower()), None)
    if not target_row: return {}
    values = [c.text.strip() for c in target_row.find_elements(By.TAG_NAME, "td")]
    return {h: v for h, v in zip(headers[1:], values[1:])}

def collect_all_company_data():
    if not safe_get(SCREEN_URL): return []
    company_list, seen_urls = [], set()
    while True:
        for a in driver.find_elements(By.CSS_SELECTOR, "a[href*='/company/']"):
            href, short_name = a.get_attribute("href"), a.text.strip()
            if href and "/company/" in href and href not in seen_urls:
                seen_urls.add(href); company_list.append({"short_name": short_name, "url": href})
        try:
            next_btn = driver.find_element(By.XPATH, "//a[contains(text(), 'Next')]")
            if "disabled" in (next_btn.get_attribute("class") or "").lower(): break
            driver.execute_script("arguments[0].click();", next_btn)
            time.sleep(1.5)
        except: break
    return company_list

def scrape_company(company_info):
    url, short_name = company_info['url'], company_info['short_name']
    if not safe_get(url): return {"Name": short_name, "Company (Page)": "Load Failed", "date": datetime.now().strftime("%Y-%m-%d")}
    try:
        try: driver.find_element(By.LINK_TEXT, "Consolidated").click(); time.sleep(1.5)
        except: pass
        return {
            "Name": short_name, "Company (Page)": get_company_full_name(),
            "Industry Group": "", "Industry": "",
            **extract_operating_profit(), **get_top_dashboard_metrics(),
            "date": datetime.now().strftime("%Y-%m-%d")
        }
    except: return {"Name": short_name, "Company (Page)": "Error", "date": datetime.now().strftime("%Y-%m-%d")}

def detect_latest_quarter(row_dict):
    non_quarter_keys = {"Sr no", "Name", "Company (Page)", "Industry Group", "Industry", "Promoter Holding", "Pledged %", "WC Days", "date", "NSE Code Exact", "Screener Name Exact"}
    quarter_candidates = []
    for k in row_dict:
        if k in non_quarter_keys or "(>70%)" in k: continue
        try: quarter_candidates.append((pd.to_datetime(k, format="%b %Y"), k))
        except: pass
    if not quarter_candidates: return None
    quarter_candidates.sort(key=lambda x: x[0], reverse=True)
    return quarter_candidates[0][0].strftime("%b %y")

def get_existing_companies_in_sheet(sheet_name):
    if not os.path.exists(RAW_FILE_PATH): return set()
    try:
        xl = pd.ExcelFile(RAW_FILE_PATH)
        if sheet_name not in xl.sheet_names: return set()
        df = pd.read_excel(RAW_FILE_PATH, sheet_name=sheet_name)
        if "Name" in df.columns: return set(df["Name"].dropna().astype(str).str.strip())
        return set()
    except: return set()

def transform_and_order_data(df):
    base_cols = ["Sr no", "date", "Name", "Company (Page)", "Industry", "Industry Group", "NSE Code Exact", "Screener Name Exact"]
    metric_cols = ["Pledged %", "Promoter Holding", "WC Days"]
    for col in base_cols + metric_cols:
        if col not in df.columns: df[col] = ""

    quarter_cols_raw = [c for c in df.columns if c not in base_cols and c not in metric_cols and "(>70%)" not in c]
    parsed_dates = []
    for c in quarter_cols_raw:
        try: parsed_dates.append((pd.to_datetime(c, format="%b %Y"), c))
        except: pass 
    parsed_dates.sort(key=lambda x: x[0], reverse=True)
    sorted_quarter_cols = [x[1] for x in parsed_dates]

    temp_df = pd.DataFrame()
    for c in sorted_quarter_cols: temp_df[c] = pd.to_numeric(df[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    matrix_cols = []
    for dt, col_current in parsed_dates:
        try:
            dt_prev_year = dt.replace(year=dt.year - 1)
            col_prev = dt_prev_year.strftime("%b %Y")
            if col_prev in temp_df.columns:
                jump_col = f"{dt.strftime('%b%y')} (>70%) {dt_prev_year.strftime('%b%y')}"
                matrix_cols.append(jump_col)
                df[jump_col] = np.where((temp_df[col_current] >= 1.7 * temp_df[col_prev]) & (temp_df[col_prev] > 0), "Yes", "")
        except: pass

    final_cols = base_cols + sorted_quarter_cols + metric_cols + matrix_cols
    df = df[final_cols].copy()
    df["Sr no"] = range(1, len(df) + 1)
    return df

def append_rows_to_raw_excel(rows, sheet_name):
    if not rows: return
    df_new = transform_and_order_data(pd.DataFrame(rows))
    if os.path.exists(RAW_FILE_PATH):
        try:
            wb = load_workbook(RAW_FILE_PATH)
            if sheet_name in wb.sheetnames:
                df_existing = pd.read_excel(RAW_FILE_PATH, sheet_name=sheet_name)
                df_final = transform_and_order_data(pd.concat([df_existing, df_new], ignore_index=True))
            else: df_final = df_new
            wb.close()
            with pd.ExcelWriter(RAW_FILE_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_final.to_excel(writer, sheet_name=sheet_name, index=False)
        except: return
    else:
        with pd.ExcelWriter(RAW_FILE_PATH, engine='openpyxl') as writer: df_new.to_excel(writer, sheet_name=sheet_name, index=False)

    wb = load_workbook(RAW_FILE_PATH)
    ws = wb[sheet_name]
    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            if cell.row == 1 or cell.col_idx == 1: cell.alignment = Alignment(horizontal='center', vertical='center')
    wb.save(RAW_FILE_PATH)

def generate_raw_json(sheet_name):
    if not os.path.exists(RAW_FILE_PATH): return
    try:
        xl = pd.ExcelFile(RAW_FILE_PATH)
        if sheet_name not in xl.sheet_names: sheet_name = xl.sheet_names[-1]
        df = pd.read_excel(RAW_FILE_PATH, sheet_name=sheet_name)
        df.columns = df.columns.astype(str).str.strip()
        df = df.fillna(0)
        df.to_json(RAW_JSON_PATH, orient="records")
        print(f"✅ Saved rawdata.json for Dashboard (Sheet: {sheet_name})")
    except Exception as e:
        print(f"⚠️ Error saving JSON: {e}")

# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════
try:
    INDUSTRY_MAP, GROUP_MAP = load_industry_maps()
    SYMBOL_MAP = load_symbol_maps()  # Load the new exact mappings
    login()
    company_data_list = collect_all_company_data()

    if company_data_list:
        first_scraped = scrape_company(company_data_list[0])
        first_row = enrich_with_data(first_scraped, INDUSTRY_MAP, GROUP_MAP, SYMBOL_MAP)
        target_sheet = detect_latest_quarter(first_row) or datetime.now().strftime("%b %y")
        
        existing_names = get_existing_companies_in_sheet(target_sheet)
        pending_rows = [] if str(first_row.get("Name", "")).strip() in existing_names else [first_row]

        remaining = [c for c in company_data_list[1:] if str(c['short_name']).strip() not in existing_names]
        
        for i, info in enumerate(remaining, start=2):
            scraped = scrape_company(info)
            enriched = enrich_with_data(scraped, INDUSTRY_MAP, GROUP_MAP, SYMBOL_MAP)
            pending_rows.append(enriched)
            if len(pending_rows) >= 10:
                append_rows_to_raw_excel(pending_rows, target_sheet)
                pending_rows = []

        if pending_rows: append_rows_to_raw_excel(pending_rows, target_sheet)
        generate_raw_json(target_sheet)

        if IS_GITHUB_ACTIONS:
            r2_upload(RAW_FILE_PATH, R2_RAW_KEY)
            r2_upload(RAW_JSON_PATH, R2_RAW_JSON_KEY)
            r2_upload(MASTER_FILE_PATH, R2_MASTER_KEY)

except Exception as e:
    print(f"\n❌ Fatal error: {e}")
    traceback.print_exc()
finally:
    try: driver.quit()
    except: pass
