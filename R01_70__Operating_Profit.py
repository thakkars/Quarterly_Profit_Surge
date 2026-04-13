# ========================= IMPORTS =========================
import time, random, traceback, os
import pandas as pd
import numpy as np
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, InvalidSessionIdException

from openpyxl.styles import Alignment, Border, Side

if os.environ.get("GITHUB_ACTIONS", "false").lower() != "true":
    from webdriver_manager.chrome import ChromeDriverManager

# ========================= ENV =========================
IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"

USERNAME = os.environ.get("SCREENER_USERNAME")
PASSWORD = os.environ.get("SCREENER_PASSWORD")

SCREEN_URL = "https://www.screener.in/screens/3391124/test/"

FOLDER_PATH = "/tmp/screener_data" if IS_GITHUB_ACTIONS else r"D:\Investment Research\Python Projects\Operating Profit more than 70% jump"

MASTER_FILE_PATH = os.path.join(FOLDER_PATH, "Master File.xlsx")
RAW_FILE_PATH    = os.path.join(FOLDER_PATH, "Raw Data.xlsx")

os.makedirs(FOLDER_PATH, exist_ok=True)

# ========================= R2 =========================
R2_BUCKET = os.environ.get("R2_BUCKET_NAME")
R2_ACCESS_KEY = os.environ.get("R2_ACCESS_KEY_ID")
R2_SECRET_KEY = os.environ.get("R2_SECRET_ACCESS_KEY")
R2_ENDPOINT = os.environ.get("R2_ENDPOINT_URL")

def get_r2():
    import boto3
    return boto3.client(
        "s3",
        endpoint_url=R2_ENDPOINT,
        aws_access_key_id=R2_ACCESS_KEY,
        aws_secret_access_key=R2_SECRET_KEY,
        region_name="auto"
    )

def r2_download(key, path):
    try:
        get_r2().download_file(R2_BUCKET, key, path)
        print(f"⬇️ {key} downloaded")
    except:
        print(f"ℹ️ {key} not found (first run)")

def r2_upload(path, key):
    if os.path.exists(path):
        get_r2().upload_file(path, R2_BUCKET, key)
        print(f"⬆️ {key} uploaded")

if IS_GITHUB_ACTIONS:
    r2_download("Raw Data.xlsx", RAW_FILE_PATH)

# ========================= DRIVER =========================
def init_driver():
    opts = Options()
    if IS_GITHUB_ACTIONS:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")

    if IS_GITHUB_ACTIONS:
        return webdriver.Chrome(options=opts)
    else:
        return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

driver = init_driver()
wait = WebDriverWait(driver, 15)

# ========================= LOGIN =========================
def login():
    driver.get("https://www.screener.in/login/")
    wait.until(EC.presence_of_element_located((By.NAME, "username")))
    driver.find_element(By.NAME, "username").send_keys(USERNAME)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    driver.find_element(By.NAME, "password").submit()
    time.sleep(2)

# ========================= HELPERS =========================
def safe_get(url):
    try:
        driver.get(url)
        return True
    except:
        return False

def get_company_full_name():
    try:
        return wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.margin-0"))).text.strip()
    except:
        return "Unknown"

def extract_operating_profit():
    try:
        table = wait.until(EC.presence_of_element_located((By.XPATH, "//section[@id='quarters']//table")))
        headers = [th.text.strip() for th in table.find_elements(By.TAG_NAME, "th")]
        rows = table.find_elements(By.TAG_NAME, "tr")

        for r in rows:
            if "operating profit" in r.text.lower():
                values = [c.text.strip() for c in r.find_elements(By.TAG_NAME, "td")]
                return dict(zip(headers[1:], values[1:]))
    except:
        pass
    return {}

def get_top_dashboard_metrics():
    out = {"Promoter Holding":0,"Pledged %":0,"WC Days":0}
    try:
        for li in driver.find_elements(By.CSS_SELECTOR,"ul#top-ratios li"):
            try:
                name = li.find_element(By.CSS_SELECTOR,"span.name").text.lower()
                val  = float(li.find_element(By.CSS_SELECTOR,"span.number").text.replace(",",""))
                if "promoter" in name: out["Promoter Holding"]=val
                elif "pledged" in name: out["Pledged %"]=val
                elif "wc" in name: out["WC Days"]=val
            except: pass
    except: pass
    return out

# ========================= COMPANY LIST =========================
def collect_all_company_data():
    safe_get(SCREEN_URL)
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,"a[href*='/company/']")))
    out, seen = [], set()

    while True:
        for a in driver.find_elements(By.CSS_SELECTOR,"a[href*='/company/']"):
            href = a.get_attribute("href")
            if href and href not in seen:
                seen.add(href)
                out.append({"short_name":a.text.strip(),"url":href})
        try:
            nxt = driver.find_element(By.XPATH,"//a[contains(text(),'Next')]")
            if "disabled" in nxt.get_attribute("class"): break
            driver.execute_script("arguments[0].click();",nxt)
            time.sleep(1)
        except: break

    print(f"📊 {len(out)} companies collected")
    return out

# ========================= SCRAPE =========================
def scrape_company(info):
    if not safe_get(info["url"]): return None

    return {
        "Name": info["short_name"],
        "Company (Page)": get_company_full_name(),
        "URL": info["url"],
        **extract_operating_profit(),
        **get_top_dashboard_metrics()
    }

# ========================= MAIN =========================
try:
    login()

    company_list = collect_all_company_data()

    # LOAD EXISTING URLS
    existing_urls = set()
    if os.path.exists(RAW_FILE_PATH):
        try:
            sheets = pd.read_excel(RAW_FILE_PATH, sheet_name=None)
            last = list(sheets.keys())[-1]
            if "URL" in sheets[last].columns:
                existing_urls = set(sheets[last]["URL"].astype(str))
        except: pass

    new_list = [c for c in company_list if c["url"] not in existing_urls]
    print(f"🆕 New: {len(new_list)}")

    data = []
    for i,c in enumerate(new_list,1):
        r = scrape_company(c)
        if r: data.append(r)
        print(f"{i}/{len(new_list)}")

    df = pd.DataFrame(data)

    # ===== DETECT QUARTER =====
    quarters = [c for c in df.columns if any(x in c for x in ["Mar","Jun","Sep","Dec"])]
    quarter = quarters[0] if quarters else datetime.now().strftime("%b %Y")

    # ===== SAVE RAW =====
    if os.path.exists(RAW_FILE_PATH):
        try:
            old = pd.read_excel(RAW_FILE_PATH, sheet_name=quarter)
        except:
            old = pd.DataFrame()
        final = pd.concat([old, df], ignore_index=True)
        mode="a"
    else:
        final=df
        mode="w"

    with pd.ExcelWriter(RAW_FILE_PATH, engine="openpyxl", mode=mode, if_sheet_exists="replace" if mode=="a" else None) as w:
        final.to_excel(w, sheet_name=quarter, index=False)

    print(f"✅ RAW saved → {quarter}")

    # ===== UPLOAD =====
    if IS_GITHUB_ACTIONS:
        r2_upload(RAW_FILE_PATH, "Raw Data.xlsx")

finally:
    try:
        driver.quit()
    except:
        pass