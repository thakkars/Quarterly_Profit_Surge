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
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side

# ═══════════════════════════════════════════════════════════
# ENVIRONMENT DETECTION
# ═══════════════════════════════════════════════════════════
IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"

# ═══════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════
USERNAME = os.environ.get("SCREENER_USERNAME", "")
PASSWORD = os.environ.get("SCREENER_PASSWORD", "")

SCREEN_URL = "https://www.screener.in/screens/742984/test/"

if IS_GITHUB_ACTIONS:
    FOLDER_PATH = "/tmp/screener_data"
else:
    FOLDER_PATH = r"D:\Investment Research\Python Projects\Operating Profit more than 70% jump"

MASTER_FILE_PATH = os.path.join(FOLDER_PATH, "Master File.xlsx")
RAW_FILE_PATH    = os.path.join(FOLDER_PATH, "Raw Data.xlsx")

os.makedirs(FOLDER_PATH, exist_ok=True)

HEADLESS = IS_GITHUB_ACTIONS
PAGE_DELAY = 1.0

# ═══════════════════════════════════════════════════════════
# R2 CONFIG
# ═══════════════════════════════════════════════════════════
R2_BUCKET   = os.environ.get("R2_BUCKET_NAME", "")
R2_ACCESS   = os.environ.get("R2_ACCESS_KEY_ID", "")
R2_SECRET   = os.environ.get("R2_SECRET_ACCESS_KEY", "")
R2_ENDPOINT = os.environ.get("R2_ENDPOINT_URL", "")

R2_MASTER_KEY = "Master File.xlsx"
R2_RAW_KEY    = "Raw Data.xlsx"

# ═══════════════════════════════════════════════════════════
# R2 HELPERS
# ═══════════════════════════════════════════════════════════
def get_r2():
    import boto3
    return boto3.client(
        "s3",
        endpoint_url=R2_ENDPOINT,
        aws_access_key_id=R2_ACCESS,
        aws_secret_access_key=R2_SECRET,
        region_name="auto"
    )

def r2_download(key, path):
    try:
        get_r2().download_file(R2_BUCKET, key, path)
        print(f"⬇️ Downloaded {key}")
        return True
    except:
        print(f"ℹ️ {key} not found in R2")
        return False

def r2_upload(path, key):
    try:
        if os.path.exists(path):
            get_r2().upload_file(path, R2_BUCKET, key)
            print(f"⬆️ Uploaded {key}")
    except Exception as e:
        print(f"⚠️ Upload failed: {e}")

# ═══════════════════════════════════════════════════════════
# DOWNLOAD FROM R2
# ═══════════════════════════════════════════════════════════
if IS_GITHUB_ACTIONS:
    print("📥 Downloading from R2...")
    r2_download(R2_MASTER_KEY, MASTER_FILE_PATH)
    r2_download(R2_RAW_KEY, RAW_FILE_PATH)

# ═══════════════════════════════════════════════════════════
# DRIVER SETUP (FIXED)
# ═══════════════════════════════════════════════════════════
def init_driver():
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")

    if IS_GITHUB_ACTIONS:
        return webdriver.Chrome(options=opts)
    else:
        from webdriver_manager.chrome import ChromeDriverManager
        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=opts
        )

driver = init_driver()
wait = WebDriverWait(driver, 15)

# ═══════════════════════════════════════════════════════════
# HELPERS
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

def safe_get(url):
    try:
        driver.get(url)
        return True
    except:
        return False

def get_company_full_name():
    try:
        return wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h1.margin-0"))
        ).text.strip()
    except:
        return "Unknown"

# ═══════════════════════════════════════════════════════════
# SCRAPING
# ═══════════════════════════════════════════════════════════
def collect_all_company_data():
    safe_get(SCREEN_URL)
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[href*='/company/']")))

    data = []
    seen = set()

    while True:
        rows = driver.find_elements(By.CSS_SELECTOR, "a[href*='/company/']")
        for a in rows:
            url = a.get_attribute("href")
            name = a.text.strip()
            if url not in seen:
                seen.add(url)
                data.append({"short_name": name, "url": url})

        try:
            nxt = driver.find_element(By.XPATH, "//a[contains(text(),'Next')]")
            if "disabled" in nxt.get_attribute("class"):
                break
            driver.execute_script("arguments[0].click();", nxt)
            time.sleep(1)
        except:
            break

    return data

def scrape_company(c):
    if not safe_get(c["url"]):
        return {}
    return {
        "Name": c["short_name"],
        "Company (Page)": get_company_full_name(),
        "Company URL": c["url"]
    }

# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════
try:
    login()

    companies = collect_all_company_data()

    data = []
    for c in companies:
        data.append(scrape_company(c))
        time.sleep(PAGE_DELAY)

    df = pd.DataFrame(data)

    # SAVE RAW
    with pd.ExcelWriter(
        RAW_FILE_PATH,
        engine="openpyxl",
        mode="a" if os.path.exists(RAW_FILE_PATH) else "w",
        if_sheet_exists="replace"
    ) as writer:
        sheet = datetime.now().strftime("%d-%m-%y_%H-%M")
        df.to_excel(writer, sheet_name=sheet, index=False)

    print("✅ RAW saved")

    # SAVE MASTER (simple version)
    df.to_excel(MASTER_FILE_PATH, index=False)
    print("✅ MASTER saved")

    # UPLOAD
    if IS_GITHUB_ACTIONS:
        print("📤 Uploading to R2...")
        r2_upload(MASTER_FILE_PATH, R2_MASTER_KEY)
        r2_upload(RAW_FILE_PATH, R2_RAW_KEY)

finally:
    driver.quit()
