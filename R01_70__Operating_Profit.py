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

if os.environ.get("GITHUB_ACTIONS", "false").lower() != "true":
    from webdriver_manager.chrome import ChromeDriverManager

IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"

USERNAME   = os.environ.get("SCREENER_USERNAME", "deepak.khatwani2712@gmail.com")
PASSWORD   = os.environ.get("SCREENER_PASSWORD", "gaurav23")

SCREEN_URL = "https://www.screener.in/screens/3391124/test/"

FOLDER_PATH = "/tmp/screener_data" if IS_GITHUB_ACTIONS else r"D:\Investment Research\Python Projects\Operating Profit more than 70% jump"

RAW_FILE_PATH = os.path.join(FOLDER_PATH, "Raw Data.xlsx")
os.makedirs(FOLDER_PATH, exist_ok=True)

HEADLESS   = IS_GITHUB_ACTIONS
PAGE_DELAY = 1.0

# DRIVER
def init_driver():
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1600,900")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")

    if IS_GITHUB_ACTIONS:
        return webdriver.Chrome(options=opts)
    else:
        return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

driver = init_driver()
wait   = WebDriverWait(driver, 15)

# LOGIN
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
        h1 = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.margin-0")))
        return h1.text.strip()
    except:
        return "Unknown"

def select_consolidated_if_exists():
    try:
        main_window = driver.current_window_handle
        tab = driver.find_element(By.LINK_TEXT, "Consolidated")
        tab.click()
        time.sleep(1)
        for w in driver.window_handles:
            if w != main_window:
                driver.switch_to.window(w)
                return main_window
        return main_window
    except:
        return driver.current_window_handle

def extract_operating_profit():
    try:
        table = wait.until(EC.presence_of_element_located((By.XPATH, "//section[@id='quarters']//table")))
    except:
        return {}

    headers = [th.text.strip() for th in table.find_elements(By.TAG_NAME, "th")]
    rows = table.find_elements(By.TAG_NAME, "tr")

    for r in rows:
        if "operating profit" in r.text.lower():
            values = [c.text.strip() for c in r.find_elements(By.TAG_NAME, "td")]
            return dict(zip(headers[1:], values[1:]))

    return {}

def get_top_dashboard_metrics():
    metrics = {"Promoter Holding": 0.0, "Pledged %": 0.0, "WC Days": 0.0}
    try:
        items = driver.find_elements(By.CSS_SELECTOR, "ul#top-ratios li")
        for li in items:
            try:
                name = li.find_element(By.CSS_SELECTOR, "span.name").text.lower()
                val  = float(li.find_element(By.CSS_SELECTOR, "span.number").text.replace(",", ""))
                if "promoter" in name:
                    metrics["Promoter Holding"] = val
                elif "pledged" in name:
                    metrics["Pledged %"] = val
                elif "wc" in name:
                    metrics["WC Days"] = val
            except:
                continue
    except:
        pass
    return metrics

def collect_all_company_data():
    safe_get(SCREEN_URL)
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[href*='/company/']")))

    company_list = []
    seen = set()

    while True:
        rows = driver.find_elements(By.CSS_SELECTOR, "a[href*='/company/']")
        for a in rows:
            href = a.get_attribute("href")
            name = a.text.strip()
            if href and href not in seen:
                seen.add(href)
                company_list.append({"short_name": name, "url": href})

        try:
            nxt = driver.find_element(By.XPATH, "//a[contains(text(),'Next')]")
            if "disabled" in nxt.get_attribute("class"):
                break
            driver.execute_script("arguments[0].click();", nxt)
            time.sleep(1)
        except:
            break

    print(f"Collected {len(company_list)} companies")
    return company_list

def scrape_company(info):
    url = info["url"]
    name = info["short_name"]

    if not safe_get(url):
        return None

    main = driver.current_window_handle

    try:
        select_consolidated_if_exists()
        full = get_company_full_name()
        op   = extract_operating_profit()
        met  = get_top_dashboard_metrics()

        return {
            "Name": name,
            "Company (Page)": full,
            "URL": url,
            **op,
            **met
        }

    except:
        return None

# MAIN
try:
    login()
    company_data_list = collect_all_company_data()

    # LOAD EXISTING RAW
    existing_urls = set()

    if os.path.exists(RAW_FILE_PATH):
        try:
            all_sheets = pd.read_excel(RAW_FILE_PATH, sheet_name=None)
            if all_sheets:
                last_sheet = list(all_sheets.keys())[-1]
                df_exist = all_sheets[last_sheet]
                if "URL" in df_exist.columns:
                    existing_urls = set(df_exist["URL"].dropna().astype(str))
                    print(f"Found {len(existing_urls)} existing companies")
        except:
            pass

    # FILTER NEW
    new_company_list = [c for c in company_data_list if c["url"] not in existing_urls]

    print(f"New companies: {len(new_company_list)}")

    data = []

    for i, comp in enumerate(new_company_list, 1):
        row = scrape_company(comp)
        if row:
            data.append(row)
        print(f"{i}/{len(new_company_list)} done")
        time.sleep(PAGE_DELAY)

    # SAVE RAW
    df_new = pd.DataFrame(data)

    sheet_name = datetime.now().strftime("%b %Y")

    if os.path.exists(RAW_FILE_PATH):
        try:
            df_old = pd.read_excel(RAW_FILE_PATH, sheet_name=sheet_name)
        except:
            df_old = pd.DataFrame()
    else:
        df_old = pd.DataFrame()

    df_final = pd.concat([df_old, df_new], ignore_index=True)

    with pd.ExcelWriter(RAW_FILE_PATH, engine="openpyxl", mode="a" if os.path.exists(RAW_FILE_PATH) else "w", if_sheet_exists="replace") as writer:
        df_final.to_excel(writer, sheet_name=sheet_name, index=False)

    print("RAW updated successfully")

finally:
    try:
        driver.quit()
    except:
        pass