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

if os.environ.get("GITHUB_ACTIONS", "false").lower() != "true":
    from webdriver_manager.chrome import ChromeDriverManager

IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"

USERNAME   = os.environ.get("SCREENER_USERNAME", "your_email")
PASSWORD   = os.environ.get("SCREENER_PASSWORD", "your_password")
SCREEN_URL = "https://www.screener.in/screens/3391124/test/"

if IS_GITHUB_ACTIONS:
    FOLDER_PATH = "/tmp/screener_data"
else:
    FOLDER_PATH = r"D:\Investment Research\Python Projects\Operating Profit more than 70% jump"

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
    wait.until(EC.presence_of_element_located((By.NAME, "username")))
    driver.find_element(By.NAME, "username").send_keys(USERNAME)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    driver.find_element(By.NAME, "password").submit()
    time.sleep(3)

def safe_text(el):
    try:
        return el.text.strip()
    except:
        return ""

def safe_get(url):
    try:
        driver.get(url)
        return True
    except:
        return False

# FULL NAME
def get_company_full_name():
    try:
        return wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "h1.margin-0"))).text.strip()
    except:
        return "Unknown"

# ✅ INDUSTRY SCRAPER (NEW)
def get_industry_info():
    industry = ""
    group = ""
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, "#company-info p")
        for r in rows:
            text = r.text.lower()

            if "industry group" in text:
                try:
                    group = r.find_element(By.TAG_NAME, "a").text.strip()
                except:
                    group = r.text.split(":")[-1].strip()

            elif "industry" in text:
                try:
                    industry = r.find_element(By.TAG_NAME, "a").text.strip()
                except:
                    industry = r.text.split(":")[-1].strip()
    except:
        pass

    return industry, group

# CONSOLIDATED
def select_consolidated_if_exists():
    try:
        tab = driver.find_element(By.LINK_TEXT, "Consolidated")
        tab.click()
        time.sleep(1.5)
    except:
        pass

# METRICS
def get_top_dashboard_metrics():
    metrics = {"Promoter Holding": 0.0, "Pledged %": 0.0, "WC Days": 0.0}
    try:
        ratio_items = driver.find_elements(By.CSS_SELECTOR, "ul#top-ratios li")
        for li in ratio_items:
            label = li.find_element(By.CSS_SELECTOR, "span.name").text.lower()
            value = float(li.find_element(By.CSS_SELECTOR, "span.number").text.replace(",", "") or 0)

            if "promoter" in label:
                metrics["Promoter Holding"] = value
            elif "pledged" in label:
                metrics["Pledged %"] = value
            elif "wc days" in label:
                metrics["WC Days"] = value
    except:
        pass
    return metrics

# OPERATING PROFIT
def extract_operating_profit():
    try:
        table = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//section[@id='quarters']//table")))
    except:
        return {}

    headers = [th.text.strip() for th in table.find_elements(By.TAG_NAME, "th")]
    rows = table.find_elements(By.TAG_NAME, "tr")

    for r in rows:
        if "operating profit" in r.text.lower():
            values = [safe_text(c) for c in r.find_elements(By.TAG_NAME, "td")]
            return {h: v for h, v in zip(headers[1:], values[1:])}

    return {}

# COMPANY LIST
def collect_all_company_data():
    driver.get(SCREEN_URL)
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[href*='/company/']")))

    company_list = []
    seen = set()

    while True:
        rows = driver.find_elements(By.CSS_SELECTOR, "a[href*='/company/']")
        for a in rows:
            href = a.get_attribute("href")
            name = a.text.strip()
            if href not in seen:
                seen.add(href)
                company_list.append({"short_name": name, "url": href})

        try:
            next_btn = driver.find_element(By.XPATH, "//a[contains(text(), 'Next')]")
            if "disabled" in next_btn.get_attribute("class"):
                break
            next_btn.click()
            time.sleep(1.5)
        except:
            break

    return company_list

# SCRAPER
def scrape_company(info):
    driver.get(info['url'])

    select_consolidated_if_exists()

    full_name = get_company_full_name()
    industry, group = get_industry_info()   # ✅ NEW
    op_data = extract_operating_profit()
    metrics = get_top_dashboard_metrics()

    return {
        "Name": info['short_name'],
        "Company": full_name,
        "Industry": industry,
        "Industry Group": group,
        **op_data,
        **metrics
    }

# SAVE
def save_to_excel(data, sheet_name):
    df = pd.DataFrame(data)

    if os.path.exists(RAW_FILE_PATH):
        with pd.ExcelWriter(RAW_FILE_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        df.to_excel(RAW_FILE_PATH, sheet_name=sheet_name, index=False)

# MAIN
try:
    login()

    companies = collect_all_company_data()
    results = []

    for i, c in enumerate(companies, 1):
        row = scrape_company(c)
        results.append(row)

        print(f"✅ [{i}/{len(companies)}] {c['short_name']} | {row['Industry']}")

        time.sleep(PAGE_DELAY + random.uniform(0.5, 1.0))

    sheet = datetime.now().strftime("%b %y")
    save_to_excel(results, sheet)

    print("\n🎉 DONE — Industry data included successfully")

except Exception as e:
    print("❌ ERROR:", e)
    traceback.print_exc()

finally:
    driver.quit()
