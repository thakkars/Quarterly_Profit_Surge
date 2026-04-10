# =========================================================
# FINAL STABLE VERSION (NO TIMEOUT + NO DUPLICATES)
# =========================================================

import time
import random
import traceback
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import os

# =========================================================
# ENV
# =========================================================
IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"

USERNAME = os.environ.get("SCREENER_USERNAME", "")
PASSWORD = os.environ.get("SCREENER_PASSWORD", "")

SCREEN_URL = "https://www.screener.in/screens/742984/top-growers/"

if IS_GITHUB_ACTIONS:
    FOLDER_PATH = "/tmp/screener_data"
else:
    FOLDER_PATH = r"D:\Investment Research\Python Projects\Operating Profit more than 70% jump"

MASTER_FILE_PATH = os.path.join(FOLDER_PATH, "Master File.xlsx")

os.makedirs(FOLDER_PATH, exist_ok=True)

# =========================================================
# DRIVER (STABLE)
# =========================================================
def init_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-blink-features=AutomationControlled")

    if IS_GITHUB_ACTIONS:
        return webdriver.Chrome(options=opts)
    else:
        from webdriver_manager.chrome import ChromeDriverManager
        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=opts
        )

driver = init_driver()
driver.set_page_load_timeout(60)
wait = WebDriverWait(driver, 25)

# =========================================================
# SAFE GET (RETRY)
# =========================================================
def safe_get(url, retries=2):
    for i in range(retries):
        try:
            driver.get(url)
            return True
        except Exception:
            print(f"Retry {i+1} for {url}")
            time.sleep(3)
    return False

# =========================================================
# LOGIN
# =========================================================
def login():
    safe_get("https://www.screener.in/login/")
    try:
        wait.until(EC.presence_of_element_located((By.NAME, "username")))
        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.NAME, "password").submit()
        time.sleep(3)
    except:
        pass

# =========================================================
# SCRAPER
# =========================================================
def collect_all_company_data():
    safe_get(SCREEN_URL)
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[href*='/company/']")))

    companies = []
    seen = set()

    while True:
        rows = driver.find_elements(By.CSS_SELECTOR, "a[href*='/company/']")
        for a in rows:
            url = a.get_attribute("href")
            name = a.text.strip()
            if url and url not in seen:
                seen.add(url)
                companies.append({"name": name, "url": url})

        try:
            nxt = driver.find_element(By.XPATH, "//a[contains(text(),'Next')]")
            if "disabled" in nxt.get_attribute("class"):
                break
            driver.execute_script("arguments[0].click();", nxt)
            time.sleep(2)
        except:
            break

    print(f"Collected {len(companies)} companies")
    return companies

# =========================================================
# SCRAPE COMPANY (SAFE)
# =========================================================
def scrape_company(c):
    if not safe_get(c["url"]):
        return None

    try:
        name = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h1.margin-0"))
        ).text.strip()
    except:
        name = "Unknown"

    return {
        "Name": c["name"],
        "Company (Page)": name,
        "Company URL": c["url"]
    }

# =========================================================
# MAIN
# =========================================================
try:
    login()

    all_companies = collect_all_company_data()

    latest_quarter = datetime.now().strftime("%b %Y")

    # LOAD EXISTING MASTER
    existing_urls = set()
    if os.path.exists(MASTER_FILE_PATH):
        try:
            df_existing = pd.read_excel(MASTER_FILE_PATH, sheet_name=latest_quarter)
            if "Company URL" in df_existing.columns:
                existing_urls = set(df_existing["Company URL"].astype(str))
        except:
            pass

    print(f"Existing companies: {len(existing_urls)}")

    # FILTER NEW
    new_companies = [c for c in all_companies if c["url"] not in existing_urls]

    print(f"New companies to scrape: {len(new_companies)}")

    if not new_companies:
        print("Nothing new. Exiting.")
        exit()

    data = []
    for c in new_companies:
        row = scrape_company(c)
        if row:
            data.append(row)

        # anti-block delay
        time.sleep(2 + random.uniform(0.5, 1.5))

    df_new = pd.DataFrame(data)

    # FINAL SAFETY DEDUP
    df_new = df_new.drop_duplicates(subset=["Company URL"])

    # SAVE MASTER
    with pd.ExcelWriter(
        MASTER_FILE_PATH,
        engine="openpyxl",
        mode="a" if os.path.exists(MASTER_FILE_PATH) else "w",
        if_sheet_exists="overlay"
    ) as writer:

        try:
            ws = writer.book[latest_quarter]
            start_row = ws.max_row
        except:
            start_row = 0

        df_new.to_excel(
            writer,
            sheet_name=latest_quarter,
            index=False,
            header=(start_row == 0),
            startrow=start_row
        )

    print(f"Added {len(df_new)} new companies")

except Exception as e:
    print("❌ ERROR OCCURRED:")
    traceback.print_exc()

finally:
    driver.quit()