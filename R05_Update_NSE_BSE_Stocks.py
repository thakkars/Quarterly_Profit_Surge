import time
import os
import glob
import shutil
import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# ═══════════════════════════════════════════════════════════
# ENVIRONMENT & CONFIG
# ═══════════════════════════════════════════════════════════
IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"

USERNAME   = os.environ.get("SCREENER_USERNAME", "")
PASSWORD   = os.environ.get("SCREENER_PASSWORD", "")
SCREEN_URL = "https://www.screener.in/screens/3650931/all-nse-bse-stocks/"

LOCAL_REPO_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR   = os.path.join(LOCAL_REPO_DIR, "downloads") if IS_GITHUB_ACTIONS else \
                 r"D:\Investment Research\Python Projects\Operating Profit more than 70% jump"

FRESH_CSV      = os.path.join(DOWNLOAD_DIR, "screener_fresh.csv")
MASTER_EXCEL_FILENAME = "All NSE BSE Stocks.xlsx"
LOCAL_MASTER   = os.path.join(LOCAL_REPO_DIR, MASTER_EXCEL_FILENAME)

os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# ═══════════════════════════════════════════════════════════
# DRIVER INIT
# ═══════════════════════════════════════════════════════════
def init_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1600,900")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_experimental_option("prefs", {
        "download.default_directory":   DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade":   True,
        "safebrowsing.enabled":         True,
    })
    if not IS_GITHUB_ACTIONS:
        from webdriver_manager.chrome import ChromeDriverManager
        from selenium.webdriver.chrome.service import Service
        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=opts
        )
    return webdriver.Chrome(options=opts)

# ═══════════════════════════════════════════════════════════
# LOGIN
# ═══════════════════════════════════════════════════════════
def login(driver, wait):
    print("🔐 Logging in to Screener.in...")
    driver.get("https://www.screener.in/login/")
    try:
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.NAME, "password").submit()
        time.sleep(3)
        print("✅ Login successful")
    except Exception as e:
        print(f"❌ Login failed: {e}")

# ═══════════════════════════════════════════════════════════
# WAIT FOR DOWNLOAD
# ═══════════════════════════════════════════════════════════
def wait_for_download(download_dir, final_path, timeout=30):
    print("⏳ Waiting for download...")
    deadline = time.time() + timeout
    while time.time() < deadline:
        if glob.glob(os.path.join(download_dir, "*.crdownload")):
            time.sleep(1)
            continue
        files = glob.glob(os.path.join(download_dir, "*.csv")) + \
                glob.glob(os.path.join(download_dir, "*.xlsx"))
        files = [f for f in files if os.path.basename(f) != MASTER_EXCEL_FILENAME]
        if files:
            latest = max(files, key=os.path.getmtime)
            if time.time() - os.path.getmtime(latest) < 60:
                if os.path.abspath(latest) != os.path.abspath(final_path):
                    shutil.move(latest, final_path)
                print(f"✅ Fresh CSV saved: {final_path}")
                return final_path
        time.sleep(1)
    print("⚠️  Download timed out")
    return None

# ═══════════════════════════════════════════════════════════
# CLICK EXPORT
# ═══════════════════════════════════════════════════════════
def click_export(driver, wait):
    print("🌐 Navigating to Screener screen...")
    driver.get(SCREEN_URL)
    time.sleep(2)

    btn = None
    for by, selector in [
        (By.XPATH,        "/html/body/main/div[2]/div[4]/div[2]/form/button"),
        (By.CSS_SELECTOR, "button[aria-label='Download Results']"),
        (By.CSS_SELECTOR, "button.button-small[data-tooltip='Download Results']"),
    ]:
        try:
            btn = wait.until(EC.presence_of_element_located((by, selector)))
            break
        except:
            continue

    if btn is None:
        print("❌ Export button not found")
        return False

    # Strategy 1: Submit parent form
    try:
        form = btn.find_element(By.XPATH, "./ancestor::form[1]")
        driver.execute_script("arguments[0].submit();", form)
        print("✅ Export triggered via form.submit()")
        return True
    except:
        pass

    # Strategy 2: JS click
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        time.sleep(0.8)
        driver.execute_script("arguments[0].click();", btn)
        print("✅ Export triggered via JS click")
        return True
    except:
        pass

    # Strategy 3: ActionChains
    try:
        driver.execute_script("document.activeElement.blur();")
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        time.sleep(0.5)
        ActionChains(driver).move_to_element(btn).pause(0.3).click().perform()
        print("✅ Export triggered via ActionChains")
        return True
    except:
        pass

    print("❌ All export strategies failed")
    return False

# ═══════════════════════════════════════════════════════════
# SESSION COOKIES: SELENIUM → REQUESTS
# ═══════════════════════════════════════════════════════════
def get_requests_session(driver):
    session = requests.Session()
    session.headers.update({
        "User-Agent":       "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Referer":          "https://www.screener.in/",
        "X-Requested-With": "XMLHttpRequest",
    })
    for cookie in driver.get_cookies():
        session.cookies.set(cookie["name"], cookie["value"])
    print("✅ Requests session ready")
    return session

# ═══════════════════════════════════════════════════════════
# FULL NAME RESOLVER VIA SCREENER SEARCH API
# ═══════════════════════════════════════════════════════════
def fetch_full_name(session, short_name, retries=2):
    url = "https://www.screener.in/api/company/search/"
    for _ in range(retries + 1):
        try:
            resp = session.get(url, params={"q": short_name, "v": 3}, timeout=5)
            if resp.status_code == 200:
                data = resp.json()
                if data and isinstance(data, list):
                    return data[0].get("name", short_name).strip()
            elif resp.status_code == 429:
                print(f"  ⚠️  Rate limited — backing off 2s...")
                time.sleep(2)
        except:
            time.sleep(0.5)
    return short_name

# ═══════════════════════════════════════════════════════════
# COMPARE: FIND NEW STOCKS
# ═══════════════════════════════════════════════════════════
def find_new_stocks(df_fresh, df_r2):
    existing_names = set(
        df_r2["Name"].dropna().astype(str).str.strip().str.lower()
    )
    mask   = ~df_fresh["Name"].astype(str).str.strip().str.lower().isin(existing_names)
    df_new = df_fresh[mask].copy()
    return df_new

# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════
def main():
    # ── Step 1: Load master file from repository ───────────────────
    print("\n📥 Step 1: Loading master file from repository...")
    if not os.path.exists(LOCAL_MASTER):
        print(f"❌ Master file not found at {LOCAL_MASTER} — aborting")
        return

    df_master = pd.read_excel(LOCAL_MASTER)
    print(f"   → Master has {len(df_master)} stocks")

    # ── Step 2: Login + Export fresh CSV ────────────────────────────
    print("\n🌐 Step 2: Scraping fresh data from Screener...")
    driver = init_driver()
    wait   = WebDriverWait(driver, 15)

    try:
        login(driver, wait)
        exported = click_export(driver, wait)

        if not exported:
            print("❌ Export failed — aborting")
            return

        fresh_path = wait_for_download(DOWNLOAD_DIR, FRESH_CSV, timeout=30)
        if not fresh_path:
            print("❌ Fresh CSV not downloaded — aborting")
            return

        df_fresh = pd.read_csv(fresh_path)
        print(f"   → Fresh Screener data has {len(df_fresh)} stocks")

        session = get_requests_session(driver)

    finally:
        driver.quit()

    # ── Step 3: Compare ───────────────────────────────────────────
    print("\n🔍 Step 3: Comparing repository master vs fresh data...")
    df_new = find_new_stocks(df_fresh, df_master)

    if df_new.empty:
        print("✅ No new stocks found — nothing to update. Done!")
        return

    print(f"   → {len(df_new)} new stock(s) found:")
    for name in df_new["Name"].tolist():
        print(f"      + {name}")

    # ── Step 4: Resolve full names ──────────────────────────────────
    print(f"\n🔎 Step 4: Resolving full company names...")
    full_names = []
    total      = len(df_new)

    for count, (idx, row) in enumerate(df_new.iterrows(), start=1):
        short_name = str(row["Name"]).strip()
        full_name  = fetch_full_name(session, short_name)
        full_names.append(full_name)
        print(f"   [{count}/{total}] {short_name:35s} → {full_name}")
        time.sleep(0.15)

    df_new = df_new.copy()
    df_new["Company"] = full_names

    # ── Step 5: Append & save to master file ───────────────────────
    print(f"\n💾 Step 5: Appending {len(df_new)} new row(s) to master file...")

    # Align columns to match master exactly
    for col in df_master.columns:
        if col not in df_new.columns:
            df_new[col] = ""
    df_new = df_new[df_master.columns]

    df_updated = pd.concat([df_master, df_new], ignore_index=True)
    df_updated.to_excel(LOCAL_MASTER, index=False)
    print(f"   → Master now has {len(df_updated)} stocks")
    print(f"\n✅ Done! {len(df_new)} new stock(s) added to master file.")

if __name__ == "__main__":
    main()
