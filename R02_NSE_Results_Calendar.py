import os
import time
import traceback
import pandas as pd
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

if os.environ.get("GITHUB_ACTIONS", "false").lower() != "true":
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.service import Service

# ═══════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════
IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"

FOLDER_PATH = "/tmp/screener_data" if IS_GITHUB_ACTIONS else \
    r"D:\Investment Research\Python Projects\Operating Profit more than 70% jump"

os.makedirs(FOLDER_PATH, exist_ok=True)

CSV_PATH  = os.path.join(FOLDER_PATH, "nse_results_calendar.csv")
JSON_PATH = os.path.join(FOLDER_PATH, "nse_results_calendar.json")

R2_BUCKET     = os.environ.get("R2_BUCKET_NAME",       "quarterly-profit-surge")
R2_ACCESS_KEY = os.environ.get("R2_ACCESS_KEY_ID",     "")
R2_SECRET_KEY = os.environ.get("R2_SECRET_ACCESS_KEY", "")
R2_ENDPOINT   = os.environ.get("R2_ENDPOINT_URL",      "")

R2_CSV_KEY  = "nse_results_calendar.csv"
R2_JSON_KEY = "nse_results_calendar.json"

NSE_XBRL_URL = "https://www.nseindia.com/companies-listing/corporate-filings-announcements-xbrl"

# ═══════════════════════════════════════════════════════════
# R2
# ═══════════════════════════════════════════════════════════
def get_r2():
    import boto3
    return boto3.client("s3",
        endpoint_url=R2_ENDPOINT,
        aws_access_key_id=R2_ACCESS_KEY,
        aws_secret_access_key=R2_SECRET_KEY,
        region_name="auto")

def r2_download(key, path):
    try:
        get_r2().download_file(R2_BUCKET, key, path)
        print(f"⬇️  R2 download OK: {key}")
        return True
    except Exception as e:
        print(f"⚠️  R2 download failed ({key}): {e}")
        return False

def r2_upload(path, key):
    try:
        if not os.path.exists(path):
            print(f"⚠️  Upload skipped — not found: {path}")
            return False
        extra = {}
        if key.endswith(".json"):
            extra = {"ContentType": "application/json",
                     "CacheControl": "no-cache, no-store, must-revalidate"}
        elif key.endswith(".csv"):
            extra = {"ContentType": "text/csv",
                     "CacheControl": "no-cache, no-store, must-revalidate"}
        get_r2().upload_file(path, R2_BUCKET, key, ExtraArgs=extra or None)
        print(f"⬆️  R2 upload OK: {key}")
        return True
    except Exception as e:
        print(f"⚠️  R2 upload failed ({key}): {e}")
        return False

# ═══════════════════════════════════════════════════════════
# DRIVER
# ═══════════════════════════════════════════════════════════
def init_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
    if IS_GITHUB_ACTIONS:
        return webdriver.Chrome(options=opts)
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=opts)

# ═══════════════════════════════════════════════════════════
# SCRAPE  — Selenium loads the page, then we intercept
# the XHR response via JS execution (most reliable approach)
# ═══════════════════════════════════════════════════════════
def scrape_xbrl_page():
    driver = init_driver()
    wait   = WebDriverWait(driver, 30)
    rows   = []
    scraped_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        print(f"  Loading NSE home page first for cookies...")
        driver.get("https://www.nseindia.com")
        time.sleep(3)

        print(f"  Loading: {NSE_XBRL_URL}")
        driver.get(NSE_XBRL_URL)
        time.sleep(6)   # wait for JS to render table

        print(f"  Page title: {driver.title}")

        # Strategy 1: intercept the JSON data NSE loaded into the page via JS
        # NSE stores the table data in a global JS variable or DataTable
        json_data = driver.execute_script("""
            try {
                // Try common DataTable patterns used by NSE
                var tables = $.fn && $.fn.dataTable ? $.fn.dataTable.tables() : [];
                if (tables.length > 0) {
                    var api = $(tables[0]).DataTable();
                    return JSON.stringify(api.rows().data().toArray());
                }
            } catch(e) {}
            // Try window-level data objects
            for (var key in window) {
                try {
                    if (Array.isArray(window[key]) && window[key].length > 0
                        && window[key][0] && typeof window[key][0] === 'object') {
                        var sample = window[key][0];
                        var keys = Object.keys(sample);
                        if (keys.some(k => k.toLowerCase().includes('symbol') || k.toLowerCase().includes('comp'))) {
                            return JSON.stringify(window[key]);
                        }
                    }
                } catch(e) {}
            }
            return null;
        """)

        if json_data:
            import json as json_lib
            try:
                data = json_lib.loads(json_data)
                print(f"  ✅ Got {len(data)} records via JS intercept")
                for item in data:
                    if not isinstance(item, dict):
                        continue
                    # Normalise keys
                    item_lower = {k.lower(): v for k, v in item.items()}
                    bcast_raw  = (
                        item_lower.get("broadcastdatetime") or
                        item_lower.get("an_dt") or
                        item_lower.get("date") or ""
                    )
                    bcast_date, bcast_time = "", ""
                    if bcast_raw:
                        parts = str(bcast_raw).strip().split(" ")
                        bcast_date = parts[0] if parts else ""
                        bcast_time = parts[1] if len(parts) > 1 else ""

                    rows.append({
                        "Symbol":         item_lower.get("symbol") or item_lower.get("nsecode") or "",
                        "Company Name":   item_lower.get("comp") or item_lower.get("company") or item_lower.get("name") or "",
                        "Series":         item_lower.get("series") or "",
                        "Broadcast Date": bcast_date,
                        "Broadcast Time": bcast_time,
                        "Subject":        item_lower.get("subject") or item_lower.get("an_subject") or "",
                        "Description":    item_lower.get("details") or item_lower.get("an_det") or item_lower.get("description") or "",
                        "XBRL Link":      item_lower.get("attchmntfile") or item_lower.get("attachment_file") or "",
                        "Scraped At":     scraped_at,
                    })
                return rows
            except Exception as e:
                print(f"  ⚠️  JS intercept parse error: {e}")

        # Strategy 2: Read the visible HTML table directly
        print("  Falling back to HTML table scraping...")
        tables = driver.find_elements(By.TAG_NAME, "table")
        print(f"  Found {len(tables)} table(s) on page")

        if not tables:
            print("  ❌ No tables found on page.")
            print(f"  Page source snippet:\n{driver.page_source[:2000]}")
            return rows

        best_table = max(tables, key=lambda t: len(t.find_elements(By.TAG_NAME, "tr")))
        all_trs    = best_table.find_elements(By.TAG_NAME, "tr")
        print(f"  Using table with {len(all_trs)} rows")

        for tr in all_trs:
            tds = tr.find_elements(By.TAG_NAME, "td")
            if not tds:
                continue
            cells = [td.text.strip() for td in tds]

            xbrl_link = ""
            for td in tds:
                try:
                    a = td.find_element(By.TAG_NAME, "a")
                    href = a.get_attribute("href") or ""
                    if href:
                        xbrl_link = href
                        break
                except:
                    pass

            # Typical XBRL page columns: Symbol | Company | Series | DateTime | Subject | Details
            bcast_date, bcast_time = "", ""
            if len(cells) > 3:
                dt_parts   = cells[3].split(" ")
                bcast_date = dt_parts[0] if dt_parts else ""
                bcast_time = dt_parts[1] if len(dt_parts) > 1 else ""

            row = {
                "Symbol":         cells[0] if len(cells) > 0 else "",
                "Company Name":   cells[1] if len(cells) > 1 else "",
                "Series":         cells[2] if len(cells) > 2 else "",
                "Broadcast Date": bcast_date,
                "Broadcast Time": bcast_time,
                "Subject":        cells[4] if len(cells) > 4 else "",
                "Description":    cells[5] if len(cells) > 5 else "",
                "XBRL Link":      xbrl_link,
                "Scraped At":     scraped_at,
            }

            if not row["Symbol"] or row["Symbol"].lower() in ("symbol", "scrip", ""):
                continue

            rows.append(row)

        print(f"  Scraped {len(rows)} rows from HTML table")

    except Exception as e:
        print(f"  ❌ Scrape error: {e}")
        traceback.print_exc()
    finally:
        try:
            driver.quit()
        except:
            pass

    return rows

# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════
def main():
    print("\n" + "="*60)
    print("  NSE Results Calendar Scraper")
    print(f"  Run time (IST): {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)

    if IS_GITHUB_ACTIONS:
        r2_download(R2_CSV_KEY, CSV_PATH)

    print("\n[1/3] Scraping NSE XBRL page...")
    new_rows = scrape_xbrl_page()

    if not new_rows:
        print("❌ No rows scraped. Exiting.")
        return

    df_new = pd.DataFrame(new_rows)

    # Merge with history, deduplicate
    if os.path.exists(CSV_PATH):
        try:
            df_existing = pd.read_csv(CSV_PATH, dtype=str).fillna("")
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_combined.drop_duplicates(
                subset=["Symbol", "Broadcast Date", "Broadcast Time", "Subject"],
                keep="last", inplace=True)
            try:
                df_combined["_dt"] = pd.to_datetime(
                    df_combined["Broadcast Date"] + " " + df_combined["Broadcast Time"],
                    dayfirst=True, errors="coerce")
                df_combined.sort_values("_dt", ascending=False, inplace=True)
                df_combined.drop(columns=["_dt"], inplace=True)
            except:
                pass
            df_final = df_combined
            print(f"  Merged: {len(df_existing)} old + {len(df_new)} new = {len(df_final)} total")
        except Exception as e:
            print(f"  ⚠️  Merge failed ({e}), using fresh data.")
            df_final = df_new
    else:
        df_final = df_new

    print(f"\n[2/3] Saving {len(df_final)} rows...")
    df_final.to_csv(CSV_PATH, index=False)
    print(f"✅ CSV: {CSV_PATH}")
    df_final.fillna("").to_json(JSON_PATH, orient="records")
    print(f"✅ JSON: {JSON_PATH}")

    print("\n[3/3] Uploading to Cloudflare R2...")
    if IS_GITHUB_ACTIONS:
        r2_upload(CSV_PATH,  R2_CSV_KEY)
        r2_upload(JSON_PATH, R2_JSON_KEY)

    print("\n✅ Done!")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
        traceback.print_exc()
