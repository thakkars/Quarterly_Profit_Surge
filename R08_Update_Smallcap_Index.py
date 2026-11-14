"""
R08_Update_Smallcap_Index.py
--------------------------------------------------------------------------------
Fetches Nifty Smallcap 100 (NSE:CNXSMALLCAP) daily OHLC historical data,
calculates 10-day EMA and 20-day EMA, and evaluates the market entry rules:
  Rule 1: Current market price (CMP/Close) > 10 EMA AND CMP > 20 EMA
  Rule 2: 10 EMA > 20 EMA
  Signal: Entry Allowed ONLY if both rules are satisfied.

Outputs to smallcap100.json for the Equity Research Terminal.
--------------------------------------------------------------------------------
"""

import os
import sys
import json
import time
import datetime
import urllib.request
import pandas as pd

FOLDER_PATH = os.path.dirname(os.path.abspath(__file__))
OUT_JSON_PATH = os.path.join(FOLDER_PATH, "smallcap100.json")
SR_EXCEL_PATH = os.path.join(FOLDER_PATH, "input", "Support and resistance.xlsx")
SR_JSON_PATH = os.path.join(FOLDER_PATH, "support_resistance.json")

def load_support_resistance_records():
    """Loads support & resistance levels from Excel or JSON."""
    records = []
    if os.path.exists(SR_EXCEL_PATH):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(SR_EXCEL_PATH)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))[1:]
            for r in rows:
                lvl = r[1]
                if lvl is not None:
                    records.append({
                        "level": int(lvl) if float(lvl).is_integer() else round(float(lvl), 2),
                        "monthly_level": int(r[0]) if r[0] is not None and float(r[0]).is_integer() else (round(float(r[0]), 2) if r[0] is not None else None),
                        "times_hit": int(r[2]) if r[2] is not None else 0,
                        "distance": int(r[3]) if r[3] is not None else 0,
                        "timeframe": str(r[4]).strip() if r[4] else "Weekly",
                        "weight": int(r[5]) if r[5] is not None else 1,
                        "strength": int(r[6]) if r[6] is not None else 1
                    })
            if records:
                # Also save / keep support_resistance.json updated
                with open(SR_JSON_PATH, "w", encoding="utf-8") as f:
                    json.dump(records, f, indent=2)
        except Exception as e:
            print(f"[WARN] Failed to read {SR_EXCEL_PATH}: {e}")

    if not records and os.path.exists(SR_JSON_PATH):
        try:
            with open(SR_JSON_PATH, "r", encoding="utf-8") as f:
                records = json.load(f)
        except Exception as e:
            print(f"[WARN] Failed to read {SR_JSON_PATH}: {e}")
    return records

def compute_nearest_sr(cmp, sr_records):
    """Computes nearest 2 weekly and monthly support & resistance levels for a given CMP."""
    if not sr_records or not cmp:
        return {}

    weekly_lvls = [r["level"] for r in sr_records if r.get("timeframe") == "Weekly"]
    monthly_lvls = [r["level"] for r in sr_records if r.get("timeframe") == "Monthly"]

    # Resistances (> cmp, ascending)
    w_res = sorted([x for x in weekly_lvls if x > cmp])[:2]
    # Supports (< cmp, descending)
    w_sup = sorted([x for x in weekly_lvls if x < cmp], reverse=True)[:2]

    m_res = sorted([x for x in monthly_lvls if x > cmp])[:2]
    m_sup = sorted([x for x in monthly_lvls if x < cmp], reverse=True)[:2]

    def fmt_item(lvl):
        diff = round(lvl - cmp, 2)
        pct = round((diff / cmp) * 100, 2)
        return {"level": lvl, "diff_pts": diff, "diff_pct": pct}

    return {
        "weekly_resistance": [fmt_item(x) for x in w_res],
        "weekly_support": [fmt_item(x) for x in w_sup],
        "monthly_resistance": [fmt_item(x) for x in m_res],
        "monthly_support": [fmt_item(x) for x in m_sup],
    }

def fetch_tradingview_quote():
    """Fetches real-time / latest quote and EMAs from TradingView Scanner API."""
    url = "https://scanner.tradingview.com/india/scan"
    headers = {"User-Agent": "Mozilla/5.0"}
    payload = {
        "symbols": {"tickers": ["NSE:CNXSMALLCAP"]},
        "columns": [
            "open", "high", "low", "close", "change",
            "volume", "EMA10", "EMA20"
        ]
    }
    try:
        req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            if data and data.get("data"):
                d = data["data"][0]["d"]
                return {
                    "open": round(float(d[0]), 2) if d[0] is not None else None,
                    "high": round(float(d[1]), 2) if d[1] is not None else None,
                    "low": round(float(d[2]), 2) if d[2] is not None else None,
                    "close": round(float(d[3]), 2) if d[3] is not None else None,
                    "change_pct": round(float(d[4]), 2) if d[4] is not None else None,
                    "volume": int(d[5]) if d[5] is not None else 0,
                    "ema10": round(float(d[6]), 2) if d[6] is not None else None,
                    "ema20": round(float(d[7]), 2) if d[7] is not None else None,
                }
    except Exception as e:
        print(f"[WARN] TradingView quote fetch failed: {e}")
    return None

def fetch_nse_historical_bars(days_back=90):
    """Fetches daily index archive records from NSE India."""
    today = datetime.date.today()
    records = []
    headers = {"User-Agent": "Mozilla/5.0"}

    # Iterate backwards through dates
    for i in range(days_back):
        d = today - datetime.timedelta(days=i)
        if d.weekday() >= 5:  # Skip weekends
            continue
        dt_str = d.strftime("%d%m%Y")
        url = f"https://nsearchives.nseindia.com/content/indices/ind_close_all_{dt_str}.csv"
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=4) as resp:
                content = resp.read().decode("utf-8", errors="ignore")
                for line in content.splitlines():
                    if "smallcap 100" in line.lower():
                        parts = line.split(",")
                        if len(parts) >= 6:
                            records.append({
                                "date": parts[1].strip(),
                                "open": round(float(parts[2].strip()), 2),
                                "high": round(float(parts[3].strip()), 2),
                                "low": round(float(parts[4].strip()), 2),
                                "close": round(float(parts[5].strip()), 2),
                                "change_pts": round(float(parts[6].strip()), 2) if len(parts) > 6 and parts[6].strip() else 0.0,
                                "change_pct": round(float(parts[7].strip()), 2) if len(parts) > 7 and parts[7].strip() else 0.0,
                                "volume": int(float(parts[8].strip())) if len(parts) > 8 and parts[8].strip() else 0
                            })
                        break
        except Exception:
            continue

    if not records:
        return []

    # Sort chronologically (oldest to newest)
    # Parse DD-MM-YYYY
    def parse_dt(x):
        try:
            return datetime.datetime.strptime(x["date"], "%d-%m-%Y")
        except Exception:
            return datetime.datetime.min

    records.sort(key=parse_dt)
    return records

def calculate_emas_and_rules(records, live_tv=None):
    """Calculates 10 EMA, 20 EMA, and evaluates entry rules for each bar."""
    if not records:
        return {}

    # If live_tv has a fresh bar matching or beyond the latest archive date, append or update
    df = pd.DataFrame(records)
    df["EMA10"] = df["close"].ewm(span=10, adjust=False).mean()
    df["EMA20"] = df["close"].ewm(span=20, adjust=False).mean()

    history = []
    for idx, r in df.iterrows():
        close_val = round(float(r["close"]), 2)
        ema10_val = round(float(r["EMA10"]), 2)
        ema20_val = round(float(r["EMA20"]), 2)

        rule1 = (close_val > ema10_val) and (close_val > ema20_val)
        rule2 = (ema10_val > ema20_val)
        entry_allowed = rule1 and rule2

        history.append({
            "date": r["date"],
            "open": r["open"],
            "high": r["high"],
            "low": r["low"],
            "close": close_val,
            "volume": r["volume"],
            "ema10": ema10_val,
            "ema20": ema20_val,
            "rule1_met": bool(rule1),
            "rule2_met": bool(rule2),
            "entry_allowed": bool(entry_allowed)
        })

    # Latest bar evaluation
    latest = history[-1] if history else {}
    
    # If live TradingView data is available, override latest with live price & indicators
    if live_tv and live_tv.get("close"):
        cmp = live_tv["close"]
        ema10 = live_tv.get("ema10") or latest.get("ema10")
        ema20 = live_tv.get("ema20") or latest.get("ema20")
        rule1 = (cmp > ema10) and (cmp > ema20) if (ema10 and ema20) else False
        rule2 = (ema10 > ema20) if (ema10 and ema20) else False
        entry_allowed = rule1 and rule2

        latest_payload = {
            "index_name": "NIFTY SMALLCAP 100",
            "symbol": "NSE:CNXSMALLCAP",
            "date": latest.get("date", datetime.date.today().strftime("%d-%m-%Y")),
            "cmp": cmp,
            "open": live_tv.get("open") or latest.get("open"),
            "high": live_tv.get("high") or latest.get("high"),
            "low": live_tv.get("low") or latest.get("low"),
            "close": cmp,
            "change_pct": live_tv.get("change_pct") or 0.0,
            "volume": live_tv.get("volume") or latest.get("volume", 0),
            "ema10": ema10,
            "ema20": ema20,
            "rule1_met": bool(rule1),
            "rule2_met": bool(rule2),
            "entry_allowed": bool(entry_allowed),
            "status_text": "ENTRY ALLOWED" if entry_allowed else "NO ENTRY",
            "rule1_desc": f"CMP ({cmp:,.2f}) > 10 EMA ({ema10:,.2f}) and > 20 EMA ({ema20:,.2f})" if (ema10 and ema20) else "N/A",
            "rule2_desc": f"10 EMA ({ema10:,.2f}) > 20 EMA ({ema20:,.2f})" if (ema10 and ema20) else "N/A",
            "updated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    else:
        cmp = latest.get("close", 0)
        ema10 = latest.get("ema10", 0)
        ema20 = latest.get("ema20", 0)
        rule1 = latest.get("rule1_met", False)
        rule2 = latest.get("rule2_met", False)
        entry_allowed = latest.get("entry_allowed", False)
        latest_payload = {
            "index_name": "NIFTY SMALLCAP 100",
            "symbol": "NSE:CNXSMALLCAP",
            "date": latest.get("date", ""),
            "cmp": cmp,
            "open": latest.get("open", 0),
            "high": latest.get("high", 0),
            "low": latest.get("low", 0),
            "close": cmp,
            "change_pct": 0.0,
            "volume": latest.get("volume", 0),
            "ema10": ema10,
            "ema20": ema20,
            "rule1_met": rule1,
            "rule2_met": rule2,
            "entry_allowed": entry_allowed,
            "status_text": "ENTRY ALLOWED" if entry_allowed else "NO ENTRY",
            "rule1_desc": f"CMP ({cmp:,.2f}) > 10 EMA ({ema10:,.2f}) and > 20 EMA ({ema20:,.2f})",
            "rule2_desc": f"10 EMA ({ema10:,.2f}) > 20 EMA ({ema20:,.2f})",
            "updated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

    return {
        "latest": latest_payload,
        "history": list(reversed(history))  # Newest bars first for table view
    }

def main():
    print("=" * 60)
    print("📊 NIFTY SMALLCAP 100 OHLC & 10/20 EMA INDICATOR UPDATER")
    print("=" * 60)

    print("1. Fetching live quote & indicators from TradingView...")
    live_tv = fetch_tradingview_quote()
    if live_tv:
        print(f"   ✅ TradingView Live Quote: Close={live_tv['close']} | EMA10={live_tv['ema10']} | EMA20={live_tv['ema20']}")
    else:
        print("   ℹ️ TradingView quote unavailable, will rely on historical archives.")

    print("\n2. Fetching historical daily bars from NSE archives...")
    records = fetch_nse_historical_bars(days_back=90)
    print(f"   ✅ Fetched {len(records)} trading days from NSE archives.")

    if not records and live_tv:
        # Fallback if NSE archive is unreachable
        records = [{
            "date": datetime.date.today().strftime("%d-%m-%Y"),
            "open": live_tv["open"],
            "high": live_tv["high"],
            "low": live_tv["low"],
            "close": live_tv["close"],
            "volume": live_tv["volume"]
        }]

    print("\n3. Calculating 10-day & 20-day EMAs and evaluating Entry Rules...")
    data = calculate_emas_and_rules(records, live_tv)

    latest = data.get("latest", {})
    print("\n" + "─" * 60)
    print(f"INDEX: {latest.get('index_name')} ({latest.get('date')})")
    print(f"CMP:   {latest.get('cmp'):,.2f}")
    print(f"10 EMA: {latest.get('ema10'):,.2f} | 20 EMA: {latest.get('ema20'):,.2f}")
    print(f"RULE 1: {latest.get('rule1_desc')} -> {'✅ PASSED' if latest.get('rule1_met') else '❌ FAILED'}")
    print(f"RULE 2: {latest.get('rule2_desc')} -> {'✅ PASSED' if latest.get('rule2_met') else '❌ FAILED'}")
    print(f"SIGNAL: {'🟢 ' + latest.get('status_text') if latest.get('entry_allowed') else '🔴 ' + latest.get('status_text')}")
    print("─" * 60)

    print(f"\n4. Saving results to {OUT_JSON_PATH}...")
    with open(OUT_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    print("   ✅ smallcap100.json updated successfully.")
    return 0

if __name__ == "__main__":
    sys.exit(main())

