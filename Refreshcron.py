import os
import base64
import json
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import cloudscraper
import time
import random
from datetime import datetime, timedelta
import pytz  # ✅ Added for proper IST handling

print("🚀 Script started...")

# ==== IST HELPER ====
def get_ist_now():
    ist = pytz.timezone("Asia/Kolkata")
    return datetime.now(ist)

# ==== DECODE CREDENTIALS FROM BASE64 SECRET ====
if "CREDENTIALS_JSON" not in os.environ:
    print("❌ ERROR: CREDENTIALS_JSON secret not found.")
    exit(1)

try:
    creds_b64 = os.environ["CREDENTIALS_JSON"]
    creds_json = base64.b64decode(creds_b64).decode("utf-8")
    with open("credentials.json", "w") as f:
        f.write(creds_json)
    print("✅ credentials.json file created successfully")
except Exception as e:
    print("❌ ERROR decoding credentials:", e)
    exit(1)

# ==== GOOGLE SHEETS SETUP ====
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

try:
    CREDS = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", SCOPE)
    CLIENT = gspread.authorize(CREDS)
    print("✅ Google Sheets authorized successfully")
except Exception as e:
    print("❌ ERROR authorizing Google Sheets:", e)
    exit(1)

try:
    SHEET = CLIENT.open("Refreshcron").sheet1
    print("✅ Sheet opened successfully")
except Exception as e:
    print("❌ ERROR opening sheet:", e)
    exit(1)

# ==== MARKET HOURS CHECK (IST) ====
now_ist = get_ist_now()
if not (datetime.strptime("09:12", "%H:%M").time() <= now_ist.time() <= datetime.strptime("15:45", "%H:%M").time()):
    print(f"⏳ Outside market hours ({now_ist.strftime('%H:%M:%S')} IST). Skipping...")

    try:
        SHEET.insert_row(
            [now_ist.strftime("%Y-%m-%d %H:%M:%S"), "Skipped (Outside Trading Hours)"],
            index=2,
            value_input_option="USER_ENTERED"
        )
        SHEET.format('A2:B2', {
            "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}  # Light gray
        })
        print("✅ Logged skip to Google Sheet")
    except Exception as e:
        print(f"⚠️ Could not log skip to Google Sheet: {e}")

    exit()

# ==== NSE URL ====
NSE_URL = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"

def fetch_nse_oi():
    scraper = cloudscraper.create_scraper(browser={"browser": "chrome", "platform": "windows", "mobile": False})
    scraper.get("https://www.nseindia.com", timeout=10)
    time.sleep(random.uniform(2, 4))
    resp = scraper.get(NSE_URL, timeout=10)
    return resp.json()

def format_num(x):
    return x if isinstance(x, (int, float)) else ""

def append_atm_to_sheet(data):
    underlying_value = data["records"]["underlyingValue"]

    expiry_date = data["records"]["expiryDates"][0] if data["records"]["expiryDates"] else "N/A"
    filtered_data = [d for d in data["records"]["data"] if d.get("expiryDate") == expiry_date]
    strikes = sorted([d["strikePrice"] for d in filtered_data if "CE" in d])

    atm_strike = min(strikes, key=lambda x: abs(x - underlying_value))
    atm_data = next((d for d in filtered_data if d.get("strikePrice") == atm_strike), {})

    ce_data = atm_data.get("CE", {})
    pe_data = atm_data.get("PE", {})

    ce_bid = ce_data.get("bidprice", "")
    ce_ask = ce_data.get("askPrice", "")
    pe_bid = pe_data.get("bidprice", "")
    pe_ask = pe_data.get("askPrice", "")

    ce_spread = ce_ask - ce_bid if isinstance(ce_ask, (int, float)) and isinstance(ce_bid, (int, float)) else ""
    pe_spread = pe_ask - pe_bid if isinstance(pe_ask, (int, float)) and isinstance(pe_bid, (int, float)) else ""

    row = [
        get_ist_now().strftime("%Y-%m-%d %H:%M:%S"),  # ✅ IST timestamp
        expiry_date,
        atm_strike,
        ce_data.get("lastPrice", ""),
        ce_data.get("impliedVolatility", ""),
        ce_data.get("totalTradedVolume", ""),
        ce_data.get("openInterest", ""),
        ce_data.get("changeinOpenInterest", ""),
        format_num(ce_bid),
        format_num(ce_ask),
        format_num(ce_spread),
        pe_data.get("lastPrice", ""),
        pe_data.get("impliedVolatility", ""),
        pe_data.get("totalTradedVolume", ""),
        pe_data.get("openInterest", ""),
        pe_data.get("changeinOpenInterest", ""),
        format_num(pe_bid),
        format_num(pe_ask),
        format_num(pe_spread)
    ]

    # Insert at top (row 2)
    SHEET.insert_row(row, index=2, value_input_option="USER_ENTERED")

    # Highlight row 2 with light yellow background
    SHEET.format('A2:S2', {
        "backgroundColor": {"red": 1, "green": 1, "blue": 0.8}  # Light yellow
    })

    print(f"✅ Logged ATM {atm_strike} at {row[0]} IST | Expiry: {expiry_date}")


if __name__ == "__main__":
    # Always set header without disturbing old data
    SHEET.update('A1:S1', [[
        "Timestamp", "Expiry Date", "Strike Price",
        "CE LTP", "CE IV", "CE Vol", "CE OI", "CE ΔOI", "CE Bid", "CE Ask", "CE Spread",
        "PE LTP", "PE IV", "PE Vol", "PE OI", "PE ΔOI", "PE Bid", "PE Ask", "PE Spread"
    ]])

    try:
        nse_data = fetch_nse_oi()
        append_atm_to_sheet(nse_data)
    except Exception as e:
        print(f"❌ Error fetching/appending data: {e}")
