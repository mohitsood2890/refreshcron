import cloudscraper
import time
import random
import gspread
import json
import os
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# ==== WRITE CREDENTIALS.JSON FROM GITHUB SECRET ====
if "CREDENTIALS_JSON" in os.environ:
    with open("credentials.json", "w") as f:
        f.write(os.environ["CREDENTIALS_JSON"])

# ==== MARKET HOURS CHECK ====
now = datetime.now()
start_time = now.replace(hour=9, minute=15, second=0, microsecond=0)
end_time = now.replace(hour=15, minute=30, second=0, microsecond=0)

if not (start_time <= now <= end_time):
    print("⏳ Outside market hours. Exiting.")
    exit()

# ==== GOOGLE SHEETS SETUP ====
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
CREDS = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", SCOPE)
CLIENT = gspread.authorize(CREDS)
SHEET = CLIENT.open("Refreshcron").sheet1

# ==== NSE URL ====
NSE_URL = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"

def fetch_nse_oi():
    scraper = cloudscraper.create_scraper(browser={"browser": "chrome", "platform": "windows", "mobile": False})
    scraper.get("https://www.nseindia.com", timeout=10)
    time.sleep(random.uniform(2, 4))
    resp = scraper.get(NSE_URL, timeout=10)
    return resp.json()

def append_atm_to_sheet(data):
    underlying_value = data["records"]["underlyingValue"]
    expiry_date = data["records"]["expiryDates"][0] if data["records"]["expiryDates"] else "N/A"

    filtered_data = [d for d in data["records"]["data"] if d.get("expiryDate") == expiry_date]
    strikes = sorted([d["strikePrice"] for d in filtered_data if "CE" in d])

    atm_strike = min(strikes, key=lambda x: abs(x - underlying_value))
    atm_data = next((d for d in filtered_data if d.get("strikePrice") == atm_strike), {})

    ce_data = atm_data.get("CE", {})
    pe_data = atm_data.get("PE", {})

    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        expiry_date,
        atm_strike,
        ce_data.get("lastPrice", ""),
        ce_data.get("impliedVolatility", ""),
        ce_data.get("totalTradedVolume", ""),
        ce_data.get("openInterest", ""),
        ce_data.get("changeinOpenInterest", ""),
        pe_data.get("lastPrice", ""),
        pe_data.get("impliedVolatility", ""),
        pe_data.get("totalTradedVolume", ""),
        pe_data.get("openInterest", ""),
        pe_data.get("changeinOpenInterest", "")
    ]

    SHEET.insert_row(row, index=2, value_input_option="USER_ENTERED")
    SHEET.format('A2:M2', {
        "backgroundColor": {"red": 1, "green": 1, "blue": 0.8}
    })

    print(f"✅ Logged ATM {atm_strike} at {row[0]} | Expiry: {expiry_date}")

if __name__ == "__main__":
    SHEET.update('A1:M1', [[
        "Timestamp", "Expiry Date", "Strike Price",
        "CE LTP", "CE IV", "CE Vol", "CE OI", "CE ΔOI",
        "PE LTP", "PE IV", "PE Vol", "PE OI", "PE ΔOI"
    ]])

    try:
        nse_data = fetch_nse_oi()
        append_atm_to_sheet(nse_data)
    except Exception as e:
        print(f"❌ Error: {e}")
