import os
import base64
import json
import time
import gspread
import cloudscraper
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import pytz  # ✅ for timezone

print("🚀 Script started...")

# ==== IST TIME HELPER ====
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
    SHEET = CLIENT.open("Refreshcron").worksheet("Sheet2")
    print("✅ Sheet2 opened successfully")
except Exception as e:
    print("❌ ERROR opening Sheet2:", e)
    exit(1)

# ==== TIME CHECK (EXIT IF OUTSIDE TRADING HOURS) ====
now = get_ist_now()
if not (datetime.strptime("09:12", "%H:%M").time() <= now.time() <= datetime.strptime("15:45", "%H:%M").time()):
    print("⏰ Outside trading hours, skipping execution...")

    try:
        SHEET.append_row([now.strftime("%Y-%m-%d %H:%M:%S"), "Skipped (Outside Trading Hours)"],
                         value_input_option="USER_ENTERED")
        print("✅ Logged skip to Google Sheet")
    except Exception as e:
        print(f"⚠️ Could not log skip to Google Sheet: {e}")

    exit()

# ==== NSE API ====
NSE_URL = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"

def fetch_nifty_spot():
    scraper = cloudscraper.create_scraper(browser={"browser": "chrome", "platform": "windows", "mobile": False})
    scraper.get("https://www.nseindia.com", timeout=10)
    resp = scraper.get(NSE_URL, timeout=10)
    data = resp.json()
    return data["records"]["underlyingValue"]  # ✅ NIFTY spot

# ==== MAIN LOOP ====
while True:
    try:
        price = fetch_nifty_spot()
        timestamp = get_ist_now().strftime("%Y-%m-%d %H:%M:%S")
        row = [timestamp, price]   # ✅ Column A = timestamp, Column B = price

        SHEET.append_row(row, value_input_option="USER_ENTERED")

        print(f"✅ Logged NIFTY at {timestamp} = {price}")
    except Exception as e:
        print(f"❌ Error: {e}")

    # Wait 3 minutes
    time.sleep(180)
