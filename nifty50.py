import os
import base64
import json
import gspread
import cloudscraper
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz

print("🚀 Script started...")

# ==== IST TIME HELPER ====
def get_ist_now():
    ist = pytz.timezone("Asia/Kolkata")
    now = datetime.now(ist)
    print("🚀 Script triggered at:", now.strftime("%Y-%m-%d %H:%M:%S"))
    return now



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
    SHEET = CLIENT.open("Refreshcron").worksheet("Sheet2")   # Always use second sheet
    print("✅ Google Sheet opened successfully")
except Exception as e:
    print("❌ ERROR opening sheet:", e)
    exit(1)

# ==== AUTO-CREATE HEADERS ====
HEADERS = ["Timestamp", "Price", "Status"]

try:
    first_row = SHEET.row_values(1)
    if first_row != HEADERS:
        SHEET.delete_rows(1)  # remove old row 1 if exists
        SHEET.insert_row(HEADERS, 1)
        print("✅ Headers set successfully:", HEADERS)
except Exception as e:
    print(f"⚠️ Could not set headers: {e}")

# ==== NSE API ====
NSE_URL = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"

def fetch_nifty_spot():
    scraper = cloudscraper.create_scraper(
        browser={"browser": "chrome", "platform": "windows", "mobile": False}
    )
    scraper.get("https://www.nseindia.com", timeout=10)
    resp = scraper.get(NSE_URL, timeout=10)
    print("🔍 NSE Response Status:", resp.status_code)
    data = resp.json()
    return data["records"]["underlyingValue"]

# ==== MAIN EXECUTION ====
now = get_ist_now()
timestamp = now.strftime("%Y-%m-%d %H:%M:%S")

try:
    if datetime.strptime("09:12", "%H:%M").time() <= now.time() <= datetime.strptime("15:45", "%H:%M").time():
        # Inside trading hours → fetch price
        price = fetch_nifty_spot()
        row = [timestamp, price, "OK"]
        print("📝 Writing to Google Sheet:", row)
        SHEET.append_row(row, value_input_option="USER_ENTERED")
        print(f"✅ Logged NIFTY at {timestamp} = {price}")
    else:
        # Outside hours → log skip
        row = [timestamp, "-", "Skipped (Outside Trading Hours)"]
        print("📝 Writing to Google Sheet:", row)
        SHEET.append_row(row, value_input_option="USER_ENTERED")
        print("⏰ Outside trading hours, logged skip")

except Exception as e:
    # Always log error
    row = [timestamp, "-", f"Error: {e}"]
    print("📝 Writing to Google Sheet:", row)
    try:
        SHEET.append_row(row, value_input_option="USER_ENTERED")
    except Exception as g:
        print(f"❌ Failed to log error to sheet: {g}")
