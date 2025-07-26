import pyotp
import pandas as pd
from SmartApi.smartConnect import SmartConnect
import requests
# import datetime
import  datetime
import yagmail
import os
from dotenv import load_dotenv

# === CONFIGURATION ===
CLIENT_CODE = "S53797011"
API_KEY = "65s3Uq8j"
MPIN = "9568"
TOTP_SECRET = "OGS5PNG2AS2SJLHMDNN6NBCIRM"
SYMBOL = "RELIANCE-EQ"
EXCHANGE = "NSE"
load_dotenv()

SYMBOLS = [
    "RELIANCE", "TCS", "INFY", "HDFCBANK", "ICICIBANK", "KOTAKBANK", "SBIN", "AXISBANK",
     "HDFCLIFE", "BAJAJFINSV", "ITC", "LT", "MARUTI", "BHARTIARTL", "TITAN",
    "ULTRACEMCO", "NESTLEIND", "HINDUNILVR", "BAJAJ-AUTO", "BAJFINANCE", "ASIANPAINT",
    "SUNPHARMA", "POWERGRID", "NTPC", "HEROMOTOCO", "DRREDDY", "DIVISLAB", "CIPLA",
    "TATASTEEL", "INDUSINDBK", "ONGC", "TATACONSUM", "JSWSTEEL", "M&M", "WIPRO",
    "TECHM", "BPCL", "COALINDIA", "IOC", "TATAMOTORS", "PIDILITIND", "GAIL", "SBILIFE",
    "EICHERMOT", "HAVELLS", "ADANIPORTS", "ADANIENT", "BANKBARODA", "PNB", "CANBK",
    "UNIONBANK", "IDFCFIRSTB", "YESBANK", "LUPIN", "BIOCON", "SRF", "PAGEIND",
    "COLPAL", "BATAINDIA", "VEDL", "TVSMOTOR", "TATAELXSI", "INDIGO", "IRCTC",
    "ICICIPRULI", "ICICIGI", "CHOLAFIN", "DALBHARAT", "MANAPPURAM", "BALRAMCHIN",
    "BALKRISIND", "AMBUJACEM", "ABB", "AUROPHARMA", "SAIL"
]


def generate_totp(secret: str) -> str:
    totp = pyotp.TOTP(secret)
    return totp.now()


def login(client_code: str, api_key: str, mpin: str, totp: str) -> SmartConnect:
    smartapi = SmartConnect(api_key=api_key)
    response = smartapi.generateSession(client_code, mpin, totp)
    if not response["status"]:
        raise Exception(f"Login failed: {response['message']}")
    print("✅ Logged in successfully")
    return smartapi


def load_instrument_dump() -> pd.DataFrame:
    url = "https://margincalculator.angelbroking.com/OpenAPI_File/files/OpenAPIScripMaster.json"
    return pd.read_json(url)


def get_token(df: pd.DataFrame, symbol: str, exchange: str) -> str:
    match = df[(df["symbol"] == f"{symbol}-EQ") & (df["exch_seg"] == exchange)]
    if match.empty:
        raise ValueError(f"❌ Symbol {symbol} not found in {exchange}")
    return str(match.iloc[0]["token"])


def fetch_candle_data(smartapi: SmartConnect, token: str, exchange: str) -> pd.DataFrame:
    from_date = (datetime.datetime.now() - datetime.timedelta(days=90)).strftime("%Y-%m-%d %H:%M")
    to_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    params = {
        "exchange": exchange,
        "symboltoken": token,
        "interval": "ONE_DAY",
        "fromdate": from_date,
        "todate": to_date
    }

    response = smartapi.getCandleData(params)
    if "data" not in response or not response["data"]:
        raise Exception("❌ Error while fetching candle data")

    df = pd.DataFrame(response["data"], columns=["timestamp", "open", "high", "low", "close", "volume"])
    df["timestamp"] = pd.to_datetime(df["timestamp"]).dt.tz_localize(None)
    return df


def calculate_metrics(df: pd.DataFrame) -> dict:
    df = df.copy()
    df["avg_ohlc"] = (df["open"] + df["high"] + df["low"] + df["close"]) / 4
    ohlc_60 = df.tail(60)["avg_ohlc"].mean()
    sma_50 = df["close"].rolling(window=50).mean().iloc[-1]
    sma_20 = df["close"].rolling(window=20).mean().iloc[-1]

    diff = ohlc_60 - sma_50
    diff_percent = (diff / sma_50) * 100 if sma_50 else 0

    return {
        "OHLC_60": round(ohlc_60, 2),
        "SMA_50": round(sma_50, 2),
        "SMA_20": round(sma_20, 2),
        "DIFF": round(diff, 2),
        "DIFF_%": round(diff_percent, 2)
    }

def is_last_thursday():
    today = datetime.date.today()
    if today.weekday() != 3:
        return False
    next_thursday = today + datetime.timedelta(days=7)
    return next_thursday.month != today.month

def send_email_with_attachment(to_email, subject, body, attachment_path):
    try:
        yag = yagmail.SMTP(
            user=os.environ["EMAIL_USER"],
            password=os.environ["EMAIL_PASS"]
        )
        yag.send(to=to_email, subject=subject, contents=body, attachments=attachment_path)
        print("📧 Email sent successfully!")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")


def main():
    if not is_last_thursday():
        print("⏭️ Not last Thursday, exiting.")
        # return
    otp = generate_totp(TOTP_SECRET)
    smartapi = login(CLIENT_CODE, API_KEY, MPIN, otp)
    instrument_df = load_instrument_dump()

    results = []
    for sym in SYMBOLS:
        try:
            print(f"\n📌 Processing {sym}...")
            token = get_token(instrument_df, sym, EXCHANGE)
            df = fetch_candle_data(smartapi, token, EXCHANGE)
            metrics = calculate_metrics(df)
            results.append({"SYMBOL": sym, **metrics})
        except Exception as e:
            print(f"❌ {sym}: {e}")

    final_df = pd.DataFrame(results)
    file_name = "bulk_stock_analysis.xlsx"
    final_df.to_excel(file_name, index=False)
    print("\n✅ Analysis complete. File saved as bulk_stock_analysis.xlsx")
    email_list = [email.strip() for email in os.environ["TO_EMAIL"].split(",") if email.strip()]
    for email in email_list:
        if email:
            send_email_with_attachment(
                to_email=email,
                subject=f"Salary Trade - Monthly Analysis Report",
                body="Please find the attached stock analysis report.",
                attachment_path=file_name
            )


if __name__ == "__main__":
    main()