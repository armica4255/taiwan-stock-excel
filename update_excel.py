import yfinance as yf
import pandas as pd
import gspread
import json
import os
from google.oauth2.service_account import Credentials
from datetime import datetime

# ===== 固定你的 Google Sheets ID（非常重要）=====
SPREADSHEET_ID = "1H3JDRbMVSWjZvIHFtFzv3NhPKPO-KRqiX_XMpLtjjZI"

# 台股標的
STOCKS = {
    "2330.TW": "2330",
    "0050.TW": "0050",
    "006208.TW": "006208"
}

START_DATE = "2020-01-02"
END_DATE = datetime.today().strftime("%Y-%m-%d")

# 讀取 GitHub Secret 裡的 Google 金鑰
creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
gc = gspread.authorize(creds)

# ✅ 用「試算表 ID」打開（不會再寫錯）
sheet = gc.open_by_key(SPREADSHEET_ID)

for ticker, sheet_name in STOCKS.items():
    print(f"Downloading {ticker}")

    df = yf.download(
        ticker,
        start=START_DATE,
        end=END_DATE,
        progress=False,
        auto_adjust=False
    )

    if df.empty:
        print(f"No data for {ticker}, skip.")
        continue

    df = df[["Open", "High", "Low", "Close"]]
    df.reset_index(inplace=True)
    df.columns = ["日期", "開盤價", "最高價", "最低價", "收盤價"]

    # 計算漲跌幅 %
    df["漲跌幅%"] = df["收盤價"].pct_change() * 100
    df["漲跌幅%"] = df["漲跌幅%"].round(2)

    ws = sheet.worksheet(sheet_name)
    ws.clear()

    ws.update(
        [df.columns.tolist()] +
        df.astype(str).values.tolist()
    )

print("Update completed")
