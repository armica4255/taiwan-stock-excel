import yfinance as yf
import pandas as pd
import gspread
import json
import os
from google.oauth2.service_account import Credentials
from datetime import datetime

STOCKS = {
    "2330.TW": "2330",
    "0050.TW": "0050",
    "006208.TW": "006208"
}

START_DATE = "2020-01-02"
END_DATE = datetime.today().strftime("%Y-%m-%d")

creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
gc = gspread.authorize(creds)

sheet = gc.open("台股每日紀錄")

for ticker, tab in STOCKS.items():
    df = yf.download(ticker, start=START_DATE, end=END_DATE)
    df = df[["Open", "High", "Low", "Close"]]
    df.reset_index(inplace=True)
    df.columns = ["日期", "開盤價", "最高價", "最低價", "收盤價"]

    df["漲跌幅%"] = df["收盤價"].pct_change() * 100
    df["漲跌幅%"] = df["漲跌幅%"].round(2)

    ws = sheet.worksheet(tab)
    ws.clear()
    ws.update([df.columns.values.tolist()] + df.values.tolist())
