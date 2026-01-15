import requests
import gspread
import json
import os
from google.oauth2.service_account import Credentials

# ===== 基本設定 =====
SPREADSHEET_ID = "1H3JDRbMVSWjZvIHFtFzv3NhPKPO-KRqiX_XMpLtjjZI"
STOCKS = {
    "2330": "2330",
    "0050": "0050",
    "006208": "006208"
}

# ===== Google Sheets 認證 =====
creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
gc = gspread.authorize(credentials)
sh = gc.open_by_key(SPREADSHEET_ID)

# ===== 價格清洗（只去逗號，不動數值）=====
def clean_price(value):
    return float(value.replace(",", ""))

# ===== 證交所 API：抓「最近一個交易日」=====
def fetch_latest_twse(stock_id):
    url = "https://www.twse.com.tw/exchangeReport/STOCK_DAY"
    params = {
        "response": "json",
        "stockNo": stock_id
    }
    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    r = requests.get(url, params=params, headers=headers, timeout=10)
    data = r.json()

    if data["stat"] != "OK" or not data["data"]:
        return None

    last = data["data"][-1]  # 最近交易日

    # 民國轉西元
    roc_date = last[0]
    y, m, d = roc_date.split("/")
    date = f"{int(y) + 1911}-{m}-{d}"

    return [
        date,
        clean_price(last[3]),  # 開盤
        clean_price(last[4]),  # 最高
        clean_price(last[5]),  # 最低
        clean_price(last[6]),  # 收盤
        int(last[1].replace(",", ""))  # 成交股數
    ]

# ===== 寫入 Google Sheets =====
for sheet_name, stock_id in STOCKS.items():
    try:
        ws = sh.worksheet(sheet_name)
    except:
        ws = sh.add_worksheet(title=sheet_name, rows="5000", cols="10")
        ws.append_row(["日期", "開盤", "最高", "最低", "收盤", "成交股數"])

    row = fetch_latest_twse(stock_id)
    if not row:
        continue

    dates = ws.col_values(1)
    if row[0] not in dates:
        ws.append_row(row)
