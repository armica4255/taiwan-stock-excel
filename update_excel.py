import requests
import gspread
import json
import os
from datetime import datetime
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = "1H3JDRbMVSWjZvIHFtFzv3NhPKPO-KRqiX_XMpLtjjZI"
STOCKS = ["2330", "0050", "006208"]

# ===== Google Sheets 認證 =====
creds = json.loads(os.environ["GOOGLE_CREDENTIALS"])
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
gc = gspread.authorize(Credentials.from_service_account_info(creds, scopes=scopes))
sh = gc.open_by_key(SPREADSHEET_ID)

HEADERS = ["日期", "開盤", "最高", "最低", "收盤", "成交股數"]

def fmt_price(v):
    return round(float(v.replace(",", "")), 2)

def fetch_month(stock, yyyymm):
    url = "https://www.twse.com.tw/exchangeReport/STOCK_DAY"
    r = requests.get(
        url,
        params={"response": "json", "date": yyyymm, "stockNo": stock},
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=10
    )
    j = r.json()
    if j["stat"] != "OK":
        return []

    rows = []
    for d in j["data"]:
        y, m, d2 = d[0].split("/")
        date = f"{int(y)+1911}-{m}-{d2}"
        rows.append([
            date,
            fmt_price(d[3]),
            fmt_price(d[4]),
            fmt_price(d[5]),
            fmt_price(d[6]),
            int(d[1].replace(",", ""))
        ])
    return rows

def month_range(start="202001"):
    y, m = int(start[:4]), int(start[4:])
    now = datetime.now()
    while (y < now.year) or (y == now.year and m <= now.month):
        yield f"{y}{m:02d}"
        m += 1
        if m == 13:
            y += 1
            m = 1

for stock in STOCKS:
    try:
        ws = sh.worksheet(stock)
        ws.clear()
    except:
        ws = sh.add_worksheet(stock, rows=6000, cols=10)

    ws.append_row(HEADERS)

    all_rows = []
    for ym in month_range():
        all_rows.extend(fetch_month(stock, ym))

    if all_rows:
        ws.append_rows(all_rows)
