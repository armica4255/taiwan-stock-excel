import requests
import gspread
import json
import os
from datetime import datetime
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

# ===== 今日日期 =====
today = datetime.now().strftime("%Y%m%d")

# ===== 證交所 API =====
def fetch_twse(stock_id):
    url = "https://www.twse.com.tw/exchangeReport/STOCK_DAY"
    params = {
        "response": "json",
        "date": today,
        "stockNo": stock_id
    }
    r = requests.get(url, params=params, timeout=10)
    data = r.json()

    if data["stat"] != "OK":
        return None

    # 只取「最後一筆（最新交易日）」
    last = data["data"][-1]

    # 日期轉西元
    roc_date = last[0]  # 112/01/15
    y, m, d = roc_date.split("/")
    date = f"{int(y)+1911}-{m}-{d}"

    return [
        date,
        int(float(last[3])),  # 開盤
        int(float(last[4])),  # 最高
        int(float(last[5])),  # 最低
        int(float(last[6])),  # 收盤
        int(last[1].replace(",", ""))  # 成交股數
    ]

# ===== 寫入 Sheet =====
for sheet_name, stock_id in STOCKS.items():
    try:
        ws = sh.worksheet(sheet_name)
    except:
        ws = sh.add_worksheet(title=sheet_name, rows="2000", cols="10")
        ws.append_row(["日期", "開盤", "最高", "最低", "收盤", "成交股數"])

    row = fetch_twse(stock_id)
    if not row:
        continue

    dates = ws.col_values(1)
    if row[0] not in dates:
        ws.append_row(row)
