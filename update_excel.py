import gspread
import json
import os
from google.oauth2.service_account import Credentials
from datetime import datetime

creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
gc = gspread.authorize(creds)

sheet = gc.open("台股每日紀錄")

ws = sheet.worksheet("2330")
ws.clear()

today = datetime.today().strftime("%Y-%m-%d")

ws.update([
    ["日期", "開盤價", "最高價", "最低價", "收盤價", "漲跌幅%"],
    [today, 600, 610, 590, 605, 0.83]
])
