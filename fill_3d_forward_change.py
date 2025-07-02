import pandas as pd
from datetime import datetime, timedelta
import yfinance as yf
from openpyxl import load_workbook

# 设定关键日期
target_day_str = "2025-06-29"
target_day = datetime.strptime(target_day_str, "%Y-%m-%d").date()
forward_day = target_day + timedelta(days=3)  # 2025-07-02

file_name = "option_activity_log.xlsx"
wb = load_workbook(file_name)
sheet_names = wb.sheetnames

def is_month_sheet(name):
    try:
        datetime.strptime(name, "%Y-%m")
        return True
    except:
        return False

# 只处理 target_day 所在的 sheet（示例只处理2025-06）
sheet_name = target_day.strftime("%Y-%m")
if sheet_name not in sheet_names:
    raise ValueError(f"找不到目标sheet {sheet_name}")

ws = wb[sheet_name]
header = [cell.value for cell in ws[1]]
date_col = header.index("Date") + 1
ticker_col = header.index("Ticker") + 1
close_col = header.index("Previous Close") + 1
change_col = header.index("3D Forward Change") + 1

# 读取所有6月29日行的ticker和价格
rows_to_update = []
tickers = []
for r in range(2, ws.max_row + 1):
    date_cell = ws.cell(row=r, column=date_col).value
    if not date_cell:
        continue
    row_date = pd.to_datetime(date_cell).date()
    if row_date != target_day:
        continue
    ticker = ws.cell(row=r, column=ticker_col).value
    price_3ago = ws.cell(row=r, column=close_col).value
    if not ticker or not price_3ago:
        continue
    ticker_str = str(ticker).strip().upper()
    tickers.append(ticker_str)
    rows_to_update.append((r, price_3ago, ticker_str))

# 用yfinance批量下载forward_day的收盘价
tickers_str = " ".join(tickers)
data = yf.download(tickers_str, start=forward_day, end=forward_day + timedelta(days=1), progress=False, group_by='ticker')

# yfinance返回数据格式复杂，兼容单一ticker和多个ticker情况
def get_close(ticker):
    try:
        if len(tickers) == 1:
            # 单ticker时data是DataFrame
            return data['Close'].iloc[0]
        else:
            return data[ticker]['Close'].iloc[0]
    except Exception:
        return None

# 计算并写入3D Forward Change
fill_count = 0
for r, price_3ago, ticker_str in rows_to_update:
    price_3later = get_close(ticker_str)
    if price_3later is None or price_3later == 0:
        print(f"⚠️ {ticker_str} {forward_day} 无收盘价，跳过")
        continue
    change = (price_3later - price_3ago) / price_3ago
    ws.cell(row=r, column=change_col).value = round(change, 4)
    ws.cell(row=r, column=change_col).number_format = "0.00%"
    fill_count += 1

print(f"✅ 补齐 {target_day} Forward Change 共 {fill_count} 条")

wb.save(file_name)
