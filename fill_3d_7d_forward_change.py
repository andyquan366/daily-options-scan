import pandas as pd
from datetime import datetime, timedelta
import yfinance as yf
from openpyxl import load_workbook

file_name = "option_activity_log.xlsx"
wb = load_workbook(file_name)

today = datetime.today().date()
base_day = today - timedelta(days=1)  # 基准日为昨天

day_3ago = base_day - timedelta(days=3)
day_7ago = base_day - timedelta(days=7)

# 准备股票列表（示例从Excel里读所有ticker，按需替换为你的纳斯达克100+标普500列表）
all_tickers = set()

# 先从所有sheet里汇总ticker
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    header = [cell.value for cell in ws[1]]
    ticker_col = header.index("Ticker") + 1
    for r in range(2, ws.max_row + 1):
        t = ws.cell(row=r, column=ticker_col).value
        if t:
            all_tickers.add(str(t).upper())

tickers_str = " ".join(all_tickers)

# 拉取三个关键日期的价格数据
dates_to_fetch = [day_3ago, day_7ago, base_day]
price_data = {}

for dt in dates_to_fetch:
    start = dt
    end = dt + timedelta(days=1)
    data = yf.download(tickers_str, start=start, end=end, progress=False, group_by='ticker')
    price_data[dt] = data

# 定位某行函数，按sheet、ticker和日期精确找行
def find_row(ws, ticker, dt):
    header = [cell.value for cell in ws[1]]
    date_col = header.index("Date") + 1
    ticker_col = header.index("Ticker") + 1
    for r in range(2, ws.max_row + 1):
        dt_cell = ws.cell(row=r, column=date_col).value
        t_cell = ws.cell(row=r, column=ticker_col).value
        if not dt_cell or not t_cell:
            continue
        # 转换日期格式
        if isinstance(dt_cell, datetime):
            dt_val = dt_cell.date()
        elif isinstance(dt_cell, str):
            dt_val = datetime.strptime(dt_cell[:10], "%Y-%m-%d").date()
        else:
            dt_val = dt_cell
        if dt_val == dt and str(t_cell).upper() == ticker:
            return r
    return None

# 确保3D和7D列存在
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    header = [cell.value for cell in ws[1]]
    if "3D Forward Change" not in header:
        ws.cell(row=1, column=ws.max_column + 1).value = "3D Forward Change"
    if "7D Forward Change" not in header:
        ws.cell(row=1, column=ws.max_column + 1).value = "7D Forward Change"

# 遍历所有sheet，更新3D和7D涨幅
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    header = [cell.value for cell in ws[1]]
    date_col = header.index("Date") + 1
    ticker_col = header.index("Ticker") + 1
    close_col = header.index("Previous Close") + 1
    change_col_3 = header.index("3D Forward Change") + 1
    change_col_7 = header.index("7D Forward Change") + 1

    for r in range(2, ws.max_row + 1):
        dt_cell = ws.cell(row=r, column=date_col).value
        t_cell = ws.cell(row=r, column=ticker_col).value
        if not dt_cell or not t_cell:
            continue
        if isinstance(dt_cell, datetime):
            dt_val = dt_cell.date()
        elif isinstance(dt_cell, str):
            dt_val = datetime.strptime(dt_cell[:10], "%Y-%m-%d").date()
        else:
            dt_val = dt_cell
        ticker_str = str(t_cell).upper()

        # 计算3D涨幅对应基准日期是dt_val+3天
        base_3d = dt_val + timedelta(days=3)
        if base_3d == base_day:
            # 从price_data里取价格计算涨幅
            price_base = None
            price_past = None
            # 取基准日期价格
            try:
                if len(all_tickers) == 1:
                    price_base = price_data[base_3d]['Close'].iloc[0]
                else:
                    price_base = price_data[base_3d][ticker_str]['Close'].iloc[0]
            except:
                pass
            # 取过去日期价格
            try:
                if len(all_tickers) == 1:
                    price_past = price_data[dt_val]['Close'].iloc[0]
                else:
                    price_past = price_data[dt_val][ticker_str]['Close'].iloc[0]
            except:
                pass
            if price_base and price_past and price_past != 0:
                change = (price_base - price_past) / price_past
                ws.cell(row=r, column=change_col_3).value = round(change, 4)
                ws.cell(row=r, column=change_col_3).number_format = "0.00%"

        # 计算7D涨幅对应基准日期是dt_val+7天
        base_7d = dt_val + timedelta(days=7)
        if base_7d == base_day:
            price_base = None
            price_past = None
            try:
                if len(all_tickers) == 1:
                    price_base = price_data[base_7d]['Close'].iloc[0]
                else:
                    price_base = price_data[base_7d][ticker_str]['Close'].iloc[0]
            except:
                pass
            try:
                if len(all_tickers) == 1:
                    price_past = price_data[dt_val]['Close'].iloc[0]
                else:
                    price_past = price_data[dt_val][ticker_str]['Close'].iloc[0]
            except:
                pass
            if price_base and price_past and price_past != 0:
                change = (price_base - price_past) / price_past
                ws.cell(row=r, column=change_col_7).value = round(change, 4)
                ws.cell(row=r, column=change_col_7).number_format = "0.00%"

wb.save(file_name)
print("补齐完成")
