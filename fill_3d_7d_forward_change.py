import time
import os
import re
from datetime import datetime, timedelta
import pytz
from openpyxl import load_workbook
import yfinance as yf

file_name = "option_activity_log.xlsx"
sheet_to_process = "2025-06"

def wait_for_file_stable(filename, wait_seconds=10, interval=1):
    """等待文件大小连续稳定 interval 秒，最多 wait_seconds 秒"""
    stable_count = 0
    last_size = -1
    max_checks = wait_seconds // interval
    for _ in range(max_checks):
        try:
            size = os.path.getsize(filename)
            if size == last_size:
                stable_count += 1
                if stable_count >= 3:  # 连续3次大小不变，认为稳定
                    return True
            else:
                stable_count = 0
            last_size = size
        except Exception as e:
            print(f"检查文件大小异常: {e}")
        time.sleep(interval)
    return False

if not wait_for_file_stable(file_name):
    print("警告：文件大小不稳定，可能未写入完成！")

wb = load_workbook(file_name)

if sheet_to_process not in wb.sheetnames:
    print(f"工作表 {sheet_to_process} 不存在，退出")
    exit()

ws = wb[sheet_to_process]

header = [str(cell.value).strip() if cell.value is not None else '' for cell in ws[1]]
print(f"工作表 {sheet_to_process} 的第一行列标题是: {header}")

required_cols = ["Date", "Ticker", "Previous Close", "3D Forward Change", "7D Forward Change"]
missing_cols = [col for col in required_cols if col not in header]
if missing_cols:
    print(f"缺失的列: {missing_cols}，退出")
    exit()

date_col = header.index("Date") + 1
ticker_col = header.index("Ticker") + 1
prev_close_col = header.index("Previous Close") + 1
col_3d = header.index("3D Forward Change") + 1
col_7d = header.index("7D Forward Change") + 1

tz = pytz.timezone("America/Toronto")
now = datetime.now(tz)
today = now.date()
base_day = today - timedelta(days=1)

yesterday_close_cache = {}

# 缓存基准日所有ticker收盘价
for r in range(2, ws.max_row + 1):
    dt_cell = ws.cell(row=r, column=date_col).value
    if not dt_cell:
        continue
    if isinstance(dt_cell, datetime):
        dt_val = dt_cell.date()
    elif isinstance(dt_cell, str):
        dt_val = datetime.strptime(dt_cell[:10], "%Y-%m-%d").date()
    else:
        dt_val = dt_cell
    if dt_val == base_day:
        ticker = str(ws.cell(row=r, column=ticker_col).value).upper()
        prev_close = ws.cell(row=r, column=prev_close_col).value
        if prev_close is not None:
            yesterday_close_cache[ticker] = prev_close

count_3d = 0
count_7d = 0

def get_price_realtime(ticker, target_date, max_lookback=14):
    for i in range(max_lookback):
        try_date = target_date - timedelta(days=i)
        if try_date.weekday() >= 5:
            continue
        try:
            stock = yf.Ticker(ticker)
            hist = stock.history(
                start=try_date.strftime("%Y-%m-%d"),
                end=(try_date + timedelta(days=1)).strftime("%Y-%m-%d"),
            )
            if not hist.empty:
                close_price = hist['Close'].iloc[0]
                print(f"获取 {ticker} 于 {try_date} 的收盘价: {close_price}")
                return close_price
        except Exception as e:
            print(f"错误: 获取 {ticker} 于 {try_date} 收盘价失败，原因: {e}")
            continue
    print(f"未找到 {ticker} 在 {target_date} 附近的有效收盘价")
    return None

for r in range(2, ws.max_row + 1):
    dt_cell = ws.cell(row=r, column=date_col).value
    if not dt_cell:
        continue
    if isinstance(dt_cell, datetime):
        dt_val = dt_cell.date()
    elif isinstance(dt_cell, str):
        dt_val = datetime.strptime(dt_cell[:10], "%Y-%m-%d").date()
    else:
        dt_val = dt_cell

    ticker = str(ws.cell(row=r, column=ticker_col).value).upper()
    prev_close = ws.cell(row=r, column=prev_close_col).value
    if prev_close is None or prev_close == 0:
        continue

    days_diff = (base_day - dt_val).days

    if days_diff == 3:
        if ticker in yesterday_close_cache:
            close_yesterday = yesterday_close_cache[ticker]
        else:
            close_yesterday = get_price_realtime(ticker, base_day)
            if close_yesterday is not None:
                yesterday_close_cache[ticker] = close_yesterday
        if close_yesterday:
            change_3d = (close_yesterday - prev_close) / prev_close
            ws.cell(row=r, column=col_3d).value = round(change_3d, 4)
            ws.cell(row=r, column=col_3d).number_format = "0.00%"
            count_3d += 1
            print(f"3D 补齐: 行 {r}，股票 {ticker}，日期 {dt_val}，涨跌幅 {change_3d:.4%}")

    if days_diff == 7:
        if ticker in yesterday_close_cache:
            close_yesterday = yesterday_close_cache[ticker]
        else:
            close_yesterday = get_price_realtime(ticker, base_day)
            if close_yesterday is not None:
                yesterday_close_cache[ticker] = close_yesterday
        if close_yesterday:
            change_7d = (close_yesterday - prev_close) / prev_close
            ws.cell(row=r, column=col_7d).value = round(change_7d, 4)
            ws.cell(row=r, column=col_7d).number_format = "0.00%"
            count_7d += 1
            print(f"7D 补齐: 行 {r}，股票 {ticker}，日期 {dt_val}，涨跌幅 {change_7d:.4%}")

print(f"工作表 {sheet_to_process} 补齐结果：3D共 {count_3d} 条，7D共 {count_7d} 条")

wb.save(file_name)

if "GITHUB_ACTIONS" in os.environ:
    os.system(f'rclone copy ./{file_name} "gdrive:/Investing/Daily top options" --drive-chunk-size 64M --progress --ignore-times')

print("补齐完成")
