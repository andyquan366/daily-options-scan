import pandas as pd
import pytz
from datetime import datetime, timedelta
import yfinance as yf
import os
from openpyxl import load_workbook

# ==== 云端自动拉取最新 Excel ====
if "GITHUB_ACTIONS" in os.environ:
    os.system('rclone copy "gdrive:/Investing/Daily top options/option_activity_log.xlsx" ./ --drive-chunk-size 64M --progress --ignore-times')

tz = pytz.timezone("America/Toronto")
now = datetime.now(tz)
today_str = now.strftime("%Y-%m-%d")
file_name = "option_activity_log.xlsx"

wb = load_workbook(file_name)
sheet_names = wb.sheetnames

def is_month_sheet(name):
    try:
        datetime.strptime(name, "%Y-%m")
        return True
    except:
        return False

min_date = input("请输入补齐起始日期 (yyyy-mm-dd，留空不限制)：").strip()
max_date = input("请输入补齐截止日期 (yyyy-mm-dd，留空不限制)：").strip()

for sheet_name in sheet_names:
    if not is_month_sheet(sheet_name):
        continue
    ws = wb[sheet_name]
    df = pd.read_excel(file_name, sheet_name=sheet_name, dtype=str)
    df_date = pd.to_datetime(df["Date"]).dt.strftime("%Y-%m-%d")

    if "Previous Close" not in df.columns:
        print(f"⚠️ Sheet {sheet_name} 缺少 'Previous Close' 列，跳过")
        continue

    header = list(df.columns)
    col_index = header.index("Previous Close") + 1  # openpyxl 从1开始
    date_index = header.index("Date")
    ticker_index = header.index("Ticker")

    fill_count = 0
    for i, row in df.iterrows():
        row_date = df_date[i]
        ticker = row["Ticker"]
        prev_close_cell = ws.cell(row=i+2, column=col_index)

        if pd.isna(ticker) or row_date == today_str or prev_close_cell.value not in [None, "", "nan"]:
            continue

        back_offset = 0
        prev_close = None
        while back_offset < 7 and prev_close is None:
            search_date = (datetime.strptime(row_date, "%Y-%m-%d") - timedelta(days=back_offset)).strftime("%Y-%m-%d")
            stock = yf.Ticker(ticker)
            try:
                hist = stock.history(
                    start=search_date,
                    end=(datetime.strptime(search_date, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")
                )
                if not hist.empty:
                    prev_close = float(hist['Close'].iloc[0])
                    print(f"✅ {ticker}: got close {prev_close} ({search_date} for {row_date})")
                else:
                    print(f"⏩ {ticker}: no price ({search_date} for {row_date}), try previous day")
            except Exception as e:
                print(f"❌ {ticker}: {e} ({search_date} for {row_date})")
            back_offset += 1

        if prev_close is not None:
            prev_close_cell.value = prev_close
            fill_count += 1

    print(f"✅ Sheet {sheet_name}: 补齐 Previous Close 共 {fill_count} 条")

# ✅ 保存不会破坏格式
wb.save(file_name)

print("✅ 所有历史 Previous Close 均已补齐（今天除外，格式无破坏）")

# ==== 云端自动上传 Excel ====
if "GITHUB_ACTIONS" in os.environ:
    os.system('rclone copy ./option_activity_log.xlsx "gdrive:/Investing/Daily top options" --drive-chunk-size 64M --progress --ignore-times')
