import pandas as pd
from datetime import datetime, timedelta
import os
import pytz
from openpyxl import load_workbook

# ==== 云端自动拉取最新 Excel ====
if "GITHUB_ACTIONS" in os.environ:
    os.system('rclone copy "gdrive:/Investing/Daily top options/option_activity_log.xlsx" ./ --drive-chunk-size 64M --progress --ignore-times')

# ✅ 全局设定 Toronto 本地时间
tz = pytz.timezone("America/Toronto")
now = datetime.now(tz)
today = now.date()
yesterday = today - timedelta(days=1)
target_day = today - timedelta(days=3)   # 优先补3天前，但回溯到最近有数据的实际交易日

file_name = "option_activity_log.xlsx"
wb = load_workbook(file_name, read_only=False, data_only=True)
sheet_names = wb.sheetnames

def is_month_sheet(name):
    try:
        datetime.strptime(name, "%Y-%m")
        return True
    except:
        return False

fill_total = 0
for sheet_name in sheet_names:
    if not is_month_sheet(sheet_name):
        continue
    df = pd.read_excel(file_name, sheet_name=sheet_name)
    if "3D Forward Change" not in df.columns:
        df["3D Forward Change"] = None

    # === 回溯到最近有数据的 target_day ===
    search_day = target_day
    mask = df["Date"] == search_day.strftime("%Y-%m-%d")
    back_offset = 0
    while df[mask].empty and back_offset < 7:
        back_offset += 1
        search_day = search_day - timedelta(days=1)
        mask = df["Date"] == search_day.strftime("%Y-%m-%d")

    fill_count = 0
    for idx, row in df[mask].iterrows():
        ticker = row["Ticker"]
        price_3ago = row["Previous Close"]
        match = df[(df["Ticker"] == ticker) & (df["Date"] == yesterday.strftime('%Y-%m-%d'))]
        if not match.empty and price_3ago and not pd.isna(price_3ago):
            price_3later = match.iloc[0]["Previous Close"]
            if price_3later and not pd.isna(price_3later):
                try:
                    change = (price_3later - price_3ago) / price_3ago
                    df.at[idx, "3D Forward Change"] = round(change, 4)
                    fill_count += 1
                except Exception as e:
                    print(f"❌ Error for {ticker}: {e}")
    fill_total += fill_count

    with pd.ExcelWriter(file_name, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"✅ Sheet {sheet_name}: 回溯补齐日期 {search_day.strftime('%Y-%m-%d')} 共 {fill_count} 条 3D Forward Change")

print(f"✅ 全部月度sheet已补齐 共 {fill_total} 条 3D Forward Change")

# ==== 云端自动上传 Excel ====
if "GITHUB_ACTIONS" in os.environ:
    os.system('rclone copy ./option_activity_log.xlsx "gdrive:/Investing/Daily top options" --drive-chunk-size 64M --progress --ignore-times')
