import os
import time
from openpyxl import load_workbook



# ==== 云端自动拉取最新 Excel ====
if "GITHUB_ACTIONS" in os.environ:
    os.system('rclone copy "gdrive:/Investing/Daily top options/option_activity_log.xlsx" ./ --drive-chunk-size 64M --progress --ignore-times')

file_name = "option_activity_log.xlsx"
sheet_name = "2025-06"

def wait_for_file_stable(filename, wait_seconds=15, interval=1):
    last_size = -1
    stable_count = 0
    for _ in range(wait_seconds // interval):
        try:
            size = os.path.getsize(filename)
            if size == last_size:
                stable_count += 1
                if stable_count >= 3:
                    return True
            else:
                stable_count = 0
            last_size = size
        except Exception as e:
            print(f"检查文件大小异常: {e}")
        time.sleep(interval)
    return False

if not wait_for_file_stable(file_name):
    print("警告：文件大小未稳定，可能未写入完成")

if not os.path.exists(file_name):
    print(f"文件 {file_name} 不存在，退出")
    exit(1)

print(f"文件大小: {os.path.getsize(file_name)} bytes")

try:
    wb = load_workbook(file_name)
except Exception as e:
    print(f"打开 Excel 文件失败: {e}")
    exit(1)

print(f"工作表列表: {wb.sheetnames}")

if sheet_name not in wb.sheetnames:
    print(f"工作表 {sheet_name} 不存在，退出")
    exit(1)

ws = wb[sheet_name]

header = [str(cell.value).strip() if cell.value is not None else '' for cell in ws[1]]
print(f"工作表 {sheet_name} 第一行列标题: {header}")
