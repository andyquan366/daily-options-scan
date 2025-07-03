import os
from openpyxl import load_workbook

file_name = "option_activity_log.xlsx"
sheet_name = "2025-06"

# ==== 云端自动拉取最新 Excel ====
if "GITHUB_ACTIONS" in os.environ:
    print("开始用 rclone 拉取最新 Excel 文件...")
    ret = os.system('rclone copy "gdrive:/Investing/Daily top options/option_activity_log.xlsx" ./ --drive-chunk-size 64M --progress --ignore-times')
    print(f"rclone 返回码: {ret}")

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

header = [cell.value for cell in ws[1]]
print(f"工作表 {sheet_name} 第一行列标题: {header}")
