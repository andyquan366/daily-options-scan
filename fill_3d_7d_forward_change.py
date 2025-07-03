import os
import time
from openpyxl import load_workbook



# ==== 云端自动拉取最新 Excel ====
if "GITHUB_ACTIONS" in os.environ:
    ret = os.system('rclone copy "gdrive:/Investing/Daily top options/option_activity_log.xlsx" ./ --drive-chunk-size 64M --progress --ignore-times')
    if ret != 0:
        print(f"rclone 命令失败，返回码: {ret}")
        exit(1)