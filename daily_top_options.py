import os
import pandas as pd

# ====== 1. 拉取云端 Excel 覆盖本地（option_activity_log.xlsx）======
os.system('rclone copy "gdrive:/Investing/Daily top options" ./option_activity_log.xlsx --drive-chunk-size 64M --progress --ignore-times')

# ====== 2. 读取 Excel，若不存在则新建 DataFrame ======
if os.path.exists('option_activity_log.xlsx'):
    try:
        df = pd.read_excel('option_activity_log.xlsx')
    except Exception as e:
        print("读取失败，创建新表。异常：", e)
        df = pd.DataFrame(columns=['Index', 'Msg'])
else:
    df = pd.DataFrame(columns=['Index', 'Msg'])

# ====== 3. 追加一行 ======
idx = len(df) + 1
df = pd.concat([df, pd.DataFrame([{'Index': idx, 'Msg': f'Hello {idx}'}])], ignore_index=True)

# ====== 4. 保存 ======
df.to_excel('option_activity_log.xlsx', index=False)

# ====== 5. 上传回 Google Drive（覆盖远端）======
os.system('rclone copy ./option_activity_log.xlsx "gdrive:/Investing/Daily top options/option_activity_log.xlsx" --drive-chunk-size 64M --progress --ignore-times')

print('✅ 云端1+1测试完成，当前总行数:', len(df))
