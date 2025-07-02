import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook

file_name = "option_activity_log.xlsx"
wb = load_workbook(file_name)

today = datetime.today().date()
yesterday = today - timedelta(days=1)
day_3ago = yesterday - timedelta(days=3)
day_7ago = yesterday - timedelta(days=7)

header_cache = {}  # 记住每个sheet的表头

def get_sheet_by_date(dt):
    sheet_name = dt.strftime("%Y-%m")
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"找不到sheet: {sheet_name}")
    return wb[sheet_name]

def get_date_row_and_close(ws, dt):
    if ws.title not in header_cache:
        header = [cell.value for cell in ws[1]]
        header_cache[ws.title] = header
    else:
        header = header_cache[ws.title]
    date_col = header.index("Date") + 1
    close_col = header.index("Previous Close") + 1
    for r in range(2, ws.max_row + 1):
        dt_cell = ws.cell(row=r, column=date_col).value
        if not dt_cell:
            continue
        if isinstance(dt_cell, datetime):
            dt_cell_date = dt_cell.date()
        elif isinstance(dt_cell, str):
            dt_cell_date = datetime.strptime(dt_cell[:10], "%Y-%m-%d").date()
        else:
            dt_cell_date = dt_cell
        if dt_cell_date == dt:
            price = ws.cell(row=r, column=close_col).value
            return r, price, header
    return None, None, header

def ensure_columns(ws, header, cols):
    # cols 是 字段名列表，如 ["3D Forward Change", "7D Forward Change"]
    for col_name in cols:
        if col_name not in header:
            ws.cell(row=1, column=ws.max_column + 1).value = col_name
            header.append(col_name)
    # 更新列索引返回
    return {name: header.index(name)+1 for name in cols}

def compute_and_write(base_price, target_ws, target_row, col_idx, target_date):
    if target_row is None or base_price is None:
        print(f"{target_ws.title} {target_date} 缺失数据，跳过")
        return
    past_price = target_ws.cell(row=target_row, column=header_cache[target_ws.title].index("Previous Close")+1).value
    if past_price is None or past_price == 0:
        print(f"{target_ws.title} {target_date} 价格缺失，跳过")
        return
    change = (base_price - past_price) / past_price
    target_ws.cell(row=target_row, column=col_idx).value = round(change, 4)
    target_ws.cell(row=target_row, column=col_idx).number_format = "0.00%"
    print(f"补齐 {target_ws.title} {target_date} 涨幅 {round(change*100,2)}%")

# 先定位昨天的sheet和昨天行及价格
ws_yesterday = get_sheet_by_date(yesterday)
row_yesterday, price_yesterday, header_yesterday = get_date_row_and_close(ws_yesterday, yesterday)
if row_yesterday is None or price_yesterday is None:
    raise ValueError(f"基准日 {yesterday} 数据缺失")

# 确保昨天sheet有涨幅列
cols_idx_yesterday = ensure_columns(ws_yesterday, header_yesterday, ["3D Forward Change", "7D Forward Change"])

# 处理3天前数据
ws_3ago = get_sheet_by_date(day_3ago)
row_3ago, price_3ago, header_3ago = get_date_row_and_close(ws_3ago, day_3ago)
cols_idx_3ago = ensure_columns(ws_3ago, header_3ago, ["3D Forward Change"])
compute_and_write(price_yesterday, ws_3ago, row_3ago, cols_idx_3ago["3D Forward Change"], day_3ago)

# 处理7天前数据
ws_7ago = get_sheet_by_date(day_7ago)
row_7ago, price_7ago, header_7ago = get_date_row_and_close(ws_7ago, day_7ago)
cols_idx_7ago = ensure_columns(ws_7ago, header_7ago, ["7D Forward Change"])
compute_and_write(price_yesterday, ws_7ago, row_7ago, cols_idx_7ago["7D Forward Change"], day_7ago)

wb.save(file_name)
print("跨月动态补齐完成")
