from openpyxl import load_workbook

file_name = "option_activity_log.xlsx"
sheet_name = "2025-06"

wb = load_workbook(file_name)
print(f"所有工作表名: {wb.sheetnames}")

if sheet_name not in wb.sheetnames:
    print(f"工作表 {sheet_name} 不存在")
    exit()

ws = wb[sheet_name]
header = [cell.value for cell in ws[1]]
print(f"工作表 {sheet_name} 第一行列标题: {header}")
