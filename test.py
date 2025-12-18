import yfinance as yf
import pandas as pd
import numpy as np

# --------------------------
# 回测参数
# --------------------------
START = "2023-01-01"
END = "2024-12-31"

RISK_BUDGET = 10000          # 每次最大亏损
SPREAD_WIDTH = 5             # put spread 宽度（点）
MAX_LOSS_PER_CONTRACT = SPREAD_WIDTH * 100
TAKE_PROFIT = 0.50           # 50% 止盈
HOLD_DAYS = 30               # 持仓天数

# --------------------------
# 下载 QQQ 日 K
# --------------------------
qqq = yf.download("QQQ", START, END)

# 修复 MultiIndex 列名问题
if isinstance(qqq.columns, pd.MultiIndex):
    qqq.columns = qqq.columns.get_level_values(0)

# 确保 Close 是 float
qqq["Close"] = pd.to_numeric(qqq["Close"])

# --------------------------
# 计算回撤
# --------------------------
qqq["HighToDate"] = qqq["Close"].cummax()
qqq["Drawdown"] = (qqq["HighToDate"] - qqq["Close"]) / qqq["HighToDate"]

# --------------------------
# 回测主循环
# --------------------------
position = None
trades = []

for i in range(len(qqq)):
    today = qqq.index[i]
    price = qqq["Close"].iloc[i]
    dd = qqq["Drawdown"].iloc[i]

    # ======================================================
    # 开仓条件：从最高点下跌 >= 8%
    # ======================================================
    if position is None and dd >= 0.08:
        qty = RISK_BUDGET // MAX_LOSS_PER_CONTRACT
        credit = 1.00 * 100 * qty   # 模拟 credit = $1 per spread
        
        position = {
            "entry_date": today,
            "entry_price": price,
            "credit": credit,
            "qty": qty,
            "exit_date": today + pd.Timedelta(days=HOLD_DAYS)
        }

        print(f"[ENTER] {today.date()} credit={credit}, qty={qty}")

    # ======================================================
    # 持仓管理
    # ======================================================
    if position is not None:

        # 模拟 debit 随价格变化
        progress = (price - position["entry_price"]) / position["entry_price"]
        debit = max(0, position["credit"] * (1 - progress))
        profit = position["credit"] - debit

        # 止盈退出
        if profit >= position["credit"] * TAKE_PROFIT:
            print(f"[TAKE PROFIT] {today.date()} profit={profit:.2f}")
            trades.append(profit)
            position = None
            continue

        # 到期退出
        if today >= position["exit_date"]:
            print(f"[TIME EXIT] {today.date()} profit={profit:.2f}")
            trades.append(profit)
            position = None
            continue

# --------------------------
# 输出结果
# --------------------------
total_profit = sum(trades)

print("\n====================================")
print(f"Total trades: {len(trades)}")
print(f"Total profit: {total_profit:.2f}")
print("====================================")
