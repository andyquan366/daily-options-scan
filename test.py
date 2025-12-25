import yfinance as yf
import pandas as pd

# 显示全部行，不省略
pd.set_option("display.max_rows", None)
pd.set_option("display.width", None)

# 下载 QQQ 数据（end 是开区间）
qqq = yf.download(
    "QQQ",
    start="2020-01-01",
    end="2020-05-01",
    progress=False
)

# 取收盘价，保留两位小数，倒序
close_prices = (
    qqq["Close"]
    .round(2)
    .sort_index(ascending=False)
    .to_numpy()   # 去掉日期索引
)

# 只打印收盘价
for price in close_prices:
    print(price)
