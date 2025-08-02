import yfinance as yf

# SOL-CAD 在 Yahoo Finance 的ticker就是 'SOL-CAD'
ticker = 'SOL-CAD'

data = yf.Ticker(ticker)
price = data.history(period="1d")["Close"][-1]
print(f"SOL-CAD 最新价格：{price}")
