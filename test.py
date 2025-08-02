import yfinance as yf

# 查询比特币和以太坊价格
tickers = ["BTC-CAD", "ETH-CAD"]
for ticker in tickers:
    try:
        data = yf.Ticker(ticker).history(period="1d")
        price = data['Close'].iloc[-1]
        print(f"{ticker}: {price}")
    except Exception as e:
        print(f"Error fetching {ticker}: {e}")
