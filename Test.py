import yfinance as yf

# 10只热门股票
tickers = ["AAPL", "MSFT", "NVDA", "TSLA", "AMZN", "META", "GOOGL", "AMD", "NFLX", "ENPH"]

for ticker in tickers:
    stock = yf.Ticker(ticker)
    # 获取所有可用期权到期日
    expirations = stock.options
    if not expirations:
        print(f"{ticker}: 无可用期权到期日")
        continue
    expiry = expirations[0]  # 取最近一期
    opt_chain = stock.option_chain(expiry)
    print(f"\n=== {ticker} 期权到期日: {expiry} ===")

    # Call
    print("Call OI (前5):")
    calls = opt_chain.calls.sort_values("openInterest", ascending=False).head(5)
    for _, row in calls.iterrows():
        print(f"  行权价: {row['strike']}, OI: {row['openInterest']}")

    # Put
    print("Put OI (前5):")
    puts = opt_chain.puts.sort_values("openInterest", ascending=False).head(5)
    for _, row in puts.iterrows():
        print(f"  行权价: {row['strike']}, OI: {row['openInterest']}")

