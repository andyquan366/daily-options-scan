import yfinance as yf

tickers = [
    "ONDO-CAD",
    "PYTH-CAD",
    "JUP-CAD",
    "UNI-CAD",
    "ENA-CAD",
    "RENDER-CAD"
]

for t in tickers:
    try:
        data = yf.Ticker(t)
        price = data.history(period="1d")["Close"].iloc[-1]
        print(f"{t}: {price}")
    except Exception as e:
        print(f"{t}: 查不到 ({e})")
