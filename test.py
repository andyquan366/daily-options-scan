import requests

tokens = {
    "PYTH-CAD": "pyth-network",
    "ONDO-CAD": "ondo-finance",
    "ENA-CAD": "ethena",
    "RENDER-CAD": "render-token",
    "JUP-CAD": "jupiter-exchange-solana",
    "UNI-CAD": "uniswap",
    "SOL-CAD": "solana",
    "LINK-CAD": "chainlink"
}

tickers = [
    "PYTH-CAD",
    "ONDO-CAD",
    "ENA-CAD",
    "RENDER-CAD",
    "JUP-CAD",
    "UNI-CAD",
    "SOL-CAD",
    "LINK-CAD"
]

def fetch_batch(tickers):
    url = "https://api.coingecko.com/api/v3/simple/price"
    ids = ",".join([tokens[t] for t in tickers])
    params = {"ids": ids, "vs_currencies": "cad"}
    data = requests.get(url, params=params, timeout=10).json()
    print("返回数据:", data)

    prices = []
    for t in tickers:
        coingecko_id = tokens[t]
        try:
            prices.append(round(data[coingecko_id]["cad"], 6))
        except Exception as e:
            prices.append(None)
            print(f"{t}: 查不到 ({e})")
    return prices

if __name__ == "__main__":
    results = fetch_batch(tickers)
    for t, v in zip(tickers, results):
        print(f"{t}: {v}")
