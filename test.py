import requests

url = "https://api.kraken.com/0/public/Depth"
params = {"pair": "SOLCAD", "count": 5}

resp = requests.get(url, params=params)
data = resp.json()

pair = list(data['result'].keys())[0]
bids = data['result'][pair]['bids']
asks = data['result'][pair]['asks']

print(f"Kraken SOL/CAD 最优买价(Bid): {bids[0][0]}, 数量: {bids[0][1]}")
print(f"Kraken SOL/CAD 最优卖价(Ask): {asks[0][0]}, 数量: {asks[0][1]}")
