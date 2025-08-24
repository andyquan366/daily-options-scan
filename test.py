# test_jup_cad.py
import requests

def get_jup_cad():
    """用 CoinGecko simple/price 获取 JUP (Solana Jupiter) → CAD"""
    url = "https://api.coingecko.com/api/v3/simple/price"
    params = {"ids": "jupiter-exchange-solana", "vs_currencies": "usd,cad"}
    data = requests.get(url, params=params, timeout=10).json()
    print("CoinGecko 返回:", data)
    usd = data["jupiter-exchange-solana"]["usd"]
    cad = data["jupiter-exchange-solana"]["cad"]
    print(f"当前 1 JUP ≈ {usd} USD / {cad} CAD")

if __name__ == "__main__":
    get_jup_cad()
