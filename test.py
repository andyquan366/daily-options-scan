# test_ondo_jup.py
import requests

def get_ondo_cad():
    """用 CoinGecko simple/price 获取 ONDO → CAD"""
    url = "https://api.coingecko.com/api/v3/simple/price"
    params = {"ids": "ondo-finance", "vs_currencies": "usd,cad"}
    data = requests.get(url, params=params, timeout=10).json()
    print("ONDO 调试返回:", data)
    usd = data["ondo-finance"]["usd"]
    cad = data["ondo-finance"]["cad"]
    print(f"当前 1 ONDO ≈ {usd} USD / {cad} CAD")

def get_jup_cad():
    """用 Binance JUP/USDT 再换算 CAD"""
    # Binance JUP/USDT
    url = "https://api.binance.com/api/v3/ticker/price"
    params = {"symbol": "JUPUSDT"}
    jup_usd = float(requests.get(url, params=params, timeout=10).json()["price"])

    # 美元 → 加元汇率 (用 open.er-api.com)
    fx_url = "https://open.er-api.com/v6/latest/USD"
    fx = requests.get(fx_url, timeout=10).json()
    cad_rate = fx["rates"]["CAD"]

    jup_cad = jup_usd * cad_rate
    print(f"当前 1 JUP ≈ {jup_usd:.4f} USD / {jup_cad:.4f} CAD")

if __name__ == "__main__":
    get_ondo_cad()
    print("-" * 50)
    get_jup_cad()
