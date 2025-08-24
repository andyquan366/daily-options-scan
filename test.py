# test_jup.py
import requests

def get_jup_cad():
    """用 Binance JUP/USDT 再换算 CAD"""
    # Binance JUP/USDT
    url = "https://api.binance.com/api/v3/ticker/price"
    params = {"symbol": "JUPUSDT"}
    jup_usd = float(requests.get(url, params=params, timeout=10).json()["price"])

    # 美元 → 加元汇率
    fx_url = "https://open.er-api.com/v6/latest/USD"
    fx = requests.get(fx_url, timeout=10).json()
    cad_rate = fx["rates"]["CAD"]

    jup_cad = jup_usd * cad_rate
    print(f"{jup_cad:.4f}")  # 只输出 JUP-CAD 价格

if __name__ == "__main__":
    get_jup_cad()
