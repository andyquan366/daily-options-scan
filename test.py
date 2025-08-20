# temp_ondo_cad.py
import requests

def get_ondo_cad():
    url = "https://api.coingecko.com/api/v3/simple/price"
    params = {
        "ids": "ondo-finance",   # 正确的 CoinGecko ID
        "vs_currencies": "cad"
    }
    try:
        response = requests.get(url, params=params, timeout=10)
        data = response.json()
        print("调试返回:", data)
        if "ondo-finance" in data and "cad" in data["ondo-finance"]:
            price = data["ondo-finance"]["cad"]
            print(f"当前 1 ONDO ≈ {price:.4f} CAD")
        else:
            print("返回结果里没有 ondo-finance，可能被限流")
    except Exception as e:
        print(f"获取 ONDO-CAD 失败: {e}")

if __name__ == "__main__":
    get_ondo_cad()
