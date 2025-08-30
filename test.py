import requests

def test_uni_price():
    url = "https://api.coingecko.com/api/v3/simple/price"
    params = {"ids": "uniswap", "vs_currencies": "cad"}
    try:
        data = requests.get(url, params=params, timeout=10).json()
        print("返回数据:", data)
        if "uniswap" in data and "cad" in data["uniswap"]:
            price = data["uniswap"]["cad"]
            print(f"UNI-CAD 当前价格: {price}")
        else:
            print("UNI-CAD: CoinGecko 没返回价格")
    except Exception as e:
        print("请求出错:", e)

if __name__ == "__main__":
    test_uni_price()
