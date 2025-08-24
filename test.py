import requests

def get_price(token_id, symbol):
    url = "https://api.coingecko.com/api/v3/simple/price"
    params = {"ids": token_id, "vs_currencies": "cad"}
    headers = {"User-Agent": "Mozilla/5.0"}  # 防止被屏蔽

    try:
        response = requests.get(url, params=params, headers=headers, timeout=10)
        data = response.json()
        print(f"{symbol} 调试返回:", data)
        if token_id in data and "cad" in data[token_id]:
            price = data[token_id]["cad"]
            print(f"当前 1 {symbol} ≈ {price:.4f} CAD")
        else:
            print(f"{symbol} 没有拿到价格，可能是 API 限流或 ID 错误")
    except Exception as e:
        print(f"获取 {symbol}-CAD 失败: {e}")

if __name__ == "__main__":
    # ONDO (正确ID: ondo-finance)
    get_price("ondo-finance", "ONDO")
    print("-" * 50)
    # JUP (CoinGecko 正确ID 可能是 jupiter，但你之前用 jupiter-exchange 返回空)
    # 建议测试两种：jupiter-exchange 和 jupiter
    get_price("jupiter-exchange", "JUP")
    get_price("jupiter", "JUP(alt)")
