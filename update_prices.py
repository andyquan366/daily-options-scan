import yfinance as yf
import requests
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

tickers = [
    "HBTE.NE",
    "HBIX.NE",
    "YTSL.NE",
    "YNVD.NE",
    "YCON.NE",
    "YAMD.NE",
    "YPLT.NE",
    "SOL-CAD",
    "ONDO-CAD",
    "POL-CAD",
    "LINK-CAD",
    "JUP-CAD"
]

def fetch_prices(tickers):
    prices = []
    for ticker in tickers:
        try:
            if ticker == "ONDO-CAD":
                # 用 CoinGecko API 获取 ONDO → CAD
                url = "https://api.coingecko.com/api/v3/simple/price"
                params = {"ids": "ondo-finance", "vs_currencies": "cad"}
                data = requests.get(url, params=params, timeout=10).json()
                price = data["ondo-finance"]["cad"]
                prices.append(round(price, 2))
                continue  # 跳过 yfinance

            if ticker == "POL-CAD":
                # 用 CoinGecko API 获取 Polygon (POL) → CAD
                url = "https://api.coingecko.com/api/v3/simple/price"
                params = {"ids": "polygon-ecosystem-token", "vs_currencies": "cad"}
                data = requests.get(url, params=params, timeout=10).json()
                price = data["polygon-ecosystem-token"]["cad"]
                prices.append(round(price, 6))
                continue  # 跳过 yfinance

            if ticker == "LINK-CAD":
                # 用 CoinGecko API 获取 Chainlink (LINK) → CAD
                url = "https://api.coingecko.com/api/v3/simple/price"
                params = {"ids": "chainlink", "vs_currencies": "cad"}
                data = requests.get(url, params=params, timeout=10).json()
                price = data["chainlink"]["cad"]
                prices.append(round(price, 2))
                continue  # 跳过 yfinance

            if ticker == "JUP-CAD":
                try:
                    # Binance JUP/USDT
                    url = "https://api.binance.com/api/v3/ticker/price"
                    params = {"symbol": "JUPUSDT"}
                    resp = requests.get(url, params=params, timeout=10).json()
                    print("Binance 返回:", resp)
                    jup_usd = float(resp["price"])

                    # 美元 → 加元汇率 (优先 open.er-api)
                    cad_rate = None
                    try:
                        fx_url = "https://open.er-api.com/v6/latest/USD"
                        fx = requests.get(fx_url, timeout=10).json()
                        print("汇率返回:", fx)
                        cad_rate = fx.get("rates", {}).get("CAD")
                    except Exception as e:
                        print("open.er-api 获取汇率失败:", e)

                    # 备用: Yahoo Finance USDCAD=X
                    if cad_rate is None:
                        try:
                            t = yf.Ticker("USDCAD=X")
                            data = t.history(period="1d")
                            cad_rate = float(data['Close'].iloc[-1])
                            print("Yahoo Finance 汇率:", cad_rate)
                        except Exception as e:
                            print("Yahoo Finance 获取汇率失败:", e)

                    if cad_rate is None:
                        raise ValueError("未找到 CAD 汇率")

                    price = jup_usd * cad_rate
                    print(f"JUP 计算: {jup_usd} USD * {cad_rate} = {price} CAD")
                    prices.append(round(price, 6))
                except Exception as e:
                    print(f"获取 JUP-CAD 失败: {e}")
                    prices.append(None)
                continue


            # 其他 ticker 默认走 yfinance
            t = yf.Ticker(ticker)
            data = t.history(period="1d")
            price = data['Close'].iloc[-1]
            prices.append(round(price, 2))

        except Exception as e:
            print(f"Error fetching {ticker}: {e}")
            prices.append(None)

    return prices


def write_prices_to_sheet_split(prices):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = 'credentials.json'  # 这里用workflow的凭证文件名

    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    SPREADSHEET_ID = '1Rfs87zMtB9hyhkRiW1UGnAuNeLjQEcb_-9yRtLjRATI'

    ranges = ["'ETF'!F14:F15", "'ETF'!F18:F22", "'ETF'!F38:F42"]
    values_list = [
        [[prices[0]], [prices[1]]],
        [[prices[2]], [prices[3]], [prices[4]], [prices[5]], [prices[6]]],
        [[prices[7]],[prices[8]],[prices[9]],[prices[10]],[prices[11]]]
    ]

    for rng, vals in zip(ranges, values_list):
        body = {'values': vals}
        result = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=rng,
            valueInputOption='RAW', body=body).execute()
        print(f"{result.get('updatedCells')} cells updated in {rng}.")

if __name__ == "__main__":
    prices = fetch_prices(tickers)
    write_prices_to_sheet_split(prices)
