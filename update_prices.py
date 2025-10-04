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
    "YPLT.NE",
    "YAMD.NE",
    "SOL-CAD",
    "ONDO-CAD",
    "ORDER-CAD",
    "PEAQ-CAD",
    "SUI-CAD",
    "LINK-CAD",
    "PYTH-CAD",
    "ENA-CAD",
    "JUP-CAD",
    "RENDER-CAD",
    "UNI-CAD",
    "UMA-CAD"
]

def fetch_prices(tickers):
    prices = []

    # 先一次性批量获取 10 个 CoinGecko 代币
    coingecko_map = {
        "ONDO-CAD": "ondo-finance",
        "ORDER-CAD": "orderly-network",
        "PEAQ-CAD": "peaq-2",
        "SUI-CAD": "sui",
        "PYTH-CAD": "pyth-network",
        "ENA-CAD": "ethena",
        "JUP-CAD": "jupiter-exchange-solana",
        "RENDER-CAD": "render-token",
        "UNI-CAD": "uniswap",
        "UMA-CAD": "uma"
    }
    ids = ",".join(coingecko_map.values())
    url = "https://api.coingecko.com/api/v3/simple/price"
    params = {"ids": ids, "vs_currencies": "cad"}
    try:
        cg_data = requests.get(url, params=params, timeout=10).json()
    except Exception as e:
        print("批量获取 CoinGecko 出错:", e)
        cg_data = {}

    # 遍历 tickers，按顺序塞进 prices
    for ticker in tickers:
        try:
            if ticker in coingecko_map:
                coin_id = coingecko_map[ticker]
                price = cg_data.get(coin_id, {}).get("cad")
                if price is not None:
                    # 每个币不同精度
                    if ticker == "ORDER-CAD":
                        prices.append(round(price, 6))
                    elif ticker == "PEAQ-CAD":
                        prices.append(round(price, 6))
                    elif ticker == "PYTH-CAD":
                        prices.append(round(price, 6))
                    elif ticker == "ENA-CAD":
                        prices.append(round(price, 6))
                    elif ticker == "JUP-CAD":
                        prices.append(round(price, 6))

                    else:
                        prices.append(round(price, 2))
                else:
                    print(f"{ticker}: CoinGecko 没返回数据")
                    prices.append(None)
                continue  # 跳过 yfinance

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

    ranges = ["'ETF'!E19:E20", "'ETF'!E23:E27", "'ETF'!E43:E54"]
    values_list = [
        [[prices[0]], [prices[1]]],
        [[prices[2]], [prices[3]], [prices[4]], [prices[5]], [prices[6]]],
        [[prices[7]], [prices[8]], [prices[9]], [prices[10]], [prices[11]], [prices[12]], [prices[13]], [prices[14]], [prices[15]], [prices[16]], [prices[17]], [prices[18]]]
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