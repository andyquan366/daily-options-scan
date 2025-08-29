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
    "LINK-CAD",
    "ONDO-CAD",
    "JUP-CAD",
    "UNI-CAD"
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

            if ticker == "JUP-CAD":
                # 用 CoinGecko API 获取 JUP → CAD
                url = "https://api.coingecko.com/api/v3/simple/price"
                params = {"ids": "jupiter-exchange-solana", "vs_currencies": "cad"}
                data = requests.get(url, params=params, timeout=10).json()
                price = data["jupi1ter-exchange-solana"]["cad"]
                prices.append(round(price, 2))
                continue

            if ticker == "UNI-CAD":
                # 用 CoinGecko API 获取 Uniswap (UNI) → CAD
                url = "https://api.coingecko.com/api/v3/simple/price"
                params = {"ids": "uniswap", "vs_currencies": "cad"}
                data = requests.get(url, params=params, timeout=10).json()
                price = data["uniswap"]["cad"]
                prices.append(round(price, 2))
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

    ranges = ["'ETF'!F15:F16", "'ETF'!F19:F23", "'ETF'!F39:F43"]
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
