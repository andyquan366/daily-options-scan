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
    "ORDER-USD",
    "PEAQ-USD",
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

    # CoinGecko 对应关系（区分 USD / CAD）
    coingecko_map = {
        "ONDO-CAD": ("ondo-finance", "cad"),
        "ORDER-USD": ("orderly-network", "usd"),
        "PEAQ-USD": ("peaq-2", "usd"),
        "SUI-CAD": ("sui", "cad"),
        "PYTH-CAD": ("pyth-network", "cad"),
        "ENA-CAD": ("ethena", "cad"),
        "JUP-CAD": ("jupiter-exchange-solana", "cad"),
        "RENDER-CAD": ("render-token", "cad"),
        "UNI-CAD": ("uniswap", "cad"),
        "UMA-CAD": ("uma", "cad")
    }

    for ticker in tickers:
        try:
            if ticker in coingecko_map:
                coin_id, currency = coingecko_map[ticker]
                url = "https://api.coingecko.com/api/v3/simple/price"
                params = {"ids": coin_id, "vs_currencies": currency}
                data = requests.get(url, params=params, timeout=10).json()
                price = data.get(coin_id, {}).get(currency)

                if price is not None:
                    # 精度统一控制
                    if ticker in ["ORDER-USD", "PEAQ-USD", "PYTH-CAD", "ENA-CAD", "JUP-CAD"]:
                        prices.append(round(price, 6))
                    else:
                        prices.append(round(price, 2))
                else:
                    print(f"{ticker}: CoinGecko 没返回数据")
                    prices.append(None)
                continue

            # 其他 ticker 用 yfinance
            t = yf.Ticker(ticker)
            data = t.history(period="1d")
            prices.append(round(data["Close"].iloc[-1], 2))

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