import yfinance as yf
import time
import requests
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

tickers = [
    "HBIX.NE",
    "HBTE.NE",
    "YCON.NE",
    "YTSL.NE",
    "YPLT.NE",
    "YNVD.NE",
    "YAMD.NE",
    "HHIS-U.TO",
    "SOL-CAD",
    "ONDO-CAD",
    "SUI-CAD",
    "LINK-CAD",
    "PYTH-CAD",
    "ENA-CAD",
    "JUP-CAD"
]

def fetch_prices(tickers):
    prices = []

    # CoinGecko 对应关系（区分 USD / CAD）
    coingecko_map = {
        "ONDO-CAD": ("ondo-finance", "cad"),
        "SUI-CAD": ("sui", "cad"),
        "PYTH-CAD": ("pyth-network", "cad"),
        "ENA-CAD": ("ethena", "cad"),
        "JUP-CAD": ("jupiter-exchange-solana", "cad")
    }

    for ticker in tickers:
        try:
            if ticker in coingecko_map:
                coin_id, currency = coingecko_map[ticker]
                url = "https://api.coingecko.com/api/v3/simple/price"
                params = {"ids": coin_id, "vs_currencies": currency}
                data = requests.get(url, params=params, timeout=10).json()
                time.sleep(1)  # ✅ 新增这一行，避免被 CoinGecko 限流
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
    import math
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = 'credentials.json'

    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    SPREADSHEET_ID = '1Rfs87zMtB9hyhkRiW1UGnAuNeLjQEcb_-9yRtLjRATI'

    ranges = ["'ETF'!E17:E18", "'ETF'!E34:E38", "'ETF'!E43", "'ETF'!E49:E55"]
    values_list = [
        [[prices[0]], [prices[1]]],
        [[prices[2]], [prices[3]], [prices[4]], [prices[5]], [prices[6]]],
        [[prices[7]]],
        [[prices[8]], [prices[9]], [prices[10]], [prices[11]], [prices[12]], [prices[13]], [prices[14]]]
    ]

    for rng, vals in zip(ranges, values_list):
        # ✅ 清理 None / NaN
        clean_vals = [["" if (v is None or (isinstance(v, float) and math.isnan(v))) else v for v in row] for row in vals]

        body = {'values': clean_vals}
        result = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=rng,
            valueInputOption='RAW',
            body=body
        ).execute()
        print(f"{result.get('updatedCells')} cells updated in {rng}.")


if __name__ == "__main__":
    prices = fetch_prices(tickers)
    write_prices_to_sheet_split(prices)