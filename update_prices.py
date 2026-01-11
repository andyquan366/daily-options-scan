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
    "RENDER-CAD",
    "UNI-CAD",
    "UMA-CAD",
    "ENA-CAD",
    "JUP-CAD",
]

def fetch_prices(tickers):
    prices = []

    # Yahoo Finance crypto 映射（统一用 USD）
    yahoo_crypto_map = {
        "SOL-CAD": "SOL-USD",
        "ONDO-CAD": "ONDO-USD",
        "SUI-CAD": "SUI-USD",
        "LINK-CAD": "LINK-USD",
        "PYTH-CAD": "PYTH-USD",
        "RENDER-CAD": "RENDER-USD",
        "UNI-CAD": "UNI-USD",
        "UMA-CAD": "UMA-USD",
        "ENA-CAD": "ENA-USD",
        "JUP-CAD": "JUP-USD"
    }

    # 1️⃣ 先取 USD → CAD 汇率（Yahoo）
    fx = yf.Ticker("USDCAD=X")
    fx_hist = fx.history(period="1d")

    if fx_hist.empty:
        raise RuntimeError("无法从 Yahoo 获取 USD→CAD 汇率")

    usd_to_cad = fx_hist["Close"].iloc[-1]

    # 2️⃣ 遍历所有 ticker
    for ticker in tickers:
        try:
            # === 加密货币：Yahoo Finance（USD → CAD）===
            if ticker in yahoo_crypto_map:
                yahoo_ticker = yahoo_crypto_map[ticker]

                t = yf.Ticker(yahoo_ticker)
                hist = t.history(period="1d")

                if hist.empty:
                    print(f"{ticker}: Yahoo USD price not available")
                    prices.append(None)
                    continue

                usd_price = hist["Close"].iloc[-1]
                cad_price = usd_price * usd_to_cad

                prices.append(round(cad_price, 6))
                continue

            # === 非加密资产（ETF / 股票）：Yahoo 原样 ===
            t = yf.Ticker(ticker)
            hist = t.history(period="1d")

            if hist.empty:
                prices.append(None)
            else:
                prices.append(round(hist["Close"].iloc[-1], 2))

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

    ranges = ["'ETF'!E21:E22", "'ETF'!E41:E45", "'ETF'!E49", "'ETF'!E56:E65"]
    values_list = [
        [[prices[0]], [prices[1]]],
        [[prices[2]], [prices[3]], [prices[4]], [prices[5]], [prices[6]]],
        [[prices[7]]],
        [[prices[8]], [prices[9]], [prices[10]], [prices[11]], [prices[12]], [prices[13]], [prices[14]], [prices[15]], [prices[16]], [prices[17]]]
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