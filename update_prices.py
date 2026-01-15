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
    "PYTH-CAD",
    "ENA-CAD",
    "LINK-CAD",
    "ONDO-CAD",
    "RENDER-CAD",
    "JUP-CAD",
    "SUI-CAD",
    "UNI-CAD",
    "UMA-CAD",
]

def fetch_prices(tickers):
    prices = []

    # === CoinGecko：只用于 Yahoo 不可靠 / 不支持的币 ===
    # 使用 USD 价格，后续统一换 CAD
    coingecko_usd_map = {
        "SUI-CAD": "sui",
        "UNI-CAD": "uniswap",
        "JUP-CAD": "jupiter-exchange-solana"
    }

    # === Yahoo Finance crypto（排除 SUI / UNI / JUP）===
    yahoo_crypto_map = {
        "SOL-CAD": "SOL-USD",
        "ONDO-CAD": "ONDO-USD",
        "LINK-CAD": "LINK-USD",
        "PYTH-CAD": "PYTH-USD",
        "RENDER-CAD": "RENDER-USD",
        "UMA-CAD": "UMA-USD",
        "ENA-CAD": "ENA-USD"
    }

    # === USD → CAD（Yahoo FX）===
    fx = yf.Ticker("USDCAD=X")
    fx_hist = fx.history(period="1d")

    if fx_hist.empty:
        raise RuntimeError("无法从 Yahoo 获取 USD→CAD 汇率")

    usd_to_cad = fx_hist["Close"].iloc[-1]

    # === 遍历 ticker ===
    for ticker in tickers:
        try:
            # ---------- CoinGecko：SUI / UNI / JUP ----------
            if ticker in coingecko_usd_map:
                coin_id = coingecko_usd_map[ticker]

                url = "https://api.coingecko.com/api/v3/simple/price"
                params = {
                    "ids": coin_id,
                    "vs_currencies": "usd"
                }

                data = requests.get(url, params=params, timeout=10).json()
                time.sleep(1)

                usd_price = data.get(coin_id, {}).get("usd")

                if usd_price is None:
                    print(f"{ticker}: CoinGecko USD price not available")
                    prices.append(None)
                    continue

                prices.append(round(usd_price * usd_to_cad, 6))
                continue

            # ---------- Yahoo crypto ----------
            if ticker in yahoo_crypto_map:
                yahoo_ticker = yahoo_crypto_map[ticker]

                t = yf.Ticker(yahoo_ticker)
                hist = t.history(period="1d")

                if hist.empty:
                    print(f"{ticker}: Yahoo USD price not available")
                    prices.append(None)
                    continue

                usd_price = hist["Close"].iloc[-1]
                prices.append(round(usd_price * usd_to_cad, 6))
                continue

            # ---------- ETF / 股票：Yahoo ----------
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

    ranges = ["'ETF'!E19:E20", "'ETF'!E40:E44", "'ETF'!E48", "'ETF'!E55:E64"]
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