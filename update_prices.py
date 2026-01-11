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

    # 所有 *-CAD 的加密货币
    # 统一规则：先取 USD → 再换算 CAD
    # ❌ 不直接使用 CoinGecko 的 CAD
    # ❌ 不兜底、不写死、不猜价格
    coingecko_usd_map = {
        "SOL-CAD": "solana",
        "ONDO-CAD": "ondo-finance",
        "SUI-CAD": "sui",
        "LINK-CAD": "chainlink",
        "PYTH-CAD": "pyth-network",
        "RENDER-CAD": "render-token",
        "UNI-CAD": "uniswap",
        "UMA-CAD": "uma",
        "ENA-CAD": "ethena",
        "JUP-CAD": "jupiter-exchange-solana"
    }

    # 1️⃣ 获取 USD → CAD 汇率（严格模式）
    # 任意失败 = 整个任务失败（避免写假数据）
    fx_url = "https://api.coingecko.com/api/v3/simple/price"
    fx_params = {
        "ids": "usd",
        "vs_currencies": "cad"
    }

    fx_data = requests.get(
        fx_url,
        params=fx_params,
        headers=headers,
        timeout=10
    ).json()

    usd_to_cad = fx_data.get("usd", {}).get("cad")

    if usd_to_cad is None:
        raise RuntimeError("无法获取 USD→CAD 汇率，已中止")

    # 2️⃣ 遍历所有 ticker
    for ticker in tickers:
        try:
            # === 加密货币：CoinGecko（USD → CAD）===
            if ticker in coingecko_usd_map:
                coin_id = coingecko_usd_map[ticker]

                price_url = "https://api.coingecko.com/api/v3/simple/price"
                price_params = {
                    "ids": coin_id,
                    "vs_currencies": "usd"
                }

                data = requests.get(
                    price_url,
                    params=price_params,
                    headers=headers,
                    timeout=10
                ).json()


                time.sleep(1)  # 防止 CoinGecko 限流

                usd_price = data.get(coin_id, {}).get("usd")

                # USD 价格缺失 → 不写入
                if usd_price is None:
                    print(f"{ticker}: USD price not available")
                    prices.append(None)
                    continue

                # USD → CAD
                cad_price = usd_price * usd_to_cad

                # crypto 使用高精度
                prices.append(round(cad_price, 6))
                continue

            # === 非加密资产：yfinance ===
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