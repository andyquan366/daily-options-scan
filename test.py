import yfinance as yf
from datetime import datetime
import pytz

tz = pytz.timezone("America/New_York")
now = datetime.now(tz).strftime('%Y-%m-%d %H:%M')
mag7 = ['NVDA', 'AAPL', 'MSFT', 'AMZN', 'GOOGL', 'META', 'TSLA']

print(f'【{now}】MAG7 主力期权IV/成交/OI')
print('-'*120)
for ticker in mag7:
    try:
        stock = yf.Ticker(ticker)
        # 获取现价
        price = None
        try:
            price = stock.history(period="1d")["Close"].iloc[-1]
        except Exception:
            price = None
        expiry_list = stock.options
        if not expiry_list:
            print(f'{ticker:<6} 无可用期权数据')
            continue
        # 取最近一期到期
        expiry = sorted(expiry_list)[0]
        chain = stock.option_chain(expiry)
        for opt_type, df in [('Call', chain.calls), ('Put', chain.puts)]:
            if df.empty:
                continue
            main = df.sort_values('volume', ascending=False).iloc[0]
            iv = main['impliedVolatility']
            oi = main['openInterest']
            vol = main['volume']
            strike = main['strike']
            csym = main['contractSymbol']
            # 异常判定
            iv_pct = round(iv * 100, 2) if iv and iv > 0 else None
            iv_tag = ""
            if iv_pct is None or iv_pct < 3 or iv_pct > 200:
                iv_tag = "⚠️异常IV"
            oi_tag = "" if oi and oi > 0 else "⚠️OI=0"
            print(f"{ticker:<6} {opt_type:<4}  Price: {price:<8.2f}  IV: {iv_pct:<7} {iv_tag:<6}  OI: {oi:<7} {oi_tag:<7}  Vol: {vol:<7}  Strike: {strike:<7}  Exp: {expiry}  [{csym}]")
    except Exception as e:
        print(f"{ticker} 发生异常：{e}")
print('-'*120)
