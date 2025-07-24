import yfinance as yf

tickers = [
    "", "COGT", "WVE", "DTSS", "SATL", "AMPH", "ULBI", "NPAC", "KITT", "WCT",
    "GLTO", "ASTS", "HCTI", "KPLT", "NTIC", "SIMA", "NMFCZ"
]

success = []
failed = []

for t in tickers:
    t = t.strip()
    if not t:
        failed.append(t)
        print(f"[空ticker] 跳过")
        continue
    try:
        df = yf.Ticker(t).history(period="7d")
        if df.empty:
            failed.append(t)
            print(f"[FAIL] {t}: 无法获取行情")
        else:
            success.append(t)
            print(f"[OK]   {t}: 行情数据行数 {len(df)}")
    except Exception as e:
        failed.append(t)
        print(f"[ERROR] {t}: {e}")

print("\n====== 拉取成功的：======\n", success)
print("\n====== 拉取失败的：======\n", failed)
