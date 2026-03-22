import yfinance as yf
import pandas as pd
from datetime import datetime

# ===== 取得対象 =====
targets = {
    "NASDAQ": "^IXIC",
    "ダウ平均": "^DJI",
    "S&P500": "^GSPC",
    "エヌビディア": "NVDA",
    "アップル": "AAPL",
    "マイクロソフト": "MSFT",
    "アルファベット": "GOOGL",
    "メタ": "META",
    "アマゾン": "AMZN",
    "テスラ": "TSLA",
    "日経平均": "^N225",
    "TOPIX": "^TOPX",
    "ドル円": "JPY=X",
    "ドルインデックス": "DX-Y.NYB",
    "米国2年国債": "^IRX",
    "米国10年国債": "^TNX",
    "DAX": "^GDAXI",
    "上海総合": "000001.SS",
    "SENSEX": "^BSESN",
    "ボベスパ": "^BVSP",
    "WTI原油": "CL=F",
    "金先物": "GC=F",
    "天然ガス": "NG=F",
    "銅先物": "HG=F"
}

data = []

for name, ticker in targets.items():
    try:
        ticker_obj = yf.Ticker(ticker)
        hist = ticker_obj.history(period="1d")

        if hist.empty:
            price = "取得失敗"
        else:
            price = round(hist["Close"].iloc[-1], 2)

    except:
        price = "取得失敗"

    print(name, ticker, price)
    data.append([name, ticker, price])

# ===== DataFrame =====
df = pd.DataFrame(data, columns=["項目", "ティッカー", "価格"])

# ===== 日付追加（実務向け）=====
today = datetime.now().strftime("%Y-%m-%d")
df["日付"] = today

# ===== 列順整理 =====
df = df[["日付", "項目", "ティッカー", "価格"]]

# ===== Excel出力 =====
file_name = "market_data.xlsx"
df.to_excel(file_name, index=False)

print("Excel出力完了:", file_name)