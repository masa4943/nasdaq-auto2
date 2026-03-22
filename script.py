import yfinance as yf
import pandas as pd
from datetime import datetime

# ===== 取得対象 =====
targets = [
    ("NASDAQ", "^IXIC"),
    ("ダウ平均", "^DJI"),
    ("S&P500", "^GSPC"),
    ("エヌビディア", "NVDA"),
    ("アップル", "AAPL"),
    ("マイクロソフト", "MSFT"),
    ("アルファベット", "GOOGL"),
    ("メタ", "META"),
    ("アマゾン", "AMZN"),
    ("テスラ", "TSLA"),
    ("日経平均", "^N225"),
    
    # ===== 日本・為替 =====
    ("TOPIX", "1306.T"),  # ETF代替
    ("ドル円", "JPY=X"),
    ("ユーロ円", "EURJPY=X"),   # ★追加
    ("ユーロドル", "EURUSD=X"), # ★追加

    # ===== 金利 =====
    ("米国2年国債", "^IRX"),
    ("米国10年国債", "^TNX"),

    # ===== 株価指数 =====
    ("DAX", "^GDAXI"),
    ("上海総合", "000001.SS"),
    ("SENSEX", "^BSESN"),
    ("ボベスパ", "^BVSP"),

    # ===== 商品（安定版）=====
    ("WTI原油", "CL=F"),
    ("金先物", "GC=F"),
    ("天然ガス", "NG=F"),
    ("銅先物", "HG=F"),
]

data = []

for name, ticker in targets:
    try:
        ticker_obj = yf.Ticker(ticker)

        # 🔥 安定化（5日分取得して最新使う）
        hist = ticker_obj.history(period="5d")

        if hist.empty:
            price = "取得失敗"
        else:
            price = round(hist["Close"].dropna().iloc[-1], 2)

    except Exception as e:
        price = "取得失敗"

    print(name, ticker, price)
    data.append([name, ticker, price])

# ===== DataFrame =====
df = pd.DataFrame(data, columns=["項目", "ティッカー", "価格"])

# ===== 日付 =====
today = datetime.now().strftime("%Y-%m-%d")
df["日付"] = today

# ===== 並び順 =====
df = df[["日付", "項目", "ティッカー", "価格"]]

# ===== Excel出力 =====
df.to_excel("market_data.xlsx", index=False)

print("Excel出力完了")
