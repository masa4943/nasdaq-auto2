import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

headers = {"User-Agent": "Mozilla/5.0"}

# ===== 取得対象（Yahoo Financeのティッカー）=====
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
    "日本10年国債": "^TNX",      # ※近似（米10年と違うので注意）
    "米国2年国債": "^IRX",
    "米国10年国債": "^TNX",
    "ドイツ10年国債": "^DE10Y",
    "DAX": "^GDAXI",
    "上海総合": "000001.SS",
    "SENSEX": "^BSESN",
    "ボベスパ": "^BVSP",
    "WTI原油": "CL=F",
    "金先物": "GC=F",
    "天然ガス": "NG=F",
    "銅先物": "HG=F"
}

# ===== 価格取得関数 =====
def get_price(ticker):
    url = f"https://finance.yahoo.com/quote/{ticker}"
    res = requests.get(url, headers=headers)
    soup = BeautifulSoup(res.text, "html.parser")

    try:
        price = soup.find("fin-streamer").text
    except:
        price = "取得失敗"

    return price

# ===== データ取得 =====
data = []

for name, ticker in targets.items():
    price = get_price(ticker)
    print(name, price)
    data.append([name, ticker, price])
    time.sleep(1)  # アクセス制限対策

# ===== DataFrame化 =====
df = pd.DataFrame(data, columns=["項目", "ティッカー", "価格"])

# ===== Excel出力 =====
df.to_excel("market_data.xlsx", index=False)

print("完了")