import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

headers = {"User-Agent": "Mozilla/5.0"}

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

def get_price(ticker):
    url = f"https://finance.yahoo.com/quote/{ticker}"
    res = requests.get(url, headers=headers)
    soup = BeautifulSoup(res.text, "html.parser")

    try:
        # 🔥 ここが修正ポイント（現在値だけ取る）
        price = soup.select_one('fin-streamer[data-field="regularMarketPrice"]').text
    except:
        price = "取得失敗"

    return price

data = []

for name, ticker in targets.items():
    price = get_price(ticker)
    print(name, price)
    data.append([name, ticker, price])
    time.sleep(1)

df = pd.DataFrame(data, columns=["項目", "ティッカー", "価格"])

df.to_excel("market_data.xlsx", index=False)

print("完了")