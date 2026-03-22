import yfinance as yf
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
import os

# ===== メール設定（GitHub Secrets）=====
EMAIL = os.getenv("EMAIL_USER")
PASSWORD = os.getenv("EMAIL_PASS")

print("EMAIL:", EMAIL)
print("PASSWORD:", PASSWORD)

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

    # 為替
    ("ドル円", "JPY=X"),
    ("ユーロ円", "EURJPY=X"),
    ("ユーロドル", "EURUSD=X"),

    # 金利
    ("米国2年国債", "^IRX"),
    ("米国10年国債", "^TNX"),

    # 指数
    ("DAX", "^GDAXI"),
    ("上海総合", "000001.SS"),
    ("SENSEX", "^BSESN"),
    ("ボベスパ", "^BVSP"),

    # 商品
    ("WTI原油", "CL=F"),
    ("金先物", "GC=F"),
    ("天然ガス", "NG=F"),
    ("銅先物", "HG=F"),
]

data = []

for name, ticker in targets:
    try:
        ticker_obj = yf.Ticker(ticker)
        hist = ticker_obj.history(period="5d")

        if hist.empty:
            price = "取得失敗"
        else:
            price = round(hist["Close"].dropna().iloc[-1], 2)

    except:
        price = "取得失敗"

    print(name, ticker, price)
    data.append([name, ticker, price])

# ===== DataFrame =====
df = pd.DataFrame(data, columns=["項目", "ティッカー", "価格"])

# ===== 日付 =====
today = datetime.now().strftime("%Y-%m-%d %H:%M")
df["日付"] = today
df = df[["日付", "項目", "ティッカー", "価格"]]

# ===== Excel出力 =====
df.to_excel("market_data.xlsx", index=False)

# ===== メール本文 =====
body = df.to_string(index=False)

msg = MIMEText(body)
msg["Subject"] = "📊 マーケットデータ"
msg["From"] = EMAIL
msg["To"] = EMAIL

# ===== メール送信 =====
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(EMAIL, PASSWORD)
    smtp.send_message(msg)

print("メール送信完了")
