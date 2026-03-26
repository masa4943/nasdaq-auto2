import yfinance as yf
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
import os

# ===== メール設定 =====
EMAIL = os.getenv("EMAIL_USER")
PASSWORD = os.getenv("EMAIL_PASS")

# ===== 銘柄（カタカナ）=====
targets = [
    ("アップル", "AAPL"),
    ("マイクロソフト", "MSFT"),
    ("アルファベット", "GOOGL"),
    ("アマゾン", "AMZN"),
    ("エヌビディア", "NVDA"),
    ("メタ", "META"),
    ("テスラ", "TSLA"),

    ("ユナイテッドヘルス", "UNH"),
    ("ゴールドマン・サックス", "GS"),
    ("ホーム・デポ", "HD"),
    ("マクドナルド", "MCD"),
    ("ビザ", "V"),
    ("ジョンソン・エンド・ジョンソン", "JNJ"),
    ("プロクター・アンド・ギャンブル", "PG"),
    ("JPモルガン", "JPM"),
    ("シェブロン", "CVX"),
    ("メルク", "MRK"),
    ("コカ・コーラ", "KO"),
    ("シスコシステムズ", "CSCO"),
    ("IBM", "IBM"),
    ("アメリカン・エキスプレス", "AXP"),
    ("キャタピラー", "CAT"),
    ("ボーイング", "BA"),
    ("ウォルマート", "WMT"),
    ("ディズニー", "DIS"),
    ("ナイキ", "NKE"),
    ("セールスフォース", "CRM"),

    ("バークシャー・ハサウェイ", "BRK-B"),
    ("イーライリリー", "LLY"),
    ("エクソンモービル", "XOM"),
    ("アッヴィ", "ABBV"),
    ("ブロードコム", "AVGO"),

    ("ASML", "ASML"),
    ("コストコ", "COST"),
    ("ペプシコ", "PEP"),
    ("ネットフリックス", "NFLX"),
    ("アドビ", "ADBE"),
]

rows = []

for name, ticker in targets:
    try:
        t = yf.Ticker(ticker)
        hist = t.history(period="1y")

        if hist.empty:
            rows.append([name, ticker, "取得失敗", "-", "-", "-"])
            continue

        close = hist["Close"].dropna()

        latest = close.iloc[-1]
        prev = close.iloc[-2] if len(close) > 1 else latest
        month = close.iloc[-22] if len(close) > 22 else close.iloc[0]
        year = close.iloc[0]

        day_ret = (latest / prev - 1) * 100
        month_ret = (latest / month - 1) * 100
        year_ret = (latest / year - 1) * 100

        rows.append([
            name,
            ticker,
            round(latest, 2),
            round(day_ret, 2),
            round(month_ret, 2),
            round(year_ret, 2)
        ])

    except Exception as e:
        rows.append([name, ticker, "取得失敗", "-", "-", "-"])

# ===== DataFrame =====
df = pd.DataFrame(rows, columns=[
    "項目", "ティッカー", "価格", "前日比%", "前月比%", "年初来%"
])

# ===== 日本時間 =====
now = datetime.utcnow() + pd.Timedelta(hours=9)
today = now.strftime("%Y-%m-%d %H:%M")
df.insert(0, "日付", today)

# ===== Excel出力 =====
df.to_excel("market_data.xlsx", index=False)

# ===== HTMLメール作成 =====
html = f"<h2>📊 マーケットデータ（{today}）</h2>"
html += "<table border='1' cellpadding='5' cellspacing='0'>"
html += "<tr><th>項目</th><th>価格</th><th>前日比</th><th>前月比</th><th>年初来</th></tr>"

def colorize(val):
    if val == "-" or val == "取得失敗":
        return val
    color = "green" if val > 0 else "red"
    return f"<span style='color:{color}'>{val}%</span>"

for _, row in df.iterrows():
    html += "<tr>"
    html += f"<td>{row['項目']}</td>"
    html += f"<td>{row['価格']}</td>"
    html += f"<td>{colorize(row['前日比%'])}</td>"
    html += f"<td>{colorize(row['前月比%'])}</td>"
    html += f"<td>{colorize(row['年初来%'])}</td>"
    html += "</tr>"

html += "</table>"

# ===== メール送信 =====
msg = MIMEText(html, "html")
msg["Subject"] = "📊 マーケットデータ（自動配信）"
msg["From"] = EMAIL
msg["To"] = EMAIL

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(EMAIL, PASSWORD)
    smtp.send_message(msg)

print("送信完了")