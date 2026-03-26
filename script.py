import yfinance as yf
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
import os

# ===== メール設定 =====
EMAIL = os.getenv("EMAIL_USER")
PASSWORD = os.getenv("EMAIL_PASS")

# ===== 指数 =====
indices = [
    ("ダウ平均", "^DJI"),
    ("S&P500", "^GSPC"),
    ("ナスダック", "^IXIC"),
    ("ラッセル2000", "^RUT"),
    ("SOX指数", "^SOX"),
]

# ===== 個別株 =====
stocks = [
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

targets = indices + stocks

rows = []

for name, ticker in targets:
    try:
        t = yf.Ticker(ticker)

        # ===== 日次データ（1年）=====
        hist = t.history(period="1y")

        if hist.empty:
            rows.append([name, ticker, "取得失敗", "-", "-", "-"])
            continue

        close = hist["Close"].dropna()

        # ===== 前日比（厳密）=====
        latest = close.iloc[-1]
        prev = close.iloc[-2] if len(close) > 1 else latest
        day_ret = (latest / prev - 1) * 100

        # ===== 前月末比（厳密）=====
        hist_month = t.history(period="6mo")
        month_close = hist_month["Close"].resample("M").last().dropna()

        if len(month_close) >= 2:
            prev_month = month_close.iloc[-2]
        else:
            prev_month = close.iloc[0]

        month_ret = (latest / prev_month - 1) * 100

        # ===== 年初来（厳密）=====
        current_year = datetime.utcnow().year
        start_of_year = f"{current_year}-01-01"

        hist_ytd = t.history(start=start_of_year)

        if not hist_ytd.empty:
            year_start = hist_ytd["Close"].dropna().iloc[0]
        else:
            year_start = close.iloc[0]

        year_ret = (latest / year_start - 1) * 100

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
    "項目", "ティッカー", "価格", "前日比%", "前月末比%", "年初来%"
])

# ===== 日本時間 =====
now = datetime.utcnow() + pd.Timedelta(hours=9)
today = now.strftime("%Y-%m-%d %H:%M")
df.insert(0, "日付", today)

# ===== Excel出力 =====
df.to_excel("market_data.xlsx", index=False)

# ===== HTMLメール =====
html = f"<h2>📊 マーケットデータ（{today}）</h2>"

def colorize(val):
    if val == "-" or val == "取得失敗":
        return val
    color = "green" if val > 0 else "red"
    return f"<span style='color:{color}'>{val}%</span>"

# ===== 指数 =====
html += "<h3>■ 指数</h3><table border='1' cellpadding='5'>"
html += "<tr><th>項目</th><th>価格</th><th>前日比</th><th>前月末比</th><th>年初来</th></tr>"

for _, row in df.iloc[:len(indices)].iterrows():
    html += "<tr>"
    html += f"<td>{row['項目']}</td>"
    html += f"<td>{row['価格']}</td>"
    html += f"<td>{colorize(row['前日比%'])}</td>"
    html += f"<td>{colorize(row['前月末比%'])}</td>"
    html += f"<td>{colorize(row['年初来%'])}</td>"
    html += "</tr>"

html += "</table>"

# ===== 個別株 =====
html += "<h3>■ 個別株</h3><table border='1' cellpadding='5'>"
html += "<tr><th>項目</th><th>価格</th><th>前日比</th><th>前月末比</th><th>年初来</th></tr>"

for _, row in df.iloc[len(indices):].iterrows():
    html += "<tr>"
    html += f"<td>{row['項目']}</td>"
    html += f"<td>{row['価格']}</td>"
    html += f"<td>{colorize(row['前日比%'])}</td>"
    html += f"<td>{colorize(row['前月末比%'])}</td>"
    html += f"<td>{colorize(row['年初来%'])}</td>"
    html += "</tr>"

html += "</table>"

# ===== メール送信 =====
msg = MIMEText(html, "html")
msg["Subject"] = "📊 マーケットデータ（完全版）"
msg["From"] = EMAIL
msg["To"] = EMAIL

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(EMAIL, PASSWORD)
    smtp.send_message(msg)

print("送信完了")