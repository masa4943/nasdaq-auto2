import requests
from bs4 import BeautifulSoup
import pandas as pd

url = "https://nikkei225jp.com/nasdaq/d"
headers = {"User-Agent": "Mozilla/5.0"}

res = requests.get(url, headers=headers)
soup = BeautifulSoup(res.text, "html.parser")

table = soup.find("table")

data = []
rows = table.find_all("tr")

for row in rows:
    cols = [col.text.strip() for col in row.find_all(["td", "th"])]
    data.append(cols)

df = pd.DataFrame(data)
df.to_excel("nasdaq.xlsx", index=False)