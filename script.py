import requests
from bs4 import BeautifulSoup
import pandas as pd

url = "https://finance.yahoo.com/quote/%5EIXIC/"
headers = {"User-Agent": "Mozilla/5.0"}

res = requests.get(url, headers=headers)
soup = BeautifulSoup(res.text, "html.parser")

price = soup.find("fin-streamer").text

df = pd.DataFrame([[price]], columns=["NASDAQ"])

df.to_excel("nasdaq.xlsx", index=False)