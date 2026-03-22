from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import pandas as pd
import time

options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

# 🔥 ここがポイント（自動でバージョン合わせる）
service = Service(ChromeDriverManager().install())

driver = webdriver.Chrome(service=service, options=options)

driver.get("https://nikkei225jp.com/nasdaq/")
time.sleep(5)

table = driver.find_element(By.TAG_NAME, "table")
rows = table.find_elements(By.TAG_NAME, "tr")

data = []
for row in rows:
    cols = [col.text for col in row.find_elements(By.TAG_NAME, "td")]
    if cols:
        data.append(cols)

df = pd.DataFrame(data)
df.to_excel("nasdaq.xlsx", index=False)

driver.quit()