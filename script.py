from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time

options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=options)

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