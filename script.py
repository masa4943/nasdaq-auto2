from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import pandas as pd

# ===== Chrome設定 =====
options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

# ===== ドライバー起動 =====
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# ===== サイトアクセス =====
url = "https://nikkei225jp.com/nasdaq/"
driver.get(url)

# ===== テーブルが表示されるまで待つ =====
wait = WebDriverWait(driver, 20)
table = wait.until(
    EC.presence_of_element_located((By.TAG_NAME, "table"))
)

# ===== データ取得 =====
rows = table.find_elements(By.TAG_NAME, "tr")

data = []
for row in rows:
    cols = [col.text.strip() for col in row.find_elements(By.TAG_NAME, "td")]
    if cols:
        data.append(cols)

# ===== DataFrame化 =====
df = pd.DataFrame(data)

# ===== Excel出力 =====
df.to_excel("nasdaq.xlsx", index=False)

# ===== 終了 =====
driver.quit()