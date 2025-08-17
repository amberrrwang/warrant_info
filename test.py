from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

options = webdriver.ChromeOptions()

# ✅ 新增：手動指定你的 Chrome.app 路徑（M1/M2 預設位置）
options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

# ✅ 如需隱藏視窗就加這行，測試時建議先註解
# options.add_argument("--headless=new")

# ✅ 指定你手動下載的 chromedriver 路徑
driver_path = "/Users/amber/Downloads/chromedriver"

# ✅ 正確啟動 ChromeDriver
driver = webdriver.Chrome(service=Service(driver_path), options=options)

driver.get("https://www.google.com")
print("✅ Google opened, title:", driver.title)

driver.quit()
