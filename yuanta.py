from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import openpyxl, re, os, time

# ======= 可改：WID 清單 =======
wid_list = [
    "03111U", "03162U", "03485U", "03616U", "03662U",
    "03281U", "03864U", "05831P", "063866", "065413", "071599",
    "07879P", "079683", "085398", "08700P", "08769P", "08992P",
    "71280U", "71286U", "71289U", "71344U", "71974U"
]

# ======= 可改：基本資料欄位（網站左欄標籤）=======
BASIC_LABELS = [
    "上市日期","最後交易日","到期日期","發行型態","最新發行張數",
    "流通在外張數/比例","最新履約價","最新行使比例",
    "買價隱波","賣價隱波","Delta","Theta",
    "剩餘天數","價內外程度","實質槓桿","買賣價差比"
]

# ======= Driver 啟動（跨平台）=======
def launch_driver(headless=False, chrome_version_hint=None):
    """
    headless: True 可背景執行
    chrome_version_hint: 若 webdriver_manager 無法自動配對，可傳 "138.0.0" 等大版本字串
    """
    options = webdriver.ChromeOptions()

    # 如需背景執行，改 True
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    # 允許用環境變數覆寫 Chrome 路徑（找不到時才需要）
    # 例：export CHROME_BIN="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    chrome_bin = os.environ.get("CHROME_BIN")
    if chrome_bin:
        options.binary_location = chrome_bin

    # 自動安裝匹配版 chromedriver（若 chrome_version_hint 提供，會用它）
    try:
        if chrome_version_hint:
            service = Service(ChromeDriverManager(version=chrome_version_hint).install())
        else:
            service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        return driver
    except WebDriverException as e:
        raise SystemExit(f"🚨 無法啟動 ChromeDriver：{e}\n"
                         f"若提示找不到 Chrome，可設定環境變數 CHROME_BIN 指向 Chrome 可執行檔。")

def text_or_blank(driver, by, sel):
    try:
        return driver.find_element(by, sel).text.strip()
    except NoSuchElementException:
        return ""

def wait_text_not_empty(driver, by, sel, timeout=25):
    WebDriverWait(driver, timeout).until(
        lambda d: d.find_element(by, sel).text.strip() != ""
    )
    return driver.find_element(by, sel).text.strip()

def find_basic_value_by_label(driver, label_text):
    """在『左為標籤、右為值』的雙欄結構中找值（多種 XPath 提升相容性）"""
    XPS = [
        f"//*[normalize-space(text())='{label_text}']/following-sibling::*[1]",
        f"//div[.//*[normalize-space(text())='{label_text}']]/*[normalize-space(text())='{label_text}']/following-sibling::*[1]",
        f"//li[.//*[normalize-space(text())='{label_text}']]//*[normalize-space(text())='{label_text}']/following::*[1]",
    ]
    for xp in XPS:
        try:
            el = driver.find_element(By.XPATH, xp)
            txt = el.text.strip()
            if txt:
                return txt
        except NoSuchElementException:
            continue
    return ""

def get_target_info(driver):
    """取『標的名稱』『標的現價』：先試 ng-bind，再以包含『標的』的抬頭正則解析"""
    name = ""
    price = ""
    for xp in [
        "//*[contains(@ng-bind, 'TAR_NAME') or contains(@ng-bind, 'FLD_TAR_NAME')]",
    ]:
        els = driver.find_elements(By.XPATH, xp)
        if els and els[0].text.strip():
            name = els[0].text.strip()
            break
    for xp in [
        "//*[contains(@ng-bind, 'TAR_PRICE') or contains(@ng-bind, 'FLD_TAR_PRICE')]",
    ]:
        els = driver.find_elements(By.XPATH, xp)
        if els and els[0].text.strip():
            price = els[0].text.strip().replace(",", "")
            break
    if name or price:
        return name, price

    # 備援：抓含「標的」的整塊字串後解析
    try:
        header = driver.find_element(By.XPATH, "//*[contains(normalize-space(.), '標的')]")
        block = header.text.strip()
    except NoSuchElementException:
        block = ""

    if block:
        after = re.split(r"標的[:：]", block, maxsplit=1)
        tail = after[1].strip() if len(after) > 1 else block
        m_name = re.match(r"([^\s(／/｜|]+)", tail)
        guess_name = m_name.group(1) if m_name else ""
        m_px = re.search(r"(\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+\.\d+)", tail)
        guess_px = m_px.group(1).replace(",", "") if m_px else ""
        return guess_name, guess_px

    return "", ""

def scrape_one_wid(driver, wid):
    url = f"https://www.warrantwin.com.tw/eyuanta/Warrant/Info.aspx?WID={wid}"
    driver.get(url)

    status = "OK"

    # 等到有價格區塊（Angular 綁定），否則標註狀態
    try:
        WebDriverWait(driver, 12).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(@ng-bind, 'WAR_BUY_PRICE')]"))
        )
        wait_text_not_empty(driver, By.XPATH, "//*[contains(@ng-bind, 'WAR_BUY_PRICE')]", timeout=20)
    except TimeoutException:
        status = "No price section / not a warrant?"

    # 優先 ng-bind，其次 class tBig
    deal = text_or_blank(driver, By.XPATH, "//*[contains(@ng-bind, 'WAR_DEAL_PRICE')]")
    buy  = text_or_blank(driver, By.XPATH, "//*[contains(@ng-bind, 'WAR_BUY_PRICE')]")
    sell = text_or_blank(driver, By.XPATH, "//*[contains(@ng-bind, 'WAR_SELL_PRICE')]")

    if not (deal and buy and sell):
        try:
            WebDriverWait(driver, 6).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "tBig")))
            prices = [e.text.strip() for e in driver.find_elements(By.CLASS_NAME, "tBig")]
            if len(prices) >= 3:
                deal = deal or prices[0]
                buy  = buy  or prices[1]
                sell = sell or prices[2]
        except TimeoutException:
            pass

    tgt_name, tgt_px = get_target_info(driver)
    basic = {lab: find_basic_value_by_label(driver, lab) for lab in BASIC_LABELS}

    if not (deal or buy or sell):
        status = "No prices"

    return {
        "WID": wid,
        "狀態": status,
        "成交價": deal,
        "買價": buy,
        "賣價": sell,
        "標的名稱": tgt_name,
        "標的現價": tgt_px,
        **basic,
        "抓取時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "來源網址": url,
    }

def save_rows_to_excel(rows, filename="yuanta_warrants.xlsx"):
    headers = list(rows[0].keys())
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "元大權證"
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h, "") for h in headers])

    # ✅ save to Desktop
    desktop = os.path.expanduser("~/Desktop")
    out_path = os.path.join(desktop, filename)

    wb.save(out_path)
    print(f"✅ 已輸出 Excel：{out_path}")
    
def main():
    driver = launch_driver(headless=False)  # 需要背景執行改 True
    rows = []
    try:
        for wid in wid_list:
            print(f"抓取 {wid} …")
            row = scrape_one_wid(driver, wid)
            print(f"→ 狀態:{row['狀態']} 成交:{row['成交價']} 買:{row['買價']} 賣:{row['賣價']} 標的:{row['標的名稱']} 價:{row['標的現價']}")
            rows.append(row)
            time.sleep(0.3)  # 禮貌性間隔，避免過快
    finally:
        driver.quit()

    if rows:
        save_rows_to_excel(rows)
    else:
        print("⚠️ 沒有任何可寫入的資料")

if __name__ == "__main__":
    main()
