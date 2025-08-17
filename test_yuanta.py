# test_yuanta.py
# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from datetime import datetime
import openpyxl, time, re

CHROME_BIN    = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
CHROMEDRIVER  = "/Users/amber/Downloads/chromedriver"

WIDS = ["03111U"]

BASIC_LABELS = [
    "上市日期","最後交易日","到期日期","發行型態","最新發行張數",
    "流通在外張數/比例","最新履約價","最新行使比例",
    "買價隱波","賣價隱波","Delta","Theta",
    "剩餘天數","價內外程度","實質槓桿","買賣價差比"
]

# ---------- 驅動 ----------
def launch_driver():
    options = webdriver.ChromeOptions()
    options.binary_location = CHROME_BIN
    # options.add_argument("--headless=new")  # 需要背景執行就打開
    options.add_argument("--disable-blink-features=AutomationControlled")
    return webdriver.Chrome(service=Service(CHROMEDRIVER), options=options)

def wait_text_not_empty(driver, by, sel, timeout=25):
    WebDriverWait(driver, timeout).until(
        lambda d: d.find_element(by, sel).text.strip() != ""
    )
    return driver.find_element(by, sel).text.strip()

def txt(driver, by, sel):
    try:
        return driver.find_element(by, sel).text.strip()
    except NoSuchElementException:
        return ""

# ---------- 基本資料 ----------
def find_basic_value_by_label(driver, label_text):
    XPS = [
        f"//*[normalize-space(text())='{label_text}']/following-sibling::*[1]",
        f"//div[.//*[normalize-space(text())='{label_text}']]/*[normalize-space(text())='{label_text}']/following-sibling::*[1]",
        f"//li[.//*[normalize-space(text())='{label_text}']]//*[normalize-space(text())='{label_text}']/following::*[1]",
    ]
    for xp in XPS:
        try:
            el = driver.find_element(By.XPATH, xp)
            t = el.text.strip()
            if t:
                return t
        except NoSuchElementException:
            pass
    return ""

# ---------- 標的資訊（三段式，最穩） ----------
def get_target_info(driver):
    # 1) 先試 ng-bind（若站上有綁定）
    name = ""
    price = ""
    for xp in [
        "//*[contains(@ng-bind, \"TAR_NAME\")]",
        "//*[contains(@ng-bind, \"FLD_TAR_NAME\")]",
    ]:
        els = driver.find_elements(By.XPATH, xp)
        if els and els[0].text.strip():
            name = els[0].text.strip()
            break
    if not price:
        for xp in [
            "//*[contains(@ng-bind, \"TAR_PRICE\")]",
            "//*[contains(@ng-bind, \"FLD_TAR_PRICE\")]",
        ]:
            els = driver.find_elements(By.XPATH, xp)
            if els and els[0].text.strip():
                price = els[0].text.strip()
                break
    if name or price:
        return name, price

    # 2) 找包含「標的」關鍵字的抬頭塊
    # 典型文字：標的：祥碩 (5269) 1785.00｜-140.00 (-7.27%)
    try:
        header = driver.find_element(By.XPATH, "//*[contains(normalize-space(.), '標的')]")
        block = header.text.strip()
    except NoSuchElementException:
        block = ""

    if block:
        # 先切出「標的：」之後的內容
        after = re.split(r"標的[:：]", block, maxsplit=1)
        tail = after[1].strip() if len(after) > 1 else block

        # 名稱：第一個非空字串，去掉括號代碼
        # 例如 "祥碩 (5269) 1785.00" -> "祥碩"
        m_name = re.match(r"([^\s(／/｜|]+)", tail)
        guess_name = m_name.group(1) if m_name else ""

        # 現價：尋找第一個像數字的小數（允許千分位）
        m_px = re.search(r"(\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+\.\d+)", tail)
        guess_px = m_px.group(1).replace(",", "") if m_px else ""

        if guess_name or guess_px:
            return guess_name, guess_px

    # 3) 全站再找可能的標的容器作為備援
    for xp in [
        "//h1[contains(., '標的')]", "//h2[contains(., '標的')]",
        "//div[contains(., '標的')]", "//p[contains(., '標的')]",
    ]:
        els = driver.find_elements(By.XPATH, xp)
        if els:
            text = els[0].text.strip()
            if text:
                after = re.split(r"標的[:：]", text, maxsplit=1)
                tail = after[1].strip() if len(after) > 1 else text
                m_name = re.match(r"([^\s(／/｜|]+)", tail)
                m_px   = re.search(r"(\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+\.\d+)", tail)
                name = m_name.group(1) if m_name else ""
                price = m_px.group(1).replace(",", "") if m_px else ""
                return name, price

    return "", ""

# ---------- 抓一筆 ----------
def scrape_one_wid(driver, wid):
    url = f"https://www.warrantwin.com.tw/eyuanta/Warrant/Info.aspx?WID={wid}"
    driver.get(url)

    # 等到價格區塊真的有文字（以買價為錨點）
    try:
        WebDriverWait(driver, 25).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(@ng-bind, 'WAR_BUY_PRICE')]"))
        )
        wait_text_not_empty(driver, By.XPATH, "//*[contains(@ng-bind, 'WAR_BUY_PRICE')]", timeout=25)
    except TimeoutException:
        time.sleep(2)

    # 成交/買/賣：先 ng-bind，後備 class tBig
    deal = txt(driver, By.XPATH, "//*[contains(@ng-bind, 'WAR_DEAL_PRICE')]")
    buy  = txt(driver, By.XPATH, "//*[contains(@ng-bind, 'WAR_BUY_PRICE')]")
    sell = txt(driver, By.XPATH, "//*[contains(@ng-bind, 'WAR_SELL_PRICE')]")

    if not (deal and buy and sell):
        try:
            WebDriverWait(driver, 8).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "tBig")))
            prices = [e.text.strip() for e in driver.find_elements(By.CLASS_NAME, "tBig")]
            if len(prices) >= 3:
                deal = deal or prices[0]
                buy  = buy  or prices[1]
                sell = sell or prices[2]
        except TimeoutException:
            pass

    # 標的名稱 / 現價（更新版三段式）
    tgt_name, tgt_px = get_target_info(driver)

    # 基本資料
    basic = {lab: find_basic_value_by_label(driver, lab) for lab in BASIC_LABELS}

    return {
        "WID": wid,
        "成交價": deal,
        "買價": buy,
        "賣價": sell,
        "標的名稱": tgt_name,
        "標的現價": tgt_px,
        **basic,
        "抓取時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "來源網址": url,
    }

# ---------- 存 Excel ----------
def save_rows_to_excel(rows):
    headers = list(rows[0].keys())
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "元大權證"
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h, "") for h in headers])

    if len(rows) == 1:
        out = f"/Users/amber/Desktop/yuanta_{rows[0]['WID']}.xlsx"
    else:
        out = "/Users/amber/Desktop/yuanta_warrants.xlsx"

    wb.save(out)
    print(f"✅ 已輸出 Excel：{out}")

# ---------- main ----------
def main():
    driver = launch_driver()
    rows = []
    try:
        for wid in WIDS:
            print(f"抓取 {wid} …")
            row = scrape_one_wid(driver, wid)
            print(f"→ 成交:{row['成交價']} 買:{row['買價']} 賣:{row['賣價']} 標的:{row['標的名稱']} 價:{row['標的現價']}")
            rows.append(row)
    finally:
        driver.quit()
    save_rows_to_excel(rows)

if __name__ == "__main__":
    main()
