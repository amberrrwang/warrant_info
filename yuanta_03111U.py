from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import openpyxl, os, re, time
import requests  # ← 新增：用來打 Yuanta API

# ======= 設定 =======
wid_list = ["03111U"]

BASIC_LABELS = [
    "上市日期","最後交易日","到期日期","發行型態","最新發行張數",
    "流通在外張數/比例","最新履約價","最新行使比例",
    "買價隱波","賣價隱波","Delta","Theta",
    "剩餘天數","價內外程度","實質槓桿","買賣價差比"
]

# 只保留「標的股價」，不再有「標的現價」
HEADER_ORDER = [
    "WID","狀態","成交價","買價","賣價",
    "標的名稱","標的股價","標的代碼",
    *BASIC_LABELS, "抓取時間","來源網址"
]

# ======= 啟動 Driver =======
def launch_driver(headless=False):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

# ======= 抓資料輔助 =======
def text_or_blank(driver, by, sel):
    try:
        return driver.find_element(by, sel).text.strip()
    except NoSuchElementException:
        return ""

def find_basic_value_by_label(driver, label_text):
    xps = [
        f"//*[normalize-space(text())='{label_text}']/following-sibling::*[1]",
        f"//div[.//*[normalize-space(text())='{label_text}']]/*[normalize-space(text())='{label_text}']/following-sibling::*[1]",
        f"//li[.//*[normalize-space(text())='{label_text}']]//*[normalize-space(text())='{label_text}']/following::*[1]",
    ]
    for xp in xps:
        try:
            txt = driver.find_element(By.XPATH, xp).text.strip()
            if txt:
                return txt
        except NoSuchElementException:
            continue
    return ""

def get_target_name_code(driver):
    """抓標的名稱/代碼（不抓價）。"""
    name, code = "", ""

    # 名稱
    for xp in ["//*[contains(@ng-bind, 'TAR_NAME') or contains(@ng-bind, 'FLD_TAR_NAME')]"]:
        els = driver.find_elements(By.XPATH, xp)
        if els and els[0].text.strip():
            name = els[0].text.strip()
            break

    # 代碼
    for xp in ["//*[contains(@ng-bind, 'TAR_CODE') or contains(@ng-bind, 'FLD_TAR_CODE')]"]:
        els = driver.find_elements(By.XPATH, xp)
        if els and els[0].text.strip():
            code = re.sub(r"\D", "", els[0].text.strip())
            break

    # 備援：從含「標的」的文字解析
    if not (name and code):
        try:
            block = driver.find_element(By.XPATH, "//*[contains(normalize-space(.), '標的')]").text.strip()
            if not name:
                m_name = re.search(r"標的[:：]\s*([^\s／/｜|()（）]+)", block)
                name = m_name.group(1) if m_name else name
            if not code:
                m_code = re.search(r"\((\d{4})\)", block) or re.search(r"[^\d](\d{4})(?:\D|$)", block)
                code = m_code.group(1) if m_code else code
        except NoSuchElementException:
            pass

    return name, code

# ======= NEW：從 Yuanta API 取「標的股價＝賣一(ask1)」 =======
def get_udly_best_ask_from_api(udly_code: str, timeout=8):
    """
    /ws/Quote.ashx?type=mem_ta5&symbol={udly_code}
    鍵位：
      101=買一, 102=賣一, 103=買二, 104=賣二, ..., 110=賣五
      113..117=買一~買五量, 118..122=賣一~賣五量
    回傳 float 或 None
    """
    if not udly_code:
        return None
    url = f"https://www.warrantwin.com.tw/eyuanta/ws/Quote.ashx?type=mem_ta5&symbol={udly_code}"
    try:
        r = requests.get(url, timeout=timeout)
        r.raise_for_status()
        data = r.json()
        items = data.get("items", {})
        ask1 = items.get("102") if isinstance(items, dict) else None
        if ask1 is None and isinstance(items, dict):  # 保險：整數鍵
            ask1 = items.get(102)
        if ask1 is None:
            return None
        try:
            return float(str(ask1).replace(",", ""))
        except Exception:
            return None
    except Exception as e:
        print("⚠️ get_udly_best_ask_from_api error:", e)
        return None

# （可留作備援）從 DOM 五檔表抓第一列賣價
def get_target_best_ask_from_dom(driver):
    try:
        WebDriverWait(driver, 6).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(normalize-space(.), '標的五檔報價')]"))
        )
        td = driver.find_element(
            By.XPATH, "//*[contains(normalize-space(.), '標的五檔報價')]/following::table[1]//tr[1]/td[3]"
        )
        return td.text.strip().replace(",", "")
    except Exception:
        return ""

def ensure_all_keys(row: dict) -> dict:
    for k in HEADER_ORDER:
        row.setdefault(k, "")
    return row

# ======= 抓單筆 =======
def scrape_one_wid(driver, wid):
    url = f"https://www.warrantwin.com.tw/eyuanta/Warrant/Info.aspx?WID={wid}"
    driver.get(url)

    try:
        WebDriverWait(driver, 12).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(@ng-bind, 'WAR_BUY_PRICE')]"))
        )
    except TimeoutException:
        return ensure_all_keys({
            "WID": wid, "狀態": "Timeout", "來源網址": url,
            "抓取時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })

    # 三價
    deal = text_or_blank(driver, By.XPATH, "//*[contains(@ng-bind, 'WAR_DEAL_PRICE')]")
    buy  = text_or_blank(driver, By.XPATH, "//*[contains(@ng-bind, 'WAR_BUY_PRICE')]")
    sell = text_or_blank(driver, By.XPATH, "//*[contains(@ng-bind, 'WAR_SELL_PRICE')]")

    # 備援：class="tBig"
    if not (deal and buy and sell):
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, "tBig"))
            )
            prices = [e.text.strip() for e in driver.find_elements(By.CLASS_NAME, "tBig")]
            if len(prices) >= 3:
                deal = deal or prices[0]
                buy  = buy  or prices[1]
                sell = sell or prices[2]
        except TimeoutException:
            pass

    # 標的名稱與代碼
    tgt_name, tgt_code = get_target_name_code(driver)

    # ★ 先用 API 取標的股價（賣一＝items['102']）
    tgt_stock_price = get_udly_best_ask_from_api(tgt_code)

    # 若 API 失敗，退回 DOM 備援
    if tgt_stock_price is None:
        dom_price = get_target_best_ask_from_dom(driver)
        tgt_stock_price = float(dom_price) if dom_price else ""

    row = {
        "WID": wid,
        "狀態": "OK",
        "成交價": deal,
        "買價": buy,
        "賣價": sell,
        "標的名稱": tgt_name,
        "標的股價": tgt_stock_price,  # ← 來自 API (items['102'])；若失敗用 DOM 備援
        "標的代碼": tgt_code,
        "來源網址": url,
        "抓取時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    for label in BASIC_LABELS:
        row[label] = find_basic_value_by_label(driver, label)

    return ensure_all_keys(row)

# ======= 寫 Excel + 試算 =======
def save_rows_to_excel(rows, filename="yuanta_warrants.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "元大權證"

    ws.append(HEADER_ORDER)
    for r in rows:
        ws.append([r.get(k, "") for k in HEADER_ORDER])

    # 試算頁（以第一筆為例）
    r0 = rows[0]
    calc = wb.create_sheet("試算")

    # 標籤
    calc["A1"] = "WID"
    calc["A2"] = "標的股價"
    calc["A3"] = "買價隱波（％）"
    calc["A4"] = "評價日"
    calc["A6"] = "無風險利率 r（年化）"
    calc["F1"] = "（以下自動帶入）"
    calc["F2"] = "履約價 K"
    calc["F3"] = "剩餘天數"
    calc["F4"] = "行使比例（數值）"

    # 小工具：轉 float
    def to_float(x):
        try:
            return float(str(x).replace(",", ""))
        except Exception:
            return x

    # 值
    calc["B1"] = r0.get("WID", "")
    calc["B2"] = to_float(r0.get("標的股價", ""))
    calc["B3"] = to_float(r0.get("買價隱波", ""))
    calc["B4"] = datetime.now().strftime("%Y/%m/%d")
    calc["B6"] = 0.01
    calc["G2"] = to_float(r0.get("最新履約價", ""))
    calc["G3"] = to_float(r0.get("剩餘天數", ""))
    calc["G4"] = to_float(r0.get("最新行使比例", ""))
    calc["C10"] = f"成交價: {r0.get('成交價', '')}"

    # 粗體 & 欄寬
    for cell in ["A1","A2","A3","A4","A6","F2","F3","F4"]:
        calc[cell].font = openpyxl.styles.Font(bold=True)
    for col, width in [("A",16),("B",14),("C",28),("F",22),("G",18)]:
        calc.column_dimensions[col].width = width

    # 儲存到桌面
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    out_path = os.path.join(desktop, filename)
    wb.save(out_path)
    print(f"✅ 已寫入 Excel：{out_path}")

# ======= 主流程 =======
def main():
    driver = launch_driver(headless=False)
    rows = []
    try:
        for wid in wid_list:
            print(f"🔎 抓取 {wid} 中...")
            row = scrape_one_wid(driver, wid)
            print(
                f"→ 成交:{row.get('成交價','')} 買:{row.get('買價','')} 賣:{row.get('賣價','')} | "
                f"標的代碼:{row.get('標的代碼','')} 標的股價(賣一):{row.get('標的股價','')}"
            )
            rows.append(row)
            time.sleep(0.3)
    finally:
        driver.quit()

    if rows:
        save_rows_to_excel(rows)
    else:
        print("⚠️ 沒有資料可寫入")

if __name__ == "__main__":
    main()
