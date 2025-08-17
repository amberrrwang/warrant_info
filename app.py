# /Users/amber/Desktop/app.py
# -*- coding: utf-8 -*-
from flask import Flask, jsonify, request, make_response
from flask import render_template_string
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime
import time, re

# ====== 你的環境：沿用你已經可用的路徑 ======
CHROME_BIN    = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
CHROMEDRIVER  = "/Users/amber/Downloads/chromedriver"

# 預設清單（可用網址參數覆蓋）
DEFAULT_WIDS = [
    "03111U", "03162U", "03485U", "03616U", "03662U",
    "03281U", "03864U", "05831P", "063866", "065413", "071599",
    "07879P", "079683", "085398", "08700P", "08769P", "08992P",
    "71280U", "71286U", "71289U", "71344U", "71974U"
]

BASIC_LABELS = [
    "上市日期","最後交易日","到期日期","發行型態","最新發行張數",
    "流通在外張數/比例","最新履約價","最新行使比例",
    "買價隱波","賣價隱波","Delta","Theta",
    "剩餘天數","價內外程度","實質槓桿","買賣價差比"
]

# ====== Selenium 基本工具 ======
def launch_driver():
    options = webdriver.ChromeOptions()
    options.binary_location = CHROME_BIN
    options.add_argument("--headless=new")             # 網頁版建議背景執行
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.page_load_strategy = "eager"
    drv = webdriver.Chrome(service=Service(CHROMEDRIVER), options=options)
    drv.set_page_load_timeout(30)
    drv.set_script_timeout(30)
    return drv

def text_or_blank(drv, by, sel):
    try:
        return drv.find_element(by, sel).text.strip()
    except NoSuchElementException:
        return ""

def find_basic_value_by_label(drv, label_text):
    XPS = [
        f"//*[normalize-space(text())='{label_text}']/following-sibling::*[1]",
        f"//div[.//*[normalize-space(text())='{label_text}']]/*[normalize-space(text())='{label_text}']/following-sibling::*[1]",
        f"//li[.//*[normalize-space(text())='{label_text}']]//*[normalize-space(text())='{label_text}']/following::*[1]",
    ]
    for xp in XPS:
        try:
            el = drv.find_element(By.XPATH, xp)
            t = el.text.strip()
            if t:
                return t
        except NoSuchElementException:
            continue
    return ""

def get_target_info(drv):
    # 先找 ng-bind
    name = ""
    price = ""
    for xp in ["//*[contains(@ng-bind,'TAR_NAME')]", "//*[contains(@ng-bind,'FLD_TAR_NAME')]"]:
        els = drv.find_elements(By.XPATH, xp)
        if els and els[0].text.strip():
            name = els[0].text.strip(); break
    for xp in ["//*[contains(@ng-bind,'TAR_PRICE')]", "//*[contains(@ng-bind,'FLD_TAR_PRICE')]"]:
        els = drv.find_elements(By.XPATH, xp)
        if els and els[0].text.strip():
            price = els[0].text.strip().replace(",", ""); break
    if name or price: return name, price

    # 再找含「標的」的抬頭字串並解析
    try:
        header = drv.find_element(By.XPATH, "//*[contains(normalize-space(.), '標的')]")
        block = header.text.strip()
        if block:
            after = re.split(r"標的[:：]", block, maxsplit=1)
            tail = after[1].strip() if len(after) > 1 else block
            m_name = re.match(r"([^\s(／/｜|]+)", tail)
            m_px   = re.search(r"(\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+\.\d+)", tail)
            return (m_name.group(1) if m_name else ""), (m_px.group(1).replace(",", "") if m_px else "")
    except NoSuchElementException:
        pass
    return "", ""

def is_warrant_code(code: str) -> bool:
    return code.endswith(("U","P")) or code.isdigit()

def scrape_one(drv, wid):
    url = f"https://www.warrantwin.com.tw/eyuanta/Warrant/Info.aspx?WID={wid}"
    drv.get(url)

    status = "OK"
    # 等待買價 ng-bind 出現且有字
    try:
        WebDriverWait(drv, 12).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(@ng-bind, 'WAR_BUY_PRICE')]"))
        )
        WebDriverWait(drv, 20).until(
            lambda d: d.find_element(By.XPATH, "//*[contains(@ng-bind, 'WAR_BUY_PRICE')]").text.strip() != ""
        )
    except TimeoutException:
        status = "No price section / slow"

    deal = text_or_blank(drv, By.XPATH, "//*[contains(@ng-bind,'WAR_DEAL_PRICE')]")
    buy  = text_or_blank(drv, By.XPATH, "//*[contains(@ng-bind,'WAR_BUY_PRICE')]")
    sell = text_or_blank(drv, By.XPATH, "//*[contains(@ng-bind,'WAR_SELL_PRICE')]")

    if not (deal and buy and sell):
        # 備援 class（常見順序：成交/買/賣）
        try:
            WebDriverWait(drv, 6).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "tBig")))
            prices = [e.text.strip() for e in drv.find_elements(By.CLASS_NAME, "tBig")]
            if len(prices) >= 3:
                deal = deal or prices[0]
                buy  = buy  or prices[1]
                sell = sell or prices[2]
        except TimeoutException:
            pass

    tgt_name, tgt_px = get_target_info(drv)
    basic = {lab: find_basic_value_by_label(drv, lab) for lab in BASIC_LABELS}

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

def scrape_batch(wids):
    # 分批重啟 driver（避免 session 掛掉）
    results = []
    batch_size = 8
    for i in range(0, len(wids), batch_size):
        drv = launch_driver()
        try:
            for wid in wids[i:i+batch_size]:
                if not is_warrant_code(wid):
                    results.append({"WID": wid, "狀態": "非權證（略過）"})
                    continue
                try:
                    row = scrape_one(drv, wid)
                except Exception as e:
                    row = {"WID": wid, "狀態": f"Error: {type(e).__name__}"}
                results.append(row)
        finally:
            try: drv.quit()
            except: pass
    return results

# ====== Flask App ======
app = Flask(__name__)

INDEX_HTML = """
<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <title>元大權證即時看板</title>
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <style>
    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial; margin:20px;}
    h1{font-size:20px;margin:0 0 12px}
    .controls{display:flex; gap:8px; align-items:center; margin-bottom:12px;}
    input,button{font-size:14px;padding:6px 10px}
    table{border-collapse:collapse; width:100%; font-size:14px}
    th,td{border:1px solid #ddd; padding:8px; text-align:left}
    th{background:#f5f5f5; position:sticky; top:0}
    tr:nth-child(even){background:#fafafa}
    .muted{color:#888}
  </style>
</head>
<body>
  <h1>元大權證即時看板</h1>
  <div class="controls">
    <input id="wids" style="flex:1" placeholder="輸入代號（逗號分隔），留空用預設清單">
    <button onclick="loadData()">更新</button>
    <span class="muted" id="ts"></span>
  </div>
  <table id="tbl">
    <thead>
      <tr>
        <th>WID</th><th>狀態</th><th>成交價</th><th>買價</th><th>賣價</th>
        <th>標的名稱</th><th>標的現價</th>
        <th>上市日期</th><th>最後交易日</th><th>到期日期</th><th>發行型態</th>
        <th>最新發行張數</th><th>流通在外張數/比例</th><th>最新履約價</th><th>最新行使比例</th>
        <th>買價隱波</th><th>賣價隱波</th><th>Delta</th><th>Theta</th>
        <th>剩餘天數</th><th>價內外程度</th><th>實質槓桿</th><th>買賣價差比</th>
        <th>抓取時間</th>
      </tr>
    </thead>
    <tbody></tbody>
  </table>
<script>
async function loadData(){
  const w = document.getElementById('wids').value.trim();
  const url = w ? '/api/warrants?wids=' + encodeURIComponent(w) : '/api/warrants';
  const res = await fetch(url, {cache: 'no-store'});
  const data = await res.json();
  const tb = document.querySelector('#tbl tbody');
  tb.innerHTML = '';
  for (const r of data.items){
    const tr = document.createElement('tr');
    const cols = [
      'WID','狀態','成交價','買價','賣價','標的名稱','標的現價',
      '上市日期','最後交易日','到期日期','發行型態','最新發行張數',
      '流通在外張數/比例','最新履約價','最新行使比例',
      '買價隱波','賣價隱波','Delta','Theta','剩餘天數','價內外程度','實質槓桿','買賣價差比',
      '抓取時間'
    ];
    for (const k of cols){
      const td = document.createElement('td');
      td.textContent = (r[k] ?? '');
      tr.appendChild(td);
    }
    tb.appendChild(tr);
  }
  document.getElementById('ts').textContent = '更新時間：' + data.generated_at;
}
// 自動每 60 秒更新一次
loadData();
setInterval(loadData, 60000);
</script>
</body>
</html>
"""

@app.route("/")
def index():
    return render_template_string(INDEX_HTML)

@app.route("/api/warrants")
def api_warrants():
    # 支援 query: ?wids=03111U,03162U,...
    q = request.args.get("wids", "")
    if q.strip():
        wids = [x.strip() for x in q.split(",") if x.strip()]
    else:
        wids = DEFAULT_WIDS
    items = scrape_batch(wids)
    payload = {"generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "count": len(items), "items": items}
    resp = make_response(jsonify(payload))
    resp.headers["Cache-Control"] = "no-store"
    return resp

if __name__ == "__main__":
    app.run(debug=True)
