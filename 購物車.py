import io
import os
import re
import json
import time
import threading
import webbrowser
import requests
from bs4 import BeautifulSoup
from PIL import Image as PILImage
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from flask import Flask, request, jsonify, render_template_string, send_file

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>吉伊卡哇跨國購物車 🛒</title>
  <style>
    * { box-sizing: border-box; font-family: "Nunito", "Noto Sans TC", -apple-system, BlinkMacSystemFont, sans-serif; }
    body {
      margin: 0;
      padding: 24px 16px 60px;
      color: #5c4742;
      background-image: url('/bg.jpg');
      background-size: cover;
      background-position: center;
      background-attachment: fixed;
      min-height: 100vh;
    }
    .container { max-width: 1150px; margin: 0 auto; }
    .header-title {
      display: inline-block;
      background: rgba(255, 255, 255, 0.92);
      padding: 10px 22px;
      border-radius: 50px;
      font-size: 22px;
      font-weight: 800;
      color: #5c4742;
      border: 3px solid #ffb3c6;
      box-shadow: 0 4px 12px rgba(255, 143, 171, 0.25);
      margin: 0 0 16px 0;
    }
    .card {
      background: rgba(255, 255, 255, 0.88);
      backdrop-filter: blur(10px);
      -webkit-backdrop-filter: blur(10px);
      border: 2px solid rgba(255, 179, 198, 0.7);
      border-radius: 20px;
      padding: 20px;
      box-shadow: 0 8px 24px rgba(92, 71, 66, 0.1);
      margin-bottom: 18px;
    }
    textarea {
      width: 100%;
      height: 110px;
      padding: 12px;
      border: 2px solid #ffe5ec;
      border-radius: 14px;
      font-size: 13px;
      resize: vertical;
      background: rgba(255, 255, 255, 0.95);
      color: #5c4742;
      outline: none;
    }
    textarea:focus { border-color: #ff8fab; }
    .controls { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; margin-top: 12px; }
    button {
      background: #ff8fab;
      color: #fff;
      border: none;
      padding: 10px 18px;
      border-radius: 50px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 700;
      box-shadow: 0 3px 8px rgba(255, 143, 171, 0.35);
      transition: transform 0.1s, background 0.2s;
    }
    button:hover { background: #fb6f92; transform: translateY(-1px); }
    button:disabled { background: #ccc; box-shadow: none; cursor: not-allowed; }
    .btn-add { background: #8ecae6; box-shadow: 0 3px 8px rgba(142, 202, 230, 0.4); }
    .btn-add:hover { background: #219ebc; }
    .btn-export { background: #70c1b3; box-shadow: 0 3px 8px rgba(112, 193, 179, 0.4); }
    .btn-export:hover { background: #4ea898; }
    .btn-rate { background: #ffd166; color: #5c4742; box-shadow: 0 3px 8px rgba(255, 209, 102, 0.4); }
    .btn-rate:hover { background: #f4be49; }

    .rates-bar {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 14px;
      padding-top: 14px;
      border-top: 2px dashed #ffe5ec;
      align-items: center;
    }
    .rate-item {
      display: flex;
      align-items: center;
      gap: 5px;
      background: #fff0f3;
      padding: 6px 12px;
      border-radius: 50px;
      font-size: 13px;
      font-weight: 700;
      border: 1px solid #ffccd5;
    }
    .rate-item input {
      width: 68px;
      padding: 3px 6px;
      border: 1px solid #ffb3c6;
      border-radius: 8px;
      text-align: right;
      font-weight: bold;
      color: #5c4742;
      background: #fff;
    }
    .summary {
      position: sticky;
      top: 12px;
      z-index: 10;
      display: flex;
      justify-content: space-between;
      align-items: center;
      background: rgba(92, 71, 66, 0.94);
      backdrop-filter: blur(8px);
      color: #fff;
      padding: 14px 24px;
      border-radius: 50px;
      margin-bottom: 18px;
      border: 2px solid #ffb3c6;
      box-shadow: 0 6px 20px rgba(92, 71, 66, 0.25);
    }
    .summary strong { font-size: 22px; color: #ffd166; margin-left: 6px; }
    .table-wrap {
      background: rgba(255, 255, 255, 0.9);
      backdrop-filter: blur(10px);
      border-radius: 20px;
      border: 2px solid rgba(255, 179, 198, 0.7);
      overflow-x: auto;
      box-shadow: 0 8px 24px rgba(92, 71, 66, 0.1);
    }
    table { width: 100%; border-collapse: collapse; }
    th, td { padding: 12px; border-bottom: 1px solid #ffe5ec; text-align: left; vertical-align: middle; }
    th { background: rgba(255, 240, 243, 0.9); font-size: 13px; color: #5c4742; font-weight: 800; }
    .thumb { width: 78px; height: 78px; object-fit: cover; border-radius: 12px; border: 2px solid #ffe5ec; background: #fff; }
    .title-input { width: 100%; padding: 6px 8px; border: 1px solid transparent; border-radius: 8px; font-size: 14px; background: transparent; color: #5c4742; font-weight: 600; }
    .title-input:hover, .title-input:focus { border-color: #ffb3c6; background: #fff; outline: none; }
    .curr-select { padding: 5px 8px; border: 1px solid #ffb3c6; border-radius: 8px; font-size: 13px; font-weight: bold; background: #fff; color: #5c4742; }
    .price-input { width: 92px; padding: 6px; border: 1px solid #ffb3c6; border-radius: 8px; text-align: right; font-size: 14px; color: #5c4742; }
    .qty-input { width: 54px; padding: 6px; border: 1px solid #ffb3c6; border-radius: 8px; text-align: center; font-size: 14px; color: #5c4742; }
    .twd-price { font-weight: 800; color: #fb6f92; white-space: nowrap; font-size: 15px; }
    .link-sm { font-size: 12px; color: #219ebc; text-decoration: none; display: inline-block; margin-top: 3px; font-weight: 600; }
    .link-sm:hover { text-decoration: underline; }
    .status { font-size: 13px; color: #5c4742; margin-left: auto; font-weight: 600; }
  </style>
</head>
<body>
  <div class="container">
    <h1 class="header-title">🛍️ 跨國購物清單與即時匯率計算機</h1>

    <div class="card">
      <textarea id="urlInput" placeholder="貼上各國商品網址（韓、日、中、台、美、澳皆可，一行一個，會自動去重）..."></textarea>
      <div class="controls">
        <button id="fetchBtn" onclick="fetchProducts()">✨ 開始抓取商品</button>
        <button class="btn-add" onclick="addManualRow()">+ 手動新增</button>
        <button id="exportBtn" class="btn-export" onclick="exportExcel()">📊 匯出含圖 Excel</button>
        <button class="btn-rate" onclick="loadLiveRates()">🔄 更新即時匯率</button>
        <span id="statusText" class="status">正在載入最新匯率...</span>
      </div>
      <div class="rates-bar">
        <div class="rate-item"><span>🇰🇷 KRW</span><input type="number" id="rate-KRW" value="0.0235" step="0.0005" oninput="recalculate()"></div>
        <div class="rate-item"><span>🇯🇵 JPY</span><input type="number" id="rate-JPY" value="0.215" step="0.005" oninput="recalculate()"></div>
        <div class="rate-item"><span>🇨🇳 CNY</span><input type="number" id="rate-CNY" value="4.50" step="0.05" oninput="recalculate()"></div>
        <div class="rate-item"><span>🇺🇸 USD</span><input type="number" id="rate-USD" value="32.5" step="0.1" oninput="recalculate()"></div>
        <div class="rate-item"><span>🇦🇺 AUD</span><input type="number" id="rate-AUD" value="21.2" step="0.1" oninput="recalculate()"></div>
        <div class="rate-item"><span>🇹🇼 TWD</span><input type="number" id="rate-TWD" value="1" disabled></div>
      </div>
    </div>

    <div class="summary">
      <div>已勾選 <strong id="selectedCount">0</strong> 件商品</div>
      <div>折合台幣總計：<strong id="totalTwd">NT$0</strong></div>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th style="width: 36px;"><input type="checkbox" id="selectAll" checked onchange="toggleAll(this)"></th>
            <th style="width: 96px;">商品照</th>
            <th>商品名稱（點擊可直接修改）</th>
            <th style="width: 85px;">幣別</th>
            <th style="width: 110px;">原幣單價</th>
            <th style="width: 105px;">約合台幣</th>
            <th style="width: 72px;">數量</th>
          </tr>
        </thead>
        <tbody id="productTable">
          <tr><td colspan="7" style="text-align:center; color:#999; padding:36px;">在上方貼上網址，點擊「開始抓取商品」就能開始逛街啦 🛒</td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <script>
    let products = [];
    const CURRENCIES = ["KRW", "JPY", "CNY", "USD", "AUD", "TWD"];

    window.addEventListener("DOMContentLoaded", loadLiveRates);

    async function loadLiveRates() {
      const status = document.getElementById("statusText");
      status.innerText = "正在抓取今日最新匯率...";
      try {
        const res = await fetch("/api/rates");
        const data = await res.json();
        if (data.ok) {
          for (const [curr, val] of Object.entries(data.rates)) {
            const el = document.getElementById(`rate-${curr}`);
            if (el) el.value = val;
          }
          recalculate();
          status.innerText = "✅ 已套用今日即時匯率";
        } else {
          status.innerText = "⚠️ 無法取得即時匯率，暫用預設值";
        }
      } catch {
        status.innerText = "⚠️ 無法取得即時匯率，暫用預設值";
      }
    }

    function getRates() {
      return {
        KRW: parseFloat(document.getElementById("rate-KRW").value) || 0,
        JPY: parseFloat(document.getElementById("rate-JPY").value) || 0,
        CNY: parseFloat(document.getElementById("rate-CNY").value) || 0,
        USD: parseFloat(document.getElementById("rate-USD").value) || 0,
        AUD: parseFloat(document.getElementById("rate-AUD").value) || 0,
        TWD: 1
      };
    }

    async function fetchProducts() {
      const rawText = document.getElementById("urlInput").value.trim();
      if (!rawText) return alert("請先貼上至少一個網址！");

      const btn = document.getElementById("fetchBtn");
      const status = document.getElementById("statusText");
      btn.disabled = true;
      btn.innerText = "抓取中...";
      status.innerText = "正在向各網站讀取商品資料，請稍候...";

      try {
        const res = await fetch("/api/fetch", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text: rawText })
        });
        const data = await res.json();
        products = data.items;
        renderTable();
        status.innerText = `✨ 抓取完成！共 ${products.length} 件不重複商品`;
      } catch (err) {
        alert("抓取發生錯誤：" + err);
        status.innerText = "抓取失敗";
      } finally {
        btn.disabled = false;
        btn.innerText = "✨ 開始抓取商品";
      }
    }

    function addManualRow() {
      products.push({ url: "", title: "手動新增商品", img: "", price: 0, currency: "TWD" });
      renderTable();
    }

    function renderTable() {
      const tbody = document.getElementById("productTable");
      if (products.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center; color:#999; padding:36px;">找不到有效網址</td></tr>';
        recalculate();
        return;
      }

      tbody.innerHTML = products.map((p, idx) => `
        <tr>
          <td><input type="checkbox" class="item-check" data-idx="${idx}" checked onchange="recalculate()"></td>
          <td>${p.img ? `<a href="${p.url || '#'}" target="_blank"><img class="thumb" src="${p.img}" alt=""></a>` : '<span style="font-size:12px;color:#aaa;">無圖片</span>'}</td>
          <td>
            <input type="text" class="title-input" id="title-${idx}" value="${(p.title || '').replace(/"/g, '&quot;')}">
            ${p.url ? `<a class="link-sm" href="${p.url}" target="_blank">開啟原網頁 ↗</a>` : ''}
          </td>
          <td>
            <select class="curr-select" id="curr-${idx}" onchange="recalculate()">
              ${CURRENCIES.map(c => `<option value="${c}" ${p.currency === c ? 'selected' : ''}>${c}</option>`).join("")}
            </select>
          </td>
          <td><input type="number" class="price-input" id="price-${idx}" value="${p.price}" step="any" min="0" oninput="recalculate()"></td>
          <td class="twd-price" id="twd-${idx}">NT$0</td>
          <td><input type="number" class="qty-input" id="qty-${idx}" value="1" min="1" oninput="recalculate()"></td>
        </tr>
      `).join("");

      document.getElementById("selectAll").checked = true;
      recalculate();
    }

    function toggleAll(master) {
      document.querySelectorAll(".item-check").forEach(cb => cb.checked = master.checked);
      recalculate();
    }

    function recalculate() {
      const rates = getRates();
      let count = 0;
      let totalTwd = 0;

      document.querySelectorAll(".item-check").forEach(cb => {
        const idx = cb.dataset.idx;
        const curr = document.getElementById(`curr-${idx}`).value;
        const price = parseFloat(document.getElementById(`price-${idx}`).value) || 0;
        const qty = parseInt(document.getElementById(`qty-${idx}`).value) || 1;
        const unitTwd = Math.round(price * (rates[curr] || 1));

        document.getElementById(`twd-${idx}`).innerText = `NT$${unitTwd.toLocaleString()}`;

        if (cb.checked) {
          count += qty;
          totalTwd += unitTwd * qty;
        }
      });

      document.getElementById("selectedCount").innerText = count;
      document.getElementById("totalTwd").innerText = `NT$${totalTwd.toLocaleString()}`;
    }

    async function exportExcel() {
      const rates = getRates();
      const selectedItems = [];

      document.querySelectorAll(".item-check").forEach(cb => {
        if (cb.checked) {
          const idx = cb.dataset.idx;
          selectedItems.push({
            title: document.getElementById(`title-${idx}`).value,
            url: products[idx].url,
            img: products[idx].img,
            currency: document.getElementById(`curr-${idx}`).value,
            price: parseFloat(document.getElementById(`price-${idx}`).value) || 0,
            qty: parseInt(document.getElementById(`qty-${idx}`).value) || 1
          });
        }
      });

      if (selectedItems.length === 0) return alert("請至少勾選一件商品再匯出！");

      const btn = document.getElementById("exportBtn");
      const status = document.getElementById("statusText");
      btn.disabled = true;
      btn.innerText = "圖片打包中...";
      status.innerText = "正在下載商品照並產生 Excel...";

      try {
        const res = await fetch("/api/export", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ items: selectedItems, rates: rates })
        });
        if (!res.ok) throw new Error("匯出失敗");

        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "global_shopping_list.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
        status.innerText = "📊 Excel 匯出完成！";
      } catch (err) {
        alert("匯出發生錯誤：" + err);
        status.innerText = "匯出失敗";
      } finally {
        btn.disabled = false;
        btn.innerText = "📊 匯出含圖 Excel";
      }
    }
  </script>
</body>
</html>
"""

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7,ja;q=0.6,ko;q=0.5"
}

SUPPORTED_CURRENCIES = {"KRW", "JPY", "CNY", "USD", "AUD", "TWD"}

@app.route("/bg.jpg")
def background_image():
    img_path = os.path.join(BASE_DIR, "123.jpg")
    if os.path.exists(img_path):
        return send_file(img_path, mimetype="image/jpeg")
    return "", 404

@app.route("/api/rates")
def api_rates():
    try:
        res = requests.get("https://open.er-api.com/v6/latest/TWD", timeout=6)
        data = res.json()
        raw = data.get("rates", {})
        rates = {
            "KRW": round(1 / raw["KRW"], 4),
            "JPY": round(1 / raw["JPY"], 3),
            "CNY": round(1 / raw["CNY"], 2),
            "USD": round(1 / raw["USD"], 2),
            "AUD": round(1 / raw["AUD"], 2),
        }
        return jsonify({"ok": True, "rates": rates})
    except Exception:
        return jsonify({"ok": False})

def extract_urls(raw_text):
    parts = re.split(r'(?=https?://)', raw_text)
    seen = set()
    urls = []
    for part in parts:
        match = re.match(r'(https?://[^\s"\'<>]+)', part.strip())
        if match:
            clean_url = match.group(1).split('#')[0]
            if clean_url not in seen:
                seen.add(clean_url)
                urls.append(clean_url)
    return urls

def guess_currency_by_url(url):
    u = url.lower()
    if any(k in u for k in [".jp", "rakuten.co.jp", "amazon.co.jp", "zozo.jp", "mercari.com/jp"]):
        return "JPY"
    if any(k in u for k in ["taobao.com", "tmall.com", "1688.com", "jd.com", ".cn"]):
        return "CNY"
    if any(k in u for k in [".com.au", ".au/", "amazon.com.au"]):
        return "AUD"
    if any(k in u for k in [".com.tw", ".tw/", "momoshop", "pchome", "shopee.tw", "pinkoi"]):
        return "TWD"
    if any(k in u for k in [".kr", "lostmanagementcities", "musinsa", "29cm", "wconcept"]):
        return "KRW"
    return "USD"

def parse_json_ld(soup):
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "")
            items = data if isinstance(data, list) else data.get("@graph", [data])
            for obj in items:
                if not isinstance(obj, dict):
                    continue
                if obj.get("@type") in ("Product", ["Product"]):
                    title = obj.get("name", "")
                    img = obj.get("image", "")
                    if isinstance(img, list) and img:
                        img = img[0]
                    if isinstance(img, dict):
                        img = img.get("url", "")
                    offers = obj.get("offers", {})
                    if isinstance(offers, list) and offers:
                        offers = offers[0]
                    price = offers.get("price") or offers.get("lowPrice") or 0
                    curr = offers.get("priceCurrency", "")
                    return title, img, float(price) if price else 0, curr.upper()
        except Exception:
            continue
    return "", "", 0, ""

def scrape_product(url):
    currency = guess_currency_by_url(url)
    try:
        res = requests.get(url, headers=HEADERS, timeout=10)
        res.encoding = res.apparent_encoding or 'utf-8'
        soup = BeautifulSoup(res.text, 'html.parser')

        title, img_url, price_num, ld_curr = parse_json_ld(soup)
        if ld_curr in SUPPORTED_CURRENCIES:
            currency = ld_curr

        if not title:
            title_meta = soup.find("meta", property="og:title")
            title = title_meta["content"].strip() if title_meta and title_meta.get("content") else (soup.title.string.strip() if soup.title and soup.title.string else url)

        if not img_url:
            img_meta = soup.find("meta", property="og:image")
            img_url = img_meta["content"].strip() if img_meta and img_meta.get("content") else ""
        if img_url.startswith("//"):
            img_url = "https:" + img_url

        curr_meta = soup.find("meta", property=re.compile(r"(product|og):price:currency"))
        if curr_meta and curr_meta.get("content") and curr_meta["content"].upper() in SUPPORTED_CURRENCIES:
            currency = curr_meta["content"].upper()

        if not price_num:
            for prop in ["product:sale_price:amount", "product:price:amount", "og:price:amount"]:
                m = soup.find("meta", property=prop)
                if m and m.get("content"):
                    num_str = re.sub(r"[^\d.]", "", m["content"])
                    if num_str:
                        price_num = float(num_str)
                        break

        if not price_num:
            price_el = soup.select_one("#span_product_price_sale, #span_product_price_text, .price, [itemprop='price']")
            if price_el:
                val = price_el.get("content") or price_el.get_text()
                num_str = re.sub(r"[^\d.]", "", val)
                if num_str:
                    price_num = float(num_str)

        return {"url": url, "title": title, "img": img_url, "price": price_num, "currency": currency}
    except Exception as e:
        return {"url": url, "title": f"抓取受阻，請手動補填 ({type(e).__name__})", "img": "", "price": 0, "currency": currency}

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/api/fetch", methods=["POST"])
def api_fetch():
    raw_text = request.json.get("text", "")
    urls = extract_urls(raw_text)
    items = []
    for url in urls:
        items.append(scrape_product(url))
        time.sleep(0.3)
    return jsonify({"items": items})

@app.route("/api/export", methods=["POST"])
def api_export():
    data = request.json
    items = data.get("items", [])
    rates = data.get("rates", {})

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "跨國購物清單"

    col_headers = ["商品照", "商品名稱", "幣別", "原幣單價", "約合台幣 (TWD)", "數量", "小計 (TWD)", "商品連結"]
    ws.append(col_headers)

    header_fill = PatternFill(start_color="5C4742", end_color="5C4742", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, len(col_headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    col_widths = {"A": 15, "B": 40, "C": 8, "D": 13, "E": 15, "F": 8, "G": 15, "H": 36}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    img_buffers = []
    total_qty = 0
    total_twd = 0

    for row_idx, item in enumerate(items, start=2):
        curr = item.get("currency", "TWD")
        rate = float(rates.get(curr, 1))
        price = float(item.get("price", 0))
        qty = int(item.get("qty", 1))
        unit_twd = round(price * rate)
        subtotal_twd = unit_twd * qty

        total_qty += qty
        total_twd += subtotal_twd

        ws.row_dimensions[row_idx].height = 78

        ws.cell(row=row_idx, column=2, value=item.get("title", "")).alignment = Alignment(vertical="center", wrap_text=True)
        ws.cell(row=row_idx, column=3, value=curr).alignment = Alignment(horizontal="center", vertical="center")

        c_price = ws.cell(row=row_idx, column=4, value=price)
        c_price.number_format = '#,##0.##'
        c_price.alignment = Alignment(horizontal="right", vertical="center")

        c_twd = ws.cell(row=row_idx, column=5, value=unit_twd)
        c_twd.number_format = '#,##0'
        c_twd.alignment = Alignment(horizontal="right", vertical="center")

        c_qty = ws.cell(row=row_idx, column=6, value=qty)
        c_qty.alignment = Alignment(horizontal="center", vertical="center")

        c_sub = ws.cell(row=row_idx, column=7, value=subtotal_twd)
        c_sub.number_format = '#,##0'
        c_sub.alignment = Alignment(horizontal="right", vertical="center")

        url_val = item.get("url", "")
        c_url = ws.cell(row=row_idx, column=8, value=url_val)
        if url_val:
            c_url.hyperlink = url_val
            c_url.font = Font(color="0563C1", underline="single")
        c_url.alignment = Alignment(vertical="center")

        img_url = item.get("img", "")
        if img_url:
            try:
                img_res = requests.get(img_url, headers=HEADERS, timeout=8)
                pil_img = PILImage.open(io.BytesIO(img_res.content))
                pil_img.thumbnail((95, 95))
                buf = io.BytesIO()
                pil_img.save(buf, format="PNG")
                buf.seek(0)
                img_buffers.append(buf)
                ws.add_image(XLImage(buf), f"A{row_idx}")
            except Exception:
                ws.cell(row=row_idx, column=1, value="圖片載入失敗").alignment = Alignment(horizontal="center", vertical="center")

    sum_row = len(items) + 2
    ws.row_dimensions[sum_row].height = 28
    ws.cell(row=sum_row, column=5, value="台幣總計").font = Font(bold=True)
    ws.cell(row=sum_row, column=5).alignment = Alignment(horizontal="right", vertical="center")

    s_qty = ws.cell(row=sum_row, column=6, value=total_qty)
    s_qty.font = Font(bold=True)
    s_qty.alignment = Alignment(horizontal="center", vertical="center")

    s_twd = ws.cell(row=sum_row, column=7, value=total_twd)
    s_twd.number_format = '#,##0'
    s_twd.font = Font(bold=True, color="FB6F92")
    s_twd.alignment = Alignment(horizontal="right", vertical="center")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="global_shopping_list.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    if not os.environ.get("RENDER"):
        threading.Timer(1.0, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    app.run(host="0.0.0.0", port=port, debug=False)