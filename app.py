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

HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>吉伊卡哇跨國團購神器 🛒</title>
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
    .container { max-width: 1220px; margin: 0 auto; }
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
      background: rgba(255, 255, 255, 0.9);
      backdrop-filter: blur(10px);
      -webkit-backdrop-filter: blur(10px);
      border: 2px solid rgba(255, 179, 198, 0.75);
      border-radius: 20px;
      padding: 18px 20px;
      box-shadow: 0 8px 24px rgba(92, 71, 66, 0.1);
      margin-bottom: 16px;
    }
    textarea {
      width: 100%;
      height: 95px;
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
    .controls { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin-top: 12px; }
    button {
      background: #ff8fab;
      color: #fff;
      border: none;
      padding: 9px 15px;
      border-radius: 50px;
      cursor: pointer;
      font-size: 13px;
      font-weight: 700;
      box-shadow: 0 3px 8px rgba(255, 143, 171, 0.35);
      transition: transform 0.1s, background 0.2s;
    }
    button:hover { background: #fb6f92; transform: translateY(-1px); }
    button:disabled { background: #ccc; box-shadow: none; cursor: not-allowed; }
    .btn-add { background: #8ecae6; box-shadow: 0 3px 8px rgba(142, 202, 230, 0.4); }
    .btn-add:hover { background: #219ebc; }
    .btn-line { background: #06c755; box-shadow: 0 3px 8px rgba(6, 199, 85, 0.35); }
    .btn-line:hover { background: #05a646; }
    .btn-export { background: #70c1b3; box-shadow: 0 3px 8px rgba(112, 193, 179, 0.4); }
    .btn-export:hover { background: #4ea898; }
    .btn-rate { background: #ffd166; color: #5c4742; box-shadow: 0 3px 8px rgba(255, 209, 102, 0.4); }
    .btn-rate:hover { background: #f4be49; }
    .btn-clear { background: #e5989b; }
    .btn-clear:hover { background: #d6686e; }

    .rates-bar, .extra-bar {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 12px;
      padding-top: 12px;
      border-top: 2px dashed #ffe5ec;
      align-items: center;
    }
    .rate-item {
      display: flex;
      align-items: center;
      gap: 5px;
      background: #fff0f3;
      padding: 5px 11px;
      border-radius: 50px;
      font-size: 13px;
      font-weight: 700;
      border: 1px solid #ffccd5;
    }
    .rate-item input[type="number"] {
      width: 66px;
      padding: 3px 5px;
      border: 1px solid #ffb3c6;
      border-radius: 8px;
      text-align: right;
      font-weight: bold;
      color: #5c4742;
      background: #fff;
    }
    .extra-item {
      display: flex;
      align-items: center;
      gap: 6px;
      background: #fef9ef;
      padding: 6px 12px;
      border-radius: 50px;
      font-size: 13px;
      font-weight: 700;
      border: 1px solid #ffd166;
    }
    .extra-item input[type="number"] {
      width: 78px;
      padding: 3px 6px;
      border: 1px solid #f4be49;
      border-radius: 8px;
      text-align: right;
      font-weight: bold;
      color: #5c4742;
    }
    .summary {
      position: sticky;
      top: 10px;
      z-index: 10;
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 10px;
      background: rgba(92, 71, 66, 0.95);
      backdrop-filter: blur(8px);
      color: #fff;
      padding: 13px 22px;
      border-radius: 50px;
      margin-bottom: 16px;
      border: 2px solid #ffb3c6;
      box-shadow: 0 6px 20px rgba(92, 71, 66, 0.25);
    }
    .summary strong { font-size: 21px; color: #ffd166; margin-left: 4px; }
    .split-box {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
    }
    .buyer-pill {
      background: #fff0f3;
      border: 2px solid #ffb3c6;
      border-radius: 14px;
      padding: 8px 14px;
      font-size: 13px;
      font-weight: 700;
      color: #5c4742;
    }
    .buyer-pill span { color: #fb6f92; font-size: 16px; margin-left: 6px; }
    .table-wrap {
      background: rgba(255, 255, 255, 0.92);
      backdrop-filter: blur(10px);
      border-radius: 20px;
      border: 2px solid rgba(255, 179, 198, 0.75);
      overflow-x: auto;
      box-shadow: 0 8px 24px rgba(92, 71, 66, 0.1);
    }
    table { width: 100%; border-collapse: collapse; }
    th, td { padding: 10px; border-bottom: 1px solid #ffe5ec; text-align: left; vertical-align: middle; }
    th { background: rgba(255, 240, 243, 0.95); font-size: 13px; color: #5c4742; font-weight: 800; white-space: nowrap; }
    .thumb-box { display: flex; flex-direction: column; align-items: center; gap: 4px; }
    .thumb { width: 72px; height: 72px; object-fit: cover; border-radius: 10px; border: 2px solid #ffe5ec; background: #fff; }
    .btn-img { font-size: 11px; padding: 2px 8px; background: #ffe5ec; color: #5c4742; box-shadow: none; }
    .btn-img:hover { background: #ffb3c6; color: #fff; }
    .title-input { width: 100%; min-width: 160px; padding: 5px 7px; border: 1px solid transparent; border-radius: 8px; font-size: 14px; background: transparent; color: #5c4742; font-weight: 600; }
    .title-input:hover, .title-input:focus { border-color: #ffb3c6; background: #fff; outline: none; }
    .buyer-input { width: 76px; padding: 5px 7px; border: 1px solid #ffb3c6; border-radius: 8px; font-size: 13px; color: #5c4742; font-weight: 700; text-align: center; background: #fff0f3; }
    .curr-select { padding: 5px 6px; border: 1px solid #ffb3c6; border-radius: 8px; font-size: 13px; font-weight: bold; background: #fff; color: #5c4742; }
    .price-input { width: 88px; padding: 5px; border: 1px solid #ffb3c6; border-radius: 8px; text-align: right; font-size: 14px; color: #5c4742; }
    .qty-input { width: 50px; padding: 5px; border: 1px solid #ffb3c6; border-radius: 8px; text-align: center; font-size: 14px; color: #5c4742; }
    .twd-price { font-weight: 800; color: #fb6f92; white-space: nowrap; font-size: 14px; }
    .btn-del { background: transparent; color: #bbb; box-shadow: none; padding: 4px 8px; font-size: 15px; }
    .btn-del:hover { background: #ffe5ec; color: #fb6f92; }
    .link-sm { font-size: 12px; color: #219ebc; text-decoration: none; display: inline-block; margin-top: 2px; font-weight: 600; }
    .link-sm:hover { text-decoration: underline; }
    .status { font-size: 13px; color: #5c4742; margin-left: auto; font-weight: 700; }
  </style>
</head>
<body>
  <div class="container">
    <h1 class="header-title">🛍️ 吉伊卡哇跨國團購神器</h1>

    <div class="card">
      <textarea id="urlInput" placeholder="貼上各國商品網址（韓、日、中、台、美、澳皆可，一行一個，會自動去重並加入下方清單）..."></textarea>
      <div class="controls">
        <button id="fetchBtn" onclick="fetchProducts()">✨ 抓取並加入清單</button>
        <button class="btn-add" onclick="addManualRow()">+ 手動新增</button>
        <button class="btn-line" onclick="copyLineText()">📋 複製 LINE 報價</button>
        <button id="exportBtn" class="btn-export" onclick="exportExcel()">📊 匯出含圖 Excel</button>
        <button class="btn-rate" onclick="loadLiveRates()">🔄 更新匯率</button>
        <button class="btn-clear" onclick="clearAll()">🗑️ 清空</button>
        <span id="statusText" class="status">已開啟自動存檔 ✨</span>
      </div>

      <div class="rates-bar">
        <div class="rate-item"><span>🇰🇷 KRW</span><input type="number" id="rate-KRW" value="0.0235" step="0.0005" oninput="recalculate()"></div>
        <div class="rate-item"><span>🇯🇵 JPY</span><input type="number" id="rate-JPY" value="0.215" step="0.005" oninput="recalculate()"></div>
        <div class="rate-item"><span>🇨🇳 CNY</span><input type="number" id="rate-CNY" value="4.50" step="0.05" oninput="recalculate()"></div>
        <div class="rate-item"><span>🇺🇸 USD</span><input type="number" id="rate-USD" value="32.5" step="0.1" oninput="recalculate()"></div>
        <div class="rate-item"><span>🇦🇺 AUD</span><input type="number" id="rate-AUD" value="21.2" step="0.1" oninput="recalculate()"></div>
        <div class="rate-item"><span>🇹🇼 TWD</span><input type="number" id="rate-TWD" value="1" disabled></div>
      </div>

      <div class="extra-bar">
        <label class="extra-item" style="cursor:pointer;">
          <input type="checkbox" id="feeCheck" onchange="recalculate()">
          <span>💳 加上 1.5% 海外刷卡手續費</span>
        </label>
        <div class="extra-item">
          <span>✈️ 國際運費總計 NT$</span>
          <input type="number" id="shippingInput" value="0" min="0" step="10" oninput="recalculate()">
          <span style="font-size:12px; color:#888;">(依勾選件數均分)</span>
        </div>
      </div>
    </div>

    <div class="summary">
      <div>已勾選 <strong id="selectedCount">0</strong> 件商品</div>
      <div>折合台幣總計：<strong id="totalTwd">NT$0</strong></div>
    </div>

    <div class="card" id="splitCard" style="display:none; padding: 14px 18px;">
      <div style="font-size:13px; font-weight:800; margin-bottom:8px; color:#5c4742;">👥 團購自動分帳（已含手續費與運費均分）：</div>
      <div class="split-box" id="splitBox"></div>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th style="width: 34px;"><input type="checkbox" id="selectAll" checked onchange="toggleAll(this)"></th>
            <th style="width: 90px; text-align:center;">商品照</th>
            <th>商品名稱（點擊可改）</th>
            <th style="width: 88px; text-align:center;">買家</th>
            <th style="width: 82px;">幣別</th>
            <th style="width: 100px;">原幣單價</th>
            <th style="width: 100px;">台幣單價</th>
            <th style="width: 62px; text-align:center;">數量</th>
            <th style="width: 38px;"></th>
          </tr>
        </thead>
        <tbody id="productTable">
          <tr><td colspan="9" style="text-align:center; color:#999; padding:36px;">在上方貼上網址，點擊「抓取並加入清單」就能開始逛街啦 🛒</td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <script>
    let products = [];
    const CURRENCIES = ["KRW", "JPY", "CNY", "USD", "AUD", "TWD"];
    const STORAGE_KEY = "chiikawa_shop_v2";

    window.addEventListener("DOMContentLoaded", () => {
      loadState();
      loadLiveRates(false);
    });

    function saveState() {
      const state = {
        products: products,
        rates: getRates(),
        feeChecked: document.getElementById("feeCheck").checked,
        shipping: parseFloat(document.getElementById("shippingInput").value) || 0
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    }

    function loadState() {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (!saved) return;
      try {
        const state = JSON.parse(saved);
        if (Array.isArray(state.products)) products = state.products;
        if (state.rates) {
          for (const [k, v] of Object.entries(state.rates)) {
            const el = document.getElementById(`rate-${k}`);
            if (el && k !== "TWD") el.value = v;
          }
        }
        if (typeof state.feeChecked === "boolean") {
          document.getElementById("feeCheck").checked = state.feeChecked;
        }
        if (typeof state.shipping === "number") {
          document.getElementById("shippingInput").value = state.shipping;
        }
        renderTable();
      } catch {}
    }

    async function loadLiveRates(showMsg = true) {
      const status = document.getElementById("statusText");
      if (showMsg) status.innerText = "正在抓取今日最新匯率...";
      try {
        const res = await fetch("/api/rates");
        const data = await res.json();
        if (data.ok) {
          for (const [curr, val] of Object.entries(data.rates)) {
            const el = document.getElementById(`rate-${curr}`);
            if (el) el.value = val;
          }
          recalculate();
          if (showMsg) status.innerText = "✅ 已套用今日即時匯率";
        }
      } catch {
        if (showMsg) status.innerText = "⚠️ 匯率更新失敗，維持原設定";
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
        const existingUrls = new Set(products.map(p => p.url).filter(Boolean));
        let added = 0;
        data.items.forEach(item => {
          if (!item.url || !existingUrls.has(item.url)) {
            products.push({ ...item, buyer: "我", qty: 1, checked: true });
            added++;
          }
        });
        document.getElementById("urlInput").value = "";
        renderTable();
        status.innerText = `✨ 新增了 ${added} 件商品（共 ${products.length} 件）`;
      } catch (err) {
        alert("抓取發生錯誤：" + err);
        status.innerText = "抓取失敗";
      } finally {
        btn.disabled = false;
        btn.innerText = "✨ 抓取並加入清單";
      }
    }

    function addManualRow() {
      products.push({ url: "", title: "手動新增商品", img: "", price: 0, currency: "TWD", buyer: "我", qty: 1, checked: true });
      renderTable();
    }

    function editImage(idx) {
      const current = products[idx].img || "";
      const newUrl = prompt("請貼上商品圖片網址 (右鍵複製圖片位址)：", current);
      if (newUrl !== null) {
        products[idx].img = newUrl.trim();
        renderTable();
      }
    }

    function deleteRow(idx) {
      products.splice(idx, 1);
      renderTable();
    }

    function clearAll() {
      if (products.length === 0) return;
      if (confirm("確定要清空所有商品清單嗎？")) {
        products = [];
        document.getElementById("shippingInput").value = 0;
        renderTable();
        document.getElementById("statusText").innerText = "已清空清單";
      }
    }

    function escapeHtml(str) {
      return String(str || "").replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

    function renderTable() {
      const tbody = document.getElementById("productTable");
      if (products.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" style="text-align:center; color:#999; padding:36px;">目前清單是空的，快去貼上網址或手動新增吧 🛒</td></tr>';
        recalculate();
        return;
      }

      tbody.innerHTML = products.map((p, idx) => `
        <tr>
          <td><input type="checkbox" class="item-check" data-idx="${idx}" ${p.checked !== false ? 'checked' : ''} onchange="syncAndCalc(${idx})"></td>
          <td>
            <div class="thumb-box">
              ${p.img ? `<a href="${p.url || p.img}" target="_blank"><img class="thumb" src="${p.img}" alt=""></a>` : '<div class="thumb" style="display:flex;align-items:center;justify-content:center;font-size:11px;color:#aaa;">無圖片</div>'}
              <button class="btn-img" onclick="editImage(${idx})">換圖</button>
            </div>
          </td>
          <td>
            <input type="text" class="title-input" id="title-${idx}" value="${escapeHtml(p.title)}" oninput="syncAndCalc(${idx})">
            ${p.url ? `<a class="link-sm" href="${p.url}" target="_blank">開啟原網頁 ↗</a>` : ''}
          </td>
          <td><input type="text" class="buyer-input" id="buyer-${idx}" value="${escapeHtml(p.buyer || '我')}" oninput="syncAndCalc(${idx})"></td>
          <td>
            <select class="curr-select" id="curr-${idx}" onchange="syncAndCalc(${idx})">
              ${CURRENCIES.map(c => `<option value="${c}" ${p.currency === c ? 'selected' : ''}>${c}</option>`).join("")}
            </select>
          </td>
          <td><input type="number" class="price-input" id="price-${idx}" value="${p.price}" step="any" min="0" oninput="syncAndCalc(${idx})"></td>
          <td class="twd-price" id="twd-${idx}">NT$0</td>
          <td><input type="number" class="qty-input" id="qty-${idx}" value="${p.qty || 1}" min="1" oninput="syncAndCalc(${idx})"></td>
          <td><button class="btn-del" title="刪除" onclick="deleteRow(${idx})">✕</button></td>
        </tr>
      `).join("");

      recalculate();
    }

    function syncAndCalc(idx) {
      const cb = document.querySelector(`.item-check[data-idx="${idx}"]`);
      if (!cb) return;
      products[idx].checked = cb.checked;
      products[idx].title = document.getElementById(`title-${idx}`).value;
      products[idx].buyer = document.getElementById(`buyer-${idx}`).value.trim() || "我";
      products[idx].currency = document.getElementById(`curr-${idx}`).value;
      products[idx].price = parseFloat(document.getElementById(`price-${idx}`).value) || 0;
      products[idx].qty = parseInt(document.getElementById(`qty-${idx}`).value) || 1;
      recalculate();
    }

    function toggleAll(master) {
      products.forEach((p, idx) => p.checked = master.checked);
      document.querySelectorAll(".item-check").forEach(cb => cb.checked = master.checked);
      recalculate();
    }

    function getBreakdown() {
      const rates = getRates();
      const feeMult = document.getElementById("feeCheck").checked ? 1.015 : 1;
      const shipping = parseFloat(document.getElementById("shippingInput").value) || 0;

      let totalQty = 0;
      let totalGoodsTwd = 0;
      const buyers = {};

      products.forEach((p, idx) => {
        const unitTwd = Math.round((p.price || 0) * (rates[p.currency] || 1) * feeMult);
        const twdEl = document.getElementById(`twd-${idx}`);
        if (twdEl) twdEl.innerText = `NT$${unitTwd.toLocaleString()}`;

        if (p.checked !== false) {
          const qty = p.qty || 1;
          const buyer = p.buyer || "我";
          totalQty += qty;
          totalGoodsTwd += unitTwd * qty;

          if (!buyers[buyer]) buyers[buyer] = { qty: 0, goodsTwd: 0, items: [] };
          buyers[buyer].qty += qty;
          buyers[buyer].goodsTwd += unitTwd * qty;
          buyers[buyer].items.push({ title: p.title, qty, unitTwd, subTwd: unitTwd * qty });
        }
      });

      const shippingPerItem = totalQty > 0 ? shipping / totalQty : 0;
      for (const b of Object.values(buyers)) {
        b.shippingShare = Math.round(shippingPerItem * b.qty);
        b.finalTwd = b.goodsTwd + b.shippingShare;
      }

      return { totalQty, totalTwd: totalGoodsTwd + Math.round(shipping), shipping, feeChecked: feeMult > 1, buyers };
    }

    function recalculate() {
      const { totalQty, totalTwd, buyers } = getBreakdown();
      document.getElementById("selectedCount").innerText = totalQty;
      document.getElementById("totalTwd").innerText = `NT$${totalTwd.toLocaleString()}`;

      const splitCard = document.getElementById("splitCard");
      const splitBox = document.getElementById("splitBox");
      const buyerEntries = Object.entries(buyers);

      if (buyerEntries.length > 0) {
        splitCard.style.display = "block";
        splitBox.innerHTML = buyerEntries.map(([name, info]) => `
          <div class="buyer-pill">
            👤 ${escapeHtml(name)} (${info.qty}件)
            <span>NT$${info.finalTwd.toLocaleString()}</span>
          </div>
        `).join("");
      } else {
        splitCard.style.display = "none";
      }

      saveState();
    }

    function copyLineText() {
      const { totalQty, totalTwd, shipping, feeChecked, buyers } = getBreakdown();
      if (totalQty === 0) return alert("請至少勾選一件商品！");

      let lines = ["🛍️ 跨國團購結帳清單"];
      if (feeChecked || shipping > 0) {
        let notes = [];
        if (feeChecked) notes.push("含1.5%刷卡費");
        if (shipping > 0) notes.push(`運費NT$${shipping}按件均分`);
        lines.push(`(${notes.join(" / ")})`);
      }
      lines.push("━━━━━━━━━━━━");

      for (const [name, info] of Object.entries(buyers)) {
        lines.push(`👤 ${name}（${info.qty}件）`);
        info.items.forEach(it => {
          lines.push(`• ${it.title} x${it.qty} — NT$${it.subTwd.toLocaleString()}`);
        });
        if (info.shippingShare > 0) {
          lines.push(`  ✈️ 運費分攤：NT$${info.shippingShare.toLocaleString()}`);
        }
        lines.push(`👉 應付小計：NT$${info.finalTwd.toLocaleString()}`);
        lines.push("");
      }

      lines.push("━━━━━━━━━━━━");
      lines.push(`💰 總計：NT$${totalTwd.toLocaleString()}（共 ${totalQty} 件）`);

      navigator.clipboard.writeText(lines.join("\n")).then(() => {
        document.getElementById("statusText").innerText = "📋 已複製 LINE 報價清單！直接去貼上吧";
      }).catch(() => {
        alert("複製失敗，請手動複製：\n\n" + lines.join("\n"));
      });
    }

    async function exportExcel() {
      const selectedItems = products.filter(p => p.checked !== false);
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
          body: JSON.stringify({
            items: selectedItems,
            rates: getRates(),
            feeChecked: document.getElementById("feeCheck").checked,
            shipping: parseFloat(document.getElementById("shippingInput").value) || 0
          })
        });
        if (!res.ok) throw new Error("匯出失敗");

        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "group_shopping_list.xlsx";
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

def extract_first_number(text):
    m = re.search(r"[\d,]+(?:\.\d+)?", str(text))
    if m:
        try:
            return float(m.group(0).replace(",", ""))
        except ValueError:
            return 0
    return 0

def scrape_product(url):
    currency = guess_currency_by_url(url)
    try:
        res = requests.get(url, headers=HEADERS, timeout=10)
        res.encoding = res.apparent_encoding or 'utf-8'
        soup = BeautifulSoup(res.text, 'html.parser')

        title, img_url, ld_price, ld_curr = parse_json_ld(soup)
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

        # 1. 優先抓「特價 / 折扣價」（Cafe24 專屬標籤與變數）
        price_num = 0
        sale_meta = soup.find("meta", property="product:sale_price:amount")
        if sale_meta and sale_meta.get("content"):
            price_num = extract_first_number(sale_meta["content"])

        if not price_num:
            sale_el = soup.select_one("#span_product_price_sale")
            if sale_el:
                price_num = extract_first_number(sale_el.get_text())

        if not price_num:
            sale_js = re.search(r"product_sale_price\s*=\s*['\"]?(\d+(?:\.\d+)?)['\"]?", res.text)
            if sale_js:
                price_num = float(sale_js.group(1))

        # 2. 沒有特價時，才改抓 JSON-LD 或原價
        if not price_num and ld_price:
            price_num = ld_price

        if not price_num:
            for prop in ["product:price:amount", "og:price:amount"]:
                m = soup.find("meta", property=prop)
                if m and m.get("content"):
                    price_num = extract_first_number(m["content"])
                    if price_num:
                        break

        if not price_num:
            price_el = soup.select_one("#span_product_price_text, .price, [itemprop='price']")
            if price_el:
                val = price_el.get("content") or price_el.get_text()
                price_num = extract_first_number(val)

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
    fee_mult = 1.015 if data.get("feeChecked") else 1.0
    shipping = float(data.get("shipping", 0))

    total_qty = sum(int(it.get("qty", 1)) for it in items)
    shipping_per_item = shipping / total_qty if total_qty > 0 else 0

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "跨國團購清單"

    col_headers = ["商品照", "買家", "商品名稱", "幣別", "原幣單價", "台幣單價", "數量", "運費分攤", "小計 (TWD)", "商品連結"]
    ws.append(col_headers)

    header_fill = PatternFill(start_color="5C4742", end_color="5C4742", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, len(col_headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    col_widths = {"A": 15, "B": 11, "C": 36, "D": 8, "E": 12, "F": 12, "G": 7, "H": 11, "I": 13, "J": 32}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    img_buffers = []
    grand_total = 0

    for row_idx, item in enumerate(items, start=2):
        curr = item.get("currency", "TWD")
        rate = float(rates.get(curr, 1))
        price = float(item.get("price", 0))
        qty = int(item.get("qty", 1))
        unit_twd = round(price * rate * fee_mult)
        ship_share = round(shipping_per_item * qty)
        subtotal_twd = unit_twd * qty + ship_share
        grand_total += subtotal_twd

        ws.row_dimensions[row_idx].height = 78

        ws.cell(row=row_idx, column=2, value=item.get("buyer", "我")).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row_idx, column=3, value=item.get("title", "")).alignment = Alignment(vertical="center", wrap_text=True)
        ws.cell(row=row_idx, column=4, value=curr).alignment = Alignment(horizontal="center", vertical="center")

        for col_i, val, fmt in [
            (5, price, '#,##0.##'),
            (6, unit_twd, '#,##0'),
            (8, ship_share, '#,##0'),
            (9, subtotal_twd, '#,##0')
        ]:
            c = ws.cell(row=row_idx, column=col_i, value=val)
            c.number_format = fmt
            c.alignment = Alignment(horizontal="right", vertical="center")

        ws.cell(row=row_idx, column=7, value=qty).alignment = Alignment(horizontal="center", vertical="center")

        url_val = item.get("url", "")
        c_url = ws.cell(row=row_idx, column=10, value=url_val)
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
    ws.cell(row=sum_row, column=6, value="總計").font = Font(bold=True)
    ws.cell(row=sum_row, column=6).alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(row=sum_row, column=7, value=total_qty).font = Font(bold=True)
    ws.cell(row=sum_row, column=7).alignment = Alignment(horizontal="center", vertical="center")

    s_ship = ws.cell(row=sum_row, column=8, value=round(shipping))
    s_ship.number_format = '#,##0'
    s_ship.font = Font(bold=True)
    s_ship.alignment = Alignment(horizontal="right", vertical="center")

    s_twd = ws.cell(row=sum_row, column=9, value=grand_total)
    s_twd.number_format = '#,##0'
    s_twd.font = Font(bold=True, color="FB6F92")
    s_twd.alignment = Alignment(horizontal="right", vertical="center")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="group_shopping_list.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    if not os.environ.get("RENDER"):
        threading.Timer(1.0, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    app.run(host="0.0.0.0", port=port, debug=False)
