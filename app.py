import io
import os
import re
import json
import time
import threading
import webbrowser
from urllib.parse import quote
import requests
from bs4 import BeautifulSoup
from PIL import Image as PILImage
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from flask import Flask, request, jsonify, render_template_string, send_file

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

FIREBASE_URL = os.environ.get("FIREBASE_URL", "").strip().rstrip("/")
ADMIN_PIN = os.environ.get("ADMIN_PIN", "1234").strip()

LOCAL_ROOMS = {}

def default_room_state():
    return {
        "products": [],
        "rates": {"KRW": 0.0235, "JPY": 0.215, "CNY": 4.50, "USD": 32.5, "AUD": 21.2, "TWD": 1},
        "feeChecked": False,
        "shipping": 0,
        "locked": False,
        "updatedAt": 0
    }

def sanitize_room(room_name):
    name = (room_name or "大廳").strip()
    safe = re.sub(r'[.$#\[\]/]', '_', name)
    return safe[:30] if safe else "大廳"

HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>吉伊卡哇多人團購神器 🛒</title>
  <style>
    * { box-sizing: border-box; font-family: "Nunito", "Noto Sans TC", -apple-system, BlinkMacSystemFont, sans-serif; }
    body {
      margin: 0;
      padding: 20px 14px 60px;
      color: #5c4742;
      background-image: url('/bg.jpg');
      background-size: cover;
      background-position: center;
      background-attachment: fixed;
      min-height: 100vh;
    }
    .container { max-width: 1280px; margin: 0 auto; }
    .top-bar {
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 10px;
      margin-bottom: 12px;
    }
    .header-title {
      display: inline-block;
      background: rgba(255, 255, 255, 0.92);
      padding: 8px 20px;
      border-radius: 50px;
      font-size: 20px;
      font-weight: 800;
      color: #5c4742;
      border: 3px solid #ffb3c6;
      box-shadow: 0 4px 12px rgba(255, 143, 171, 0.25);
      margin: 0;
    }
    .user-bar {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 8px;
      background: rgba(255, 255, 255, 0.92);
      padding: 6px 14px;
      border-radius: 50px;
      border: 2px solid #ffb3c6;
      font-size: 13px;
      font-weight: 700;
    }
    .user-bar input {
      width: 82px;
      padding: 4px 8px;
      border: 1px solid #ffb3c6;
      border-radius: 10px;
      font-weight: 700;
      color: #5c4742;
      text-align: center;
    }
    .room-bar {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      background: rgba(255, 255, 255, 0.92);
      backdrop-filter: blur(8px);
      padding: 10px 16px;
      border-radius: 18px;
      border: 2px solid #ffb3c6;
      margin-bottom: 14px;
      font-size: 13px;
      font-weight: 700;
    }
    .room-controls { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; }
    .room-controls input {
      width: 130px;
      padding: 5px 10px;
      border: 2px solid #ffe5ec;
      border-radius: 12px;
      font-weight: 700;
      color: #5c4742;
      outline: none;
    }
    .room-controls input:focus { border-color: #ff8fab; }
    .online-box {
      display: flex;
      align-items: center;
      gap: 6px;
      background: #e8f8f5;
      color: #1b7a68;
      padding: 5px 12px;
      border-radius: 50px;
      border: 1px solid #a3e4d7;
      font-size: 12px;
      font-weight: 800;
    }
    .lock-banner {
      background: #ffccd5;
      color: #800f2f;
      border: 2px solid #ff8fab;
      border-radius: 14px;
      padding: 10px 16px;
      margin-bottom: 14px;
      font-size: 14px;
      font-weight: 800;
      text-align: center;
      display: none;
    }
    .card {
      background: rgba(255, 255, 255, 0.9);
      backdrop-filter: blur(10px);
      -webkit-backdrop-filter: blur(10px);
      border: 2px solid rgba(255, 179, 198, 0.75);
      border-radius: 20px;
      padding: 16px 18px;
      box-shadow: 0 8px 24px rgba(92, 71, 66, 0.1);
      margin-bottom: 14px;
    }
    textarea {
      width: 100%;
      height: 88px;
      padding: 11px;
      border: 2px solid #ffe5ec;
      border-radius: 14px;
      font-size: 13px;
      resize: vertical;
      background: rgba(255, 255, 255, 0.95);
      color: #5c4742;
      outline: none;
    }
    textarea:focus { border-color: #ff8fab; }
    .controls { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin-top: 10px; }
    button {
      background: #ff8fab;
      color: #fff;
      border: none;
      padding: 8px 14px;
      border-radius: 50px;
      cursor: pointer;
      font-size: 13px;
      font-weight: 700;
      box-shadow: 0 3px 8px rgba(255, 143, 171, 0.35);
      transition: transform 0.1s, background 0.2s;
    }
    button:hover { background: #fb6f92; transform: translateY(-1px); }
    button:disabled { background: #ccc; box-shadow: none; cursor: not-allowed; transform: none; }
    .btn-sm { padding: 6px 12px; font-size: 12px; }
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
    .btn-admin { background: #5c4742; color: #ffd166; padding: 6px 12px; font-size: 12px; }
    .btn-lock-on { background: #e63946; padding: 6px 12px; font-size: 12px; }
    .btn-lock-off { background: #2a9d8f; padding: 6px 12px; font-size: 12px; }

    .rates-bar, .extra-bar {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 10px;
      padding-top: 10px;
      border-top: 2px dashed #ffe5ec;
      align-items: center;
    }
    .rate-item {
      display: flex;
      align-items: center;
      gap: 5px;
      background: #fff0f3;
      padding: 4px 10px;
      border-radius: 50px;
      font-size: 12px;
      font-weight: 700;
      border: 1px solid #ffccd5;
    }
    .rate-item input[type="number"] {
      width: 62px;
      padding: 2px 4px;
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
      padding: 5px 12px;
      border-radius: 50px;
      font-size: 13px;
      font-weight: 700;
      border: 1px solid #ffd166;
    }
    .extra-item input[type="number"] {
      width: 74px;
      padding: 2px 5px;
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
      padding: 12px 22px;
      border-radius: 50px;
      margin-bottom: 14px;
      border: 2px solid #ffb3c6;
      box-shadow: 0 6px 20px rgba(92, 71, 66, 0.25);
    }
    .summary strong { font-size: 20px; color: #ffd166; margin-left: 4px; }
    .split-box { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
    .buyer-pill {
      background: #fff0f3;
      border: 2px solid #ffb3c6;
      border-radius: 14px;
      padding: 6px 12px;
      font-size: 13px;
      font-weight: 700;
      color: #5c4742;
    }
    .buyer-pill span { color: #fb6f92; font-size: 15px; margin-left: 5px; }
    .table-wrap {
      background: rgba(255, 255, 255, 0.92);
      backdrop-filter: blur(10px);
      border-radius: 20px;
      border: 2px solid rgba(255, 179, 198, 0.75);
      overflow-x: auto;
      box-shadow: 0 8px 24px rgba(92, 71, 66, 0.1);
    }
    table { width: 100%; border-collapse: collapse; }
    th, td { padding: 9px 8px; border-bottom: 1px solid #ffe5ec; text-align: left; vertical-align: middle; }
    th { background: rgba(255, 240, 243, 0.95); font-size: 12px; color: #5c4742; font-weight: 800; white-space: nowrap; }
    .thumb-box { display: flex; flex-direction: column; align-items: center; gap: 3px; }
    .thumb { width: 68px; height: 68px; object-fit: cover; border-radius: 10px; border: 2px solid #ffe5ec; background: #fff; }
    .btn-img { font-size: 11px; padding: 2px 7px; background: #ffe5ec; color: #5c4742; box-shadow: none; }
    .btn-img:hover { background: #ffb3c6; color: #fff; }
    .title-input { width: 100%; min-width: 150px; padding: 4px 6px; border: 1px solid transparent; border-radius: 8px; font-size: 13px; background: transparent; color: #5c4742; font-weight: 600; }
    .title-input:hover, .title-input:focus { border-color: #ffb3c6; background: #fff; outline: none; }
    .spec-stack { display: flex; flex-direction: column; gap: 5px; }
    .spec-input, .spec-select {
      width: 145px;
      padding: 5px 6px;
      border: 1px solid #ffb3c6;
      border-radius: 8px;
      font-size: 12px;
      color: #5c4742;
      background: #fff;
    }
    .note-input {
      width: 145px;
      padding: 4px 6px;
      border: 1px dashed #ffb3c6;
      border-radius: 8px;
      font-size: 12px;
      color: #5c4742;
      background: #fff8fa;
    }
    .buyer-input { width: 68px; padding: 5px; border: 1px solid #ffb3c6; border-radius: 8px; font-size: 12px; color: #5c4742; font-weight: 700; text-align: center; background: #fff0f3; }
    .curr-select { padding: 4px 5px; border: 1px solid #ffb3c6; border-radius: 8px; font-size: 12px; font-weight: bold; background: #fff; color: #5c4742; }
    .price-input { width: 82px; padding: 5px; border: 1px solid #ffb3c6; border-radius: 8px; text-align: right; font-size: 13px; color: #5c4742; }
    .qty-input { width: 46px; padding: 5px; border: 1px solid #ffb3c6; border-radius: 8px; text-align: center; font-size: 13px; color: #5c4742; }
    .twd-price { font-weight: 800; color: #fb6f92; white-space: nowrap; font-size: 13px; }
    .btn-del { background: transparent; color: #bbb; box-shadow: none; padding: 3px 6px; font-size: 14px; }
    .btn-del:hover { background: #ffe5ec; color: #fb6f92; }
    .link-sm { font-size: 11px; color: #219ebc; text-decoration: none; display: inline-block; margin-top: 2px; font-weight: 600; }
    .link-sm:hover { text-decoration: underline; }
    .status { font-size: 12px; color: #5c4742; margin-left: auto; font-weight: 700; }
  </style>
</head>
<body>
  <div class="container">
    <div class="top-bar">
      <h1 class="header-title">🛍️ 吉伊卡哇多人團購神器</h1>
      <div class="user-bar">
        <span>👤 你的名字：</span>
        <input type="text" id="myNameInput" placeholder="輸入暱稱" oninput="saveMyName()">
        <span id="syncBadge" style="font-size:12px; color:#2a9d8f;">☁️ 連線中...</span>
        <button id="adminBtn" class="btn-admin" onclick="toggleAdminLogin()">👑 團主管理</button>
        <button id="lockToggleBtn" style="display:none;" onclick="toggleLock()"></button>
      </div>
    </div>

    <div class="room-bar">
      <div class="room-controls">
        <span>🚪 目前房間：</span>
        <input type="text" id="roomInput" placeholder="輸入房名">
        <button class="btn-sm btn-add" onclick="switchRoom()">切換 / 開新房</button>
        <button class="btn-sm btn-rate" onclick="copyRoomLink()">🔗 複製此房間連結</button>
      </div>
      <div class="online-box" id="onlineBox">🟢 目前在線：讀取中...</div>
    </div>

    <div id="lockBanner" class="lock-banner">
      🔒 團主已鎖定結單，目前為「僅供檢視」模式（無法新增或修改商品）
    </div>

    <div class="card">
      <textarea id="urlInput" placeholder="貼上各國商品網址（韓、日、中、台、美、澳皆可，一行一個，會自動帶入你的名字與規格庫存）..."></textarea>
      <div class="controls">
        <button id="fetchBtn" class="edit-only" onclick="fetchProducts()">✨ 抓取並加入清單</button>
        <button class="btn-add edit-only" onclick="addManualRow()">+ 手動新增</button>
        <button class="btn-line" onclick="copyLineText()">📋 複製 LINE 報價</button>
        <button id="exportBtn" class="btn-export" onclick="exportExcel()">📊 匯出含圖 Excel</button>
        <button class="btn-rate edit-only" onclick="loadLiveRates()">🔄 更新匯率</button>
        <button class="btn-clear edit-only" onclick="clearAll()">🗑️ 清空</button>
        <span id="statusText" class="status">準備就緒 ✨</span>
      </div>

      <div class="rates-bar">
        <div class="rate-item"><span>🇰🇷 KRW</span><input type="number" class="edit-only" id="rate-KRW" value="0.0235" step="0.0005" onchange="pushState()"></div>
        <div class="rate-item"><span>🇯🇵 JPY</span><input type="number" class="edit-only" id="rate-JPY" value="0.215" step="0.005" onchange="pushState()"></div>
        <div class="rate-item"><span>🇨🇳 CNY</span><input type="number" class="edit-only" id="rate-CNY" value="4.50" step="0.05" onchange="pushState()"></div>
        <div class="rate-item"><span>🇺🇸 USD</span><input type="number" class="edit-only" id="rate-USD" value="32.5" step="0.1" onchange="pushState()"></div>
        <div class="rate-item"><span>🇦🇺 AUD</span><input type="number" class="edit-only" id="rate-AUD" value="21.2" step="0.1" onchange="pushState()"></div>
        <div class="rate-item"><span>🇹🇼 TWD</span><input type="number" id="rate-TWD" value="1" disabled></div>
      </div>

      <div class="extra-bar">
        <label class="extra-item" style="cursor:pointer;">
          <input type="checkbox" class="edit-only" id="feeCheck" onchange="pushState()">
          <span>💳 加上 1.5% 海外刷卡手續費</span>
        </label>
        <div class="extra-item">
          <span>✈️ 國際運費總計 NT$</span>
          <input type="number" class="edit-only" id="shippingInput" value="0" min="0" step="10" onchange="pushState()">
          <span style="font-size:12px; color:#888;">(依勾選件數均分)</span>
        </div>
      </div>
    </div>

    <div class="summary">
      <div>已勾選 <strong id="selectedCount">0</strong> 件商品</div>
      <div>折合台幣總計：<strong id="totalTwd">NT$0</strong></div>
    </div>

    <div class="card" id="splitCard" style="display:none; padding: 12px 18px;">
      <div style="font-size:13px; font-weight:800; margin-bottom:8px; color:#5c4742;">👥 團購自動分帳（已含手續費與運費均分）：</div>
      <div class="split-box" id="splitBox"></div>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th style="width: 32px;"><input type="checkbox" class="edit-only" id="selectAll" checked onchange="toggleAll(this)"></th>
            <th style="width: 82px; text-align:center;">商品照</th>
            <th>商品名稱</th>
            <th style="width: 155px;">規格 / 顏色備註</th>
            <th style="width: 76px; text-align:center;">買家</th>
            <th style="width: 76px;">幣別</th>
            <th style="width: 90px;">原幣單價</th>
            <th style="width: 90px;">台幣單價</th>
            <th style="width: 54px; text-align:center;">數量</th>
            <th style="width: 34px;"></th>
          </tr>
        </thead>
        <tbody id="productTable">
          <tr><td colspan="10" style="text-align:center; color:#999; padding:36px;">載入雲端清單中...</td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <script>
    let products = [];
    let isLocked = false;
    let isAdmin = false;
    let adminPin = "";
    let lastUpdatedAt = 0;
    let isEditingNow = false;
    const CURRENCIES = ["KRW", "JPY", "CNY", "USD", "AUD", "TWD"];

    let clientId = sessionStorage.getItem("chiikawa_cid");
    if (!clientId) {
      clientId = "u_" + Math.random().toString(36).slice(2, 10);
      sessionStorage.setItem("chiikawa_cid", clientId);
    }

    const urlParams = new URLSearchParams(window.location.search);
    let currentRoom = (urlParams.get("room") || "大廳").trim();

    window.addEventListener("DOMContentLoaded", async () => {
      const savedName = localStorage.getItem("chiikawa_my_name") || "訪客";
      document.getElementById("myNameInput").value = savedName;
      document.getElementById("roomInput").value = currentRoom;
      await pullState(true);
      setInterval(() => {
        if (!isEditingNow) pullState(false);
      }, 4000);
    });

    function saveMyName() {
      localStorage.setItem("chiikawa_my_name", document.getElementById("myNameInput").value.trim() || "訪客");
    }

    function getMyName() {
      return document.getElementById("myNameInput").value.trim() || "訪客";
    }

    function switchRoom() {
      const newRoom = document.getElementById("roomInput").value.trim() || "大廳";
      const nextUrl = `${window.location.pathname}?room=${encodeURIComponent(newRoom)}`;
      window.location.href = nextUrl;
    }

    function copyRoomLink() {
      const roomName = document.getElementById("roomInput").value.trim() || currentRoom;
      const shareUrl = `${window.location.origin}${window.location.pathname}?room=${encodeURIComponent(roomName)}`;
      navigator.clipboard.writeText(shareUrl).then(() => {
        document.getElementById("statusText").innerText = `🔗 已複製「${roomName}」房間連結！快傳給朋友吧`;
      }).catch(() => {
        prompt("請手動複製此房間連結：", shareUrl);
      });
    }

    function canEdit() {
      return !isLocked || isAdmin;
    }

    function applyPermissionsUI() {
      const editable = canEdit();
      document.getElementById("lockBanner").style.display = isLocked ? "block" : "none";
      document.getElementById("urlInput").disabled = !editable;
      document.querySelectorAll(".edit-only").forEach(el => { el.disabled = !editable; });

      const lockBtn = document.getElementById("lockToggleBtn");
      if (isAdmin) {
        lockBtn.style.display = "inline-block";
        lockBtn.className = isLocked ? "btn-lock-on" : "btn-lock-off";
        lockBtn.innerText = isLocked ? "🔒 目前已鎖定（點此開放）" : "🔓 開放編輯中（點此鎖定）";
        document.getElementById("adminBtn").innerText = "👑 團主已登入";
      } else {
        lockBtn.style.display = "none";
        document.getElementById("adminBtn").innerText = "👑 團主管理";
      }
    }

    async function toggleAdminLogin() {
      if (isAdmin) {
        if (confirm("要登出團主模式嗎？")) {
          isAdmin = false;
          adminPin = "";
          applyPermissionsUI();
          renderTable();
        }
        return;
      }
      const pin = prompt("請輸入團主密碼（預設 1234）：");
      if (!pin) return;
      const res = await fetch("/api/verify_admin", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pin })
      });
      const data = await res.json();
      if (data.ok) {
        isAdmin = true;
        adminPin = pin;
        applyPermissionsUI();
        renderTable();
        document.getElementById("statusText").innerText = "👑 已進入團主模式，可控制鎖定權限";
      } else {
        alert("密碼錯誤！");
      }
    }

    async function toggleLock() {
      if (!isAdmin) return;
      isLocked = !isLocked;
      await pushState();
      applyPermissionsUI();
      renderTable();
    }

    async function pullState(force = false) {
      try {
        const qs = new URLSearchParams({
          room: currentRoom,
          cid: clientId,
          name: getMyName()
        });
        const res = await fetch(`/api/state?${qs.toString()}`);
        const data = await res.json();
        const badge = document.getElementById("syncBadge");
        badge.innerText = data.cloud ? "☁️ 雲端同步中" : "⚠️ 單機暫存";
        badge.style.color = data.cloud ? "#2a9d8f" : "#e76f51";

        const online = data.online || [];
        const onlineBox = document.getElementById("onlineBox");
        onlineBox.innerText = `🟢 目前在線 (${online.length})：${online.join("、") || getMyName()}`;

        const st = data.state;
        if (!st) return;
        if (!force && st.updatedAt && st.updatedAt <= lastUpdatedAt) return;

        lastUpdatedAt = st.updatedAt || Date.now();
        products = Array.isArray(st.products) ? st.products : [];
        isLocked = !!st.locked;

        if (st.rates) {
          for (const [k, v] of Object.entries(st.rates)) {
            const el = document.getElementById(`rate-${k}`);
            if (el && k !== "TWD") el.value = v;
          }
        }
        document.getElementById("feeCheck").checked = !!st.feeChecked;
        document.getElementById("shippingInput").value = st.shipping || 0;

        applyPermissionsUI();
        renderTable();
      } catch {}
    }

    async function pushState() {
      if (!canEdit()) return;
      const payload = {
        room: currentRoom,
        products: products,
        rates: getRates(),
        feeChecked: document.getElementById("feeCheck").checked,
        shipping: parseFloat(document.getElementById("shippingInput").value) || 0,
        locked: isLocked,
        adminPin: adminPin
      };
      try {
        const res = await fetch("/api/state", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });
        const data = await res.json();
        if (data.ok) {
          lastUpdatedAt = data.updatedAt;
          recalculate();
        } else if (data.error) {
          alert(data.error);
          pullState(true);
        }
      } catch {}
    }

    async function loadLiveRates() {
      if (!canEdit()) return;
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
          await pushState();
          status.innerText = "✅ 已套用今日即時匯率並同步";
        }
      } catch {
        status.innerText = "⚠️ 匯率更新失敗";
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

    function extractUrlsClient(rawText) {
      const parts = rawText.split(/(?=https?:\/\/)/);
      const seen = new Set();
      const urls = [];
      for (const part of parts) {
        const m = part.trim().match(/^(https?:\/\/[^\s"'<>]+)/);
        if (m) {
          const clean = m[1].split("#")[0];
          if (!seen.has(clean)) {
            seen.add(clean);
            urls.push(clean);
          }
        }
      }
      return urls;
    }

    async function fetchProducts() {
      if (!canEdit()) return alert("團主已鎖定結單！");
      const rawText = document.getElementById("urlInput").value.trim();
      if (!rawText) return alert("請先貼上至少一個網址！");

      const urlsToFetch = extractUrlsClient(rawText);
      const btn = document.getElementById("fetchBtn");
      const status = document.getElementById("statusText");
      const myName = getMyName();
      btn.disabled = true;
      isEditingNow = true;

      let added = 0;
      for (let i = 0; i < urlsToFetch.length; i++) {
        const url = urlsToFetch[i];
        btn.innerText = `抓取中 (${i + 1}/${urlsToFetch.length})`;
        status.innerText = `正在抓取第 ${i + 1} / ${urlsToFetch.length} 件商品與庫存...`;

        try {
          const res = await fetch("/api/fetch", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ text: url })
          });
          if (!res.ok) throw new Error("err");
          const data = await res.json();
          if (data.items && data.items.length > 0) {
            const it = data.items[0];
            const defaultSpec = (it.options && it.options.length > 0) ? it.options[0].label : "";
            products.push({ ...it, spec: defaultSpec, note: "", buyer: myName, qty: 1, checked: true });
            added++;
            renderTable();
            await pushState();
          }
        } catch {
          products.push({ url: url, title: "抓取逾時或受阻，請手動補填", img: "", price: 0, currency: "TWD", options: [], spec: "", note: "", buyer: myName, qty: 1, checked: true });
          added++;
          renderTable();
          await pushState();
        }
      }

      document.getElementById("urlInput").value = "";
      btn.disabled = false;
      btn.innerText = "✨ 抓取並加入清單";
      status.innerText = `✨ 已為「${myName}」新增 ${added} 件商品並同步`;
      isEditingNow = false;
    }

    async function addManualRow() {
      if (!canEdit()) return;
      products.push({ url: "", title: "手動新增商品", img: "", price: 0, currency: "TWD", options: [], spec: "", note: "", buyer: getMyName(), qty: 1, checked: true });
      renderTable();
      await pushState();
    }

    async function editImage(idx) {
      if (!canEdit()) return;
      const newUrl = prompt("請貼上商品圖片網址 (右鍵複製圖片位址)：", products[idx].img || "");
      if (newUrl !== null) {
        products[idx].img = newUrl.trim();
        renderTable();
        await pushState();
      }
    }

    async function deleteRow(idx) {
      if (!canEdit()) return;
      products.splice(idx, 1);
      renderTable();
      await pushState();
    }

    async function clearAll() {
      if (!canEdit() || products.length === 0) return;
      if (confirm(`確定要清空「${currentRoom}」房間的所有商品嗎？`)) {
        products = [];
        document.getElementById("shippingInput").value = 0;
        renderTable();
        await pushState();
        document.getElementById("statusText").innerText = "已清空此房間清單";
      }
    }

    function escapeHtml(str) {
      return String(str || "").replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

    function formatFullSpec(p) {
      return [p.spec || "", p.note || ""].map(s => s.trim()).filter(Boolean).join(" / ");
    }

    function renderSpecCell(p, idx, disabledAttr) {
      if (Array.isArray(p.options) && p.options.length > 0) {
        const optsHtml = p.options.map(o => {
          const tag = o.in_stock === false ? "❌ 缺貨" : "✅ 有貨";
          const selected = p.spec === o.label ? "selected" : "";
          return `<option value="${escapeHtml(o.label)}" ${selected}>${escapeHtml(o.label)} (${tag})</option>`;
        }).join("");
        return `
          <div class="spec-stack">
            <select class="spec-select" id="spec-${idx}" ${disabledAttr} onchange="syncAndPush(${idx})">${optsHtml}</select>
            <input type="text" class="note-input" id="note-${idx}" placeholder="顏色 / 備註" value="${escapeHtml(p.note || '')}" ${disabledAttr} onfocus="isEditingNow=true" onblur="isEditingNow=false;syncAndPush(${idx})">
          </div>
        `;
      }
      return `<input type="text" class="spec-input" id="spec-${idx}" placeholder="顏色 / 尺寸 / 備註" value="${escapeHtml(p.spec || '')}" ${disabledAttr} onfocus="isEditingNow=true" onblur="isEditingNow=false;syncAndPush(${idx})">`;
    }

    function renderTable() {
      const tbody = document.getElementById("productTable");
      const editable = canEdit();
      const dis = editable ? "" : "disabled";

      if (products.length === 0) {
        tbody.innerHTML = `<tr><td colspan="10" style="text-align:center; color:#999; padding:36px;">「${escapeHtml(currentRoom)}」目前是空的，輸入名字後貼上網址就能一起買啦 🛒</td></tr>`;
        recalculate();
        return;
      }

      tbody.innerHTML = products.map((p, idx) => `
        <tr>
          <td><input type="checkbox" class="item-check" data-idx="${idx}" ${p.checked !== false ? 'checked' : ''} ${dis} onchange="syncAndPush(${idx})"></td>
          <td>
            <div class="thumb-box">
              ${p.img ? `<a href="${p.url || p.img}" target="_blank"><img class="thumb" src="${p.img}" alt=""></a>` : '<div class="thumb" style="display:flex;align-items:center;justify-content:center;font-size:11px;color:#aaa;">無圖片</div>'}
              ${editable ? `<button class="btn-img" onclick="editImage(${idx})">換圖</button>` : ''}
            </div>
          </td>
          <td>
            <input type="text" class="title-input" id="title-${idx}" value="${escapeHtml(p.title)}" ${dis} onfocus="isEditingNow=true" onblur="isEditingNow=false;syncAndPush(${idx})">
            ${p.url ? `<a class="link-sm" href="${p.url}" target="_blank">開啟原網頁 ↗</a>` : ''}
          </td>
          <td>${renderSpecCell(p, idx, dis)}</td>
          <td><input type="text" class="buyer-input" id="buyer-${idx}" value="${escapeHtml(p.buyer || '訪客')}" ${dis} onfocus="isEditingNow=true" onblur="isEditingNow=false;syncAndPush(${idx})"></td>
          <td>
            <select class="curr-select" id="curr-${idx}" ${dis} onchange="syncAndPush(${idx})">
              ${CURRENCIES.map(c => `<option value="${c}" ${p.currency === c ? 'selected' : ''}>${c}</option>`).join("")}
            </select>
          </td>
          <td><input type="number" class="price-input" id="price-${idx}" value="${p.price}" step="any" min="0" ${dis} onfocus="isEditingNow=true" onblur="isEditingNow=false;syncAndPush(${idx})"></td>
          <td class="twd-price" id="twd-${idx}">NT$0</td>
          <td><input type="number" class="qty-input" id="qty-${idx}" value="${p.qty || 1}" min="1" ${dis} onchange="syncAndPush(${idx})"></td>
          <td>${editable ? `<button class="btn-del" title="刪除" onclick="deleteRow(${idx})">✕</button>` : ''}</td>
        </tr>
      `).join("");

      recalculate();
    }

    async function syncAndPush(idx) {
      if (!canEdit()) return;
      const cb = document.querySelector(`.item-check[data-idx="${idx}"]`);
      if (!cb) return;
      products[idx].checked = cb.checked;
      products[idx].title = document.getElementById(`title-${idx}`).value;
      products[idx].spec = document.getElementById(`spec-${idx}`).value;
      const noteEl = document.getElementById(`note-${idx}`);
      products[idx].note = noteEl ? noteEl.value : "";
      products[idx].buyer = document.getElementById(`buyer-${idx}`).value.trim() || "訪客";
      products[idx].currency = document.getElementById(`curr-${idx}`).value;
      products[idx].price = parseFloat(document.getElementById(`price-${idx}`).value) || 0;
      products[idx].qty = parseInt(document.getElementById(`qty-${idx}`).value) || 1;
      recalculate();
      await pushState();
    }

    async function toggleAll(master) {
      if (!canEdit()) return;
      products.forEach(p => p.checked = master.checked);
      document.querySelectorAll(".item-check").forEach(cb => cb.checked = master.checked);
      recalculate();
      await pushState();
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
          const buyer = p.buyer || "訪客";
          totalQty += qty;
          totalGoodsTwd += unitTwd * qty;

          if (!buyers[buyer]) buyers[buyer] = { qty: 0, goodsTwd: 0, items: [] };
          buyers[buyer].qty += qty;
          buyers[buyer].goodsTwd += unitTwd * qty;
          buyers[buyer].items.push({ title: p.title, fullSpec: formatFullSpec(p), qty, unitTwd, subTwd: unitTwd * qty });
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
    }

    function copyLineText() {
      const { totalQty, totalTwd, shipping, feeChecked, buyers } = getBreakdown();
      if (totalQty === 0) return alert("請至少勾選一件商品！");

      let lines = [`🛍️ 跨國團購結帳清單【${currentRoom}】`];
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
          const specStr = it.fullSpec ? ` [${it.fullSpec}]` : "";
          lines.push(`• ${it.title}${specStr} x${it.qty} — NT$${it.subTwd.toLocaleString()}`);
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
        document.getElementById("statusText").innerText = "📋 已複製 LINE 報價清單！";
      }).catch(() => {
        alert("複製失敗，請手動複製：\n\n" + lines.join("\n"));
      });
    }

    async function exportExcel() {
      const selectedItems = products.filter(p => p.checked !== false).map(p => ({
        ...p,
        fullSpec: formatFullSpec(p)
      }));
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
            room: currentRoom,
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
        a.download = `shopping_${currentRoom}.xlsx`;
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

def get_room_bundle(room_id):
    if not FIREBASE_URL:
        if room_id not in LOCAL_ROOMS:
            LOCAL_ROOMS[room_id] = {"state": default_room_state(), "presence": {}}
        return LOCAL_ROOMS[room_id]
    try:
        res = requests.get(f"{FIREBASE_URL}/rooms/{quote(room_id)}.json", timeout=6)
        data = res.json()
        if isinstance(data, dict):
            return {
                "state": data.get("state") or default_room_state(),
                "presence": data.get("presence") or {}
            }
    except Exception:
        pass
    return {"state": default_room_state(), "presence": {}}

def update_presence(room_id, cid, name):
    now = int(time.time())
    if not FIREBASE_URL:
        if room_id not in LOCAL_ROOMS:
            LOCAL_ROOMS[room_id] = {"state": default_room_state(), "presence": {}}
        LOCAL_ROOMS[room_id]["presence"][cid] = {"name": name, "ts": now}
        return
    try:
        requests.patch(
            f"{FIREBASE_URL}/rooms/{quote(room_id)}/presence.json",
            json={cid: {"name": name, "ts": now}},
            timeout=5
        )
    except Exception:
        pass

def save_room_state(room_id, state):
    if not FIREBASE_URL:
        if room_id not in LOCAL_ROOMS:
            LOCAL_ROOMS[room_id] = {"state": default_room_state(), "presence": {}}
        LOCAL_ROOMS[room_id]["state"] = state
        return True
    try:
        res = requests.put(f"{FIREBASE_URL}/rooms/{quote(room_id)}/state.json", json=state, timeout=6)
        return res.ok
    except Exception:
        return False

@app.route("/bg.jpg")
def background_image():
    img_path = os.path.join(BASE_DIR, "123.jpg")
    if os.path.exists(img_path):
        return send_file(img_path, mimetype="image/jpeg")
    return "", 404

@app.route("/api/verify_admin", methods=["POST"])
def api_verify_admin():
    pin = str(request.json.get("pin", "")).strip()
    return jsonify({"ok": pin == ADMIN_PIN})

@app.route("/api/state", methods=["GET", "POST"])
def api_state():
    if request.method == "GET":
        room_id = sanitize_room(request.args.get("room", "大廳"))
        cid = request.args.get("cid", "").strip()
        name = (request.args.get("name", "訪客").strip() or "訪客")[:15]

        if cid:
            update_presence(room_id, cid, name)

        bundle = get_room_bundle(room_id)
        now = int(time.time())
        online_names = []
        for _, info in (bundle.get("presence") or {}).items():
            if isinstance(info, dict) and now - info.get("ts", 0) <= 18:
                online_names.append(info.get("name", "訪客"))

        return jsonify({
            "cloud": bool(FIREBASE_URL),
            "state": bundle["state"],
            "online": online_names
        })

    data = request.json or {}
    room_id = sanitize_room(data.get("room", "大廳"))
    current = get_room_bundle(room_id)["state"]
    req_pin = str(data.get("adminPin", "")).strip()
    is_admin = (req_pin == ADMIN_PIN)

    if current.get("locked") and not is_admin:
        return jsonify({"ok": False, "error": "團主已鎖定結單，目前無法修改！"}), 403

    new_locked = bool(data.get("locked", current.get("locked", False))) if is_admin else bool(current.get("locked", False))
    new_state = {
        "products": data.get("products", []),
        "rates": data.get("rates", current.get("rates", {})),
        "feeChecked": bool(data.get("feeChecked", False)),
        "shipping": float(data.get("shipping", 0)),
        "locked": new_locked,
        "updatedAt": int(time.time() * 1000)
    }
    save_room_state(room_id, new_state)
    return jsonify({"ok": True, "updatedAt": new_state["updatedAt"]})

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

def parse_options_and_stock(url, html_text, soup):
    options = []
    m = re.search(r"option_stock_data\s*=\s*'(\{.*?\})'", html_text)
    if not m:
        m = re.search(r'option_stock_data\s*=\s*"(\{.*?\})"', html_text)
    if m:
        try:
            raw_json = m.group(1).replace(r'\"', '"').replace(r"\/", "/")
            stock_map = json.loads(raw_json)
            for _, info in stock_map.items():
                if not isinstance(info, dict):
                    continue
                label = info.get("option_value") or info.get("option_name") or ""
                if not label:
                    continue
                is_soldout = (
                    info.get("is_selling") == "F"
                    or info.get("stock_number") == 0
                    or str(info.get("stock_number")) == "0"
                    or "품절" in label
                    or "SOLD OUT" in label.upper()
                )
                clean_label = re.sub(r"\s*\[?(품절|SOLD OUT)\]?", "", label, flags=re.I).strip()
                options.append({"label": clean_label, "in_stock": not is_soldout})
            if options:
                return options
        except Exception:
            pass

    if "/products/" in url:
        try:
            base = url.split("?")[0].rstrip("/")
            s_res = requests.get(f"{base}.json", headers=HEADERS, timeout=5)
            if s_res.ok:
                prod = s_res.json().get("product", {})
                for v in prod.get("variants", []):
                    title = v.get("title", "")
                    if title and title != "Default Title":
                        in_stock = v.get("available", True)
                        options.append({"label": title, "in_stock": bool(in_stock)})
                if options:
                    return options
        except Exception:
            pass

    for sel in soup.select("select[id*='product_option'], select[name*='option']"):
        for opt in sel.find_all("option"):
            val = opt.get("value", "").strip()
            txt = opt.get_text(strip=True)
            if not val or val in ("*", "**") or any(k in txt for k in ["선택", "Select", "請選擇", "---"]):
                continue
            is_soldout = ("품절" in txt) or ("SOLD OUT" in txt.upper()) or opt.has_attr("disabled")
            clean_txt = re.sub(r"\s*\[?(품절|SOLD OUT)\]?", "", txt, flags=re.I).strip()
            options.append({"label": clean_txt, "in_stock": not is_soldout})
        if options:
            break

    return options

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

        options = parse_options_and_stock(url, res.text, soup)

        return {"url": url, "title": title, "img": img_url, "price": price_num, "currency": currency, "options": options}
    except Exception as e:
        return {"url": url, "title": f"抓取受阻，請手動補填 ({type(e).__name__})", "img": "", "price": 0, "currency": currency, "options": []}

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/api/fetch", methods=["POST"])
def api_fetch():
    raw_text = request.json.get("text", "")
    urls = extract_urls(raw_text)
    items = [scrape_product(url) for url in urls]
    return jsonify({"items": items})

@app.route("/api/export", methods=["POST"])
def api_export():
    data = request.json
    room_name = sanitize_room(data.get("room", "大廳"))
    items = data.get("items", [])
    rates = data.get("rates", {})
    fee_mult = 1.015 if data.get("feeChecked") else 1.0
    shipping = float(data.get("shipping", 0))

    total_qty = sum(int(it.get("qty", 1)) for it in items)
    shipping_per_item = shipping / total_qty if total_qty > 0 else 0

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = room_name[:28]

    col_headers = ["商品照", "買家", "商品名稱", "規格 / 顏色備註", "幣別", "原幣單價", "台幣單價", "數量", "運費分攤", "小計 (TWD)", "商品連結"]
    ws.append(col_headers)

    header_fill = PatternFill(start_color="5C4742", end_color="5C4742", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, len(col_headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    col_widths = {"A": 15, "B": 10, "C": 32, "D": 20, "E": 7, "F": 11, "G": 11, "H": 7, "I": 10, "J": 12, "K": 30}
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

        spec_text = item.get("fullSpec") or item.get("spec", "")
        ws.cell(row=row_idx, column=2, value=item.get("buyer", "訪客")).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row_idx, column=3, value=item.get("title", "")).alignment = Alignment(vertical="center", wrap_text=True)
        ws.cell(row=row_idx, column=4, value=spec_text).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(row=row_idx, column=5, value=curr).alignment = Alignment(horizontal="center", vertical="center")

        for col_i, val, fmt in [
            (6, price, '#,##0.##'),
            (7, unit_twd, '#,##0'),
            (9, ship_share, '#,##0'),
            (10, subtotal_twd, '#,##0')
        ]:
            c = ws.cell(row=row_idx, column=col_i, value=val)
            c.number_format = fmt
            c.alignment = Alignment(horizontal="right", vertical="center")

        ws.cell(row=row_idx, column=8, value=qty).alignment = Alignment(horizontal="center", vertical="center")

        url_val = item.get("url", "")
        c_url = ws.cell(row=row_idx, column=11, value=url_val)
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
    ws.cell(row=sum_row, column=7, value="總計").font = Font(bold=True)
    ws.cell(row=sum_row, column=7).alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(row=sum_row, column=8, value=total_qty).font = Font(bold=True)
    ws.cell(row=sum_row, column=8).alignment = Alignment(horizontal="center", vertical="center")

    s_ship = ws.cell(row=sum_row, column=9, value=round(shipping))
    s_ship.number_format = '#,##0'
    s_ship.font = Font(bold=True)
    s_ship.alignment = Alignment(horizontal="right", vertical="center")

    s_twd = ws.cell(row=sum_row, column=10, value=grand_total)
    s_twd.number_format = '#,##0'
    s_twd.font = Font(bold=True, color="FB6F92")
    s_twd.alignment = Alignment(horizontal="right", vertical="center")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"shopping_{room_name}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    if not os.environ.get("RENDER"):
        threading.Timer(1.0, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    app.run(host="0.0.0.0", port=port, debug=False)
