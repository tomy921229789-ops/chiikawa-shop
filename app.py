import io
import os
import re
import json
import time
import hmac
import hashlib
import threading
import webbrowser
from datetime import datetime, timezone, timedelta
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
SIGN_SECRET = (os.environ.get("SECRET_KEY") or f"{ADMIN_PIN}_chiikawa_sign_v1").encode("utf-8")

LOCAL_ROOMS = {}
FAILED_ADMIN_LOG = {}  # {ip: {"count": int, "until": float}}
TW_TZ = timezone(timedelta(hours=8))

def default_room_state():
    return {
        "products": [],
        "rates": {"KRW": 0.0235, "JPY": 0.215, "CNY": 4.50, "USD": 32.5, "AUD": 21.2, "TWD": 1},
        "feeChecked": False,
        "shipping": 0,
        "shippingMode": "qty",
        "freeShipThreshold": 0,
        "bankInfo": "",
        "announcement": "",
        "deadline": "",
        "paidBuyers": {},
        "locked": False,
        "allowedEditors": [],
        "roomPassword": "",
        "updatedAt": 0
    }

def sanitize_room(room_name):
    name = (room_name or "大廳").strip()
    safe = re.sub(r'[.$#\[\]/]', '_', name)
    return safe[:30] if safe else "大廳"

def sign_price(url, price, currency):
    msg = f"{url or ''}|{float(price or 0):.4f}|{str(currency or 'TWD').upper()}".encode("utf-8")
    return hmac.new(SIGN_SECRET, msg, hashlib.sha256).hexdigest()[:16]

def is_deadline_passed(deadline_str):
    if not deadline_str:
        return False
    try:
        dt = datetime.strptime(deadline_str[:16], "%Y-%m-%dT%H:%M").replace(tzinfo=TW_TZ)
        return datetime.now(TW_TZ) >= dt
    except Exception:
        return False

def get_client_ip():
    fwd = request.headers.get("X-Forwarded-For", "")
    return fwd.split(",")[0].strip() if fwd else (request.remote_addr or "unknown")

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
    .container { max-width: 1340px; margin: 0 auto; }
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
      margin-bottom: 12px;
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
    .announce-bar {
      display: none;
      background: rgba(255, 248, 220, 0.96);
      border: 2px solid #ffd166;
      border-radius: 14px;
      padding: 10px 16px;
      margin-bottom: 12px;
      font-size: 14px;
      font-weight: 800;
      color: #6c4f3d;
    }
    .countdown-bar {
      display: none;
      background: linear-gradient(90deg, #ff8fab, #fb6f92);
      color: #fff;
      border-radius: 14px;
      padding: 9px 16px;
      margin-bottom: 12px;
      font-size: 14px;
      font-weight: 800;
      text-align: center;
      box-shadow: 0 4px 12px rgba(251, 111, 146, 0.3);
    }
    .countdown-bar.expired { background: #6c757d; box-shadow: none; }
    .admin-panel {
      display: none;
      background: rgba(92, 71, 66, 0.96);
      color: #fff;
      border: 2px solid #ffd166;
      border-radius: 18px;
      padding: 14px 18px;
      margin-bottom: 14px;
      font-size: 13px;
      font-weight: 700;
    }
    .admin-row { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-top: 9px; }
    .admin-panel input {
      padding: 5px 10px;
      border: 1px solid #ffd166;
      border-radius: 10px;
      font-size: 13px;
      color: #5c4742;
      font-weight: 700;
    }
    .rooms-list { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 6px; }
    .room-chip {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: rgba(255, 255, 255, 0.15);
      border: 1px solid #ffd166;
      color: #fff;
      padding: 3px 10px;
      border-radius: 50px;
      font-size: 12px;
      cursor: pointer;
    }
    .room-chip:hover { background: rgba(255, 209, 102, 0.3); }
    .room-chip-del { color: #ff8fab; font-weight: 900; margin-left: 2px; padding: 0 3px; }
    .room-chip-del:hover { color: #ff4d6d; }
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
      height: 86px;
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
    .extra-item input[type="number"], .extra-item select {
      padding: 2px 6px;
      border: 1px solid #f4be49;
      border-radius: 8px;
      font-weight: bold;
      color: #5c4742;
      background: #fff;
    }
    .extra-item input[type="number"] { width: 74px; text-align: right; }

    .freeship-card {
      display: none;
      background: rgba(255, 255, 255, 0.94);
      border: 2px solid #ffd166;
      border-radius: 18px;
      padding: 12px 18px;
      margin-bottom: 14px;
    }
    .freeship-text {
      display: flex;
      justify-content: space-between;
      font-size: 13px;
      font-weight: 800;
      margin-bottom: 6px;
      color: #5c4742;
    }
    .progress-track { width: 100%; height: 14px; background: #ffe5ec; border-radius: 50px; overflow: hidden; }
    .progress-fill { height: 100%; background: linear-gradient(90deg, #ff8fab, #ffd166); border-radius: 50px; transition: width 0.4s ease; }
    .progress-fill.reached { background: linear-gradient(90deg, #2a9d8f, #70c1b3); }

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
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: #fff0f3;
      border: 2px solid #ffb3c6;
      border-radius: 14px;
      padding: 6px 12px;
      font-size: 13px;
      font-weight: 700;
      color: #5c4742;
      cursor: pointer;
      transition: all 0.15s;
    }
    .buyer-pill:hover { border-color: #fb6f92; transform: translateY(-1px); }
    .buyer-pill.active-filter { background: #ff8fab; color: #fff; border-color: #fb6f92; }
    .buyer-pill.active-filter .pill-amt { color: #ffd166; }
    .pill-amt { color: #fb6f92; font-size: 15px; }
    .paid-tag { font-size: 11px; padding: 2px 7px; border-radius: 50px; font-weight: 800; background: #ffe5ec; color: #c9184a; }
    .paid-tag.is-paid { background: #d8f3dc; color: #1b7a68; }

    .table-wrap {
      background: rgba(255, 255, 255, 0.92);
      backdrop-filter: blur(10px);
      border-radius: 20px;
      border: 2px solid rgba(255, 179, 198, 0.75);
      overflow-x: auto;
      box-shadow: 0 8px 24px rgba(92, 71, 66, 0.1);
    }
    table { width: 100%; border-collapse: collapse; }
    th, td { padding: 9px 7px; border-bottom: 1px solid #ffe5ec; text-align: left; vertical-align: middle; }
    th { background: rgba(255, 240, 243, 0.95); font-size: 12px; color: #5c4742; font-weight: 800; white-space: nowrap; }
    tr.row-other { background: rgba(245, 240, 242, 0.55); }
    .thumb-box { display: flex; flex-direction: column; align-items: center; gap: 3px; }
    .thumb { width: 64px; height: 64px; object-fit: cover; border-radius: 10px; border: 2px solid #ffe5ec; background: #fff; }
    .btn-img { font-size: 11px; padding: 2px 7px; background: #ffe5ec; color: #5c4742; box-shadow: none; }
    .btn-img:hover { background: #ffb3c6; color: #fff; }
    .title-input { width: 100%; min-width: 140px; padding: 4px 6px; border: 1px solid transparent; border-radius: 8px; font-size: 13px; background: transparent; color: #5c4742; font-weight: 600; }
    .title-input:hover:not(:disabled), .title-input:focus { border-color: #ffb3c6; background: #fff; outline: none; }
    .spec-stack { display: flex; flex-direction: column; gap: 5px; }
    .spec-input, .spec-select { width: 136px; padding: 5px 6px; border: 1px solid #ffb3c6; border-radius: 8px; font-size: 12px; color: #5c4742; background: #fff; }
    .note-input { width: 136px; padding: 4px 6px; border: 1px dashed #ffb3c6; border-radius: 8px; font-size: 12px; color: #5c4742; background: #fff8fa; }
    .buyer-input { width: 64px; padding: 5px; border: 1px solid #ffb3c6; border-radius: 8px; font-size: 12px; color: #5c4742; font-weight: 700; text-align: center; background: #fff0f3; }
    .status-select { padding: 4px 5px; border: 1px solid #ffb3c6; border-radius: 8px; font-size: 12px; font-weight: 700; background: #fff; color: #5c4742; }
    .status-badge { display: inline-block; padding: 3px 8px; border-radius: 50px; font-size: 11px; font-weight: 800; white-space: nowrap; }
    .st-wait { background: #f1f3f5; color: #495057; }
    .st-ordered { background: #e7f5ff; color: #1971c2; }
    .st-arrived { background: #d8f3dc; color: #1b7a68; }
    .st-out { background: #ffe3e3; color: #c92a2a; }
    input:disabled, select:disabled { background: #f3ecee !important; border-color: #e2d5d8 !important; color: #7d6b66 !important; cursor: not-allowed; }
    .curr-select { padding: 4px 5px; border: 1px solid #ffb3c6; border-radius: 8px; font-size: 12px; font-weight: bold; background: #fff; color: #5c4742; }
    .price-input { width: 78px; padding: 5px; border: 1px solid #ffb3c6; border-radius: 8px; text-align: right; font-size: 13px; color: #5c4742; }
    .qty-input { width: 44px; padding: 5px; border: 1px solid #ffb3c6; border-radius: 8px; text-align: center; font-size: 13px; color: #5c4742; }
    .twd-price { font-weight: 800; color: #fb6f92; white-space: nowrap; font-size: 13px; }
    .action-cell { display: flex; align-items: center; gap: 4px; }
    .btn-plus1 { background: #ffd166; color: #5c4742; padding: 4px 8px; font-size: 11px; box-shadow: none; }
    .btn-plus1:hover { background: #f4be49; }
    .btn-del { background: transparent; color: #bbb; box-shadow: none; padding: 3px 6px; font-size: 14px; }
    .btn-del:hover { background: #ffe5ec; color: #fb6f92; }
    .link-sm { font-size: 11px; color: #219ebc; text-decoration: none; display: inline-block; margin-top: 2px; font-weight: 600; }
    .link-sm:hover { text-decoration: underline; }
    .status { font-size: 12px; color: #5c4742; margin-left: auto; font-weight: 700; }

    .modal-overlay {
      position: fixed;
      inset: 0;
      background: rgba(92, 71, 66, 0.82);
      backdrop-filter: blur(10px);
      display: none;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    .modal-box {
      background: #fff;
      border: 3px solid #ffb3c6;
      border-radius: 24px;
      padding: 28px 24px;
      max-width: 360px;
      width: 100%;
      text-align: center;
      box-shadow: 0 12px 32px rgba(0,0,0,0.25);
    }
    .modal-box h2 { margin: 0 0 8px; font-size: 18px; color: #5c4742; }
    .modal-box input {
      width: 100%;
      padding: 10px 14px;
      border: 2px solid #ffb3c6;
      border-radius: 12px;
      font-size: 15px;
      text-align: center;
      margin: 12px 0;
      outline: none;
    }
  </style>
</head>
<body>
  <div id="pwOverlay" class="modal-overlay" style="z-index: 1000;">
    <div class="modal-box">
      <h2>🔐 此房間設有入房密碼</h2>
      <div style="font-size:13px; color:#888;">請輸入密碼以進入「<span id="pwRoomName"></span>」</div>
      <input type="password" id="pwInput" placeholder="請輸入入房密碼" onkeydown="if(event.key==='Enter') submitRoomPassword()">
      <div id="pwError" style="color:#e63946; font-size:12px; font-weight:700; min-height:18px; margin-bottom:8px;"></div>
      <div style="display:flex; gap:8px; justify-content:center;">
        <button onclick="submitRoomPassword()">進入房間</button>
        <button class="btn-admin" onclick="toggleAdminLogin()">👑 團主登入</button>
      </div>
    </div>
  </div>

  <div id="nameOverlay" class="modal-overlay" style="z-index: 990;">
    <div class="modal-box">
      <h2>👋 請先輸入你的名字</h2>
      <div style="font-size:13px; color:#888;">系統會自動綁定你的裝置金鑰，防冒名修改你的商品</div>
      <input type="text" id="entryNameInput" placeholder="例如：小明、小美" maxlength="15" onkeydown="if(event.key==='Enter') submitEntryName()">
      <div id="nameError" style="color:#e63946; font-size:12px; font-weight:700; min-height:18px; margin-bottom:8px;"></div>
      <button onclick="submitEntryName()" style="width:100%; padding:10px;">✨ 確認並進入房間</button>
    </div>
  </div>

  <div class="container">
    <div class="top-bar">
      <h1 class="header-title">🛍️ 吉伊卡哇多人團購神器</h1>
      <div class="user-bar">
        <span>👤 你的名字：</span>
        <input type="text" id="myNameInput" placeholder="輸入暱稱" oninput="saveMyName()">
        <span id="syncBadge" style="font-size:12px; color:#2a9d8f;">☁️ 連線中...</span>
        <button id="adminBtn" class="btn-admin" onclick="toggleAdminLogin()">👑 團主管理</button>
        <button id="topLockBtn" style="display:none;" onclick="toggleLock()"></button>
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

    <div id="announceBar" class="announce-bar"></div>
    <div id="countdownBar" class="countdown-bar"></div>

    <div id="adminPanel" class="admin-panel">
      <div style="color:#ffd166; font-size:14px;">👑 團主控制台（已啟用防冒名、防竄改價格與後端結單保護）</div>
      <div class="admin-row">
        <button id="lockToggleBtn" onclick="toggleLock()"></button>
        <span>｜ 授權名單：</span>
        <input type="text" id="allowedEditorsInput" placeholder="例：小明, 小美" style="width:150px;">
        <span>｜ 入房密碼：</span>
        <input type="text" id="roomPwSettingInput" placeholder="留空免密碼" style="width:105px;">
        <span>｜ 免運門檻 NT$：</span>
        <input type="number" id="freeShipInput" placeholder="0=關閉" min="0" step="100" style="width:88px;">
        <span>｜ 結單倒數：</span>
        <input type="datetime-local" id="deadlineInput" style="width:175px;">
      </div>
      <div class="admin-row">
        <span>📢 置頂公告：</span>
        <input type="text" id="announceInput" placeholder="例：LMC 版型偏大建議拿小一號，周五晚上10點準時收單！" style="flex:1; min-width:220px;">
      </div>
      <div class="admin-row">
        <span>🏦 匯款資訊：</span>
        <input type="text" id="bankInfoInput" placeholder="例：中信(822) 1234-5678-9012 / LINE Pay（會附在 LINE 報價）" style="flex:1; min-width:220px;">
        <button class="btn-sm btn-rate" onclick="saveAdminSettings()">💾 套用全部設定</button>
      </div>
      <div class="admin-row" style="align-items:flex-start; flex-direction:column; gap:4px; margin-top:10px; padding-top:8px; border-top:1px dashed rgba(255,209,102,0.4);">
        <div style="display:flex; align-items:center; gap:8px;">
          <span>📂 所有房間管理：</span>
          <button class="btn-sm" style="background:rgba(255,255,255,0.2); padding:2px 9px;" onclick="loadAllRooms()">🔄 重新整理</button>
        </div>
        <div class="rooms-list" id="roomsListBox">載入中...</div>
      </div>
    </div>

    <div id="lockBanner" class="lock-banner"></div>

    <div id="freeShipCard" class="freeship-card">
      <div class="freeship-text">
        <span id="freeShipMsg">🔥 湊免運進度</span>
        <span id="freeShipPct">0%</span>
      </div>
      <div class="progress-track">
        <div id="freeShipFill" class="progress-fill" style="width:0%;"></div>
      </div>
    </div>

    <div class="card">
      <textarea id="urlInput" placeholder="貼上各國商品網址（韓、日、中、台、美、澳皆可，一行一個；價格與身分均受後端安全保護）..."></textarea>
      <div class="controls">
        <button id="fetchBtn" class="edit-only" onclick="fetchProducts()">✨ 抓取並加入清單</button>
        <button class="btn-add edit-only" onclick="addManualRow()">+ 手動新增</button>
        <button class="btn-line" onclick="copyLineText()">📋 複製 LINE 報價</button>
        <button id="exportBtn" class="btn-export" onclick="exportExcel()">📊 匯出含圖 Excel</button>
        <button class="btn-rate admin-only" onclick="loadLiveRates()">🔄 更新匯率</button>
        <button class="btn-clear admin-only" onclick="clearAll()">🗑️ 清空</button>
        <span id="statusText" class="status">🛡️ 安全防護運作中</span>
      </div>

      <div class="rates-bar">
        <div class="rate-item"><span>🇰🇷 KRW</span><input type="number" class="admin-only" id="rate-KRW" value="0.0235" step="0.0005" onchange="pushState()"></div>
        <div class="rate-item"><span>🇯🇵 JPY</span><input type="number" class="admin-only" id="rate-JPY" value="0.215" step="0.005" onchange="pushState()"></div>
        <div class="rate-item"><span>🇨🇳 CNY</span><input type="number" class="admin-only" id="rate-CNY" value="4.50" step="0.05" onchange="pushState()"></div>
        <div class="rate-item"><span>🇺🇸 USD</span><input type="number" class="admin-only" id="rate-USD" value="32.5" step="0.1" onchange="pushState()"></div>
        <div class="rate-item"><span>🇦🇺 AUD</span><input type="number" class="admin-only" id="rate-AUD" value="21.2" step="0.1" onchange="pushState()"></div>
        <div class="rate-item"><span>🇹🇼 TWD</span><input type="number" id="rate-TWD" value="1" disabled></div>
      </div>

      <div class="extra-bar">
        <label class="extra-item" style="cursor:pointer;">
          <input type="checkbox" class="admin-only" id="feeCheck" onchange="pushState()">
          <span>💳 加上 1.5% 海外刷卡手續費</span>
        </label>
        <div class="extra-item">
          <span>✈️ 國際運費 NT$</span>
          <input type="number" class="admin-only" id="shippingInput" value="0" min="0" step="10" onchange="pushState()">
          <select id="shippingModeSelect" class="admin-only" onchange="pushState()">
            <option value="qty">依件數均分</option>
            <option value="amount">依金額比例</option>
          </select>
        </div>
      </div>
    </div>

    <div class="summary">
      <div>已勾選 <strong id="selectedCount">0</strong> 件商品</div>
      <div>折合台幣總計：<strong id="totalTwd">NT$0</strong></div>
    </div>

    <div class="card" id="splitCard" style="display:none; padding: 12px 18px;">
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-bottom:8px;">
        <span style="font-size:13px; font-weight:800; color:#5c4742;">👥 團購自動分帳（點名字可篩選商品；團主點標籤可標記已付款）：</span>
        <button id="clearFilterBtn" class="btn-sm btn-add" style="display:none;" onclick="setFilterBuyer(null)">顯示全部商品</button>
      </div>
      <div class="split-box" id="splitBox"></div>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th style="width: 30px;"><input type="checkbox" class="admin-only" id="selectAll" checked onchange="toggleAll(this)"></th>
            <th style="width: 78px; text-align:center;">商品照</th>
            <th>商品名稱</th>
            <th style="width: 146px;">規格 / 顏色備註</th>
            <th style="width: 70px; text-align:center;">買家🔒</th>
            <th style="width: 82px; text-align:center;">進度</th>
            <th style="width: 70px;">幣別🔒</th>
            <th style="width: 84px;">單價🔒</th>
            <th style="width: 84px;">台幣單價</th>
            <th style="width: 50px; text-align:center;">數量</th>
            <th style="width: 72px; text-align:center;">操作</th>
          </tr>
        </thead>
        <tbody id="productTable">
          <tr><td colspan="11" style="text-align:center; color:#999; padding:36px;">載入雲端清單中...</td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <script>
    let products = [];
    let isLocked = false;
    let allowedEditors = [];
    let roomPasswordSetting = "";
    let freeShipThreshold = 0;
    let bankInfo = "";
    let announcement = "";
    let deadlineStr = "";
    let paidBuyers = {};
    let filterBuyer = null;
    let isAdmin = false;
    let adminPin = "";
    let lastUpdatedAt = 0;
    let isEditingNow = false;
    const CURRENCIES = ["KRW", "JPY", "CNY", "USD", "AUD", "TWD"];
    const ITEM_STATUSES = ["⏳ 待下單", "🛍️ 已下單", "📦 已到貨", "❌ 斷貨"];

    // 瀏覽器專屬防冒名金鑰（永久存於本機）
    let userToken = localStorage.getItem("chiikawa_utoken");
    if (!userToken) {
      userToken = "tk_" + Math.random().toString(36).slice(2) + Date.now().toString(36);
      localStorage.setItem("chiikawa_utoken", userToken);
    }

    let clientId = sessionStorage.getItem("chiikawa_cid");
    if (!clientId) {
      clientId = "u_" + Math.random().toString(36).slice(2, 10);
      sessionStorage.setItem("chiikawa_cid", clientId);
    }

    const urlParams = new URLSearchParams(window.location.search);
    let currentRoom = (urlParams.get("room") || "大廳").trim();

    function getSavedRoomPw() { return sessionStorage.getItem("chiikawa_rpw_" + currentRoom) || ""; }
    function getSessionName() { return sessionStorage.getItem("chiikawa_session_name") || ""; }
    function makeId() { return "p_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 6); }

    function isDeadlineExpired() {
      if (!deadlineStr) return false;
      const t = new Date(deadlineStr).getTime();
      return !isNaN(t) && Date.now() >= t;
    }

    window.addEventListener("DOMContentLoaded", async () => {
      document.getElementById("roomInput").value = currentRoom;
      document.getElementById("pwRoomName").innerText = currentRoom;

      const savedPin = sessionStorage.getItem("chiikawa_admin_pin");
      if (savedPin) {
        try {
          const res = await fetch("/api/verify_admin", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ pin: savedPin })
          });
          const d = await res.json();
          if (d.ok) {
            isAdmin = true;
            adminPin = savedPin;
          } else {
            sessionStorage.removeItem("chiikawa_admin_pin");
          }
        } catch {}
      }

      const sName = getSessionName();
      if (sName) {
        document.getElementById("myNameInput").value = sName;
      } else {
        document.getElementById("nameOverlay").style.display = "flex";
        setTimeout(() => document.getElementById("entryNameInput").focus(), 100);
      }

      await pullState(true);
      if (isAdmin) loadAllRooms();

      setInterval(() => {
        if (!isEditingNow && getSessionName()) pullState(false);
      }, 4000);
      setInterval(updateCountdownUI, 1000);
    });

    function updateCountdownUI() {
      const bar = document.getElementById("countdownBar");
      if (!deadlineStr) {
        bar.style.display = "none";
        return;
      }
      const target = new Date(deadlineStr).getTime();
      if (isNaN(target)) {
        bar.style.display = "none";
        return;
      }
      bar.style.display = "block";
      const diff = target - Date.now();
      if (diff <= 0) {
        if (!bar.classList.contains("expired")) {
          bar.classList.add("expired");
          bar.innerText = "⏰ 結單時間已到，系統已自動鎖定房間！";
          applyPermissionsUI();
          renderTable();
        }
        return;
      }
      bar.classList.remove("expired");
      const d = Math.floor(diff / 86400000);
      const h = Math.floor((diff % 86400000) / 3600000);
      const m = Math.floor((diff % 3600000) / 60000);
      const s = Math.floor((diff % 60000) / 1000);
      bar.innerText = `⏳ 距離結單倒數：${d > 0 ? d + '天 ' : ''}${String(h).padStart(2,'0')}時 ${String(m).padStart(2,'0')}分 ${String(s).padStart(2,'0')}秒`;
    }

    function submitEntryName() {
      const val = document.getElementById("entryNameInput").value.trim();
      if (!val) {
        document.getElementById("nameError").innerText = "請輸入名字才能進入喔！";
        return;
      }
      sessionStorage.setItem("chiikawa_session_name", val);
      document.getElementById("myNameInput").value = val;
      document.getElementById("nameOverlay").style.display = "none";
      applyPermissionsUI();
      renderTable();
      pullState(true);
    }

    function saveMyName() {
      const val = document.getElementById("myNameInput").value.trim();
      if (val) sessionStorage.setItem("chiikawa_session_name", val);
      applyPermissionsUI();
      renderTable();
    }

    function getMyName() {
      return document.getElementById("myNameInput").value.trim() || getSessionName() || "訪客";
    }

    function switchRoom(targetName) {
      const newRoom = (targetName || document.getElementById("roomInput").value).trim() || "大廳";
      window.location.href = `${window.location.pathname}?room=${encodeURIComponent(newRoom)}`;
    }

    function copyRoomLink() {
      const roomName = document.getElementById("roomInput").value.trim() || currentRoom;
      const shareUrl = `${window.location.origin}${window.location.pathname}?room=${encodeURIComponent(roomName)}`;
      navigator.clipboard.writeText(shareUrl).then(() => {
        document.getElementById("statusText").innerText = `🔗 已複製「${roomName}」房間連結！`;
      }).catch(() => {
        prompt("請手動複製此房間連結：", shareUrl);
      });
    }

    function isEffectivelyLocked() {
      return isLocked || isDeadlineExpired();
    }

    function canEdit() {
      if (isAdmin) return true;
      if (!isEffectivelyLocked()) return true;
      return allowedEditors.includes(getMyName());
    }

    function canEditRow(p) {
      if (isAdmin) return true;
      if (!canEdit()) return false;
      return !!p.isMine;
    }

    function setFilterBuyer(name) {
      filterBuyer = (filterBuyer === name) ? null : name;
      document.getElementById("clearFilterBtn").style.display = filterBuyer ? "inline-block" : "none";
      renderTable();
    }

    async function togglePaid(event, buyerName) {
      event.stopPropagation();
      if (!isAdmin) return;
      paidBuyers[buyerName] = !paidBuyers[buyerName];
      recalculate();
      await pushState();
    }

    function applyPermissionsUI() {
      const editable = canEdit();
      const lockedNow = isEffectivelyLocked();
      const banner = document.getElementById("lockBanner");

      const annBar = document.getElementById("announceBar");
      if (announcement) {
        annBar.style.display = "block";
        annBar.innerText = `📢 團主公告：${announcement}`;
      } else {
        annBar.style.display = "none";
      }

      if (lockedNow) {
        banner.style.display = "block";
        if (isAdmin) {
          banner.innerText = `🔒 房間已鎖定：目前僅限團主與授權名單（${allowedEditors.join("、") || "無"}）可編輯各自商品`;
        } else if (editable) {
          banner.innerText = `✨ 你已被團主個別授權編輯自己的商品（其他未授權成員僅供檢視）`;
        } else {
          banner.innerText = `🔒 房間已鎖定結單，目前為「僅供檢視」模式`;
        }
      } else {
        banner.style.display = "none";
      }

      document.getElementById("urlInput").disabled = !editable;
      document.querySelectorAll(".edit-only").forEach(el => { el.disabled = !editable; });
      document.querySelectorAll(".admin-only").forEach(el => { el.disabled = !isAdmin; });

      const adminPanel = document.getElementById("adminPanel");
      const lockBtn = document.getElementById("lockToggleBtn");
      const topLockBtn = document.getElementById("topLockBtn");

      if (isAdmin) {
        adminPanel.style.display = "block";
        lockBtn.className = isLocked ? "btn-lock-on" : "btn-lock-off";
        lockBtn.innerText = isLocked ? "🔒 僅限授權名單（點此全體開放）" : "🔓 全體開放中（點此鎖定房間）";

        topLockBtn.style.display = "inline-block";
        topLockBtn.className = isLocked ? "btn-lock-on" : "btn-lock-off";
        topLockBtn.innerText = isLocked ? "🔒 已鎖定" : "🔓 鎖定房間";

        document.getElementById("adminBtn").innerText = "👑 團主登出";
      } else {
        adminPanel.style.display = "none";
        topLockBtn.style.display = "none";
        document.getElementById("adminBtn").innerText = "👑 團主管理";
      }
    }

    async function loadAllRooms() {
      if (!isAdmin) return;
      const box = document.getElementById("roomsListBox");
      box.innerText = "讀取中...";
      try {
        const res = await fetch(`/api/rooms?adminPin=${encodeURIComponent(adminPin)}`);
        const data = await res.json();
        if (!data.ok || !data.rooms.length) {
          box.innerText = "目前沒有其他房間";
          return;
        }
        box.innerHTML = data.rooms.map(r => `
          <span class="room-chip" onclick="switchRoom('${escapeHtml(r)}')">
            🚪 ${escapeHtml(r)}
            <span class="room-chip-del" title="刪除此房間" onclick="deleteRoom(event, '${escapeHtml(r)}')">✕</span>
          </span>
        `).join("");
      } catch {
        box.innerText = "讀取失敗";
      }
    }

    async function deleteRoom(event, roomName) {
      event.stopPropagation();
      if (!isAdmin) return;
      if (!confirm(`確定要永久刪除房間「${roomName}」與裡面的所有資料嗎？`)) return;
      try {
        const res = await fetch("/api/rooms", {
          method: "DELETE",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ room: roomName, adminPin })
        });
        const data = await res.json();
        if (data.ok) {
          loadAllRooms();
          if (roomName === currentRoom) switchRoom("大廳");
        }
      } catch {}
    }

    async function submitRoomPassword() {
      const pw = document.getElementById("pwInput").value.trim();
      sessionStorage.setItem("chiikawa_rpw_" + currentRoom, pw);
      document.getElementById("pwError").innerText = "驗證中...";
      const ok = await pullState(true);
      if (!ok) {
        document.getElementById("pwError").innerText = "❌ 密碼錯誤，請再試一次";
      } else {
        document.getElementById("pwError").innerText = "";
        document.getElementById("pwOverlay").style.display = "none";
      }
    }

    async function toggleAdminLogin() {
      if (isAdmin) {
        if (confirm("要登出團主模式嗎？")) {
          isAdmin = false;
          adminPin = "";
          sessionStorage.removeItem("chiikawa_admin_pin");
          sessionStorage.removeItem("chiikawa_rpw_" + currentRoom);
          applyPermissionsUI();
          renderTable();
          await pullState(true);
        }
        return;
      }
      const pin = prompt("請輸入團主密碼：");
      if (!pin) return;
      try {
        const res = await fetch("/api/verify_admin", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ pin })
        });
        const data = await res.json();
        if (data.ok) {
          isAdmin = true;
          adminPin = pin;
          sessionStorage.setItem("chiikawa_admin_pin", pin);
          document.getElementById("pwOverlay").style.display = "none";
          applyPermissionsUI();
          renderTable();
          await pullState(true);
          loadAllRooms();
          document.getElementById("statusText").innerText = "👑 已進入團主模式";
        } else {
          alert(data.error || "密碼錯誤！");
        }
      } catch {
        alert("驗證失敗，請稍後再試");
      }
    }

    async function toggleLock() {
      if (!isAdmin) return;
      isLocked = !isLocked;
      applyPermissionsUI();
      await pushState();
      renderTable();
    }

    async function saveAdminSettings() {
      if (!isAdmin) return;
      const rawNames = document.getElementById("allowedEditorsInput").value;
      allowedEditors = rawNames.split(/[,，\s]+/).map(s => s.trim()).filter(Boolean);
      roomPasswordSetting = document.getElementById("roomPwSettingInput").value.trim();
      freeShipThreshold = parseFloat(document.getElementById("freeShipInput").value) || 0;
      deadlineStr = document.getElementById("deadlineInput").value || "";
      announcement = document.getElementById("announceInput").value.trim();
      bankInfo = document.getElementById("bankInfoInput").value.trim();
      await pushState();
      applyPermissionsUI();
      updateCountdownUI();
      recalculate();
      document.getElementById("statusText").innerText = "💾 已套用全部團主設定！";
    }

    async function pullState(force = false) {
      try {
        const qs = new URLSearchParams({
          room: currentRoom,
          cid: clientId,
          utoken: userToken,
          name: getMyName(),
          rpw: getSavedRoomPw(),
          adminPin: adminPin
        });
        const res = await fetch(`/api/state?${qs.toString()}`);
        const data = await res.json();

        if (data.needPassword) {
          document.getElementById("pwOverlay").style.display = "flex";
          return false;
        }
        document.getElementById("pwOverlay").style.display = "none";

        const badge = document.getElementById("syncBadge");
        badge.innerText = data.cloud ? "☁️ 雲端同步中" : "⚠️ 單機暫存";
        badge.style.color = data.cloud ? "#2a9d8f" : "#e76f51";

        const online = data.online || [];
        document.getElementById("onlineBox").innerText = `🟢 目前在線 (${online.length})：${online.join("、") || getMyName()}`;

        const st = data.state;
        if (!st) return true;
        if (!force && st.updatedAt && st.updatedAt <= lastUpdatedAt) return true;

        lastUpdatedAt = st.updatedAt || Date.now();
        products = Array.isArray(st.products) ? st.products : [];
        isLocked = !!st.locked;
        allowedEditors = Array.isArray(st.allowedEditors) ? st.allowedEditors : [];
        freeShipThreshold = parseFloat(st.freeShipThreshold) || 0;
        bankInfo = st.bankInfo || "";
        announcement = st.announcement || "";
        deadlineStr = st.deadline || "";
        paidBuyers = (st.paidBuyers && typeof st.paidBuyers === "object") ? st.paidBuyers : {};

        if (isAdmin) {
          if (typeof st.roomPassword === "string") {
            roomPasswordSetting = st.roomPassword;
            document.getElementById("roomPwSettingInput").value = roomPasswordSetting;
          }
          document.getElementById("allowedEditorsInput").value = allowedEditors.join(", ");
          document.getElementById("freeShipInput").value = freeShipThreshold || "";
          document.getElementById("deadlineInput").value = deadlineStr;
          document.getElementById("announceInput").value = announcement;
          document.getElementById("bankInfoInput").value = bankInfo;
        }

        if (st.rates) {
          for (const [k, v] of Object.entries(st.rates)) {
            const el = document.getElementById(`rate-${k}`);
            if (el && k !== "TWD") el.value = v;
          }
        }
        document.getElementById("feeCheck").checked = !!st.feeChecked;
        document.getElementById("shippingInput").value = st.shipping || 0;
        document.getElementById("shippingModeSelect").value = st.shippingMode || "qty";

        applyPermissionsUI();
        updateCountdownUI();
        renderTable();
        return true;
      } catch {
        return false;
      }
    }

    async function pushState() {
      if (!canEdit()) return;
      const payload = {
        room: currentRoom,
        userName: getMyName(),
        utoken: userToken,
        rpw: getSavedRoomPw(),
        products: products,
        rates: getRates(),
        feeChecked: document.getElementById("feeCheck").checked,
        shipping: parseFloat(document.getElementById("shippingInput").value) || 0,
        shippingMode: document.getElementById("shippingModeSelect").value || "qty",
        freeShipThreshold: freeShipThreshold,
        bankInfo: bankInfo,
        announcement: announcement,
        deadline: deadlineStr,
        paidBuyers: paidBuyers,
        locked: isLocked,
        allowedEditors: allowedEditors,
        roomPassword: roomPasswordSetting,
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
          if (Array.isArray(data.products)) {
            products = data.products;
            renderTable();
          } else {
            recalculate();
          }
        } else if (data.error) {
          alert(data.error);
          pullState(true);
        }
      } catch {}
    }

    async function loadLiveRates() {
      if (!isAdmin) return;
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
      if (!canEdit()) return alert("目前沒有編輯權限！");
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
            products.push({ id: makeId(), ...it, spec: defaultSpec, note: "", status: "⏳ 待下單", buyer: myName, isMine: true, qty: 1, checked: true });
            added++;
            await pushState();
          }
        } catch {
          products.push({ id: makeId(), url: url, title: "抓取逾時或受阻，請通知團主補價", img: "", price: 0, currency: "TWD", options: [], spec: "", note: "", status: "⏳ 待下單", buyer: myName, isMine: true, qty: 1, checked: true });
          added++;
          await pushState();
        }
      }

      document.getElementById("urlInput").value = "";
      btn.disabled = false;
      btn.innerText = "✨ 抓取並加入清單";
      status.innerText = `✨ 已為「${myName}」新增 ${added} 件商品`;
      isEditingNow = false;
    }

    async function addManualRow() {
      if (!canEdit()) return;
      products.push({ id: makeId(), url: "", title: "手動新增商品（待團主核價）", img: "", price: 0, currency: "TWD", options: [], spec: "", note: "", status: "⏳ 待下單", buyer: getMyName(), isMine: true, qty: 1, checked: true });
      await pushState();
    }

    async function duplicateItem(idx) {
      if (!canEdit()) return alert("目前沒有編輯權限！");
      const src = products[idx];
      const copy = JSON.parse(JSON.stringify(src));
      copy.id = makeId();
      copy.buyer = getMyName();
      copy.isMine = true;
      copy.qty = 1;
      copy.status = "⏳ 待下單";
      copy.checked = true;
      products.push(copy);
      await pushState();
      document.getElementById("statusText").innerText = `✨ 已幫「${copy.buyer}」跟團 +1！記得選自己的尺寸顏色喔`;
    }

    async function editImage(idx) {
      if (!canEditRow(products[idx])) return;
      const newUrl = prompt("請貼上商品圖片網址 (右鍵複製圖片位址)：", products[idx].img || "");
      if (newUrl !== null) {
        products[idx].img = newUrl.trim();
        await pushState();
      }
    }

    async function deleteRow(idx) {
      if (!canEditRow(products[idx])) return;
      products.splice(idx, 1);
      await pushState();
    }

    async function clearAll() {
      if (!isAdmin || products.length === 0) return;
      if (confirm(`確定要清空「${currentRoom}」房間的所有商品嗎？`)) {
        products = [];
        document.getElementById("shippingInput").value = 0;
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

    function renderStatusCell(p, idx) {
      const st = p.status || "⏳ 待下單";
      if (isAdmin) {
        return `<select class="status-select" id="status-${idx}" onchange="syncAndPush(${idx})">
          ${ITEM_STATUSES.map(s => `<option value="${s}" ${st === s ? 'selected' : ''}>${s}</option>`).join("")}
        </select>`;
      }
      let cls = "st-wait";
      if (st.includes("已下單")) cls = "st-ordered";
      else if (st.includes("已到貨")) cls = "st-arrived";
      else if (st.includes("斷貨")) cls = "st-out";
      return `<span class="status-badge ${cls}">${escapeHtml(st)}</span>`;
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
      const adminDis = isAdmin ? "" : "disabled";
      const canPlusOne = canEdit();

      if (products.length === 0) {
        tbody.innerHTML = `<tr><td colspan="11" style="text-align:center; color:#999; padding:36px;">「${escapeHtml(currentRoom)}」目前是空的，貼上網址就能一起買啦 🛒</td></tr>`;
        recalculate();
        return;
      }

      const rowsHtml = products.map((p, idx) => {
        if (filterBuyer && (p.buyer || "訪客") !== filterBuyer) return "";

        const rowEditable = canEditRow(p);
        const rowDis = rowEditable ? "" : "disabled";
        const rowClass = rowEditable ? "" : "row-other";

        return `
        <tr class="${rowClass}">
          <td><input type="checkbox" class="item-check" data-idx="${idx}" ${p.checked !== false ? 'checked' : ''} ${rowDis} onchange="syncAndPush(${idx})"></td>
          <td>
            <div class="thumb-box">
              ${p.img ? `<a href="${p.url || p.img}" target="_blank"><img class="thumb" src="${p.img}" alt=""></a>` : '<div class="thumb" style="display:flex;align-items:center;justify-content:center;font-size:11px;color:#aaa;">無圖片</div>'}
              ${rowEditable ? `<button class="btn-img" onclick="editImage(${idx})">換圖</button>` : ''}
            </div>
          </td>
          <td>
            <input type="text" class="title-input" id="title-${idx}" value="${escapeHtml(p.title)}" ${rowDis} onfocus="isEditingNow=true" onblur="isEditingNow=false;syncAndPush(${idx})">
            ${p.url ? `<a class="link-sm" href="${p.url}" target="_blank">開啟原網頁 ↗</a>` : ''}
          </td>
          <td>${renderSpecCell(p, idx, rowDis)}</td>
          <td><input type="text" class="buyer-input" id="buyer-${idx}" value="${escapeHtml(p.buyer || '訪客')}" ${adminDis} title="僅團主可修改買家" onfocus="isEditingNow=true" onblur="isEditingNow=false;syncAndPush(${idx})"></td>
          <td style="text-align:center;">${renderStatusCell(p, idx)}</td>
          <td>
            <select class="curr-select" id="curr-${idx}" ${adminDis} title="僅團主可修改幣別" onchange="syncAndPush(${idx})">
              ${CURRENCIES.map(c => `<option value="${c}" ${p.currency === c ? 'selected' : ''}>${c}</option>`).join("")}
            </select>
          </td>
          <td><input type="number" class="price-input" id="price-${idx}" value="${p.price}" step="any" min="0" ${adminDis} title="僅團主可修改價格" onfocus="isEditingNow=true" onblur="isEditingNow=false;syncAndPush(${idx})"></td>
          <td class="twd-price" id="twd-${idx}">NT$0</td>
          <td><input type="number" class="qty-input" id="qty-${idx}" value="${p.qty || 1}" min="1" ${rowDis} onchange="syncAndPush(${idx})"></td>
          <td>
            <div class="action-cell">
              ${canPlusOne ? `<button class="btn-plus1" title="跟團 +1（複製給自己）" onclick="duplicateItem(${idx})">+1</button>` : ''}
              ${rowEditable ? `<button class="btn-del" title="刪除" onclick="deleteRow(${idx})">✕</button>` : ''}
            </div>
          </td>
        </tr>
      `;
      }).join("");

      tbody.innerHTML = rowsHtml || `<tr><td colspan="11" style="text-align:center; color:#999; padding:30px;">沒有符合篩選條件的商品</td></tr>`;
      recalculate();
    }

    async function syncAndPush(idx) {
      if (!canEditRow(products[idx])) return;
      const cb = document.querySelector(`.item-check[data-idx="${idx}"]`);
      if (!cb) return;
      products[idx].checked = cb.checked;
      products[idx].title = document.getElementById(`title-${idx}`).value;
      products[idx].spec = document.getElementById(`spec-${idx}`).value;
      const noteEl = document.getElementById(`note-${idx}`);
      products[idx].note = noteEl ? noteEl.value : "";
      products[idx].qty = parseInt(document.getElementById(`qty-${idx}`).value) || 1;

      if (isAdmin) {
        products[idx].buyer = document.getElementById(`buyer-${idx}`).value.trim() || "訪客";
        products[idx].currency = document.getElementById(`curr-${idx}`).value;
        products[idx].price = parseFloat(document.getElementById(`price-${idx}`).value) || 0;
        const stEl = document.getElementById(`status-${idx}`);
        if (stEl) products[idx].status = stEl.value;
      }
      await pushState();
    }

    async function toggleAll(master) {
      if (!isAdmin) return;
      products.forEach(p => p.checked = master.checked);
      document.querySelectorAll(".item-check").forEach(cb => cb.checked = master.checked);
      await pushState();
    }

    function getBreakdown() {
      const rates = getRates();
      const feeMult = document.getElementById("feeCheck").checked ? 1.015 : 1;
      const shipping = parseFloat(document.getElementById("shippingInput").value) || 0;
      const shippingMode = document.getElementById("shippingModeSelect").value || "qty";

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
          buyers[buyer].items.push({ title: p.title, fullSpec: formatFullSpec(p), status: p.status || "", qty, unitTwd, subTwd: unitTwd * qty });
        }
      });

      for (const b of Object.values(buyers)) {
        if (shippingMode === "amount") {
          b.shippingShare = totalGoodsTwd > 0 ? Math.round(shipping * (b.goodsTwd / totalGoodsTwd)) : 0;
        } else {
          b.shippingShare = totalQty > 0 ? Math.round((shipping / totalQty) * b.qty) : 0;
        }
        b.finalTwd = b.goodsTwd + b.shippingShare;
      }

      return { totalQty, totalGoodsTwd, totalTwd: totalGoodsTwd + Math.round(shipping), shipping, shippingMode, feeChecked: feeMult > 1, buyers };
    }

    function recalculate() {
      const { totalQty, totalGoodsTwd, totalTwd, buyers } = getBreakdown();
      document.getElementById("selectedCount").innerText = totalQty;
      document.getElementById("totalTwd").innerText = `NT$${totalTwd.toLocaleString()}`;

      const fsCard = document.getElementById("freeShipCard");
      if (freeShipThreshold > 0) {
        fsCard.style.display = "block";
        const pct = Math.min(100, Math.round((totalGoodsTwd / freeShipThreshold) * 100));
        const diff = Math.max(0, Math.round(freeShipThreshold - totalGoodsTwd));
        document.getElementById("freeShipPct").innerText = `${pct}%`;
        const fill = document.getElementById("freeShipFill");
        fill.style.width = `${pct}%`;
        if (diff === 0) {
          fill.classList.add("reached");
          document.getElementById("freeShipMsg").innerText = `🎉 太棒了！商品總額已達 NT$${freeShipThreshold.toLocaleString()} 免運門檻！`;
        } else {
          fill.classList.remove("reached");
          document.getElementById("freeShipMsg").innerText = `🔥 距離免運門檻（NT$${freeShipThreshold.toLocaleString()}）還差 NT$${diff.toLocaleString()}，快揪朋友 +1！`;
        }
      } else {
        fsCard.style.display = "none";
      }

      const splitCard = document.getElementById("splitCard");
      const splitBox = document.getElementById("splitBox");
      const buyerEntries = Object.entries(buyers);

      if (buyerEntries.length > 0) {
        splitCard.style.display = "block";
        splitBox.innerHTML = buyerEntries.map(([name, info]) => {
          const isPaid = !!paidBuyers[name];
          const activeCls = filterBuyer === name ? "active-filter" : "";
          const paidBadge = `<span class="paid-tag ${isPaid ? 'is-paid' : ''}" onclick="togglePaid(event, '${escapeHtml(name)}')" title="${isAdmin ? '點此切換付款狀態' : ''}">${isPaid ? '✅ 已付' : '💰 未付'}</span>`;
          return `
            <div class="buyer-pill ${activeCls}" onclick="setFilterBuyer('${escapeHtml(name)}')">
              <span>👤 ${escapeHtml(name)} (${info.qty}件)</span>
              <span class="pill-amt">NT$${info.finalTwd.toLocaleString()}</span>
              ${paidBadge}
            </div>
          `;
        }).join("");
      } else {
        splitCard.style.display = "none";
      }
    }

    function copyLineText() {
      const { totalQty, totalTwd, shipping, shippingMode, feeChecked, buyers } = getBreakdown();
      if (totalQty === 0) return alert("請至少勾選一件商品！");

      let lines = [`🛍️ 跨國團購結帳清單【${currentRoom}】`];
      if (feeChecked || shipping > 0) {
        let notes = [];
        if (feeChecked) notes.push("含1.5%刷卡費");
        if (shipping > 0) {
          const modeTxt = shippingMode === "amount" ? "依金額比例分攤" : "依件數均分";
          notes.push(`運費NT$${shipping}(${modeTxt})`);
        }
        lines.push(`(${notes.join(" / ")})`);
      }
      lines.push("━━━━━━━━━━━━");

      for (const [name, info] of Object.entries(buyers)) {
        const paidTxt = paidBuyers[name] ? " ✅已付款" : " 💰未付款";
        lines.push(`👤 ${name}（${info.qty}件）${paidTxt}`);
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
      if (bankInfo) {
        lines.push(`🏦 匯款資訊：${bankInfo}`);
      }

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
            shipping: parseFloat(document.getElementById("shippingInput").value) || 0,
            shippingMode: document.getElementById("shippingModeSelect").value || "qty"
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
            st = default_room_state()
            if isinstance(data.get("state"), dict):
                st.update(data["state"])
            return {
                "state": st,
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

def sanitize_products_for_client(prods, utoken, is_admin):
    """隱藏 ownerToken 不外洩，只回傳 isMine 給前端判斷編輯權"""
    safe_list = []
    for p in (prods or []):
        cp = dict(p)
        tok = cp.pop("ownerToken", None)
        cp["isMine"] = True if is_admin else (bool(utoken) and tok == utoken)
        safe_list.append(cp)
    return safe_list

@app.route("/bg.jpg")
def background_image():
    img_path = os.path.join(BASE_DIR, "123.jpg")
    if os.path.exists(img_path):
        return send_file(img_path, mimetype="image/jpeg")
    return "", 404

@app.route("/api/verify_admin", methods=["POST"])
def api_verify_admin():
    ip = get_client_ip()
    now = time.time()
    rec = FAILED_ADMIN_LOG.get(ip, {"count": 0, "until": 0})
    if now < rec["until"]:
        wait_sec = int(rec["until"] - now)
        return jsonify({"ok": False, "error": f"嘗試次數過多，請 {wait_sec} 秒後再試"}), 429

    pin = str((request.json or {}).get("pin", "")).strip()
    if hmac.compare_digest(pin, ADMIN_PIN):
        FAILED_ADMIN_LOG.pop(ip, None)
        return jsonify({"ok": True})

    time.sleep(2)
    rec["count"] += 1
    if rec["count"] >= 5:
        rec["until"] = time.time() + 300
        rec["count"] = 0
    FAILED_ADMIN_LOG[ip] = rec
    return jsonify({"ok": False, "error": "團主密碼錯誤！"})

@app.route("/api/rooms", methods=["GET", "DELETE"])
def api_rooms():
    if request.method == "GET":
        req_pin = request.args.get("adminPin", "").strip()
        if not hmac.compare_digest(req_pin, ADMIN_PIN):
            return jsonify({"ok": False}), 403
        if not FIREBASE_URL:
            return jsonify({"ok": True, "rooms": list(LOCAL_ROOMS.keys())})
        try:
            res = requests.get(f"{FIREBASE_URL}/rooms.json?shallow=true", timeout=6)
            data = res.json()
            rooms = list(data.keys()) if isinstance(data, dict) else []
            return jsonify({"ok": True, "rooms": rooms})
        except Exception:
            return jsonify({"ok": False, "rooms": []})

    data = request.json or {}
    if not hmac.compare_digest(str(data.get("adminPin", "")).strip(), ADMIN_PIN):
        return jsonify({"ok": False}), 403
    room_id = sanitize_room(data.get("room", ""))
    if not room_id:
        return jsonify({"ok": False})
    if not FIREBASE_URL:
        LOCAL_ROOMS.pop(room_id, None)
        return jsonify({"ok": True})
    try:
        requests.delete(f"{FIREBASE_URL}/rooms/{quote(room_id)}.json", timeout=6)
        return jsonify({"ok": True})
    except Exception:
        return jsonify({"ok": False})

@app.route("/api/state", methods=["GET", "POST"])
def api_state():
    if request.method == "GET":
        room_id = sanitize_room(request.args.get("room", "大廳"))
        cid = request.args.get("cid", "").strip()
        utoken = request.args.get("utoken", "").strip()
        name = (request.args.get("name", "訪客").strip() or "訪客")[:15]
        rpw = request.args.get("rpw", "").strip()
        req_pin = request.args.get("adminPin", "").strip()
        is_admin = hmac.compare_digest(req_pin, ADMIN_PIN)

        bundle = get_room_bundle(room_id)
        st = dict(bundle["state"])
        room_pw = str(st.get("roomPassword", "")).strip()

        if room_pw and not is_admin and not hmac.compare_digest(rpw, room_pw):
            return jsonify({"needPassword": True, "cloud": bool(FIREBASE_URL)})

        if cid:
            update_presence(room_id, cid, name)
            bundle = get_room_bundle(room_id)

        now = int(time.time())
        online_names = []
        for _, info in (bundle.get("presence") or {}).items():
            if isinstance(info, dict) and now - info.get("ts", 0) <= 18:
                online_names.append(info.get("name", "訪客"))

        st["products"] = sanitize_products_for_client(st.get("products"), utoken, is_admin)
        if not is_admin:
            st.pop("roomPassword", None)

        return jsonify({
            "needPassword": False,
            "cloud": bool(FIREBASE_URL),
            "state": st,
            "online": online_names
        })

    data = request.json or {}
    room_id = sanitize_room(data.get("room", "大廳"))
    current = get_room_bundle(room_id)["state"]
    req_pin = str(data.get("adminPin", "")).strip()
    is_admin = hmac.compare_digest(req_pin, ADMIN_PIN)
    user_name = (str(data.get("userName", "")).strip() or "訪客")[:15]
    utoken = str(data.get("utoken", "")).strip()
    rpw = str(data.get("rpw", "")).strip()

    room_pw = str(current.get("roomPassword", "")).strip()
    if room_pw and not is_admin and not hmac.compare_digest(rpw, room_pw):
        return jsonify({"ok": False, "error": "入房密碼驗證失敗！"}), 403

    allowed_list = current.get("allowedEditors") or []
    expired = is_deadline_passed(current.get("deadline", ""))
    if (current.get("locked") or expired) and not is_admin and user_name not in allowed_list:
        msg = "結單時間已到，系統已自動鎖定！" if expired else "團主已鎖定此房間，你目前沒有編輯權限！"
        return jsonify({"ok": False, "error": msg}), 403

    existing_prods = current.get("products") or []
    existing_by_id = {p.get("id"): p for p in existing_prods if p.get("id")}
    incoming_products = data.get("products", [])

    if is_admin:
        merged_products = []
        for p in incoming_products:
            cp = dict(p)
            cp.pop("isMine", None)
            old = existing_by_id.get(cp.get("id"))
            # 保留原本商品的 ownerToken，避免團主修改時覆蓋掉原買家金鑰
            cp["ownerToken"] = old.get("ownerToken") if old else utoken
            cp["sig"] = sign_price(cp.get("url", ""), cp.get("price", 0), cp.get("currency", "TWD"))
            merged_products.append(cp)

        new_rates = data.get("rates", current.get("rates", {}))
        new_fee = bool(data.get("feeChecked", False))
        new_shipping = float(data.get("shipping", 0))
        new_ship_mode = data.get("shippingMode", "qty")
        new_free_ship = float(data.get("freeShipThreshold", 0))
        new_bank = str(data.get("bankInfo", "")).strip()
        new_ann = str(data.get("announcement", "")).strip()
        new_deadline = str(data.get("deadline", "")).strip()
        new_paid = data.get("paidBuyers", current.get("paidBuyers", {}))
        new_locked = bool(data.get("locked", current.get("locked", False)))
        new_allowed = data.get("allowedEditors", allowed_list)
        new_rpw = str(data.get("roomPassword", room_pw)).strip()
    else:
        if not utoken:
            return jsonify({"ok": False, "error": "缺少使用者識別金鑰，請重新整理網頁"}), 400

        # 嚴格依 ownerToken 區分：非本人的商品完全不動
        other_items = [p for p in existing_prods if p.get("ownerToken") != utoken]
        my_items = []
        for p in incoming_products:
            pid = p.get("id")
            old = existing_by_id.get(pid)
            if old:
                # 既有商品：必須屬於自己才能改，且價格/幣別/狀態/買家強制維持原值
                if old.get("ownerToken") != utoken:
                    continue
                cp = dict(old)
                cp["checked"] = bool(p.get("checked", True))
                cp["title"] = str(p.get("title", old.get("title", "")))[:120]
                cp["spec"] = str(p.get("spec", ""))[:80]
                cp["note"] = str(p.get("note", ""))[:80]
                cp["img"] = str(p.get("img", old.get("img", "")))[:400]
                cp["qty"] = max(1, min(99, int(p.get("qty", 1) or 1)))
                my_items.append(cp)
            else:
                # 新商品：必須驗證後端價格簽章，否則價格強制歸零待團主核價
                cp = dict(p)
                cp.pop("isMine", None)
                url_val = str(cp.get("url", ""))
                price_val = float(cp.get("price", 0) or 0)
                curr_val = str(cp.get("currency", "TWD")).upper()
                expected_sig = sign_price(url_val, price_val, curr_val)
                if not hmac.compare_digest(str(cp.get("sig", "")), expected_sig):
                    price_val = 0
                    curr_val = "TWD"
                    cp["sig"] = sign_price(url_val, 0, "TWD")

                cp["price"] = price_val
                cp["currency"] = curr_val
                cp["buyer"] = user_name
                cp["ownerToken"] = utoken
                cp["status"] = "⏳ 待下單"
                cp["qty"] = max(1, min(99, int(cp.get("qty", 1) or 1)))
                my_items.append(cp)

        merged_products = other_items + my_items
        new_rates = current.get("rates", {})
        new_fee = bool(current.get("feeChecked", False))
        new_shipping = float(current.get("shipping", 0))
        new_ship_mode = current.get("shippingMode", "qty")
        new_free_ship = float(current.get("freeShipThreshold", 0))
        new_bank = current.get("bankInfo", "")
        new_ann = current.get("announcement", "")
        new_deadline = current.get("deadline", "")
        new_paid = current.get("paidBuyers", {})
        new_locked = bool(current.get("locked", False))
        new_allowed = allowed_list
        new_rpw = room_pw

    new_state = {
        "products": merged_products,
        "rates": new_rates,
        "feeChecked": new_fee,
        "shipping": new_shipping,
        "shippingMode": new_ship_mode,
        "freeShipThreshold": new_free_ship,
        "bankInfo": new_bank,
        "announcement": new_ann,
        "deadline": new_deadline,
        "paidBuyers": new_paid,
        "locked": new_locked,
        "allowedEditors": new_allowed,
        "roomPassword": new_rpw,
        "updatedAt": int(time.time() * 1000)
    }
    save_room_state(room_id, new_state)
    return jsonify({
        "ok": True,
        "updatedAt": new_state["updatedAt"],
        "products": sanitize_products_for_client(merged_products, utoken, is_admin)
    })

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
        sig = sign_price(url, price_num, currency)

        return {"url": url, "title": title, "img": img_url, "price": price_num, "currency": currency, "options": options, "sig": sig}
    except Exception as e:
        sig = sign_price(url, 0, currency)
        return {"url": url, "title": f"抓取受阻，請通知團主補價 ({type(e).__name__})", "img": "", "price": 0, "currency": currency, "options": [], "sig": sig}

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/api/fetch", methods=["POST"])
def api_fetch():
    raw_text = (request.json or {}).get("text", "")
    urls = extract_urls(raw_text)
    items = [scrape_product(url) for url in urls]
    return jsonify({"items": items})

@app.route("/api/export", methods=["POST"])
def api_export():
    data = request.json or {}
    room_name = sanitize_room(data.get("room", "大廳"))
    items = data.get("items", [])
    rates = data.get("rates", {})
    fee_mult = 1.015 if data.get("feeChecked") else 1.0
    shipping = float(data.get("shipping", 0))
    shipping_mode = data.get("shippingMode", "qty")

    total_qty = sum(int(it.get("qty", 1)) for it in items)
    total_goods_twd = sum(
        round(float(it.get("price", 0)) * float(rates.get(it.get("currency", "TWD"), 1)) * fee_mult) * int(it.get("qty", 1))
        for it in items
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = room_name[:28]

    col_headers = ["商品照", "買家", "採購進度", "商品名稱", "規格 / 顏色備註", "幣別", "原幣單價", "台幣單價", "數量", "運費分攤", "小計 (TWD)", "商品連結"]
    ws.append(col_headers)

    header_fill = PatternFill(start_color="5C4742", end_color="5C4742", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, len(col_headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    col_widths = {"A": 15, "B": 10, "C": 12, "D": 30, "E": 20, "F": 7, "G": 11, "H": 11, "I": 7, "J": 10, "K": 12, "L": 30}
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
        goods_sub = unit_twd * qty

        if shipping_mode == "amount":
            ship_share = round(shipping * (goods_sub / total_goods_twd)) if total_goods_twd > 0 else 0
        else:
            ship_share = round((shipping / total_qty) * qty) if total_qty > 0 else 0

        subtotal_twd = goods_sub + ship_share
        grand_total += subtotal_twd

        ws.row_dimensions[row_idx].height = 78

        spec_text = item.get("fullSpec") or item.get("spec", "")
        ws.cell(row=row_idx, column=2, value=item.get("buyer", "訪客")).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row_idx, column=3, value=item.get("status", "⏳ 待下單")).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row_idx, column=4, value=item.get("title", "")).alignment = Alignment(vertical="center", wrap_text=True)
        ws.cell(row=row_idx, column=5, value=spec_text).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(row=row_idx, column=6, value=curr).alignment = Alignment(horizontal="center", vertical="center")

        for col_i, val, fmt in [
            (7, price, '#,##0.##'),
            (8, unit_twd, '#,##0'),
            (10, ship_share, '#,##0'),
            (11, subtotal_twd, '#,##0')
        ]:
            c = ws.cell(row=row_idx, column=col_i, value=val)
            c.number_format = fmt
            c.alignment = Alignment(horizontal="right", vertical="center")

        ws.cell(row=row_idx, column=9, value=qty).alignment = Alignment(horizontal="center", vertical="center")

        url_val = item.get("url", "")
        c_url = ws.cell(row=row_idx, column=12, value=url_val)
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
    ws.cell(row=sum_row, column=8, value="總計").font = Font(bold=True)
    ws.cell(row=sum_row, column=8).alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(row=sum_row, column=9, value=total_qty).font = Font(bold=True)
    ws.cell(row=sum_row, column=9).alignment = Alignment(horizontal="center", vertical="center")

    s_ship = ws.cell(row=sum_row, column=10, value=round(shipping))
    s_ship.number_format = '#,##0'
    s_ship.font = Font(bold=True)
    s_ship.alignment = Alignment(horizontal="right", vertical="center")

    s_twd = ws.cell(row=sum_row, column=11, value=grand_total)
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
