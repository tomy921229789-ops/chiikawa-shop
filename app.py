import io
import os
import re
import json
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
    .
