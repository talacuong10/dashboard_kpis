"""
HDBank - Web Dashboard Tổng Quan KPIs & Lương Phụ Trội QHKHDN
============================================================
Chuẩn hóa 100% theo dữ liệu thực tế và tài liệu Prompt_Dac_ta_Cong_thuc_Dashboard_KPIs_LPT.md:
- 2 nguồn dữ liệu độc lập: Data import KPIs (sheet "Dashboard KPIs và LPT") + Data import số ĐVKD (sheet "Dashboard số liệu toàn hàng")
- 3 Slicer phong cách PowerBI: KHU VỰC (KV1-KV6 + ĐVKD + Nhân sự), SỐ THÁNG (Theo tháng cụ thể), KHỐI CHỈ TIÊU (Bảng A-I)
- 6 Thẻ KPI Pill chuẩn màu & chỉ tiêu thật: Tổng NS, Đạt KPIs, Thỏa LPT, Tổng ĐVKD, KHDN mới, KHTD tăng ròng
- Delta tự động tính toán từ dữ liệu thật so với tháng trước đó (không hardcode)
- 3 Biểu đồ: Donut Đạt KPIs, Bar Tỷ lệ đạt KPIs theo 6 Khu vực (%), Bar Tỷ lệ thỏa LPT theo 6 Khu vực (%)
- Bảng ma trận chi tiết 9 Khối chỉ tiêu (theo từng tháng) & Xuất báo cáo Excel 9 sheet
"""

import base64
import hashlib
import io
import json
import os
import secrets
import socket
import sys
import unicodedata
from datetime import datetime
from pathlib import Path
import random
from typing import Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ──────────────────────────────────────────────────────────────────────────────
# CẤU HÌNH TRANG VÀ CSS ĐẶC TẢ GIAO DIỆN HDBANK
# ──────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="HDBank - Kết Quả Tổng Quan KPIs & LPT",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

HDBANK_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@400;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
}

.block-container {
    padding-top: 0.8rem !important;
    padding-bottom: 1.5rem !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
    max-width: 98% !important;
}

/* Header Container */
.hdbank-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 8px 6px 14px 6px;
    border-bottom: 1.5px solid #e5e7eb;
    margin-bottom: 12px;
    gap: 16px;
}
.hdbank-brand {
    display: flex;
    align-items: center;
    flex-shrink: 0;
}
.hdbank-brand img {
    height: 85px !important;
    max-height: 90px !important;
    width: auto !important;
    object-fit: contain !important;
}
.hdbank-title-center {
    text-align: center;
    flex-grow: 1;
    padding: 0 10px;
}
.hdbank-main-title {
    font-size: 20px;
    font-weight: 800;
    color: #111827;
    letter-spacing: 0.2px;
    margin: 0;
    text-transform: uppercase;
}
.hdbank-sub-title {
    font-size: 13.5px;
    color: #4b5563;
    margin-top: 4px;
    font-weight: 500;
}
.hdbank-date-right {
    text-align: right;
    font-size: 12px;
    color: #374151;
    line-height: 1.6;
    flex-shrink: 0;
}

/* ─────────────────────────────────────────────────────────────────────────
   SLICER BOX STYLES & EQUAL HEIGHT (Streamlit 1.42+ & legacy compatible)
   ───────────────────────────────────────────────────────────────────────── */
div[data-testid="stHorizontalBlock"]:has(.slicer-card-header) {
    align-items: stretch !important;
}
div[data-testid="stHorizontalBlock"]:has(.slicer-card-header) > div[data-testid="stColumn"] {
    display: flex !important;
    flex-direction: column !important;
}
div[data-testid="stHorizontalBlock"]:has(.slicer-card-header) > div[data-testid="stColumn"] > div[data-testid="stVerticalBlock"] {
    height: 100% !important;
    flex: 1 1 auto !important;
    display: flex !important;
    flex-direction: column !important;
}
div[data-testid="stHorizontalBlock"]:has(.slicer-card-header) div[data-testid="stLayoutWrapper"] {
    height: 100% !important;
    flex: 1 1 auto !important;
    display: flex !important;
    flex-direction: column !important;
}

/* Slicer container card — viền trên pastel riêng theo từng panel (xem 3 rule bên dưới) */
div[data-testid="stHorizontalBlock"]:has(.slicer-card-header) div[data-testid="stLayoutWrapper"]:has(.slicer-card-header) > div[data-testid="stVerticalBlock"],
div[data-testid="stVerticalBlockBorderWrapper"]:has(.slicer-card-header) {
    border: 1px solid #e2e8f0 !important;
    border-top: 3.5px solid #cbd5e1 !important;
    border-radius: 6px !important;
    background: #ffffff !important;
    padding: 10px 12px 12px 12px !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04) !important;
    height: 100% !important;
    flex: 1 1 auto !important;
    display: flex !important;
    flex-direction: column !important;
    justify-content: flex-start !important;
}
/* Viền trên pastel riêng theo từng Slicer: KHU VỰC = tím, SỐ THÁNG = xanh dương, KHỐI CHỈ TIÊU = hồng */
div[data-testid="stHorizontalBlock"]:has(.slicer-card-header) div[data-testid="stLayoutWrapper"]:has(.slicer-card-header-kv) > div[data-testid="stVerticalBlock"],
div[data-testid="stVerticalBlockBorderWrapper"]:has(.slicer-card-header-kv) {
    border-top: 3.5px solid #8b7ff5 !important;
}
div[data-testid="stHorizontalBlock"]:has(.slicer-card-header) div[data-testid="stLayoutWrapper"]:has(.slicer-card-header-thang) > div[data-testid="stVerticalBlock"],
div[data-testid="stVerticalBlockBorderWrapper"]:has(.slicer-card-header-thang) {
    border-top: 3.5px solid #60a5fa !important;
}
div[data-testid="stHorizontalBlock"]:has(.slicer-card-header) div[data-testid="stLayoutWrapper"]:has(.slicer-card-header-ct) > div[data-testid="stVerticalBlock"],
div[data-testid="stVerticalBlockBorderWrapper"]:has(.slicer-card-header-ct) {
    border-top: 3.5px solid #ec4899 !important;
}

/* Row gap inside slicer cards */
div[data-testid="stColumn"]:has(.slicer-card-header) div[data-testid="stHorizontalBlock"] {
    gap: 6px !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[data-testid="stVerticalBlock"] {
    gap: 8px !important;
}
/* Đảm bảo mỗi block markdown (header/label) và mỗi hàng nút là 1 khối tách bạch,
   không bị co lại/chồng lên nhau do Streamlit collapse margin của element rỗng */
div[data-testid="stColumn"]:has(.slicer-card-header) div[data-testid="stElementContainer"] {
    width: 100% !important;
}

.slicer-card-header {
    font-size: 11.5px !important;
    font-weight: 800 !important;
    color: #111827 !important;
    letter-spacing: 0.3px !important;
    text-transform: uppercase !important;
    padding-bottom: 5px !important;
    border-bottom: 1px solid #f1f5f9 !important;
    margin-bottom: 4px !important;
    display: flex !important;
    justify-content: space-between !important;
    align-items: center !important;
}

.slicer-sub-label {
    font-size: 9.5px !important;
    font-weight: 700 !important;
    letter-spacing: 0.4px !important;
    text-transform: uppercase !important;
    color: #64748b !important;
    width: 100% !important;
    margin: 4px 0 8px 0 !important;
    padding: 0 !important;
    display: flex !important;
    align-items: center !important;
    gap: 5px !important;
    line-height: 1.4 !important;
    position: relative !important;
    z-index: 1 !important;
}
.sub-label-indicator {
    display: inline-block;
    width: 6px;
    height: 6px;
    border-radius: 50%;
}
.indicator-navy { background-color: #1e3a8a; }
.indicator-slate { background-color: #64748b; }
.indicator-gold { background-color: #f59e0b; }
.slicer-divider {
    height: 1px;
    background-color: #e5e7eb;
    margin: 5px 0 4px 0;
}

/* Slicer button overrides inside the slicer containers */
div[data-testid="stColumn"]:has(.slicer-card-header) .stButton > button {
    width: 100% !important;
    border-radius: 6px !important;
    margin: 0 !important;
    transition: all 0.15s ease !important;
}

/* PRIMARY (ACTIVE/SELECTED) BUTTON — mặc định trung tính, các nhóm pastel bên dưới override theo panel */
div[data-testid="stColumn"]:has(.slicer-card-header) .stButton > button[kind="primary"] {
    border: none !important;
    color: #ffffff !important;
    font-weight: 700 !important;
}

/* KHU VỰC (nút "Toàn Hàng" + KV1-KV6 + ĐVKD con): Tím pastel */
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_all_kv"] .stButton > button[kind="primary"],
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_kv"] .stButton > button[kind="primary"],
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_dv_"] .stButton > button[kind="primary"] {
    background-color: #8b7ff5 !important;
    box-shadow: 0 1px 3px rgba(139, 127, 245, 0.35) !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_all_kv"] .stButton > button[kind="primary"]:hover,
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_kv"] .stButton > button[kind="primary"]:hover,
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_dv_"] .stButton > button[kind="primary"]:hover {
    background-color: #7566e8 !important;
}

/* SỐ THÁNG (nút "Tất cả các tháng" + Kỳ lũy kế + Theo tháng cụ thể): Xanh dương pastel */
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_all_months"] .stButton > button[kind="primary"],
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_k_"] .stButton > button[kind="primary"],
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_m_"] .stButton > button[kind="primary"] {
    background-color: #60a5fa !important;
    box-shadow: 0 1px 3px rgba(96, 165, 250, 0.35) !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_all_months"] .stButton > button[kind="primary"]:hover,
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_k_"] .stButton > button[kind="primary"]:hover,
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_m_"] .stButton > button[kind="primary"]:hover {
    background-color: #3b82f6 !important;
}

/* KHỐI CHỈ TIÊU (nút "Tất cả chỉ tiêu" + từng chỉ tiêu): Hồng pastel */
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_all_ct"] .stButton > button[kind="primary"],
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_ct_"] .stButton > button[kind="primary"] {
    background-color: #ec4899 !important;
    box-shadow: 0 1px 3px rgba(236, 72, 153, 0.35) !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_all_ct"] .stButton > button[kind="primary"]:hover,
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_ct_"] .stButton > button[kind="primary"]:hover {
    background-color: #db2777 !important;
}

/* 1. KHU VỰC block: nhóm nút giãn đều lấp đầy chiều cao container (fix mất cân đối v2.7) */
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_kv"] {
    flex: 1 1 auto !important;
    display: flex !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_kv"] > div {
    flex: 1 1 auto !important;
    display: flex !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_kv"] .stButton {
    flex: 1 1 auto !important;
    display: flex !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_kv"] .stButton > button {
    flex: 1 1 auto !important;
    min-height: 52px !important;
    height: 100% !important;
    font-size: 11.5px !important;
    font-weight: 700 !important;
    padding: 2px 6px !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_kv"] .stButton > button[kind="secondary"] {
    background-color: #ffffff !important;
    border: 1px solid #d1d5db !important;
    color: #1f2937 !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_kv"] .stButton > button[kind="secondary"]:hover {
    background-color: #f3f4f6 !important;
    border-color: #9ca3af !important;
    color: #111827 !important;
}
/* Các cột r1_c1 / r1_c2 chứa 3 nút KV: giãn đều theo chiều dọc để lấp hết khối */
div[data-testid="stColumn"]:has(.slicer-card-header) div[data-testid="stHorizontalBlock"]:has(div[class*="st-key-btn_kv"]) {
    flex: 1 1 auto !important;
    height: 100% !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[data-testid="stHorizontalBlock"]:has(div[class*="st-key-btn_kv"]) div[data-testid="stColumn"] {
    display: flex !important;
    flex-direction: column !important;
    height: 100% !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[data-testid="stHorizontalBlock"]:has(div[class*="st-key-btn_kv"]) div[data-testid="stVerticalBlock"] {
    display: flex !important;
    flex-direction: column !important;
    justify-content: space-between !important;
    height: 100% !important;
    flex: 1 1 auto !important;
    gap: 6px !important;
}

/* 1b. KHU VỰC: nút ĐVKD (MA_DV) con — xổ ra khi đã chọn 1 Khu vực, đa lựa chọn */
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_dv_"] .stButton > button {
    height: 30px !important;
    min-height: 30px !important;
    max-height: 30px !important;
    font-size: 10px !important;
    font-weight: 600 !important;
    padding: 2px 5px !important;
    white-space: normal !important;
    line-height: 1.15 !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_dv_"] .stButton > button[kind="secondary"] {
    background-color: #f5f3ff !important;
    border: 1px solid #ddd6fe !important;
    color: #5b21b6 !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_dv_"] .stButton > button[kind="secondary"]:hover {
    background-color: #ede9fe !important;
    border-color: #c4b5fd !important;
    color: #4c1d95 !important;
}

/* 2. SỐ THÁNG buttons (Kỳ lũy kế: Soft Navy tint) */
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_k_"] .stButton > button {
    height: 28px !important;
    min-height: 28px !important;
    max-height: 28px !important;
    font-size: 9.5px !important;
    font-weight: 600 !important;
    padding: 2px 4px !important;
    white-space: nowrap !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_k_"] .stButton > button[kind="secondary"] {
    background-color: #f0f4f9 !important;
    border: 1px solid #cbd5e1 !important;
    color: #1e3a8a !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_k_"] .stButton > button[kind="secondary"]:hover {
    background-color: #e2eaf4 !important;
    border-color: #94a3b8 !important;
    color: #0f2c59 !important;
}

/* 3. SỐ THÁNG buttons (Theo tháng: Soft Slate tint) */
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_m_"] .stButton > button {
    height: 28px !important;
    min-height: 28px !important;
    max-height: 28px !important;
    font-size: 10.5px !important;
    font-weight: 600 !important;
    padding: 2px 5px !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_m_"] .stButton > button[kind="secondary"] {
    background-color: #f8fafc !important;
    border: 1px solid #e2e8f0 !important;
    color: #475569 !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_m_"] .stButton > button[kind="secondary"]:hover {
    background-color: #f1f5f9 !important;
    border-color: #cbd5e1 !important;
    color: #1e293b !important;
}

/* 3b. ĐỒNG BỘ ĐỘ RỘNG 6 CỘT trong mọi hàng của khối SỐ THÁNG (kể cả hàng lẻ chỉ có
   1 nút như "Kỳ 7T" hay "T07/2026") — ép mọi stColumn trong các hàng này có cùng
   flex-basis, tránh trường hợp cột trống bị co lại khiến nút còn lại phình to ra. */
div[data-testid="stColumn"]:has(.slicer-card-header) div[data-testid="stHorizontalBlock"]:has(div[class*="st-key-btn_k_"]) > div[data-testid="stColumn"],
div[data-testid="stColumn"]:has(.slicer-card-header) div[data-testid="stHorizontalBlock"]:has(div[class*="st-key-btn_m_"]) > div[data-testid="stColumn"] {
    flex: 1 1 0px !important;
    min-width: 0 !important;
    max-width: none !important;
    width: auto !important;
}
/* Các hàng nút Kỳ/Tháng luôn cách đều nhau kể cả khi hàng chỉ có 1 nút thật sự */
div[data-testid="stColumn"]:has(.slicer-card-header) div[data-testid="stHorizontalBlock"]:has(div[class*="st-key-btn_k_"]),
div[data-testid="stColumn"]:has(.slicer-card-header) div[data-testid="stHorizontalBlock"]:has(div[class*="st-key-btn_m_"]) {
    width: 100% !important;
    margin-bottom: 2px !important;
}

/* 4. KHỐI CHỈ TIÊU buttons (Fixed 44px height, multi-line wrap, perfectly aligned rows) */
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_ct_"] .stButton > button {
    height: 44px !important;
    min-height: 44px !important;
    max-height: 44px !important;
    padding: 3px 6px !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    text-align: center !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_ct_"] .stButton > button,
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_ct_"] .stButton > button p {
    font-size: 10.5px !important;
    font-weight: 600 !important;
    line-height: 1.22 !important;
    white-space: normal !important;
    word-break: break-word !important;
    margin: 0 !important;
    padding: 0 !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_ct_"] .stButton > button[kind="secondary"] {
    background-color: #ffffff !important;
    border: 1px solid #d1d5db !important;
    color: #1f2937 !important;
}
div[data-testid="stColumn"]:has(.slicer-card-header) div[class*="st-key-btn_ct_"] .stButton > button[kind="secondary"]:hover {
    background-color: #f3f4f6 !important;
    border-color: #9ca3af !important;
    color: #111827 !important;
}

/* 6 KPI Cards (Pill shape) — nền pastel nhạt riêng theo từng thẻ + icon badge tròn màu đậm hơn */
.kpi-pill {
    border: none;
    border-radius: 20px;
    padding: 10px 8px;
    text-align: center;
    box-shadow: 0 2px 6px rgba(15, 23, 42, 0.06);
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    height: 158px;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
    position: relative;
}
.kpi-pill:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.12);
}
.kpi-icon-badge {
    position: absolute;
    top: 10px;
    left: 12px;
    width: 26px;
    height: 26px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 13px;
}
.kpi-bg-purple { background: #f1efff; }
.kpi-bg-green { background: #ecfdf5; }
.kpi-bg-orange { background: #fff7ed; }
.kpi-bg-blue { background: #eff6ff; }
.kpi-bg-pink { background: #fdf2f8; }
.kpi-bg-amber { background: #fffbeb; }
.kpi-icon-purple { background: #8b7ff5; color: #ffffff; }
.kpi-icon-green { background: #34d399; color: #ffffff; }
.kpi-icon-orange { background: #fb923c; color: #ffffff; }
.kpi-icon-blue { background: #60a5fa; color: #ffffff; }
.kpi-icon-pink { background: #f472b6; color: #ffffff; }
.kpi-icon-amber { background: #fbbf24; color: #ffffff; }
.kpi-title {
    font-size: 14.5px;
    font-weight: 800;
    color: #1f2937;
    line-height: 1.3;
    min-height: 36px;
    display: flex;
    align-items: center;
    justify-content: center;
    padding-top: 14px;
}
.kpi-value-container {
    display: flex;
    align-items: baseline;
    justify-content: center;
    gap: 8px;
    margin: 4px 0;
}
.kpi-big-num {
    font-size: 40px;
    font-weight: 800;
    line-height: 1;
    letter-spacing: -0.5px;
}
.kpi-sub-pct {
    font-size: 16.5px;
    font-weight: 800;
    color: #1f2937;
}
.kpi-delta {
    font-size: 13.5px;
    font-weight: 800;
}

/* Delta color */
.delta-up { color: #059669; }
.delta-down { color: #e11d48; }
.delta-neutral { color: #6b7280; }

/* KPI value colors — pastel, đồng bộ với nền .kpi-bg-* tương ứng */
.color-blue { color: #3b82f6; }
.color-green { color: #10b981; }
.color-orange { color: #fb923c; }
.color-navy { color: #8b7ff5; }
.color-magenta { color: #ec4899; }
.color-red { color: #fbbf24; }

/* Chart Container Card */
.chart-box {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 8px;
    padding: 8px 10px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}

/* ─────────────────────────────────────────────────────────────────────────
   (v4.27) 2 TAB CẤP CAO NHẤT (Dashboard KPIs và LPT / Dashboard số liệu ĐVKD)
   — làm nổi bật rõ ràng (nền thẻ + chữ đậm + gạch chân màu đỏ thương hiệu khi
   đang chọn) thay vì kiểu tab mặc định của Streamlit (chữ nhỏ, ít tương phản,
   dễ bị đọc nhầm là "trống"). Marker ẩn `.top-nav-marker` (render ngay TRƯỚC
   `st.tabs()` cấp cao nhất trong main()) + `:has(> div ...)` (chỉ khớp đúng
   1 cấp con trực tiếp) đảm bảo CSS này CHỈ áp dụng cho ĐÚNG 2 tab cấp cao
   nhất, không lan sang các `st.tabs()` con khác trong app (vd 3 tab "Dashboard
   Tổng Quan/Báo Cáo/Xuất Excel" bên trong tab KPIs) — theo đúng kỹ thuật
   `:has()` đã dùng cho 3 khối Slicer phía trên.
   ───────────────────────────────────────────────────────────────────────── */
div[data-testid="stVerticalBlock"]:has(> div .top-nav-marker) > div [data-testid="stTabs"] [data-baseweb="tab-list"] {
    gap: 8px !important;
    border-bottom: 2px solid #e2e8f0 !important;
    background: #f8fafc !important;
    padding: 6px 6px 0 6px !important;
    border-radius: 10px 10px 0 0 !important;
    margin-bottom: 4px !important;
}
div[data-testid="stVerticalBlock"]:has(> div .top-nav-marker) > div [data-testid="stTabs"] [data-baseweb="tab"] {
    height: 48px !important;
    padding: 0 22px !important;
    background: transparent !important;
    border-radius: 8px 8px 0 0 !important;
}
div[data-testid="stVerticalBlock"]:has(> div .top-nav-marker) > div [data-testid="stTabs"] [data-baseweb="tab"] p {
    font-size: 15px !important;
    font-weight: 700 !important;
    color: #64748b !important;
}
div[data-testid="stVerticalBlock"]:has(> div .top-nav-marker) > div [data-testid="stTabs"] [data-baseweb="tab"]:hover p {
    color: #334155 !important;
}
div[data-testid="stVerticalBlock"]:has(> div .top-nav-marker) > div [data-testid="stTabs"] [aria-selected="true"] {
    background: #ffffff !important;
    box-shadow: 0 -1px 3px rgba(0,0,0,0.05) !important;
}
div[data-testid="stVerticalBlock"]:has(> div .top-nav-marker) > div [data-testid="stTabs"] [aria-selected="true"] p {
    color: #ed1c24 !important;
    font-weight: 800 !important;
}
div[data-testid="stVerticalBlock"]:has(> div .top-nav-marker) > div [data-testid="stTabs"] [data-baseweb="tab-highlight"] {
    background-color: #ed1c24 !important;
    height: 3px !important;
}
</style>
"""

st.markdown(HDBANK_CSS, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# QUẢN LÝ SESSION STATE BỘ LỌC VÀ DỮ LIỆU
# ──────────────────────────────────────────────────────────────────────────────

if "filter_kv" not in st.session_state:
    # v4.12: đổi từ chuỗi đơn sang list (đa lựa chọn) — [] = Toàn Hàng, danh sách 1 phần tử = chọn
    # đúng 1 Khu vực (hành vi cũ), danh sách nhiều phần tử = đa lựa chọn mới (KPI/ĐVKD/Nhân sự gộp
    # theo TẤT CẢ Khu vực đã chọn).
    st.session_state.filter_kv = []

if "filter_thang" not in st.session_state:
    # v4.12: đổi từ chuỗi đơn sang list (đa lựa chọn) — [] = Toàn bộ dữ liệu, danh sách 1 phần tử =
    # chọn đúng 1 tháng (hành vi cũ, có Delta so tháng trước), danh sách nhiều phần tử = đa lựa
    # chọn mới (giống "Tất cả các tháng" nhưng chỉ lọc đúng các tháng đã chọn, không tính Delta).
    st.session_state.filter_thang = []

if "filter_chi_tieu" not in st.session_state:
    # v4.11: đổi từ chuỗi đơn sang list (đa lựa chọn) — [] = chưa chọn gì, danh sách 1 phần tử =
    # chọn đúng 1 chỉ tiêu (hành vi cũ), danh sách nhiều phần tử = đa lựa chọn mới, đủ 9 phần tử =
    # tương đương "Tất cả chỉ tiêu" (xem is_all_ct trong main()).
    st.session_state.filter_chi_tieu = []

if "filter_ma_dv" not in st.session_state:
    st.session_state.filter_ma_dv = []

if "filter_nhan_su" not in st.session_state:
    st.session_state.filter_nhan_su = []

if "filter_dvkd_thang" not in st.session_state:
    st.session_state.filter_dvkd_thang = []

if "filter_dvkd_chi_tieu" not in st.session_state:
    # (v4.18) Đổi từ 1 giá trị đơn (int|None) sang LIST — hộp CHỈ TIÊU nay là multiselect (đa lựa
    # chọn + nút "Tất cả"), theo đúng pattern đã dùng cho filter_chi_tieu/filter_ma_dv/filter_kv.
    st.session_state.filter_dvkd_chi_tieu = []

if "active_top_tab" not in st.session_state:
    # (v4.18) Theo dõi tab cấp cao (0=KPIs/LPT, 1=ĐVKD) đang xem — st.tabs() không tự lưu tab đang
    # chọn phía server nên phải tự quản lý, dùng để tự bấm lại đúng tab đó qua JS sau mỗi rerun.
    st.session_state.active_top_tab = 0

if "filter_dvkd_don_vi" not in st.session_state:
    # v4.15: trang "Dashboard số liệu ĐVKD" nay có nhiều Đơn vị kinh doanh (trước đó chỉ có "Toàn
    # hàng" duy nhất) — mặc định chọn "Toàn hàng". Dùng literal "Toàn hàng" ở đây (không tham chiếu
    # hằng số DVKD_DEFAULT_DON_VI) vì khối init này chạy Ở CẤP MODULE ngay khi nạp file, TRƯỚC khi
    # hằng số đó được định nghĩa xa hơn ở dưới — tham chiếu sớm sẽ ném NameError khi Streamlit nạp
    # script. `DVKD_DEFAULT_DON_VI` vẫn dùng an toàn bên trong các HÀM (gọi sau khi module đã nạp
    # xong), chỉ tránh dùng tại các câu lệnh module-level nằm TRƯỚC vị trí định nghĩa nó.
    st.session_state.filter_dvkd_don_vi = "Toàn hàng"

if "uploaded_df" not in st.session_state:
    st.session_state.uploaded_df = None

if "uploaded_file_name" not in st.session_state:
    st.session_state.uploaded_file_name = ""

if "uploaded_dvkd_items" not in st.session_state:
    st.session_state.uploaded_dvkd_items = None

if "uploaded_file_name_dvkd" not in st.session_state:
    st.session_state.uploaded_file_name_dvkd = ""

if "use_demo" not in st.session_state:
    st.session_state.use_demo = False

# Hỗ trợ URL query params để test và bookmark trạng thái bộ lọc
params = st.query_params
if params.get("demo") in ["1", "true", "True"]:
    st.session_state.use_demo = True
if "kv" in params:
    st.session_state.filter_kv = [params.get("kv")] if params.get("kv") else []
if "thang" in params:
    st.session_state.filter_thang = [params.get("thang")] if params.get("thang") else []
if "ct" in params:
    st.session_state.filter_chi_tieu = [params.get("ct")] if params.get("ct") else []


# ──────────────────────────────────────────────────────────────────────────────
# XỬ LÝ TIỀN DỮ LIỆU FILE EXCEL (Sheet 'Data import')
# ──────────────────────────────────────────────────────────────────────────────

REQUIRED_COLUMNS = [
    "SYM_RUN_DATE",
    "MA_KV",
    "TEN_KV",
    "MA_DV",
    "TEN_DV",
    "TEN_CV",
    "KET_QUA_KPIS",
    "KET_QUA_LPT",
    "SL_KHDN_MOI_N",
    "SL_KHDN_TANG_RONG",
    "SL_KH_TD_TANG_RONG",
    "Hệ số LPT thực nhận",
    "TONG_CV_THANG_N",
    "TONG_CV_THANG_N_1",
    "TANG_RONG_CV",
    "TONG_HD_THANG_N",
    "TONG_HD_THANG_N_1",
    "TANG_RONG_HD",
    "DU_CHI_HUY_DONG",
    "DU_THU_CHO_VAY",
    "THU_TU_THANH_TOAN",
    "THU_QUAN_LY_TIEN_MAT",
    "THU_NGAN_HANG_DIEN_TU",
    "THU_NGAN_QUY",
    "THU_TAI_TRO_THUONG_MAI",
    "THU_BAO_LANH",
    "THU_TIN_DUNG",
    "THU_MUA_BAN_NGOAI_TE",
]

# 16 cột số liệu cấp ĐVKD cho Khối G (Dư nợ) / H (Huy động) / I (TOI) — CẢNH BÁO: các cột này
# lặp lại y hệt trên mọi dòng nhân sự cùng 1 (MA_DV, SYM_RUN_DATE). TUYỆT ĐỐI KHÔNG SUM trực
# tiếp theo dòng — luôn dùng agg_unit_level_metric() (dedup trước khi sum). Xem Mục 6 memory_project.md.
UNIT_LEVEL_NUMERIC_COLUMNS = [
    "TONG_CV_THANG_N",
    "TONG_CV_THANG_N_1",
    "TANG_RONG_CV",
    "TONG_HD_THANG_N",
    "TONG_HD_THANG_N_1",
    "TANG_RONG_HD",
    "DU_CHI_HUY_DONG",
    "DU_THU_CHO_VAY",
    "THU_TU_THANH_TOAN",
    "THU_QUAN_LY_TIEN_MAT",
    "THU_NGAN_HANG_DIEN_TU",
    "THU_NGAN_QUY",
    "THU_TAI_TRO_THUONG_MAI",
    "THU_BAO_LANH",
    "THU_TIN_DUNG",
    "THU_MUA_BAN_NGOAI_TE",
]

# 6 Dải % Hệ số LPT thực nhận chuẩn hóa theo Mục 0.4 trong Đặc tả
LPT_BINS = [
    ("< 0%", lambda x: x < 0),
    ("0% ≤ … < 20%", lambda x: (x >= 0.0) & (x < 0.20)),
    ("20% ≤ … < 50%", lambda x: (x >= 0.20) & (x < 0.50)),
    ("50% ≤ … < 70%", lambda x: (x >= 0.50) & (x < 0.70)),
    ("70% ≤ … < 100%", lambda x: (x >= 0.70) & (x < 1.00)),
    ("= 100%", lambda x: x >= 1.00),
]

STANDARD_KV_ORDER = [
    "Khu vực 1 (KV Hà Nội)",
    "Khu vực 2 (KV Miền Bắc)",
    "Khu vực 3 (KV Miền Trung)",
    "Khu vực 4 (KV TPHCM)",
    "Khu vực 5 (KV Đồng Nai)",
    "Khu vực 6 (KV Tây Nam Bộ)",
]

# ──────────────────────────────────────────────────────────────────────────────
# BẢNG MÀU PASTEL — MỖI KHU VỰC / ĐƠN VỊ KINH DOANH 1 MÀU ĐẠI DIỆN RIÊNG
# ──────────────────────────────────────────────────────────────────────────────

# 6 màu pastel cố định cho 6 Khu vực chuẩn — dùng xuyên suốt mọi biểu đồ so sánh theo KV
# (khớp đúng nhãn "kv_label" được tạo ra trong compute_overview_metrics / calc_monthly_series_by_kv).
KV_COLOR_MAP = {
    "KV 1 (Hà Nội)": "#8b7ff5",
    "KV 2 (Miền Bắc)": "#f472b6",
    "KV 3 (Miền Trung)": "#fb923c",
    "KV 4 (TPHCM)": "#60a5fa",
    "KV 5 (Đồng Nai)": "#34d399",
    "KV 6 (Tây Nam Bộ)": "#22d3ee",
}

# Bảng màu pastel mở rộng cho Đơn vị kinh doanh (MA_DV) — lặp vòng theo chỉ số khi số ĐVKD > 6.
UNIT_COLOR_PALETTE = [
    "#8b7ff5", "#f472b6", "#fb923c", "#60a5fa", "#34d399", "#22d3ee",
    "#fbbf24", "#fb7185", "#4ade80", "#c084fc", "#38bdf8", "#fdba74",
]

# Cặp màu pastel semantic dùng cho biểu đồ phân kỳ âm/dương (thay cho đỏ/xanh đậm trước đây)
PASTEL_POSITIVE = "#34d399"
PASTEL_NEGATIVE = "#fb7185"
PASTEL_TARGET_LINE = "#f43f5e"


def get_kv_color(kv_label: str, idx: int = 0) -> str:
    """Trả về màu pastel đại diện cho 1 Khu vực (theo nhãn hiển thị dạng 'KV n (...)')."""
    return KV_COLOR_MAP.get(kv_label, UNIT_COLOR_PALETTE[idx % len(UNIT_COLOR_PALETTE)])


def get_unit_color(idx: int) -> str:
    """Trả về màu pastel đại diện cho 1 Đơn vị kinh doanh, lặp vòng theo chỉ số thứ tự."""
    return UNIT_COLOR_PALETTE[idx % len(UNIT_COLOR_PALETTE)]


def _hex_to_rgba(hex_color: str, alpha: float) -> str:
    """Chuyển màu hex '#rrggbb' sang chuỗi 'rgba(r, g, b, alpha)' — dùng làm màu tô vùng (fill) nhạt
    cho Area Chart, giữ đúng tông màu pastel gốc của nhóm nhưng giảm độ đậm (v4.8)."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r}, {g}, {b}, {alpha})"


def find_data_sheet(xls: pd.ExcelFile) -> Optional[str]:
    """Tìm sheet 'Data import' hoặc sheet đầu tiên có SYM_RUN_DATE."""
    for name in xls.sheet_names:
        if name.strip().lower() == "data import":
            return name
    for name in xls.sheet_names:
        try:
            sample = pd.read_excel(xls, sheet_name=name, nrows=5)
            if "SYM_RUN_DATE" in sample.columns:
                return name
        except Exception:
            continue
    return None


def parse_lpt_value(val) -> Optional[float]:
    """
    Chuẩn hóa Hệ số LPT thực nhận thành float (0.0 - 1.0).
    Trả về None nếu rỗng/NULL để không tính vào bất kỳ dải nào trong Bảng C.
    """
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s in ("-", "", "N/A", "nan", "None", "NULL"):
        return None
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


@st.cache_data(show_spinner="Đang đọc và phân tích dữ liệu...")
def load_and_preprocess_excel(file_bytes: bytes) -> pd.DataFrame:
    """Đọc và chuẩn hóa dữ liệu đầu vào theo đúng đặc tả công thức."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    sheet = find_data_sheet(xls)
    if sheet is None:
        raise ValueError("Không tìm thấy sheet 'Data import' hoặc sheet có cột SYM_RUN_DATE.")

    # Ép MA_DV đọc dạng chuỗi ngay từ đầu — tránh pandas tự suy luận sang số nguyên làm mất số 0
    # ở đầu mã (vd '012' -> 12) đối với các mã ĐVKD dạng số.
    df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl", dtype={"MA_DV": str})

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Thiếu các cột bắt buộc: {', '.join(missing)}")

    # Chuẩn hóa LPT: giữ NaN cho RAW, fill 0 cho tính toán thông thường
    df["HE_SO_LPT_RAW"] = df["Hệ số LPT thực nhận"].apply(parse_lpt_value)
    df["HE_SO_LPT"] = df["HE_SO_LPT_RAW"].fillna(0.0)

    # Trích xuất thời gian
    df["SYM_RUN_DATE"] = pd.to_datetime(df["SYM_RUN_DATE"], errors="coerce")
    df["YEAR_MONTH"] = df["SYM_RUN_DATE"].dt.to_period("M")
    df["THANG"] = df["SYM_RUN_DATE"].dt.month
    df["NAM"] = df["SYM_RUN_DATE"].dt.year
    df["THANG_STR"] = df.apply(
        lambda r: f"T{r['THANG']:02d}/{r['NAM']}" if pd.notna(r["THANG"]) and pd.notna(r["NAM"]) else "(blank)",
        axis=1,
    )

    # Chuẩn hóa tên KV
    df["TEN_KV"] = df["TEN_KV"].astype(str).str.strip()

    # Chuẩn hóa Mã ĐVKD: giữ nguyên dạng chuỗi, loại đuôi ".0" nếu file gốc lưu dạng số thực,
    # và đệm lại đủ 3 chữ số cho các mã thuần số (vd '12' -> '012') — khớp đúng định dạng data_import.
    df["MA_DV"] = df["MA_DV"].fillna("").astype(str).str.strip()
    df["MA_DV"] = df["MA_DV"].str.replace(r"\.0$", "", regex=True)
    df["MA_DV"] = df["MA_DV"].apply(lambda v: v.zfill(3) if v.isdigit() else v)

    # Chuẩn hóa họ tên nhân sự (tầng lọc thứ 3: Khu vực -> ĐVKD -> Nhân sự)
    df["TEN_CV"] = df["TEN_CV"].fillna("").astype(str).str.strip()

    # Chuẩn hóa số lượng
    for col in ["SL_KHDN_MOI_N", "SL_KHDN_TANG_RONG", "SL_KH_TD_TANG_RONG"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Chuẩn hóa 16 cột số liệu cấp ĐVKD cho Khối G/H/I (xem cảnh báo dedup ở UNIT_LEVEL_NUMERIC_COLUMNS)
    for col in UNIT_LEVEL_NUMERIC_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


# ──────────────────────────────────────────────────────────────────────────────
# XỬ LÝ TIỀN DỮ LIỆU SHEET "DATA IMPORT SỐ ĐVKD" (Dashboard số liệu toàn hàng)
# ──────────────────────────────────────────────────────────────────────────────

# Cấu trúc cột cố định theo đúng đặc tả (bảng pivot 2 dòng header, không dò tên cột động).
# (v4.15) File nguồn đổi cấu trúc: từ 1 block "Toàn hàng" duy nhất (67 dòng) sang 4 block lặp lại
# CÙNG bộ chỉ tiêu, phân biệt bởi cột D = TÊN ĐVKD (vd "Toàn hàng", "CN Phú Yên", "CN TTKD",
# "CN Hoàn Kiếm" — mỗi block ~63-64 dòng, tổng 257 dòng). Cột nhận diện chỉ tiêu bên trong mỗi
# block dịch sang phải đúng 1 cột so với bản cũ (D,E,F,G -> E,F,G,H — tận dụng đúng cột H vốn là
# 1 cột trống/gap giữa G và cột "baseline 31/12" ở I trong file cũ, nên các cột SỐ LIỆU từ J trở
# đi KHÔNG đổi vị trí). Cột A/B/C không có ý nghĩa phân loại (rỗng hoặc lặp lại tên ĐVKD ở block
# đầu), không dùng tới trong code.
# E=STT, F=STT phụ, G=Tên chỉ tiêu, H=ĐVT; I=baseline 31/12 (không dùng); J-U=Thực hiện Tháng
# (12 tháng theo vị trí, GIỮ NGUYÊN vị trí so với file cũ); V-AG=Kế hoạch Tháng (T1-T11 + Năm 2026
# ở cuối, KHÔNG có T12); AT-BD=% hoàn thành (T2-T11 + Năm 2026 ở cuối, KHÔNG có T1 và T12).
DVKD_COL_DVKD_TEN = 3  # D — MỚI (v4.15): tên ĐVKD, key phân biệt block
DVKD_COL_STT = 4
DVKD_COL_STT_PHU = 5
DVKD_COL_TEN_CHI_TIEU = 6
DVKD_COL_DVT = 7
DVKD_COL_CHU_KY = 8  # I — MỚI (v4.21): "Hàng ngày"/"Hàng tháng", nằm ngay sau ĐVT
DVKD_COL_THUC_HIEN = list(range(9, 21))  # J-U — không đổi so với file cũ
DVKD_COL_KE_HOACH = list(range(21, 33))  # V-AG — không đổi so với file cũ
DVKD_COL_PCT_HT = list(range(45, 56))  # AT-BD — không đổi so với file cũ
DVKD_COL_TODO_THANG = 64  # BM — MỚI (v4.22): "Todo tháng" (So với KH Tháng)
DVKD_COL_PCT_THANG = 65  # BN — MỚI (v4.22): "% hoàn thành T{tháng hiện tại}/2026"
DVKD_COL_TODO_NAM = 66  # BO — MỚI (v4.22): "Todo năm" (So với KH Năm)
DVKD_COL_PCT_NAM_2 = 67  # BP — MỚI (v4.22): "% hoàn thành/năm 2026"

DVKD_SECTION_TITLES = {"A. Thông tin chung", "A. Chỉ tiêu Chính", "B. Chỉ tiêu Phụ"}
DVKD_DEFAULT_DON_VI = "Toàn hàng"


def find_data_sheet_dvkd(xls: pd.ExcelFile) -> Optional[str]:
    """Tìm sheet 'Data import số ĐVKD' (bảng pivot 2 dòng header, số liệu toàn hàng)."""

    def _norm(s: str) -> str:
        # "Đ/đ" KHÔNG có phân rã NFD (không giống các nguyên âm có dấu khác) nên phải thay riêng.
        s = s.replace("Đ", "D").replace("đ", "d")
        s = unicodedata.normalize("NFD", s)
        s = "".join(c for c in s if unicodedata.category(c) != "Mn")
        return s.strip().lower()

    for name in xls.sheet_names:
        if _norm(name) == "data import so dvkd":
            return name
    for name in xls.sheet_names:
        norm = _norm(name)
        if "dvkd" in norm and "data import" in norm:
            return name
    return None


def _dvkd_parse_number(val) -> Optional[float]:
    """Chuẩn hóa 1 ô số liệu trong bảng pivot ĐVKD: hỗ trợ số âm dạng ngoặc đơn '(1,234)',
    dấu phẩy phân cách nghìn, hậu tố '%' (quy về dạng phân số), và các giá trị rỗng/lỗi."""
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s in ("-", "", "N/A", "nan", "None", "NULL", "#N/A", "#DIV/0!"):
        return None
    is_negative = s.startswith("(") and s.endswith(")")
    if is_negative:
        s = s[1:-1].strip()
    is_pct = s.endswith("%")
    s = s.rstrip("%").strip().replace(",", "")
    try:
        v = float(s)
    except ValueError:
        return None
    if is_pct:
        v = v / 100
    return -v if is_negative else v


def _assign_chart_type(label: str, dvt: str, section: str) -> str:
    """
    Gán loại biểu đồ mặc định cho 1 chỉ tiêu của trang 'Dashboard số liệu toàn hàng' (v4.8, phân
    loại chi tiết hơn ở v4.9) — dựa trên ĐVT/section/từ khóa trong tên chỉ tiêu (không có metadata
    chart_type tường minh trong file Excel nguồn, nên suy luận theo bản chất số liệu), theo thứ tự
    ưu tiên: ĐVT `%` -> đường (tỷ lệ liên tục); nhóm "A. Thông tin chung" -> cột đứng (đơn giản, ít
    chỉ tiêu); chỉ tiêu tăng/giảm hoặc nợ xấu/nợ nhóm/bán nợ (có thể âm/dương) -> cột ngang (đọc rõ
    dấu hơn cột đứng); chỉ tiêu ngoại tệ/doanh số/thanh toán quốc tế (tên thường dài) -> cột ngang;
    chỉ tiêu số lượng KH/SP/TK -> vùng tô (thể hiện quy mô tích lũy); chỉ tiêu huy động/cho vay/TOI/
    thu phí chính (có cặp Thực hiện+Kế hoạch cùng quy mô lớn) -> Combo Chart (giữ nguyên hành vi
    trước v4.8); còn lại (mặc định) -> đường.
    """
    if dvt == "%":
        return "line"
    if section == "A. Thông tin chung":
        return "bar"
    if any(kw in label for kw in ["Tăng", "Giảm", "tăng", "giảm", "NQH", "Nợ xấu", "Nợ nhóm", "Bán nợ"]):
        return "hbar"
    if any(kw in label for kw in ["MBNT", "TTQT", "Bảo hiểm", "Doanh số", "USD", "Triệu"]):
        return "hbar"
    if any(kw in label for kw in ["Số lượng KH", "SLKH", "SL KH", "SL SP", "SL TK", "KH active", "KH Payroll"]):
        return "area"
    if any(kw in label for kw in ["Huy động", "Cho vay", "Tổng cho", "TOI", "thu phí", "Casa"]):
        return "combo"
    return "line"


@st.cache_data(show_spinner="Đang đọc và phân tích dữ liệu số liệu ĐVKD...")
def load_and_preprocess_dvkd_excel(file_bytes: bytes) -> list[dict]:
    """
    Đọc và chuẩn hóa sheet 'Data import số ĐVKD' — bảng pivot 2 dòng header, GỒM NHIỀU BLOCK lặp
    lại cùng bộ chỉ tiêu, mỗi block ứng với 1 Đơn vị kinh doanh (v4.15 — trước đó chỉ có 1 block
    "Toàn hàng" duy nhất). Mỗi chỉ tiêu trả về 1 dict:
    {row_idx, dvkd, section, label, dvt, stt_display, chu_ky, chart_type, thuc_hien: {1..12: v},
    ke_hoach: {1..11: v}, ke_hoach_nam, pct_ht: {2..11: v}, pct_ht_nam, todo_thang, pct_ht_thang,
    todo_nam, pct_ht_nam2} — tháng thiếu dữ liệu không có key (UI hiển thị '—'). `dvkd` (v4.15, mới)
    = tên Đơn vị kinh doanh của block chứa dòng đó (vd "Toàn hàng", "CN Phú Yên"...), đọc từ cột D —
    `row_idx` vẫn là khóa DUY NHẤT toàn cục (không trùng giữa các block dù cùng `label`). `chart_type`
    (v4.8, xem `_assign_chart_type()`) quyết định loại biểu đồ. `stt_display` (v4.21, mới) = đúng
    chuỗi STT/STT phụ dùng làm tiền tố `label` (KHÔNG kèm dấu chấm) — dùng riêng cho cột "STT" độc
    lập trong bảng thống kê ĐVKD, tách biệt khỏi `label` (vẫn giữ dạng ghép "STT. Tên chỉ tiêu" như
    cũ cho các chỗ dùng khác). `chu_ky` (v4.21, mới) = "Hàng ngày"/"Hàng tháng"..., đọc từ cột I
    (`DVKD_COL_CHU_KY`, ngay sau ĐVT). `todo_thang`/`pct_ht_thang`/`todo_nam`/`pct_ht_nam2` (v4.22,
    mới) = 4 GIÁ TRỊ ĐƠN (không phải dict theo tháng như `thuc_hien`/`ke_hoach`/`pct_ht`, vì file gốc
    chỉ có ĐÚNG 1 cột mỗi loại, luôn ứng với "tháng hiện tại"/"năm hiện tại" của kỳ báo cáo, không
    lặp lại theo từng tháng) — đọc từ 4 cột cuối cùng `DVKD_COL_TODO_THANG/PCT_THANG/TODO_NAM/PCT_NAM_2`.
    """
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    sheet = find_data_sheet_dvkd(xls)
    if sheet is None:
        raise ValueError("Không tìm thấy sheet 'Data import số ĐVKD'.")

    raw = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl", header=None)

    items = []
    current_section = ""
    # (v4.15) Cột D (tên ĐVKD) thường chỉ có giá trị ở 1 vài dòng đầu mỗi block (merged cell —
    # Excel/pandas chỉ đọc được giá trị ở dòng đầu tiên của vùng merge, các dòng còn lại trả về
    # NaN) — dùng forward-fill thủ công: giữ nguyên `current_dvkd` cho tới khi gặp 1 giá trị MỚI,
    # KHÁC giá trị đang giữ, ở cột D.
    current_dvkd = DVKD_DEFAULT_DON_VI
    for row_idx in range(2, len(raw)):
        row = raw.iloc[row_idx]

        dvkd_val = row[DVKD_COL_DVKD_TEN]
        if pd.notna(dvkd_val):
            dvkd_str = str(dvkd_val).strip()
            if dvkd_str and dvkd_str != current_dvkd:
                current_dvkd = dvkd_str

        ten_chi_tieu = row[DVKD_COL_TEN_CHI_TIEU]
        if pd.isna(ten_chi_tieu):
            continue
        ten_chi_tieu = str(ten_chi_tieu).strip()
        if not ten_chi_tieu:
            continue

        if ten_chi_tieu in DVKD_SECTION_TITLES:
            current_section = ten_chi_tieu
            continue

        def _fmt_stt(v) -> str:
            # STT đọc qua pandas có thể bị suy luận thành float (vd 1 -> "1.0") khi cột có
            # lẫn ô rỗng/NaN — bỏ đuôi ".0" cho số nguyên, giữ nguyên STT phụ dạng "1.1".
            s = str(v).strip()
            if s.endswith(".0") and s[:-2].lstrip("-").isdigit():
                return s[:-2]
            return s

        stt = row[DVKD_COL_STT]
        stt_phu = row[DVKD_COL_STT_PHU]
        if pd.notna(stt) and _fmt_stt(stt):
            prefix = _fmt_stt(stt)
        elif pd.notna(stt_phu) and _fmt_stt(stt_phu):
            prefix = _fmt_stt(stt_phu)
        else:
            prefix = ""
        label = f"{prefix}. {ten_chi_tieu}" if prefix else ten_chi_tieu

        dvt_val = row[DVKD_COL_DVT]
        dvt = str(dvt_val).strip() if pd.notna(dvt_val) else ""

        chu_ky_val = row[DVKD_COL_CHU_KY]
        chu_ky = str(chu_ky_val).strip() if pd.notna(chu_ky_val) else ""

        thuc_hien = {}
        for pos, col in enumerate(DVKD_COL_THUC_HIEN):
            v = _dvkd_parse_number(row[col])
            if v is not None:
                thuc_hien[pos + 1] = v

        ke_hoach = {}
        ke_hoach_nam = None
        for pos, col in enumerate(DVKD_COL_KE_HOACH):
            v = _dvkd_parse_number(row[col])
            if pos < 11:
                if v is not None:
                    ke_hoach[pos + 1] = v
            else:
                ke_hoach_nam = v

        pct_ht = {}
        pct_ht_nam = None
        for pos, col in enumerate(DVKD_COL_PCT_HT):
            v = _dvkd_parse_number(row[col])
            if pos < 10:
                if v is not None:
                    pct_ht[pos + 2] = v
            else:
                pct_ht_nam = v

        # (v4.22) 4 cột cuối cùng — So với Kế hoạch Tháng/Năm (không theo dạng {thang: v} như
        # thuc_hien/ke_hoach/pct_ht vì đây là 4 giá trị ĐƠN, tự thân đã ứng với "tháng hiện tại"/
        # "năm hiện tại" của file, không lặp lại theo từng tháng trong sel_months).
        todo_thang = _dvkd_parse_number(row[DVKD_COL_TODO_THANG])
        pct_ht_thang = _dvkd_parse_number(row[DVKD_COL_PCT_THANG])
        todo_nam = _dvkd_parse_number(row[DVKD_COL_TODO_NAM])
        pct_ht_nam2 = _dvkd_parse_number(row[DVKD_COL_PCT_NAM_2])

        items.append(
            {
                "row_idx": row_idx,
                "dvkd": current_dvkd,
                "section": current_section,
                "label": label,
                "dvt": dvt,
                "stt_display": prefix,
                "chu_ky": chu_ky,
                "chart_type": _assign_chart_type(label, dvt, current_section),
                "thuc_hien": thuc_hien,
                "ke_hoach": ke_hoach,
                "ke_hoach_nam": ke_hoach_nam,
                "pct_ht": pct_ht,
                "pct_ht_nam": pct_ht_nam,
                "todo_thang": todo_thang,
                "pct_ht_thang": pct_ht_thang,
                "todo_nam": todo_nam,
                "pct_ht_nam2": pct_ht_nam2,
            }
        )

    return items


# ──────────────────────────────────────────────────────────────────────────────
# TÍNH TOÁN KỲ N THÁNG GẦN NHẤT & GROUP KHU VỰC
# ──────────────────────────────────────────────────────────────────────────────


def get_monthly_periods(df: pd.DataFrame, max_n: int = 12) -> list[dict]:
    """
    Danh sách TỪNG THÁNG THỰC TẾ riêng lẻ (không cộng dồn/lũy kế), dùng cho 9 Khối chỉ tiêu
    (Tab Báo cáo Khối chỉ tiêu). Mỗi phần tử {"label", "yms"} — "label" là tên tháng hiển thị
    (vd "T03/2026"), "yms" là danh sách 1 YEAR_MONTH duy nhất ứng với tháng đó.
    """
    unique_yms = sorted(df["YEAR_MONTH"].dropna().unique())[-max_n:]
    return [{"label": f"T{ym.month:02d}/{ym.year}", "yms": [ym]} for ym in unique_yms]


def _region_groups(df: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    """
    Nhóm theo TEN_KV (không dùng MA_KV).
    Đảm bảo 6 khu vực chuẩn được sắp xếp đúng thứ tự + 1 dòng Tổng (cộng dồn toàn bộ).
    """
    groups = []
    avail = df["TEN_KV"].unique().tolist()
    used = set()

    for kv in STANDARD_KV_ORDER:
        matched = None
        for a in avail:
            if a not in used:
                if a == kv or kv.lower().startswith(a.lower()) or a.lower().startswith(kv.lower()) or kv[:9].lower() == a[:9].lower():
                    matched = a
                    break
        if matched:
            groups.append((matched, df[df["TEN_KV"] == matched]))
            used.add(matched)
        elif kv in avail:
            groups.append((kv, df[df["TEN_KV"] == kv]))
            used.add(kv)

    # Các khu vực khác có trong data nếu có
    for a in avail:
        if a not in used:
            groups.append((a, df[df["TEN_KV"] == a]))
            used.add(a)

    # Dòng Tổng: cộng dồn toàn bộ data (không lọc theo khu vực)
    groups.append(("Tổng", df))
    return groups


# ──────────────────────────────────────────────────────────────────────────────
# TÍNH TOÁN 6 KHỐI CHỈ TIÊU (Đặc tả công thức mục 3)
# ──────────────────────────────────────────────────────────────────────────────


def calc_block1(df: pd.DataFrame, periods: Optional[list] = None) -> pd.DataFrame:
    """BẢNG A — Tỷ lệ đạt KPIs (theo từng tháng thực tế)."""
    if periods is None:
        periods = get_monthly_periods(df)
    groups = _region_groups(df)
    rows = []
    for kv_name, subset in groups:
        row = {"Khu vực": kv_name}
        for p in periods:
            sm = subset[subset["YEAR_MONTH"].isin(p["yms"])]
            tot = len(sm)
            dat = len(sm[sm["KET_QUA_KPIS"] == "DAT_KPIS"])
            pct = (dat / tot * 100) if tot > 0 else 0.0
            col_key = p["label"]
            row[("Tổng NS đánh giá KPIs", col_key)] = tot
            row[("SL nhân sự đạt KPIs", col_key)] = dat
            row[("% nhân sự đạt KPIs", col_key)] = round(pct, 1)
        rows.append(row)

    res = pd.DataFrame(rows).set_index("Khu vực")
    res.columns = pd.MultiIndex.from_tuples(res.columns)
    return res


def calc_block2(df: pd.DataFrame, periods: Optional[list] = None) -> pd.DataFrame:
    """BẢNG B — Tỷ lệ thỏa LPT (theo từng tháng thực tế)."""
    if periods is None:
        periods = get_monthly_periods(df)
    groups = _region_groups(df)
    rows = []
    for kv_name, subset in groups:
        row = {"Khu vực": kv_name}
        for p in periods:
            sm = subset[subset["YEAR_MONTH"].isin(p["yms"])]
            tot = len(sm)
            thoa = len(sm[sm["KET_QUA_LPT"] == "THOA_LPT"])
            pct = (thoa / tot * 100) if tot > 0 else 0.0
            col_key = p["label"]
            row[("Tổng NS đánh giá KPIs", col_key)] = tot
            row[("SL nhân sự thỏa LPT", col_key)] = thoa
            row[("% nhân sự thỏa LPT", col_key)] = round(pct, 1)
        rows.append(row)

    res = pd.DataFrame(rows).set_index("Khu vực")
    res.columns = pd.MultiIndex.from_tuples(res.columns)
    return res


def calc_block3(df: pd.DataFrame, periods: Optional[list] = None) -> pd.DataFrame:
    """BẢNG C — Phân bổ 6 dải Hệ số LPT thực nhận (NULL không tính vào dải nào; theo từng tháng thực tế)."""
    if periods is None:
        periods = get_monthly_periods(df)
    groups = _region_groups(df)
    rows = []
    for kv_name, subset in groups:
        row = {"Khu vực": kv_name}
        for p in periods:
            sm = subset[subset["YEAR_MONTH"].isin(p["yms"])]
            vals = sm["HE_SO_LPT_RAW"].dropna()
            col_key = p["label"]
            # Mẫu số của Khối C giống Khối A/B — Tổng NS đánh giá KPIs (tot = len(sm), KHÔNG phải
            # len(vals)): người có Hệ số LPT NULL vẫn nằm trong mẫu số, chỉ không tính vào dải nào.
            tot = len(sm)
            row[("Tổng NS đánh giá KPIs", col_key)] = tot
            for label, cond_fn in LPT_BINS:
                row[(label, col_key)] = int(cond_fn(vals).sum())
        rows.append(row)

    res = pd.DataFrame(rows).set_index("Khu vực")
    res.columns = pd.MultiIndex.from_tuples(res.columns)
    return res


def _calc_block_kh(df: pd.DataFrame, sum_col: str, metric_name: str, periods: Optional[list] = None) -> pd.DataFrame:
    """Helper cho Bảng D, E, F: Mẫu số COUNT(DISTINCT MA_DV), Tử số SUM (theo từng tháng thực tế)."""
    if periods is None:
        periods = get_monthly_periods(df)
    groups = _region_groups(df)
    rows = []
    for kv_name, subset in groups:
        row = {"Khu vực": kv_name}
        for p in periods:
            sm = subset[subset["YEAR_MONTH"].isin(p["yms"])]
            n_dv = sm["MA_DV"].nunique()
            total = int(sm[sum_col].sum())
            avg = (total / n_dv) if n_dv > 0 else 0.0
            col_key = p["label"]
            row[("Tổng ĐVKD", col_key)] = n_dv
            row[(f"Tổng {metric_name}", col_key)] = total
            row[(f"{metric_name} / ĐVKD", col_key)] = round(avg, 1)
        rows.append(row)

    res = pd.DataFrame(rows).set_index("Khu vực")
    res.columns = pd.MultiIndex.from_tuples(res.columns)
    return res


def calc_block4(df: pd.DataFrame, periods: Optional[list] = None) -> pd.DataFrame:
    """BẢNG D — % phát triển KHDN mới / ĐVKD."""
    return _calc_block_kh(df, "SL_KHDN_MOI_N", "KHDN mới", periods)


def calc_block5(df: pd.DataFrame, periods: Optional[list] = None) -> pd.DataFrame:
    """BẢNG E — % phát triển KHTD tăng ròng / ĐVKD (SUM SL_KH_TD_TANG_RONG)."""
    return _calc_block_kh(df, "SL_KH_TD_TANG_RONG", "KHTD tăng ròng", periods)


def calc_block6(df: pd.DataFrame, periods: Optional[list] = None) -> pd.DataFrame:
    """BẢNG F — % phát triển KHDN tăng ròng / ĐVKD (SUM SL_KHDN_TANG_RONG)."""
    return _calc_block_kh(df, "SL_KHDN_TANG_RONG", "KHDN tăng ròng", periods)


# ──────────────────────────────────────────────────────────────────────────────
# TÍNH TOÁN 3 KHỐI CHỈ TIÊU MỚI G/H/I (Dư nợ / Huy động / TOI) — CẤP ĐƠN VỊ (ĐVKD)
# ──────────────────────────────────────────────────────────────────────────────
# ⚠️ CẢNH BÁO DỮ LIỆU QUAN TRỌNG: 16 cột nguồn của Khối G/H/I (UNIT_LEVEL_NUMERIC_COLUMNS) là
# giá trị CẤP ĐVKD nhưng bị LẶP LẠI y hệt trên MỌI dòng nhân sự cùng 1 (MA_DV, SYM_RUN_DATE).
# TUYỆT ĐỐI KHÔNG SUM trực tiếp theo dòng như các cột cấp-nhân-sự của Khối A-F (sẽ nhân sai giá
# trị lên gấp N lần, N = số nhân sự của ĐVKD đó trong tháng). LUÔN dùng agg_unit_level_metric()
# (dedup theo MA_DV+SYM_RUN_DATE trước khi sum) cho 16 cột này — không tự ý đổi lại thành .sum().

TOI_SOURCE_COLS = {
    "TOI huy động": "DU_CHI_HUY_DONG",
    "TOI cho vay": "DU_THU_CHO_VAY",
    "TOI thu thanh toán": "THU_TU_THANH_TOAN",
    "TOI QLTM": "THU_QUAN_LY_TIEN_MAT",
    "TOI ngân quỹ": "THU_NGAN_QUY",
    "TOI ngân hàng điện tử": "THU_NGAN_HANG_DIEN_TU",
    "TOI TTTM": "THU_TAI_TRO_THUONG_MAI",
    "TOI bảo lãnh": "THU_BAO_LANH",
    "TOI tín dụng": "THU_TIN_DUNG",
    "TOI MBNT": "THU_MUA_BAN_NGOAI_TE",
}


def agg_unit_level_metric(df: pd.DataFrame, col: str) -> float:
    """
    Cộng 1 cột giá trị CẤP ĐVKD — BẮT BUỘC dedup theo (MA_DV, SYM_RUN_DATE) TRƯỚC khi sum, vì
    giá trị bị lặp lại y hệt trên mọi dòng nhân sự cùng ĐVKD/tháng (xem cảnh báo phía trên).
    """
    dedup = df.drop_duplicates(subset=["MA_DV", "SYM_RUN_DATE"])
    return float(pd.to_numeric(dedup[col], errors="coerce").fillna(0).sum())


def _calc_block_unit_level(df: pd.DataFrame, periods: list[dict], metric_cols: dict[str, str]) -> pd.DataFrame:
    """
    Khung tính dùng chung cho Khối G/H/I: HÀNG = Khu vực + Tổng (như A-F), CỘT = MultiIndex
    (tên chỉ số, tháng). Mỗi chỉ số trong metric_cols PHẢI qua agg_unit_level_metric (dedup) —
    không SUM trực tiếp (xem cảnh báo ở đầu Mục này).
    """
    groups = _region_groups(df)
    rows = []
    for kv_name, subset in groups:
        row = {"Khu vực": kv_name}
        for p in periods:
            sm = subset[subset["YEAR_MONTH"].isin(p["yms"])]
            for metric_name, col in metric_cols.items():
                row[(metric_name, p["label"])] = round(agg_unit_level_metric(sm, col), 2)
        rows.append(row)

    res = pd.DataFrame(rows).set_index("Khu vực")
    res.columns = pd.MultiIndex.from_tuples(res.columns)
    return res


def calc_block7(df: pd.DataFrame, periods: Optional[list] = None) -> pd.DataFrame:
    """KHỐI G — Chi tiết dư nợ: Tổng cho vay tháng N/N-1, Tăng ròng cho vay (cấp ĐVKD, đã dedup)."""
    if periods is None:
        periods = get_monthly_periods(df)
    return _calc_block_unit_level(df, periods, {
        "Tổng cho vay tháng N": "TONG_CV_THANG_N",
        "Tổng cho vay tháng N-1": "TONG_CV_THANG_N_1",
        "Tăng ròng cho vay": "TANG_RONG_CV",
    })


def calc_block8(df: pd.DataFrame, periods: Optional[list] = None) -> pd.DataFrame:
    """KHỐI H — Chi tiết huy động: Tổng huy động tháng N/N-1, Tăng ròng huy động (cấp ĐVKD, đã dedup)."""
    if periods is None:
        periods = get_monthly_periods(df)
    return _calc_block_unit_level(df, periods, {
        "Tổng huy động tháng N": "TONG_HD_THANG_N",
        "Tổng huy động tháng N-1": "TONG_HD_THANG_N_1",
        "Tăng ròng huy động": "TANG_RONG_HD",
    })


def calc_block9(df: pd.DataFrame, periods: Optional[list] = None) -> pd.DataFrame:
    """KHỐI I — Chi tiết TOI: 10 nguồn thu (cấp ĐVKD, đã dedup) + Tổng TOI = SUM 10 dòng trên."""
    if periods is None:
        periods = get_monthly_periods(df)
    res = _calc_block_unit_level(df, periods, TOI_SOURCE_COLS)
    for p in periods:
        res[("Tổng TOI", p["label"])] = sum(res[(name, p["label"])] for name in TOI_SOURCE_COLS)
    return res


# ──────────────────────────────────────────────────────────────────────────────
# TỔNG HỢP CHỈ SỐ METRICS TỔNG QUAN TỪ DỮ LIỆU THẬT
# ──────────────────────────────────────────────────────────────────────────────


def apply_thang_filter(
    df: pd.DataFrame,
    thang_filter,
    unique_yms: Optional[list] = None,
) -> tuple[pd.DataFrame, pd.DataFrame, str]:
    """
    Lọc dữ liệu theo bộ lọc SỐ THÁNG đang chọn — `thang_filter` là **list** (đa lựa chọn, v4.12;
    vẫn nhận `str` đơn lẻ để tương thích ngược). Trả về (df_thang, prev_df, prev_date_str) dùng
    chung cho KPI, biểu đồ và drill-down.
    - 0 tháng (rỗng): không lọc, không tính Delta (như cũ khi `thang_filter=""`).
    - Đúng 1 tháng: lọc `== tháng đó` + tính Delta so tháng liền trước (hành vi gốc, không đổi).
    - ≥2 tháng: lọc `.isin(...)` các tháng đã chọn, KHÔNG tính Delta (nhiều tháng không liên tục thì
      "tháng trước" không còn ý nghĩa rõ ràng — xem quyết định người dùng khi thêm đa lựa chọn).
    """
    if unique_yms is None:
        unique_yms = sorted(df["YEAR_MONTH"].dropna().unique())

    months = thang_filter if isinstance(thang_filter, list) else ([thang_filter] if thang_filter else [])
    months = [m for m in months if m and m != "(blank)"]

    df_thang = df
    prev_df = pd.DataFrame()
    prev_date_str = "Toàn bộ"

    if len(months) == 1:
        df_thang = df[df["THANG_STR"] == months[0]]
        if not df_thang.empty:
            curr_ym = df_thang["YEAR_MONTH"].iloc[0]
            if curr_ym in unique_yms:
                idx = unique_yms.index(curr_ym)
                if idx > 0:
                    prev_ym = unique_yms[idx - 1]
                    prev_df = df[df["YEAR_MONTH"] == prev_ym]
                    prev_date_str = f"T{prev_ym.month:02d}/{prev_ym.year}"
    elif len(months) > 1:
        df_thang = df[df["THANG_STR"].isin(months)]
    return df_thang, prev_df, prev_date_str


def kv_full_to_label(kv_full: str) -> str:
    """Chuyển tên đầy đủ 'Khu vực N (KV Xyz)' -> nhãn hiển thị ngắn 'KV N (Xyz)' dùng trên mọi biểu đồ."""
    return (
        kv_full.replace("Khu vực ", "KV ")
        .replace("KV Hà Nội", "Hà Nội")
        .replace("KV Miền Bắc", "Miền Bắc")
        .replace("KV Miền Trung", "Miền Trung")
        .replace("KV TPHCM", "TPHCM")
        .replace("KV Đồng Nai", "Đồng Nai")
        .replace("KV Tây Nam Bộ", "Tây Nam Bộ")
    )


def apply_kv_filter(df: pd.DataFrame, kv_filter) -> pd.DataFrame:
    """
    Lọc dữ liệu theo bộ lọc KHU VỰC đang chọn — `kv_filter` là **list** (đa lựa chọn, v4.12; vẫn
    nhận `str` đơn lẻ để tương thích ngược). Danh sách rỗng = Toàn Hàng (không lọc). Nhiều Khu vực
    được chọn -> OR các điều kiện khớp của TỪNG Khu vực lại (kết quả là HỢP/union dữ liệu của tất
    cả Khu vực đã chọn), không phải AND.
    """
    kv_list = kv_filter if isinstance(kv_filter, list) else ([kv_filter] if kv_filter else [])
    kv_list = [k for k in kv_list if k and k != "(blank)"]
    if not kv_list:
        return df
    mask = pd.Series(False, index=df.index)
    for kv in kv_list:
        target = kv.replace("KV ", "Khu vực ").replace("KV", "Khu vực ")
        mask = (
            mask
            | df["TEN_KV"].str.contains(target, case=False, na=False)
            | (df["TEN_KV"] == kv)
            | (df["MA_KV"].astype(str) == kv)
        )
    return df[mask]


def apply_ma_dv_filter(df: pd.DataFrame, ma_dv_filter: Optional[list] = None) -> pd.DataFrame:
    """Lọc dữ liệu theo danh sách MA_DV đang chọn (đa lựa chọn trong Slicer KHU VỰC sau khi chọn 1 KV)."""
    if not ma_dv_filter:
        return df
    return df[df["MA_DV"].isin(ma_dv_filter)]


def apply_nhan_su_filter(df: pd.DataFrame, ns_filter: Optional[list] = None) -> pd.DataFrame:
    """Lọc dữ liệu theo danh sách TEN_CV đang chọn — tầng lọc thứ 3 (Khu vực -> ĐVKD -> Nhân sự)."""
    if not ns_filter:
        return df
    return df[df["TEN_CV"].isin(ns_filter)]


def get_kv_units(df: pd.DataFrame, kv_filter) -> list[tuple[str, str]]:
    """Danh sách (MA_DV, TEN_DV) duy nhất trong (các) Khu vực đang chọn (list, v4.12), sắp xếp theo MA_DV."""
    d = apply_kv_filter(df, kv_filter)
    if d.empty:
        return []
    pairs = d[["MA_DV", "TEN_DV"]].drop_duplicates().sort_values("MA_DV")
    return list(pairs.itertuples(index=False, name=None))


def get_dv_staff(df: pd.DataFrame, kv_filter, ma_dv_filter: Optional[list] = None) -> list[str]:
    """Danh sách TEN_CV duy nhất, đã sort, trong phạm vi (các) Khu vực (list, v4.12) + ĐVKD đang chọn."""
    d = apply_ma_dv_filter(apply_kv_filter(df, kv_filter), ma_dv_filter)
    if d.empty:
        return []
    names = d["TEN_CV"].dropna().astype(str).str.strip()
    return sorted(n for n in names.unique() if n)


def compute_overview_metrics(
    df: pd.DataFrame,
    kv_filter=None,
    thang_filter=None,
    chi_tieu_filter: str = "",
    ma_dv_filter: Optional[list] = None,
    ns_filter: Optional[list] = None,
) -> dict:
    """
    Tính toán các chỉ số trên 6 thẻ KPI & các biểu đồ từ dữ liệu thật.
    - Donut & các Bar Chart theo Khu vực: LUÔN tính trên toàn bộ data theo kỳ tháng đang chọn (không bị lọc theo KV).
    - 6 Thẻ KPI Pill: tính theo kỳ tháng, khu vực VÀ đơn vị kinh doanh đang chọn (nếu có).
    - Delta: so sánh với kỳ (N-1)T hoặc tháng liền trước (không hardcode, in log ra console).
    `kv_filter`/`thang_filter` là **list** (đa lựa chọn, v4.12) — vẫn nhận `str`/`None` để tương thích ngược.
    """
    unique_yms = sorted(df["YEAR_MONTH"].dropna().unique())

    # 1. Lọc theo SỐ THÁNG (theo tháng cụ thể) trên TOÀN BỘ DATA (dùng cho Donut & các Bar Chart theo KV)
    df_thang, prev_df_thang, prev_date_str = apply_thang_filter(df, thang_filter, unique_yms)

    # 2. TÍNH TOÁN 3 BIỂU ĐỒ (LUÔN TÍNH TRÊN TOÀN BỘ DATA ĐÃ LỌC THEO THÁNG - KHÔNG LỌC THEO KV)
    # Donut Chart: Toàn hàng
    donut_dat = len(df_thang[df_thang["KET_QUA_KPIS"] == "DAT_KPIS"])
    donut_khong_dat = len(df_thang) - donut_dat

    # 2 Bar Charts: Luôn đủ 6 Khu vực chuẩn
    kv_charts = []
    for kv in STANDARD_KV_ORDER:
        matched_df = df_thang[df_thang["TEN_KV"] == kv]
        if matched_df.empty:
            for a in df_thang["TEN_KV"].unique():
                if kv[:9].lower() == a[:9].lower():
                    matched_df = df_thang[df_thang["TEN_KV"] == a]
                    break

        kv_tot = len(matched_df)
        kv_dat = len(matched_df[matched_df["KET_QUA_KPIS"] == "DAT_KPIS"]) if kv_tot > 0 else 0
        kv_pct_dat = (kv_dat / kv_tot * 100) if kv_tot > 0 else 0.0

        kv_thoa = len(matched_df[matched_df["KET_QUA_LPT"] == "THOA_LPT"]) if kv_tot > 0 else 0
        kv_pct_thoa = (kv_thoa / kv_tot * 100) if kv_tot > 0 else 0.0

        # Bảng D/E/F: mẫu số COUNT(DISTINCT MA_DV), tử số SUM, chỉ tiêu bình quân/ĐVKD
        n_dv = matched_df["MA_DV"].nunique() if kv_tot > 0 else 0
        khdn_moi_sum = int(matched_df["SL_KHDN_MOI_N"].sum()) if kv_tot > 0 else 0
        avg_khdn_moi = (khdn_moi_sum / n_dv) if n_dv > 0 else 0.0
        khtd_tang_sum = int(matched_df["SL_KH_TD_TANG_RONG"].sum()) if kv_tot > 0 else 0
        avg_khtd_tang = (khtd_tang_sum / n_dv) if n_dv > 0 else 0.0
        khdn_tang_sum = int(matched_df["SL_KHDN_TANG_RONG"].sum()) if kv_tot > 0 else 0
        avg_khdn_tang = (khdn_tang_sum / n_dv) if n_dv > 0 else 0.0

        # Khối G/H/I: giá trị CẤP ĐVKD — BẮT BUỘC dùng agg_unit_level_metric (dedup theo
        # MA_DV+SYM_RUN_DATE trước khi sum), KHÔNG dùng .sum() trực tiếp như các cột trên.
        tong_cv_n_sum = agg_unit_level_metric(matched_df, "TONG_CV_THANG_N") if kv_tot > 0 else 0.0
        tong_cv_n1_sum = agg_unit_level_metric(matched_df, "TONG_CV_THANG_N_1") if kv_tot > 0 else 0.0
        tang_rong_cv_sum = agg_unit_level_metric(matched_df, "TANG_RONG_CV") if kv_tot > 0 else 0.0
        tong_hd_n_sum = agg_unit_level_metric(matched_df, "TONG_HD_THANG_N") if kv_tot > 0 else 0.0
        tong_hd_n1_sum = agg_unit_level_metric(matched_df, "TONG_HD_THANG_N_1") if kv_tot > 0 else 0.0
        tang_rong_hd_sum = agg_unit_level_metric(matched_df, "TANG_RONG_HD") if kv_tot > 0 else 0.0
        tong_toi_sum = (
            sum(agg_unit_level_metric(matched_df, col) for col in TOI_SOURCE_COLS.values()) if kv_tot > 0 else 0.0
        )

        # Bảng C: % phân bổ 6 dải Hệ số LPT thực nhận (NULL gộp thành "Chưa xác định")
        lpt_vals = matched_df["HE_SO_LPT_RAW"].dropna() if kv_tot > 0 else pd.Series(dtype=float)
        lpt_bin_pct = {}
        for label, cond_fn in LPT_BINS:
            cnt = int(cond_fn(lpt_vals).sum()) if kv_tot > 0 else 0
            lpt_bin_pct[label] = round((cnt / kv_tot * 100) if kv_tot > 0 else 0.0, 1)
        null_cnt = kv_tot - len(lpt_vals)
        lpt_bin_pct["Chưa xác định"] = round((null_cnt / kv_tot * 100) if kv_tot > 0 else 0.0, 1)

        kv_label = kv_full_to_label(kv)

        kv_charts.append({
            "kv_full": kv,
            "kv_label": kv_label,
            "pct_dat": round(kv_pct_dat, 1),
            "pct_thoa": round(kv_pct_thoa, 1),
            "khdn_moi_sum": khdn_moi_sum,
            "avg_khdn_moi": round(avg_khdn_moi, 1),
            "khtd_tang_sum": khtd_tang_sum,
            "avg_khtd_tang": round(avg_khtd_tang, 1),
            "khdn_tang_sum": khdn_tang_sum,
            "avg_khdn_tang": round(avg_khdn_tang, 1),
            "lpt_bin_pct": lpt_bin_pct,
            "tong_cv_n_sum": round(tong_cv_n_sum, 2),
            "tong_cv_n1_sum": round(tong_cv_n1_sum, 2),
            "tang_rong_cv_sum": round(tang_rong_cv_sum, 2),
            "tong_hd_n_sum": round(tong_hd_n_sum, 2),
            "tong_hd_n1_sum": round(tong_hd_n1_sum, 2),
            "tang_rong_hd_sum": round(tang_rong_hd_sum, 2),
            "tong_toi_sum": round(tong_toi_sum, 2),
        })

    # 3. TÍNH TOÁN 6 THẺ KPI PILL (ÁP DỤNG KV FILTER + ĐVKD FILTER + NHÂN SỰ FILTER NẾU CÓ)
    curr_df = apply_nhan_su_filter(apply_ma_dv_filter(apply_kv_filter(df_thang, kv_filter), ma_dv_filter), ns_filter)
    prev_df = (
        apply_nhan_su_filter(apply_ma_dv_filter(apply_kv_filter(prev_df_thang, kv_filter), ma_dv_filter), ns_filter)
        if not prev_df_thang.empty
        else prev_df_thang.copy()
    )

    tot_ns = len(curr_df)
    dat_kpi = len(curr_df[curr_df["KET_QUA_KPIS"] == "DAT_KPIS"])
    pct_dat = (dat_kpi / tot_ns * 100) if tot_ns > 0 else 0.0
    kdat_kpi = tot_ns - dat_kpi
    pct_kdat = (kdat_kpi / tot_ns * 100) if tot_ns > 0 else 0.0

    thoa_lpt = len(curr_df[curr_df["KET_QUA_LPT"] == "THOA_LPT"])
    pct_thoa = (thoa_lpt / tot_ns * 100) if tot_ns > 0 else 0.0

    tot_dv = curr_df["MA_DV"].nunique() if tot_ns > 0 else 0
    tot_khdn_moi = int(curr_df["SL_KHDN_MOI_N"].sum()) if tot_ns > 0 else 0
    avg_khdn_moi = (tot_khdn_moi / tot_dv) if tot_dv > 0 else 0.0

    tot_khtd_tang = int(curr_df["SL_KH_TD_TANG_RONG"].sum()) if tot_ns > 0 else 0
    avg_khtd_tang = (tot_khtd_tang / tot_dv) if tot_dv > 0 else 0.0

    tot_khdn_tang = int(curr_df["SL_KHDN_TANG_RONG"].sum()) if tot_ns > 0 else 0
    avg_khdn_tang = (tot_khdn_tang / tot_dv) if tot_dv > 0 else 0.0

    # (v4.28) Khi chọn TỪ 2 THÁNG TRỞ LÊN, 3 thẻ "Tổng NS đánh giá KPIs"/"Đạt KPIs"/"Thỏa LPT" hiển
    # thị BÌNH QUÂN theo tháng (tổng cộng dồn các tháng đã chọn / số tháng đã chọn) thay vì tổng cộng
    # dồn thô — theo đúng yêu cầu người dùng (vd chọn T06+T07+T08 -> hiển thị = (NS_T06+NS_T07+NS_T08)/3).
    # Đúng 1 tháng (hoặc chưa chọn tháng nào — `n_months_sel == 0`): giữ NGUYÊN hành vi cũ (hiển thị
    # = tổng thô, không chia). `pct_dat`/`pct_kdat`/`pct_thoa` (đã tính Ở TRÊN, dùng tử/mẫu số THÔ)
    # KHÔNG cần đổi — tỷ lệ % không đổi dù chia cả tử và mẫu cho cùng 1 số tháng hay không.
    _months_sel_raw = thang_filter if isinstance(thang_filter, list) else ([thang_filter] if thang_filter else [])
    _months_sel = [m for m in _months_sel_raw if m and m != "(blank)"]
    n_months_sel = len(_months_sel)

    disp_ns = round(tot_ns / n_months_sel) if n_months_sel > 1 else tot_ns
    disp_dat_kpi = round(dat_kpi / n_months_sel) if n_months_sel > 1 else dat_kpi
    disp_kdat_kpi = round(kdat_kpi / n_months_sel) if n_months_sel > 1 else kdat_kpi
    disp_thoa_lpt = round(thoa_lpt / n_months_sel) if n_months_sel > 1 else thoa_lpt

    # 4. Tính Delta so với prev_df (so sánh kỳ trước)
    def _calc_delta(curr_val, prev_val):
        if prev_df.empty or prev_val is None:
            return "—"
        diff = curr_val - prev_val
        pct = (diff / abs(prev_val) * 100) if prev_val != 0 else 0.0
        sign = "▲" if diff >= 0 else "▼"
        fmt_diff = f"{abs(diff):,}".replace(",", ".")
        return f"{sign} {fmt_diff} ({pct:+.1f}%)"

    p_ns = len(prev_df) if not prev_df.empty else None
    p_dat = len(prev_df[prev_df["KET_QUA_KPIS"] == "DAT_KPIS"]) if not prev_df.empty else None
    p_thoa = len(prev_df[prev_df["KET_QUA_LPT"] == "THOA_LPT"]) if not prev_df.empty else None
    p_dv = prev_df["MA_DV"].nunique() if not prev_df.empty else None
    p_moi = int(prev_df["SL_KHDN_MOI_N"].sum()) if not prev_df.empty else None
    p_td = int(prev_df["SL_KH_TD_TANG_RONG"].sum()) if not prev_df.empty else None
    p_dntg = int(prev_df["SL_KHDN_TANG_RONG"].sum()) if not prev_df.empty else None

    delta_ns = _calc_delta(tot_ns, p_ns)
    delta_dat = _calc_delta(dat_kpi, p_dat)
    delta_thoa = _calc_delta(thoa_lpt, p_thoa)
    delta_dv = _calc_delta(tot_dv, p_dv)
    delta_moi = _calc_delta(tot_khdn_moi, p_moi)
    delta_td = _calc_delta(tot_khtd_tang, p_td)
    delta_dntg = _calc_delta(tot_khdn_tang, p_dntg)

    # In log chi tiết ra console để kiểm tra tử số / mẫu số
    try:
        _thang_log = ", ".join(thang_filter) if thang_filter else "All"
        _kv_log = ", ".join(kv_filter) if kv_filter else "All"
        print(f"[KPI LOG] THANG='{_thang_log}', KV='{_kv_log}'")
        print(f"  Curr: NS={tot_ns}, Đạt={dat_kpi}/{tot_ns} ({pct_dat:.1f}%), Thỏa={thoa_lpt}/{tot_ns} ({pct_thoa:.1f}%), ĐVKD={tot_dv}, KHDN_moi={tot_khdn_moi}, KHTD={tot_khtd_tang}")
        if not prev_df.empty:
            print(f"  Prev ({prev_date_str}): NS={p_ns}, Đạt={p_dat}, Thỏa={p_thoa}, ĐVKD={p_dv}, KHDN_moi={p_moi}, KHTD={p_td}")
            print(f"  Delta: NS={delta_ns}, Đạt={delta_dat}, Thỏa={delta_thoa}, ĐVKD={delta_dv}, KHDN_moi={delta_moi}, KHTD={delta_td}")
        else:
            print("  Prev: Không có (xem toàn bộ hoặc không đủ data so sánh) -> Delta = '—'")
    except Exception:
        pass

    max_date = curr_df["SYM_RUN_DATE"].max() if not curr_df.empty else df["SYM_RUN_DATE"].max()
    run_date_str = max_date.strftime("%d/%m/%Y") if pd.notna(max_date) else "31/07/2026"

    return {
        "run_date": run_date_str,
        "prev_date": prev_date_str,
        "total_personnel": disp_ns,
        "delta_personnel": delta_ns,
        "kpi_dat": disp_dat_kpi,
        "pct_kpi_dat": round(pct_dat, 1),
        "delta_kpi_dat": delta_dat,
        "kpi_khong_dat": disp_kdat_kpi,
        "pct_kpi_khong_dat": round(pct_kdat, 1),
        "thoa_lpt": disp_thoa_lpt,
        "pct_thoa_lpt": round(pct_thoa, 1),
        "delta_thoa_lpt": delta_thoa,
        "total_dvkd": tot_dv,
        "delta_dvkd": delta_dv,
        "total_khdn_moi": tot_khdn_moi,
        "avg_khdn_moi": round(avg_khdn_moi, 1),
        "delta_khdn_moi": delta_moi,
        "total_khtd_tang": tot_khtd_tang,
        "avg_khtd_tang": round(avg_khtd_tang, 1),
        "delta_khtd_tang": delta_td,
        "total_khdn_tang": tot_khdn_tang,
        "avg_khdn_tang": round(avg_khdn_tang, 1),
        "delta_khdn_tang": delta_dntg,
        "donut_kpi_dat": donut_dat,
        "donut_kpi_khong_dat": donut_khong_dat,
        "kv_charts": kv_charts,
    }


def calc_monthly_trend(
    df: pd.DataFrame,
    kv_filter=None,
    ma_dv_filter: Optional[list] = None,
    ns_filter: Optional[list] = None,
) -> pd.DataFrame:
    """
    Xu hướng % đạt KPIs & % thỏa LPT theo TỪNG THÁNG THỰC TẾ (không lũy kế), sắp xếp theo thời gian.
    Dùng cho biểu đồ đường (Line Chart) — khác bản chất với các Bảng A-F vốn tính theo kỳ lũy kế N tháng.
    `kv_filter` là list (đa lựa chọn, v4.12).
    """
    d = apply_nhan_su_filter(apply_ma_dv_filter(apply_kv_filter(df, kv_filter), ma_dv_filter), ns_filter)
    rows = []
    for ym, sub in d.groupby("YEAR_MONTH"):
        if pd.isna(ym):
            continue
        tot = len(sub)
        dat = len(sub[sub["KET_QUA_KPIS"] == "DAT_KPIS"])
        thoa = len(sub[sub["KET_QUA_LPT"] == "THOA_LPT"])
        rows.append({
            "YEAR_MONTH": ym,
            "THANG_STR": sub["THANG_STR"].iloc[0] if not sub.empty else str(ym),
            "pct_dat": round(dat / tot * 100, 1) if tot > 0 else 0.0,
            "pct_thoa": round(thoa / tot * 100, 1) if tot > 0 else 0.0,
            "tot_ns": tot,
        })
    if not rows:
        return pd.DataFrame(columns=["YEAR_MONTH", "THANG_STR", "pct_dat", "pct_thoa", "tot_ns"])
    return pd.DataFrame(rows).sort_values("YEAR_MONTH").reset_index(drop=True)


# Ánh xạ Khối chỉ tiêu -> tên cột trong DataFrame dài (long-form) trả về bởi
# calc_monthly_series_by_kv() / calc_monthly_series_by_unit(), dùng cho chế độ "Tất cả các tháng"
# (so sánh biến động nhiều tháng cùng lúc, thay cho 1 biểu đồ snapshot theo kỳ/tháng đang chọn).
# "chart_type" quyết định loại biểu đồ, dispatch qua build_trend_chart_by_type() — dùng `unit_key`
# khi vẽ theo ĐVKD/Nhân sự (mặc định), hoặc `kv_key` khi vẽ theo 6 Khu vực (truyền
# `value_col_key="kv_key"`, từ v4.10 — TRƯỚC v4.10 tầng Khu vực hardcode Bar riêng, không còn đúng).
# (v4.13) "chart_type" nay CỐ ĐỊNH theo từng khối chỉ tiêu — GIỐNG HỆT loại biểu đồ dùng ở cấp Khu
# vực (CHI_TIEU_CHART_MAP) và cấp ĐVKD (CHI_TIEU_UNIT_META), bất kể đang lọc Toàn Hàng/Khu vực/ĐVKD/
# Nhân sự/tháng nào: "hbar" (Tỷ lệ đạt KPIs/Thỏa LPT), "pie" (Phân bổ dải LPT — tỷ trọng đóng góp cả
# kỳ), "lollipop" (KHDN mới/KHTD tăng ròng/KHDN tăng ròng — `build_multi_lollipop_trend_chart()`,
# mới, nhiều cụm que+chấm tròn theo tháng), "bar" (Chi tiết dư nợ/huy động/TOI — cột dọc, giá trị âm
# hiện tự nhiên qua đường zero-line, không cần tô màu dương/âm riêng như ở cấp Khu vực/ĐVKD vì mỗi
# tháng có nhiều nhóm cạnh nhau — tô theo dấu sẽ làm mất khả năng phân biệt nhóm).
CHI_TIEU_TREND_META = {
    "Tỷ lệ đạt KPIs": {"kv_key": "pct_dat", "unit_key": "pct_dat", "suffix": "%", "title": "TỶ LỆ ĐẠT KPIs", "chart_type": "hbar"},
    "Tỷ lệ thỏa LPT": {"kv_key": "pct_thoa", "unit_key": "pct_thoa", "suffix": "%", "title": "TỶ LỆ THỎA LƯƠNG PHỤ TRỘI", "chart_type": "hbar"},
    "Phân bổ dải LPT thực nhận": {"kv_key": "lpt_tb", "unit_key": "lpt_tb", "suffix": "%", "title": "HỆ SỐ LPT THỰC NHẬN BÌNH QUÂN", "chart_type": "pie"},
    "KHDN mới/ĐVKD": {"kv_key": "avg_khdn_moi", "unit_key": "khdn_moi", "suffix": "", "title": "KHDN MỚI", "chart_type": "lollipop"},
    "KHTD tăng ròng/ĐVKD": {"kv_key": "avg_khtd_tang", "unit_key": "khtd_tang", "suffix": "", "title": "KHTD TĂNG RÒNG", "chart_type": "lollipop"},
    "KHDN tăng ròng/ĐVKD": {"kv_key": "avg_khdn_tang", "unit_key": "khdn_tang", "suffix": "", "title": "KHDN TĂNG RÒNG", "chart_type": "lollipop"},
    "Chi tiết dư nợ": {"kv_key": "tang_rong_cv", "unit_key": "tang_rong_cv", "suffix": "", "title": "TĂNG/GIẢM RÒNG CHO VAY", "chart_type": "bar"},
    "Chi tiết huy động": {"kv_key": "tang_rong_hd", "unit_key": "tang_rong_hd", "suffix": "", "title": "TĂNG/GIẢM RÒNG HUY ĐỘNG", "chart_type": "bar"},
    "Chi tiết TOI": {"kv_key": "tong_toi", "unit_key": "tong_toi", "suffix": "", "title": "TỔNG TOI", "chart_type": "bar"},
}
DEFAULT_TREND_META = {"kv_key": "pct_dat", "unit_key": "pct_dat", "suffix": "%", "title": "TỶ LỆ ĐẠT KPIs", "chart_type": "hbar"}


def calc_monthly_series_by_kv(df: pd.DataFrame, months_filter: Optional[list] = None) -> pd.DataFrame:
    """
    Chuỗi thời gian theo TỪNG THÁNG THỰC TẾ cho MỖI Khu vực (6 KV) — DataFrame dài (long-form),
    dùng khi bật 'Tất cả các tháng' để so sánh biến động của 1 chỉ tiêu giữa các Khu vực qua nhiều tháng.
    `months_filter` (v4.12, đa lựa chọn THÁNG): nếu có, chỉ giữ lại các tháng nằm trong danh sách
    này (lọc theo `THANG_STR`) — dùng khi người dùng chọn NHIỀU tháng cụ thể (không phải toàn bộ).
    """
    unique_yms = sorted(df["YEAR_MONTH"].dropna().unique())
    rows = []
    for ym in unique_yms:
        month_df = df[df["YEAR_MONTH"] == ym]
        thang_str = month_df["THANG_STR"].iloc[0] if not month_df.empty else str(ym)
        for kv_name, subset in _region_groups(month_df):
            if kv_name == "Tổng":
                continue
            tot = len(subset)
            dat = len(subset[subset["KET_QUA_KPIS"] == "DAT_KPIS"])
            thoa = len(subset[subset["KET_QUA_LPT"] == "THOA_LPT"])
            n_dv = subset["MA_DV"].nunique()
            khdn_moi = int(subset["SL_KHDN_MOI_N"].sum())
            khtd_tang = int(subset["SL_KH_TD_TANG_RONG"].sum())
            khdn_tang = int(subset["SL_KHDN_TANG_RONG"].sum())
            lpt_vals = subset["HE_SO_LPT_RAW"].dropna()
            rows.append({
                "YEAR_MONTH": ym,
                "THANG_STR": thang_str,
                "kv_label": kv_full_to_label(kv_name),
                "pct_dat": round(dat / tot * 100, 1) if tot > 0 else 0.0,
                "pct_thoa": round(thoa / tot * 100, 1) if tot > 0 else 0.0,
                "avg_khdn_moi": round(khdn_moi / n_dv, 1) if n_dv > 0 else 0.0,
                "avg_khtd_tang": round(khtd_tang / n_dv, 1) if n_dv > 0 else 0.0,
                "avg_khdn_tang": round(khdn_tang / n_dv, 1) if n_dv > 0 else 0.0,
                "lpt_tb": round(lpt_vals.mean() * 100, 1) if not lpt_vals.empty else None,
                "tong_cv_n": round(agg_unit_level_metric(subset, "TONG_CV_THANG_N"), 2),
                "tang_rong_cv": round(agg_unit_level_metric(subset, "TANG_RONG_CV"), 2),
                "tong_hd_n": round(agg_unit_level_metric(subset, "TONG_HD_THANG_N"), 2),
                "tang_rong_hd": round(agg_unit_level_metric(subset, "TANG_RONG_HD"), 2),
                "tong_toi": round(sum(agg_unit_level_metric(subset, c) for c in TOI_SOURCE_COLS.values()), 2),
            })
    if not rows:
        return pd.DataFrame()
    result = pd.DataFrame(rows).sort_values("YEAR_MONTH").reset_index(drop=True)
    if months_filter:
        result = result[result["THANG_STR"].isin(months_filter)].reset_index(drop=True)
    return result


def calc_monthly_series_by_unit(
    df: pd.DataFrame,
    kv_filter,
    ma_dv_filter: Optional[list] = None,
    ns_filter: Optional[list] = None,
    months_filter: Optional[list] = None,
) -> pd.DataFrame:
    """
    Chuỗi thời gian theo TỪNG THÁNG THỰC TẾ cho MỖI Đơn vị kinh doanh (MA_DV) trong (các) Khu vực đã
    chọn (`kv_filter` list, v4.12) — DataFrame dài (long-form), dùng khi bật 'Tất cả các tháng' để so
    sánh biến động giữa các ĐVKD. Nếu có `ma_dv_filter` (đa lựa chọn), chỉ tính cho các ĐVKD đó thay
    vì toàn bộ ĐVKD trong (các) Khu vực. `months_filter` (v4.12): chỉ giữ lại các tháng đã chọn.
    """
    d = apply_nhan_su_filter(apply_ma_dv_filter(apply_kv_filter(df, kv_filter), ma_dv_filter), ns_filter)
    unique_yms = sorted(d["YEAR_MONTH"].dropna().unique())
    rows = []
    for ym in unique_yms:
        month_df = d[d["YEAR_MONTH"] == ym]
        thang_str = month_df["THANG_STR"].iloc[0] if not month_df.empty else str(ym)
        for (ma_dv, ten_dv), sub in month_df.groupby(["MA_DV", "TEN_DV"], sort=True):
            tot = len(sub)
            dat = len(sub[sub["KET_QUA_KPIS"] == "DAT_KPIS"])
            thoa = len(sub[sub["KET_QUA_LPT"] == "THOA_LPT"])
            lpt_vals = sub["HE_SO_LPT_RAW"].dropna()
            rows.append({
                "YEAR_MONTH": ym,
                "THANG_STR": thang_str,
                "unit_label": f"{ma_dv} - {ten_dv}",
                "pct_dat": round(dat / tot * 100, 1) if tot > 0 else 0.0,
                "pct_thoa": round(thoa / tot * 100, 1) if tot > 0 else 0.0,
                "khdn_moi": int(sub["SL_KHDN_MOI_N"].sum()),
                "khtd_tang": int(sub["SL_KH_TD_TANG_RONG"].sum()),
                "khdn_tang": int(sub["SL_KHDN_TANG_RONG"].sum()),
                "lpt_tb": round(lpt_vals.mean() * 100, 1) if not lpt_vals.empty else None,
                "tong_cv_n": round(agg_unit_level_metric(sub, "TONG_CV_THANG_N"), 2),
                "tang_rong_cv": round(agg_unit_level_metric(sub, "TANG_RONG_CV"), 2),
                "tong_hd_n": round(agg_unit_level_metric(sub, "TONG_HD_THANG_N"), 2),
                "tang_rong_hd": round(agg_unit_level_metric(sub, "TANG_RONG_HD"), 2),
                "tong_toi": round(sum(agg_unit_level_metric(sub, c) for c in TOI_SOURCE_COLS.values()), 2),
            })
    if not rows:
        return pd.DataFrame()
    result = pd.DataFrame(rows).sort_values("YEAR_MONTH").reset_index(drop=True)
    if months_filter:
        result = result[result["THANG_STR"].isin(months_filter)].reset_index(drop=True)
    return result


def calc_monthly_series_by_staff(
    df: pd.DataFrame,
    kv_filter,
    ma_dv_filter: Optional[list] = None,
    ns_filter: Optional[list] = None,
    months_filter: Optional[list] = None,
) -> pd.DataFrame:
    """
    Chuỗi thời gian theo TỪNG THÁNG THỰC TẾ cho MỖI Nhân sự (TEN_CV) — DataFrame dài (long-form),
    dùng thay cho calc_monthly_series_by_unit() khi đã chọn ≥1 Nhân sự cụ thể (tầng lọc thứ 3),
    để biểu đồ 'Tất cả các tháng' tách riêng theo từng người thay vì gộp chung 1 cột theo ĐVKD.
    `kv_filter` là list (đa lựa chọn Khu vực, v4.12). `months_filter` (v4.12): chỉ giữ lại các tháng
    đã chọn (đa lựa chọn THÁNG). Lưu ý: các chỉ tiêu cấp ĐVKD (tong_cv_n, tong_hd_n, tong_toi...) sẽ
    hiển thị giá trị GIỐNG NHAU cho mọi nhân sự cùng 1 ĐVKD/tháng — đúng bản chất dữ liệu (xem
    UNIT_LEVEL_NUMERIC_COLUMNS), không phải lỗi tính toán.
    """
    d = apply_nhan_su_filter(apply_ma_dv_filter(apply_kv_filter(df, kv_filter), ma_dv_filter), ns_filter)
    unique_yms = sorted(d["YEAR_MONTH"].dropna().unique())
    rows = []
    for ym in unique_yms:
        month_df = d[d["YEAR_MONTH"] == ym]
        thang_str = month_df["THANG_STR"].iloc[0] if not month_df.empty else str(ym)
        for ten_cv, sub in month_df.groupby("TEN_CV", sort=True):
            if not ten_cv:
                continue
            tot = len(sub)
            dat = len(sub[sub["KET_QUA_KPIS"] == "DAT_KPIS"])
            thoa = len(sub[sub["KET_QUA_LPT"] == "THOA_LPT"])
            lpt_vals = sub["HE_SO_LPT_RAW"].dropna()
            rows.append({
                "YEAR_MONTH": ym,
                "THANG_STR": thang_str,
                "staff_label": ten_cv,
                "pct_dat": round(dat / tot * 100, 1) if tot > 0 else 0.0,
                "pct_thoa": round(thoa / tot * 100, 1) if tot > 0 else 0.0,
                "khdn_moi": int(sub["SL_KHDN_MOI_N"].sum()),
                "khtd_tang": int(sub["SL_KH_TD_TANG_RONG"].sum()),
                "khdn_tang": int(sub["SL_KHDN_TANG_RONG"].sum()),
                "lpt_tb": round(lpt_vals.mean() * 100, 1) if not lpt_vals.empty else None,
                "tong_cv_n": round(agg_unit_level_metric(sub, "TONG_CV_THANG_N"), 2),
                "tang_rong_cv": round(agg_unit_level_metric(sub, "TANG_RONG_CV"), 2),
                "tong_hd_n": round(agg_unit_level_metric(sub, "TONG_HD_THANG_N"), 2),
                "tang_rong_hd": round(agg_unit_level_metric(sub, "TANG_RONG_HD"), 2),
                "tong_toi": round(sum(agg_unit_level_metric(sub, c) for c in TOI_SOURCE_COLS.values()), 2),
            })
    if not rows:
        return pd.DataFrame()
    result = pd.DataFrame(rows).sort_values("YEAR_MONTH").reset_index(drop=True)
    if months_filter:
        result = result[result["THANG_STR"].isin(months_filter)].reset_index(drop=True)
    return result


# ──────────────────────────────────────────────────────────────────────────────
# BIỂU ĐỒ TRỰC QUAN (DONUT & 2 BARS CHUẨN THIẾT KẾ HDBANK)
# ──────────────────────────────────────────────────────────────────────────────


def build_donut_chart(kpi_dat: int, kpi_khong_dat: int):
    """Biểu đồ tròn khoét lỗ (Donut Chart) Tỷ lệ đạt KPIs toàn hàng."""
    fig = go.Figure(
        data=[
            go.Pie(
                labels=["Đạt KPIs", "Không đạt KPIs"],
                values=[kpi_dat, kpi_khong_dat],
                hole=0.58,
                marker=dict(
                    colors=["#34d399", "#fb923c"],
                    line=dict(color="#ffffff", width=2.5),
                ),
                textinfo="percent",
                textposition="inside",
                insidetextorientation="horizontal",
                sort=False,
                textfont=dict(size=17, color="#ffffff", family="Segoe UI, Arial"),
            )
        ]
    )

    fig.update_layout(
        title=dict(
            text="<b>TỶ LỆ ĐẠT KPIs TOÀN HÀNG</b>",
            x=0.5,
            y=0.96,
            xanchor="center",
            yanchor="top",
            font=dict(size=17, color="#111827", family="Segoe UI, Arial"),
        ),
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="left",
            x=1.02,
            font=dict(size=15, color="#374151"),
        ),
        margin=dict(l=10, r=100, t=35, b=15),
        height=290,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def build_trend_line_chart(trend_df: pd.DataFrame, title_suffix: str = "Toàn hàng"):
    """Line Chart — Xu hướng % đạt KPIs & % thỏa LPT theo từng tháng thực tế (không lũy kế)."""
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=trend_df["THANG_STR"],
            y=trend_df["pct_dat"],
            mode="lines+markers",
            name="% Đạt KPIs",
            line=dict(color="#60a5fa", width=2.5),
            marker=dict(size=7, color="#60a5fa"),
        )
    )
    fig.add_trace(
        go.Scatter(
            x=trend_df["THANG_STR"],
            y=trend_df["pct_thoa"],
            mode="lines+markers",
            name="% Thỏa LPT",
            line=dict(color="#34d399", width=2.5),
            marker=dict(size=7, color="#34d399"),
        )
    )
    fig.add_hline(y=80, line_width=1, line_dash="dot", line_color="#60a5fa", opacity=0.45)
    fig.add_hline(y=70, line_width=1, line_dash="dot", line_color="#34d399", opacity=0.45)

    fig.update_layout(
        title=dict(
            text=f"<b>XU HƯỚNG % ĐẠT KPIs & % THỎA LPT THEO THÁNG — {title_suffix}</b>",
            x=0.5,
            y=0.95,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        yaxis=dict(range=[0, 105], ticksuffix="%", gridcolor="#f0f2f5", tickfont=dict(size=14, color="#6b7280")),
        xaxis=dict(tickfont=dict(size=14.5, color="#374151")),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5, font=dict(size=14.5)),
        margin=dict(l=10, r=20, t=55, b=25),
        height=280,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def build_multi_bar_trend_chart(
    long_df: pd.DataFrame,
    group_col: str,
    value_col: str,
    title: str,
    y_suffix: str = "",
    show_legend: bool = True,
):
    """
    Bar Chart nhiều cột nhóm theo tháng (mỗi Khu vực/ĐVKD/Nhân sự 1 cột màu pastel riêng trong từng
    nhóm tháng) — dùng cho: (1) so sánh 6 Khu vực khi bật 'Tất cả các tháng' (giữ Bar theo phản hồi
    người dùng, xem v3.11); (2) các chỉ tiêu SỐ LƯỢNG/TIỀN TỆ (`chart_type="bar"` trong
    CHI_TIEU_TREND_META, v4.3) khi so sánh nhiều ĐVKD/Nhân sự — cột rõ hơn Line cho giá trị rời rạc,
    có thể âm/bằng 0. Chỉ tiêu TỶ LỆ % dùng `build_multi_line_trend_chart()` thay vì hàm này.
    Màu tự nhận diện theo `group_col`: "kv_label" -> màu theo Khu vực, còn lại -> màu ĐVKD/Nhân sự.
    `show_legend=False` (v4.7) tắt legend riêng của biểu đồ này khi được vẽ trong 1 lưới nhiều biểu
    đồ dùng chung 1 legend tổng hợp bên ngoài (xem lời gọi trong main()).
    """
    fig = go.Figure()
    group_labels = list(dict.fromkeys(long_df[group_col]))
    for idx, grp in enumerate(group_labels):
        sub = long_df[long_df[group_col] == grp].sort_values("YEAR_MONTH")
        if sub.empty:
            continue
        color = get_kv_color(grp, idx) if group_col == "kv_label" else get_unit_color(idx)
        # (v4.31) Hiển thị nhãn giá trị ngay trên từng cột — % dùng 1 chữ số thập phân, số tuyệt đối
        # dùng dấu chấm phân cách nghìn (đúng quy ước đã dùng ở các hàm cột dọc khác trong file).
        if y_suffix == "%":
            bar_text = [f"{v:.1f}%" for v in sub[value_col]]
        else:
            bar_text = [f"{v:,.0f}".replace(",", ".") for v in sub[value_col]]
        fig.add_trace(
            go.Bar(
                x=sub["THANG_STR"],
                y=sub[value_col],
                name=grp,
                marker=dict(color=color),
                text=bar_text,
                textposition="outside",
                textfont=dict(size=11, color="#111827", family="Segoe UI, Arial"),
                cliponaxis=False,
            )
        )

    fig.update_layout(
        barmode="group",
        bargap=0.2,
        bargroupgap=0.08,
        title=dict(
            text=f"<b>{title}</b>",
            x=0.5,
            y=0.96,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        yaxis=dict(
            ticksuffix=y_suffix,
            gridcolor="#f0f2f5",
            tickfont=dict(size=14, color="#6b7280"),
            zeroline=True,
            zerolinecolor="#9ca3af",
            zerolinewidth=1.5,
        ),
        xaxis=dict(tickfont=dict(size=14, color="#374151")),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(size=13.5)),
        showlegend=show_legend,
        margin=dict(l=10, r=20, t=45, b=70),
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def build_multi_line_trend_chart(
    long_df: pd.DataFrame,
    group_col: str,
    value_col: str,
    title: str,
    y_suffix: str = "",
    show_legend: bool = True,
):
    """
    Line Chart nhiều đường theo tháng (mỗi Đơn vị kinh doanh / Nhân sự 1 đường màu pastel riêng,
    có marker tại từng điểm) — dùng cho so sánh xu hướng nhiều ĐVKD/Nhân sự qua các tháng khi bật
    'Tất cả các tháng': đọc rõ xu hướng tăng/giảm hơn Bar Chart khi so sánh nhiều đối tượng cùng lúc,
    đặc biệt khi chọn nhiều Nhân sự. Từ v4.10, cũng dùng được cho biểu đồ theo 6 Khu vực khi
    `chart_type` của chỉ tiêu là "line" (xem `build_trend_chart_by_type(..., value_col_key="kv_key")`).
    `show_legend=False` (v4.7): tắt legend riêng khi vẽ trong lưới dùng chung 1 legend tổng hợp.
    Mỗi đường (v4.7) hiển thị data label tại điểm ĐẦU và điểm CUỐI (các điểm giữa để trống) — giúp
    đọc nhanh giá trị mốc đầu/cuối kỳ mà không làm rối biểu đồ như label ở mọi điểm. Màu tự nhận
    diện theo `group_col` giống `build_multi_bar_trend_chart()`: "kv_label" -> màu theo Khu vực,
    còn lại -> màu ĐVKD/Nhân sự.
    """
    fig = go.Figure()
    group_labels = list(dict.fromkeys(long_df[group_col]))
    for idx, grp in enumerate(group_labels):
        sub = long_df[long_df[group_col] == grp].sort_values("YEAR_MONTH")
        if sub.empty:
            continue
        color = get_kv_color(grp, idx) if group_col == "kv_label" else get_unit_color(idx)
        vals = list(sub[value_col])
        text_labels = ["" for _ in vals]
        if vals:
            text_labels[0] = f"{vals[0]:.0f}{y_suffix}"
            text_labels[-1] = f"{vals[-1]:.0f}{y_suffix}"
        fig.add_trace(
            go.Scatter(
                x=sub["THANG_STR"],
                y=sub[value_col],
                name=grp,
                mode="lines+markers+text",
                line=dict(color=color, width=2.5, shape="linear"),
                marker=dict(color=color, size=7, line=dict(color="white", width=1)),
                text=text_labels,
                textposition="top center",
                textfont=dict(size=13, color=color, family="Segoe UI, Arial"),
                hovertemplate=f"<b>{grp}</b><br>%{{x}}: %{{y}}{y_suffix}<extra></extra>",
            )
        )

    # Chỉ tiêu % (Tỷ lệ đạt KPIs, Tỷ lệ thỏa LPT, Hệ số LPT TB) bị bó buộc 0-100% — cố định trục Y
    # thay vì để Plotly auto-scale sát dữ liệu, tránh đường gần như phẳng (vd quanh 100%) trông
    # "gãy khúc" giả tạo do trục Y co quá hẹp.
    yaxis_cfg = dict(
        ticksuffix=y_suffix,
        gridcolor="#f0f2f5",
        tickfont=dict(size=14, color="#6b7280"),
        zeroline=True,
        zerolinecolor="#9ca3af",
        zerolinewidth=1.5,
    )
    if y_suffix == "%":
        yaxis_cfg["range"] = [0, 105]

    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b>",
            x=0.5,
            y=0.96,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        yaxis=yaxis_cfg,
        xaxis=dict(tickfont=dict(size=14, color="#374151")),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(size=13.5)),
        showlegend=show_legend,
        margin=dict(l=10, r=20, t=45, b=70),
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
        hovermode="x unified",
    )
    return fig


def build_multi_area_trend_chart(
    long_df: pd.DataFrame,
    group_col: str,
    value_col: str,
    title: str,
    y_suffix: str = "",
    show_legend: bool = True,
):
    """
    Area Chart (biểu đồ vùng) nhiều nhóm theo tháng — cấu trúc giống `build_multi_line_trend_chart()`
    nhưng tô vùng dưới đường (`fill="tozeroy"`, màu nhạt 30% opacity theo đúng màu của nhóm) thay vì
    vẽ đường + marker — dùng để khác biệt thị giác với "line" (vd Tỷ lệ thỏa LPT, cùng là % liên tục
    nhưng khác "Tỷ lệ đạt KPIs") hoặc thể hiện khối lượng tích lũy (Chi tiết dư nợ/huy động), v4.8.
    Màu tự nhận diện theo `group_col` giống `build_multi_bar_trend_chart()`: "kv_label" -> màu theo
    Khu vực, còn lại -> màu ĐVKD/Nhân sự.
    """
    fig = go.Figure()
    group_labels = list(dict.fromkeys(long_df[group_col]))
    for idx, grp in enumerate(group_labels):
        sub = long_df[long_df[group_col] == grp].sort_values("YEAR_MONTH")
        if sub.empty:
            continue
        color = get_kv_color(grp, idx) if group_col == "kv_label" else get_unit_color(idx)
        fig.add_trace(
            go.Scatter(
                x=sub["THANG_STR"],
                y=sub[value_col],
                name=grp,
                mode="lines",
                fill="tozeroy",
                fillcolor=_hex_to_rgba(color, 0.3),
                line=dict(color=color, width=2),
                hovertemplate=f"<b>{grp}</b><br>%{{x}}: %{{y}}{y_suffix}<extra></extra>",
            )
        )

    yaxis_cfg = dict(
        ticksuffix=y_suffix,
        gridcolor="#f0f2f5",
        tickfont=dict(size=14, color="#6b7280"),
        zeroline=True,
        zerolinecolor="#9ca3af",
        zerolinewidth=1.5,
    )
    if y_suffix == "%":
        yaxis_cfg["range"] = [0, 105]

    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b>",
            x=0.5,
            y=0.96,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        yaxis=yaxis_cfg,
        xaxis=dict(tickfont=dict(size=14, color="#374151")),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(size=13.5)),
        showlegend=show_legend,
        margin=dict(l=10, r=20, t=45, b=70),
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
        hovermode="x unified",
    )
    return fig


def build_multi_hbar_trend_chart(
    long_df: pd.DataFrame,
    group_col: str,
    value_col: str,
    title: str,
    y_suffix: str = "",
    show_legend: bool = True,
):
    """
    Bar Chart NGANG nhiều nhóm theo tháng (trục tháng nằm DỌC — mỗi tháng 1 hàng, mỗi ĐVKD/Nhân sự
    1 cột màu riêng trong hàng đó) — biến thể xoay trục của `build_multi_bar_trend_chart()`, đa
    dạng hóa thêm loại "cột ngang" ngay trong ngữ cảnh so sánh nhiều đối tượng theo tháng (v4.6).
    `show_legend=False` (v4.7) tắt legend riêng khi vẽ trong lưới dùng chung 1 legend tổng hợp. Màu
    tự nhận diện theo `group_col` giống `build_multi_bar_trend_chart()` (v4.10).
    """
    fig = go.Figure()
    group_labels = list(dict.fromkeys(long_df[group_col]))
    for idx, grp in enumerate(group_labels):
        sub = long_df[long_df[group_col] == grp].sort_values("YEAR_MONTH")
        if sub.empty:
            continue
        color = get_kv_color(grp, idx) if group_col == "kv_label" else get_unit_color(idx)
        # (v4.32) Hiển thị nhãn giá trị ngay trên từng cột ngang — cùng quy ước định dạng đã dùng ở
        # build_multi_bar_trend_chart() (v4.31): % dùng 1 chữ số thập phân, số tuyệt đối dùng dấu
        # chấm phân cách nghìn.
        if y_suffix == "%":
            bar_text = [f"{v:.1f}%" for v in sub[value_col]]
        else:
            bar_text = [f"{v:,.0f}".replace(",", ".") for v in sub[value_col]]
        fig.add_trace(
            go.Bar(
                y=sub["THANG_STR"],
                x=sub[value_col],
                name=grp,
                orientation="h",
                marker=dict(color=color),
                text=bar_text,
                textposition="outside",
                textfont=dict(size=11, color="#111827", family="Segoe UI, Arial"),
                cliponaxis=False,
            )
        )

    fig.update_layout(
        barmode="group",
        bargap=0.2,
        bargroupgap=0.08,
        title=dict(
            text=f"<b>{title}</b>",
            x=0.5,
            y=0.96,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        xaxis=dict(
            ticksuffix=y_suffix,
            gridcolor="#f0f2f5",
            tickfont=dict(size=14, color="#6b7280"),
            zeroline=True,
            zerolinecolor="#9ca3af",
            zerolinewidth=1.5,
        ),
        yaxis=dict(autorange="reversed", tickfont=dict(size=14, color="#374151")),
        legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="center", x=0.5, font=dict(size=13.5)),
        showlegend=show_legend,
        margin=dict(l=10, r=20, t=45, b=55),
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def build_multi_lollipop_trend_chart(
    long_df: pd.DataFrame,
    group_col: str,
    value_col: str,
    title: str,
    y_suffix: str = "",
    show_legend: bool = True,
):
    """
    Lollipop (kẹo mút) NHIỀU NHÓM theo tháng — mỗi tháng 1 vị trí trên trục X, mỗi ĐVKD/Nhân sự/Khu
    vực 1 que (từ 0 tới giá trị) + đầu tròn, xếp lệch cạnh nhau trong cùng tháng (giống nguyên lý
    Bar Chart nhóm cột nhưng dùng que+chấm tròn thay vì khối đặc) — v4.13, cố định cho Khối D/E/F ở
    MỌI ngữ cảnh, kể cả khi so sánh nhiều tháng cùng lúc. Trục X dùng vị trí số (không phải chuỗi
    tháng trực tiếp) để có thể dịch lệch từng nhóm trong cùng 1 tháng — nhãn tháng gắn qua `tickvals`.
    """
    month_order = list(dict.fromkeys(long_df.sort_values("YEAR_MONTH")["THANG_STR"]))
    month_index = {m: i for i, m in enumerate(month_order)}
    group_labels = list(dict.fromkeys(long_df[group_col]))
    n_groups = max(len(group_labels), 1)
    slot_width = 0.8 / n_groups

    fig = go.Figure()
    for idx, grp in enumerate(group_labels):
        sub = long_df[long_df[group_col] == grp].sort_values("YEAR_MONTH")
        if sub.empty:
            continue
        color = get_kv_color(grp, idx) if group_col == "kv_label" else get_unit_color(idx)
        offset = (idx - (n_groups - 1) / 2) * slot_width
        xs = [month_index[m] + offset for m in sub["THANG_STR"]]
        vals = list(sub[value_col])

        stem_x, stem_y = [], []
        for x, v in zip(xs, vals):
            stem_x += [x, x, None]
            stem_y += [0, v, None]
        fig.add_trace(go.Scatter(x=stem_x, y=stem_y, mode="lines", line=dict(color=color, width=1.5), opacity=0.55, showlegend=False, hoverinfo="skip"))
        # (v4.31) Hiển thị rõ giá trị ngay tại từng đầu que — tên tháng (trước đây gán vào `text` để
        # ghép hover) chuyển sang `customdata` để KHÔNG mất thông tin tháng khi hover, nhường `text`
        # cho nhãn giá trị hiển thị trực tiếp trên biểu đồ.
        point_labels = [
            (f"{v:.1f}%" if y_suffix == "%" else f"{v:,.0f}".replace(",", ".")) if v is not None else None
            for v in vals
        ]
        fig.add_trace(
            go.Scatter(
                x=xs,
                y=vals,
                mode="markers+text",
                name=grp,
                marker=dict(color=color, size=10, line=dict(color="white", width=1)),
                text=point_labels,
                textposition="top center",
                textfont=dict(size=10, color=color, family="Segoe UI, Arial"),
                customdata=list(sub["THANG_STR"]),
                hovertemplate=f"<b>{grp}</b><br>%{{customdata}}: %{{y}}{y_suffix}<extra></extra>",
            )
        )

    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b>",
            x=0.5,
            y=0.96,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        xaxis=dict(
            tickmode="array",
            tickvals=list(range(len(month_order))),
            ticktext=month_order,
            tickfont=dict(size=14, color="#374151"),
        ),
        yaxis=dict(
            ticksuffix=y_suffix,
            gridcolor="#f0f2f5",
            tickfont=dict(size=14, color="#6b7280"),
            zeroline=True,
            zerolinecolor="#9ca3af",
            zerolinewidth=1.5,
        ),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(size=13.5)),
        showlegend=show_legend,
        margin=dict(l=10, r=20, t=45, b=70),
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
        hovermode="closest",
    )
    return fig


def build_trend_contribution_pie(long_df: pd.DataFrame, group_col: str, value_col: str, title: str, show_legend: bool = True):
    """
    Donut Chart tỷ trọng đóng góp — SUM `value_col` theo `group_col` trên TOÀN BỘ giai đoạn đang
    xem (không tách theo từng tháng) — góc nhìn bổ sung "cơ cấu đóng góp cả kỳ", đặt cạnh các biểu
    đồ theo-tháng khác để đa dạng loại biểu đồ (v4.6), tương tự cách Khối C có cả Stacked Bar+Donut.
    `show_legend=False` (v4.7) tắt legend riêng khi vẽ trong lưới dùng chung 1 legend tổng hợp.
    Nếu chỉ đúng 1 nhóm có giá trị dương (v4.7): donut vẫn vẽ (100% 1 màu) nhưng ghi rõ tên nhóm đó
    ngay giữa vòng tròn bằng annotation — tránh nhìn như 1 vòng tròn trống/lỗi khi chỉ có 1 màu.
    """
    grouped = long_df.groupby(group_col)[value_col].sum()
    grouped = grouped[grouped > 0].sort_values(ascending=False)
    labels = list(grouped.index)
    vals = list(grouped.values)
    colors = [get_kv_color(lbl, i) if group_col == "kv_label" else get_unit_color(i) for i, lbl in enumerate(labels)]

    if not labels:
        fig = go.Figure()
        fig.update_layout(
            title=dict(text=f"<b>{title} — TỶ TRỌNG ĐÓNG GÓP CẢ KỲ</b>", x=0.5, font=dict(size=16.5, color="#111827", family="Segoe UI, Arial")),
            annotations=[dict(text="Không có dữ liệu dương để tính tỷ trọng", showarrow=False, font=dict(size=15, color="#6b7280"))],
            height=340,
            plot_bgcolor="white",
            paper_bgcolor="white",
        )
        return fig

    fig = go.Figure(
        data=[
            go.Pie(
                labels=labels,
                values=vals,
                hole=0.55,
                marker=dict(colors=colors, line=dict(color="#ffffff", width=2)),
                textinfo="percent",
                textposition="inside",
                textfont=dict(size=15, color="#ffffff", family="Segoe UI, Arial"),
            )
        ]
    )
    single_group_annotations = []
    if len(labels) == 1:
        single_group_annotations = [dict(text=labels[0], x=0.5, y=0.5, showarrow=False, font=dict(size=15, color="#111827", family="Segoe UI, Arial"))]
    fig.update_layout(
        title=dict(
            text=f"<b>{title} — TỶ TRỌNG ĐÓNG GÓP CẢ KỲ</b>",
            x=0.5,
            y=0.96,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        annotations=single_group_annotations,
        legend=dict(orientation="h", yanchor="top", y=-0.05, xanchor="center", x=0.5, font=dict(size=13.5)),
        showlegend=show_legend,
        margin=dict(l=10, r=10, t=45, b=45),
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def build_trend_chart_by_type(
    monthly_df: pd.DataFrame,
    group_col: str,
    tmeta: dict,
    title_suffix: str = "THEO THÁNG",
    show_legend: bool = True,
    value_col_key: str = "unit_key",
):
    """
    Điều phối loại biểu đồ theo tháng dựa trên `tmeta['chart_type']`
    (line/area/bar/hbar/pie/lollipop, v4.13). `show_legend=False`: dùng khi vẽ trong 1 lưới nhiều
    biểu đồ đã có 1 legend tổng hợp riêng bên ngoài (xem main()), tránh lặp legend 9 lần.
    `value_col_key` (v4.10): tên cột giá trị trong `tmeta` cần đọc — `"unit_key"` (mặc định, dùng
    với `monthly_df` group theo ĐVKD/Nhân sự, cột `unit_label`/`staff_label`) hoặc `"kv_key"` (dùng
    với `monthly_kv_df` group theo Khu vực, cột `kv_label` — tên cột trong 2 DataFrame này KHÁC
    NHAU cho 3 chỉ tiêu D/E/F, vd `avg_khdn_moi` (kv) vs `khdn_moi` (unit), nên bắt buộc chọn đúng
    key tương ứng, không thể dùng chung 1 key mặc định cho cả 2 ngữ cảnh).
    "pie": trước khi vẽ, GỘP `monthly_df` (nhiều dòng/tháng mỗi nhóm) thành 1 dòng/nhóm bằng SUM
    `value_col` theo `group_col` — thể hiện tỷ trọng đóng góp của từng nhóm trong TOÀN BỘ kỳ đang
    xem (không phải theo từng tháng), rồi mới gọi `build_trend_contribution_pie()`.
    "lollipop" (v4.13): `build_multi_lollipop_trend_chart()` — giữ ĐÚNG loại biểu đồ Lollipop cố
    định của Khối D/E/F ngay cả khi xem theo tháng (nhiều nhóm, mỗi tháng 1 cụm que+chấm tròn).
    """
    chart_type = tmeta.get("chart_type", "bar")
    value_col = tmeta[value_col_key]
    title = f"{tmeta['title']} {title_suffix}"
    if chart_type == "line":
        return build_multi_line_trend_chart(monthly_df, group_col, value_col, title, tmeta["suffix"], show_legend=show_legend)
    if chart_type == "area":
        return build_multi_area_trend_chart(monthly_df, group_col, value_col, title, tmeta["suffix"], show_legend=show_legend)
    if chart_type == "hbar":
        return build_multi_hbar_trend_chart(monthly_df, group_col, value_col, title, tmeta["suffix"], show_legend=show_legend)
    if chart_type == "lollipop":
        return build_multi_lollipop_trend_chart(monthly_df, group_col, value_col, title, tmeta["suffix"], show_legend=show_legend)
    if chart_type == "pie":
        agg_df = monthly_df.groupby(group_col, as_index=False)[value_col].sum()
        return build_trend_contribution_pie(agg_df, group_col, value_col, tmeta["title"], show_legend=show_legend)
    return build_multi_bar_trend_chart(monthly_df, group_col, value_col, title, tmeta["suffix"], show_legend=show_legend)


def render_shared_group_legend(labels: list[str], colors: list[str]) -> None:
    """
    Legend dùng chung (chấm tròn màu + tên) cho 1 lưới nhiều biểu đồ theo tháng (v4.7) — khi cả 9
    biểu đồ trong lưới đều lặp lại cùng 1 tập nhóm màu (6 Khu vực / các ĐVKD / các Nhân sự đang lọc),
    vẽ legend chung 1 lần thay vì để mỗi biểu đồ tự vẽ legend riêng (lặp lại 9 lần, dư thừa).
    """
    items_html = "".join(
        f'<span style="display:inline-flex; align-items:center; gap:4px; margin-right:14px;">'
        f'<span style="width:9px; height:9px; border-radius:50%; background:{color}; display:inline-block; flex-shrink:0;"></span>'
        f'<span style="font-size:11px; color:#374151; font-weight:600;">{label}</span></span>'
        for label, color in zip(labels, colors)
    )
    st.markdown(
        f'<div style="display:flex; flex-wrap:wrap; align-items:center; padding:2px 2px 10px 2px;">{items_html}</div>',
        unsafe_allow_html=True,
    )


def build_kv_kpi_bar(kv_charts: list[dict]):
    """Biểu đồ thanh ngang SO SÁNH TỶ LỆ ĐẠT KPIs THEO 6 KHU VỰC (%) — mỗi Khu vực 1 màu pastel riêng."""
    labels = [k["kv_label"] for k in kv_charts]
    vals = [k["pct_dat"] for k in kv_charts]
    bar_colors = [get_kv_color(lbl, i) for i, lbl in enumerate(labels)]

    fig = go.Figure(
        go.Bar(
            x=vals,
            y=labels,
            orientation="h",
            marker=dict(color=bar_colors),
            text=[f"{v:.1f}%" for v in vals],
            textposition="outside",
            textfont=dict(size=15.5, color="#111827", family="Segoe UI, Arial"),
            cliponaxis=False,
        )
    )

    fig.update_layout(
        title=dict(
            text="<b>TỶ LỆ ĐẠT KPIs THEO 6 KHU VỰC (%)</b>",
            x=0.5,
            y=0.98,
            xanchor="center",
            yanchor="top",
            font=dict(size=17, color="#111827", family="Segoe UI, Arial"),
        ),
        yaxis=dict(
            autorange="reversed",
            tickfont=dict(size=14.5, color="#374151"),
        ),
        xaxis=dict(
            range=[0, 112],
            tickvals=[0, 20, 40, 60, 80, 100],
            ticktext=["0%", "20%", "40%", "60%", "80%", "100%"],
            gridcolor="#f0f2f5",
            tickfont=dict(size=14, color="#6b7280"),
        ),
        margin=dict(l=10, r=40, t=44, b=25),
        height=290,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )

    fig.add_vline(
        x=80,
        line_width=1.5,
        line_dash="dash",
        line_color=PASTEL_TARGET_LINE,
        annotation_text="Target 80%",
        annotation_position="top right",
        annotation=dict(
            yref="paper",
            y=1.04,
            yanchor="bottom",
            font=dict(size=14, color=PASTEL_TARGET_LINE, family="Segoe UI, Arial"),
        ),
    )
    return fig


def build_kv_lpt_bar(kv_charts: list[dict]):
    """Biểu đồ thanh ngang SO SÁNH TỶ LỆ THỎA LƯƠNG PHỤ TRỘI THEO 6 KHU VỰC (%) — mỗi Khu vực 1 màu pastel riêng."""
    labels = [k["kv_label"] for k in kv_charts]
    vals = [k["pct_thoa"] for k in kv_charts]
    bar_colors = [get_kv_color(lbl, i) for i, lbl in enumerate(labels)]

    fig = go.Figure(
        go.Bar(
            x=vals,
            y=labels,
            orientation="h",
            marker=dict(color=bar_colors),
            text=[f"{v:.1f}%" for v in vals],
            textposition="outside",
            textfont=dict(size=15.5, color="#111827", family="Segoe UI, Arial"),
            cliponaxis=False,
        )
    )

    fig.update_layout(
        title=dict(
            text="<b>TỶ LỆ THỎA LƯƠNG PHỤ TRỘI THEO 6 KHU VỰC (%)</b>",
            x=0.5,
            y=0.98,
            xanchor="center",
            yanchor="top",
            font=dict(size=17, color="#111827", family="Segoe UI, Arial"),
        ),
        yaxis=dict(
            autorange="reversed",
            tickfont=dict(size=14.5, color="#374151"),
        ),
        xaxis=dict(
            range=[0, 112],
            tickvals=[0, 20, 40, 60, 80, 100],
            ticktext=["0%", "20%", "40%", "60%", "80%", "100%"],
            gridcolor="#f0f2f5",
            tickfont=dict(size=14, color="#6b7280"),
        ),
        margin=dict(l=10, r=40, t=44, b=25),
        height=290,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )

    fig.add_vline(
        x=70,
        line_width=1.5,
        line_dash="dash",
        line_color=PASTEL_TARGET_LINE,
        annotation_text="Target 70%",
        annotation_position="top right",
        annotation=dict(
            yref="paper",
            y=1.04,
            yanchor="bottom",
            font=dict(size=14, color=PASTEL_TARGET_LINE, family="Segoe UI, Arial"),
        ),
    )
    return fig


# Màu pastel 6 dải Hệ số LPT thực nhận (hồng nhạt -> xanh ngọc) + dải "Chưa xác định" (NULL/gạch nối)
LPT_BIN_COLORS = {
    "< 0%": "#f43f5e",
    "0% ≤ … < 20%": "#ec4899",
    "20% ≤ … < 50%": "#fb923c",
    "50% ≤ … < 70%": "#fbbf24",
    "70% ≤ … < 100%": "#a3e635",
    "= 100%": "#10b981",
    "Chưa xác định": "#cbd5e1",
}


def build_lpt_distribution_chart(kv_charts: list[dict]):
    """BẢNG C — Stacked Bar 100% phân bổ dải Hệ số LPT thực nhận theo 6 Khu vực."""
    labels = [k["kv_label"] for k in kv_charts]

    fig = go.Figure()
    for label, color in LPT_BIN_COLORS.items():
        vals = [k["lpt_bin_pct"].get(label, 0.0) for k in kv_charts]
        fig.add_trace(
            go.Bar(
                name=label,
                x=vals,
                y=labels,
                orientation="h",
                marker=dict(color=color, cornerradius=6, line=dict(color="#ffffff", width=2)),
                text=[f"{v:.0f}%" if v >= 8 else "" for v in vals],
                textposition="inside",
                textfont=dict(size=13.5, color="#ffffff", family="Segoe UI, Arial", weight=700),
                hovertemplate="%{y}: %{x:.1f}%<extra>" + label + "</extra>",
            )
        )

    fig.update_layout(
        barmode="stack",
        bargap=0.35,
        title=dict(
            text="<b>🍩 PHÂN BỔ DẢI HỆ SỐ LPT THỰC NHẬN THEO 6 KHU VỰC (%)</b>",
            x=0.5,
            y=0.98,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        yaxis=dict(autorange="reversed", tickfont=dict(size=14.5, color="#374151"), ticksuffix="  "),
        xaxis=dict(
            range=[0, 100],
            ticksuffix="%",
            gridcolor="#f1f0fb",
            tickfont=dict(size=14, color="#9ca3af"),
        ),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.18,
            xanchor="center",
            x=0.5,
            font=dict(size=13, color="#374151"),
        ),
        margin=dict(l=10, r=20, t=40, b=75),
        height=340,
        plot_bgcolor="#fdfcff",
        paper_bgcolor="#fdfcff",
    )
    return fig


def build_lpt_distribution_pie(df: pd.DataFrame, thang_filter=None):
    """Pie Chart — Cơ cấu 6 dải Hệ số LPT thực nhận TOÀN HÀNG (bổ sung góc nhìn cho Stacked Bar theo KV). `thang_filter` là list (đa lựa chọn, v4.12)."""
    df_thang, _, _ = apply_thang_filter(df, thang_filter)
    tot = len(df_thang)
    vals = df_thang["HE_SO_LPT_RAW"].dropna() if tot > 0 else pd.Series(dtype=float)

    labels, counts, colors = [], [], []
    for label, cond_fn in LPT_BINS:
        labels.append(label)
        counts.append(int(cond_fn(vals).sum()) if tot > 0 else 0)
        colors.append(LPT_BIN_COLORS[label])
    labels.append("Chưa xác định")
    counts.append(tot - len(vals))
    colors.append(LPT_BIN_COLORS["Chưa xác định"])

    fig = go.Figure(
        go.Pie(
            labels=labels,
            values=counts,
            hole=0.55,
            marker=dict(colors=colors, line=dict(color="#ffffff", width=2.5)),
            textinfo="percent",
            textposition="inside",
            insidetextorientation="horizontal",
            textfont=dict(size=15, color="#ffffff", family="Segoe UI, Arial", weight=700),
            sort=False,
        )
    )
    fig.update_layout(
        title=dict(
            text="<b>🍩 CƠ CẤU DẢI LPT THỰC NHẬN — TOÀN HÀNG</b>",
            x=0.5,
            y=0.96,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=13.5, color="#374151")),
        margin=dict(l=10, r=110, t=35, b=15),
        height=320,
        plot_bgcolor="#fdfcff",
        paper_bgcolor="#fdfcff",
    )
    return fig


def _build_kv_vertical_bar_chart(kv_charts: list[dict], value_key: str, title: str, y_title: str, diverging: bool = False):
    """
    Cột dọc đơn giản (1 chỉ số) theo 6 Khu vực. Mặc định mỗi Khu vực 1 màu pastel riêng; khi
    `diverging=True` (v4.13, dùng cho Khối G/H/I — "cột dọc dương/âm") tô màu theo dấu giá trị
    (xanh lá = dương, đỏ nhạt = âm) thay vì theo Khu vực, để thể hiện rõ tăng/giảm.
    """
    labels = [k["kv_label"] for k in kv_charts]
    vals = [k[value_key] for k in kv_charts]
    if diverging:
        bar_colors = [(PASTEL_POSITIVE if v >= 0 else PASTEL_NEGATIVE) for v in vals]
    else:
        bar_colors = [get_kv_color(lbl, i) for i, lbl in enumerate(labels)]

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=labels,
            y=vals,
            name=y_title,
            marker=dict(color=bar_colors),
            text=[f"{v:,.0f}".replace(",", ".") for v in vals],
            textposition="outside",
            textfont=dict(size=14.5, color="#111827", family="Segoe UI, Arial"),
        )
    )
    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b>",
            x=0.5,
            y=0.98,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        xaxis=dict(tickfont=dict(size=14, color="#374151"), tickangle=-12),
        yaxis=dict(
            title=dict(text=y_title, font=dict(size=14.5)),
            gridcolor="#f0f2f5",
            tickfont=dict(size=14, color="#6b7280"),
            zeroline=True,
            zerolinecolor="#9ca3af",
            zerolinewidth=1.5,
        ),
        showlegend=False,
        margin=dict(l=10, r=10, t=55, b=45),
        height=300,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def _build_kv_lollipop_chart(kv_charts: list[dict], value_key: str, title: str, y_title: str, diverging: bool = False):
    """
    Biểu đồ Lollipop (kẹo mút) qua 6 Khu vực — que mảnh từ 0 tới giá trị + đầu tròn tại giá trị
    (v4.13, cố định cho Khối D/E/F ở cấp Khu vực bất kể bộ lọc đang chọn). Mặc định mỗi Khu vực 1
    màu pastel riêng; `diverging=True` tô màu theo dấu giá trị (dùng khi chỉ tiêu có thể âm).
    """
    labels = [k["kv_label"] for k in kv_charts]
    vals = [k[value_key] for k in kv_charts]
    if diverging:
        dot_colors = [(PASTEL_POSITIVE if v >= 0 else PASTEL_NEGATIVE) for v in vals]
    else:
        dot_colors = [get_kv_color(lbl, i) for i, lbl in enumerate(labels)]

    stem_x, stem_y = [], []
    for lbl, v in zip(labels, vals):
        stem_x += [lbl, lbl, None]
        stem_y += [0, v, None]

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=stem_x, y=stem_y, mode="lines", line=dict(color="#cbd5e1", width=2), showlegend=False, hoverinfo="skip"))
    fig.add_trace(
        go.Scatter(
            x=labels,
            y=vals,
            mode="markers+text",
            marker=dict(color=dot_colors, size=18, line=dict(color="white", width=1.5)),
            text=[f"{v:,.0f}".replace(",", ".") for v in vals],
            textposition="top center",
            textfont=dict(size=14, color="#111827", family="Segoe UI, Arial"),
            showlegend=False,
        )
    )
    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b>",
            x=0.5,
            y=0.98,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        xaxis=dict(tickfont=dict(size=14, color="#374151"), tickangle=-12),
        yaxis=dict(
            title=dict(text=y_title, font=dict(size=14.5)),
            gridcolor="#f0f2f5",
            tickfont=dict(size=14, color="#6b7280"),
            zeroline=True,
            zerolinecolor="#9ca3af",
            zerolinewidth=1.5,
        ),
        showlegend=False,
        margin=dict(l=10, r=25, t=55, b=45),
        height=300,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def build_khdn_moi_bar(kv_charts: list[dict]):
    """BẢNG D — Lollipop: Tổng SL KHDN mới theo 6 Khu vực — mỗi Khu vực 1 màu pastel (v4.13, cố định)."""
    return _build_kv_lollipop_chart(
        kv_charts, "khdn_moi_sum", "KHDN MỚI: TỔNG SL THEO 6 KHU VỰC", "Tổng SL KHDN mới"
    )


def build_khtd_tang_bar(kv_charts: list[dict]):
    """BẢNG E — Lollipop: Tổng SL KHTD tăng ròng theo 6 Khu vực (có thể âm, v4.13, cố định)."""
    return _build_kv_lollipop_chart(
        kv_charts, "khtd_tang_sum", "KHTD TĂNG RÒNG THEO 6 KHU VỰC", "Tổng SL KHTD tăng ròng", diverging=True
    )


def build_khdn_tang_bar(kv_charts: list[dict]):
    """BẢNG F — Lollipop: Tổng SL KHDN tăng ròng theo 6 Khu vực (có thể âm, v4.13, cố định)."""
    return _build_kv_lollipop_chart(
        kv_charts, "khdn_tang_sum", "KHDN TĂNG RÒNG THEO 6 KHU VỰC", "Tổng SL KHDN tăng ròng", diverging=True
    )


def build_khoi_g_bar(kv_charts: list[dict]):
    """KHỐI G — Cột dọc dương/âm: Tăng/giảm ròng cho vay theo 6 Khu vực (v4.13, đổi từ Tổng cho vay)."""
    return _build_kv_vertical_bar_chart(
        kv_charts, "tang_rong_cv_sum", "TĂNG/GIẢM RÒNG CHO VAY THEO 6 KHU VỰC", "Tăng/giảm ròng cho vay", diverging=True
    )


def build_khoi_h_bar(kv_charts: list[dict]):
    """KHỐI H — Cột dọc dương/âm: Tăng/giảm ròng huy động theo 6 Khu vực (v4.13, đổi từ Tổng huy động)."""
    return _build_kv_vertical_bar_chart(
        kv_charts, "tang_rong_hd_sum", "TĂNG/GIẢM RÒNG HUY ĐỘNG THEO 6 KHU VỰC", "Tăng/giảm ròng huy động", diverging=True
    )


def build_khoi_i_bar(kv_charts: list[dict]):
    """KHỐI I — Cột dọc dương/âm: Tổng TOI theo 6 Khu vực (v4.13, tô màu theo dấu dù TOI thường dương)."""
    return _build_kv_vertical_bar_chart(kv_charts, "tong_toi_sum", "TOI: TỔNG TOI THEO 6 KHU VỰC", "Tổng TOI", diverging=True)


def build_dvkd_combo3(
    months_labels: list[str],
    thuc_hien_vals: list[float],
    ke_hoach_vals: list[float],
    pct_ht_vals: list,
    title: str,
    dvt: str = "",
):
    """
    (v4.20) Cột kết hợp đường — 1 trong 3 dạng biểu đồ của trang 'Dashboard số liệu ĐVKD', thể hiện
    CẢ 3 số liệu cùng lúc: Cột nhóm (Thực hiện `#60a5fa` + Kế hoạch `#34d399`, trục Y chính, cùng
    ĐVT) + Đường (% hoàn thành `#fb923c`, trục Y PHỤ bên phải 0-105% — `yaxis2`, lần đầu dùng trong
    trang này vì %HT khác thang đo hẳn với 2 số liệu tiền tệ/số lượng kia).
    """
    # (v4.31) Hiển thị rõ giá trị ngay trên/trong từng cột Thực hiện/Kế hoạch (dấu chấm phân cách
    # nghìn, "—" cho tháng thiếu dữ liệu) + nhãn % ngay tại từng điểm của đường %HT — theo yêu cầu
    # người dùng cho ví dụ cụ thể "biểu đồ Số lượng nhân sự" (chính là hàm này).
    def _fmt_bar_label(v):
        return "—" if v is None else f"{v:,.0f}".replace(",", ".")

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=months_labels,
            y=thuc_hien_vals,
            name="Thực hiện",
            marker=dict(color="#60a5fa"),
            text=[_fmt_bar_label(v) for v in thuc_hien_vals],
            textposition="outside",
            textfont=dict(size=11, color="#1e40af", family="Segoe UI, Arial"),
            cliponaxis=False,
        )
    )
    fig.add_trace(
        go.Bar(
            x=months_labels,
            y=ke_hoach_vals,
            name="Kế hoạch",
            marker=dict(color="#34d399"),
            text=[_fmt_bar_label(v) for v in ke_hoach_vals],
            textposition="outside",
            textfont=dict(size=11, color="#065f46", family="Segoe UI, Arial"),
            cliponaxis=False,
        )
    )
    pct_ht_pct = [None if v is None else v * 100 for v in pct_ht_vals]
    fig.add_trace(
        go.Scatter(
            x=months_labels,
            y=pct_ht_pct,
            name="% hoàn thành",
            mode="lines+markers+text",
            line=dict(color="#fb923c", width=2.5),
            marker=dict(color="#fb923c", size=8),
            text=[None if v is None else f"{v:.0f}%" for v in pct_ht_pct],
            textposition="top center",
            textfont=dict(size=11, color="#c2410c", family="Segoe UI, Arial"),
            yaxis="y2",
            connectgaps=False,
        )
    )
    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b>",
            x=0.5,
            y=0.98,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        xaxis=dict(tickfont=dict(size=14, color="#374151")),
        yaxis=dict(ticksuffix=f" {dvt}" if dvt else "", gridcolor="#f0f2f5", tickfont=dict(size=14, color="#6b7280")),
        yaxis2=dict(overlaying="y", side="right", range=[0, 105], ticksuffix="%", tickfont=dict(size=14, color="#fb923c"), showgrid=False),
        legend=dict(orientation="h", yanchor="bottom", y=1.005, xanchor="right", x=1, font=dict(size=14)),
        margin=dict(l=10, r=35, t=55, b=35),
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
        barmode="group",
    )
    return fig


def build_dvkd_pie3(
    months_labels: list[str],
    thuc_hien_vals: list[float],
    ke_hoach_vals: list[float],
    pct_ht_vals: list,
    title: str,
    dvt: str = "",
):
    """
    (v4.20) Hình tròn — donut 2 lát Thực hiện/Kế hoạch cộng dồn qua các tháng đang xem (bỏ qua giá
    trị thiếu), annotation giữa hiển thị %HT TỔNG HỢP = tổng Thực hiện / tổng Kế hoạch — hình tròn
    không thể hiện biến động theo từng tháng (đặc thù của dạng biểu đồ này), phù hợp xem tổng quan
    tỷ trọng thay vì xu hướng.
    """
    sum_th = sum(v for v in thuc_hien_vals if v is not None)
    sum_kh = sum(v for v in ke_hoach_vals if v is not None)
    if sum_th <= 0 and sum_kh <= 0:
        fig = go.Figure()
        fig.update_layout(
            title=dict(text=f"<b>{title}</b>", x=0.5, font=dict(size=16.5, color="#111827", family="Segoe UI, Arial")),
            annotations=[dict(text="Không có dữ liệu", showarrow=False, font=dict(size=15, color="#6b7280"))],
            height=340,
            plot_bgcolor="white",
            paper_bgcolor="white",
        )
        return fig

    labels = ["Thực hiện", "Kế hoạch"]
    vals = [max(sum_th, 0), max(sum_kh, 0)]
    pct_overall = (sum_th / sum_kh) if sum_kh else None
    center_text = f"{pct_overall * 100:.1f}%" if pct_overall is not None else "—"

    fig = go.Figure(
        data=[
            go.Pie(
                labels=labels,
                values=vals,
                hole=0.55,
                marker=dict(colors=["#60a5fa", "#34d399"], line=dict(color="#ffffff", width=2)),
                textinfo="percent",
                textposition="inside",
                textfont=dict(size=15, color="#ffffff", family="Segoe UI, Arial"),
            )
        ]
    )
    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b>",
            x=0.5,
            y=0.96,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        annotations=[dict(text=f"%HT<br><b>{center_text}</b>", x=0.5, y=0.5, showarrow=False, font=dict(size=17, color="#111827", family="Segoe UI, Arial"))],
        legend=dict(orientation="h", yanchor="top", y=-0.05, xanchor="center", x=0.5, font=dict(size=13.5)),
        margin=dict(l=10, r=10, t=45, b=45),
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def build_dvkd_lollipop3(
    months_labels: list[str],
    thuc_hien_vals: list[float],
    ke_hoach_vals: list[float],
    pct_ht_vals: list,
    title: str,
    dvt: str = "",
):
    """
    (v4.20) Kẹo mút — 2 que Thực hiện/Kế hoạch lệch nhau trong từng tháng, cùng kỹ thuật offset-
    theo-nhóm của `build_multi_lollipop_trend_chart()` (trục X số + `tickvals/ticktext`). % hoàn
    thành hiển thị dạng nhãn nhỏ phía dưới trục mỗi tháng (không vẽ thành que riêng vì khác thang
    đo hẳn với 2 số liệu kia).
    """
    n = len(months_labels)
    slot_width = 0.8 / 2
    offsets = [-slot_width / 2, slot_width / 2]
    series = [("Thực hiện", thuc_hien_vals, "#60a5fa"), ("Kế hoạch", ke_hoach_vals, "#34d399")]

    fig = go.Figure()
    for (name, vals, color), offset in zip(series, offsets):
        xs = [i + offset for i in range(n)]
        stem_x, stem_y = [], []
        for x, v in zip(xs, vals):
            stem_x += [x, x, None]
            stem_y += [0, v, None]
        fig.add_trace(go.Scatter(x=stem_x, y=stem_y, mode="lines", line=dict(color=color, width=1.5), opacity=0.6, showlegend=False, hoverinfo="skip"))
        # (v4.31) Hiển thị rõ giá trị ngay tại từng đầu que (dấu chấm phân cách nghìn) — theo yêu cầu
        # người dùng hiển thị số liệu chi tiết ngay trong biểu đồ kẹo mút.
        fig.add_trace(
            go.Scatter(
                x=xs,
                y=vals,
                mode="markers+text",
                name=name,
                marker=dict(color=color, size=12, line=dict(color="white", width=1)),
                text=[None if v is None else f"{v:,.0f}".replace(",", ".") for v in vals],
                textposition="top center",
                textfont=dict(size=10.5, color=color, family="Segoe UI, Arial"),
            )
        )

    annotations = [
        dict(
            x=i,
            y=-0.14,
            xref="x",
            yref="paper",
            text=(f"{v * 100:.0f}%" if v is not None else "—"),
            showarrow=False,
            font=dict(size=13, color="#fb923c", family="Segoe UI, Arial"),
        )
        for i, v in enumerate(pct_ht_vals)
    ]

    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b>",
            x=0.5,
            y=0.96,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        xaxis=dict(tickmode="array", tickvals=list(range(n)), ticktext=months_labels, tickfont=dict(size=14, color="#374151")),
        yaxis=dict(
            ticksuffix=f" {dvt}" if dvt else "",
            gridcolor="#f0f2f5",
            tickfont=dict(size=14, color="#6b7280"),
            zeroline=True,
            zerolinecolor="#9ca3af",
            zerolinewidth=1.5,
        ),
        legend=dict(orientation="h", yanchor="bottom", y=1.005, xanchor="right", x=1, font=dict(size=14)),
        annotations=annotations,
        margin=dict(l=10, r=20, t=55, b=75),
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def build_dvkd_summary_table(dvkd_items_scoped: list[dict], sel_months: list[int]):
    """
    (v4.20, mở rộng v4.21) Bảng thống kê TOÀN BỘ chỉ tiêu của 1 Đơn vị kinh doanh, mirror đúng layout
    hàng/cột của file Excel gốc — mỗi dòng 1 chỉ tiêu (đúng thứ tự xuất hiện gốc trong file, không
    sắp lại), cột MultiIndex 2 tầng (nhóm số liệu, tháng/"Cả năm"). Giữ NGUYÊN số liệu đã parse —
    không tính lại gì khác. Chèn thêm các dòng tiêu đề SECTION (sẽ tô đậm khi style) mỗi khi gặp
    `section` mới trong lúc duyệt tuần tự, giống hệt các dòng tiêu đề nhóm có trong file Excel gốc
    (vd "A. Chỉ tiêu Chính"). Trả về `(DataFrame, set các vị trí dòng 0-based là dòng tiêu đề section)`.

    (v4.21) Thêm cột đơn `("", "STT")`/`("", "Chu kỳ")` (từ field `stt_display`/`chu_ky` của item —
    xem `load_and_preprocess_dvkd_excel()`), và nhóm cột `"Tăng/Giảm"` mỗi tháng = Thực hiện tháng
    đang xét TRỪ Thực hiện tháng liền trước (`—` nếu 1 trong 2 tháng không có dữ liệu). Số format
    theo chuẩn Việt Nam (dấu `.` phân cách nghìn, dấu `,` phân cách thập phân — vd `10.776,40`).

    (v4.22) Thêm 2 nhóm cột CUỐI CÙNG (sau các nhóm theo tháng, KHÔNG lặp theo `sel_months` — chỉ
    xuất hiện ĐÚNG 1 lần mỗi bảng, vì nguồn là 4 giá trị đơn của file gốc, không phải dict theo
    tháng): `"SO VỚI KH THÁNG"` (`"Todo tháng"` từ `todo_thang`, `"% hoàn thành"` từ `pct_ht_thang`)
    và `"SO VỚI KH NĂM"` (`"Todo năm"` từ `todo_nam`, `"% hoàn thành/năm"` từ `pct_ht_nam2`).

    (v4.23) ĐẢO thứ tự tuple key của các cột theo tháng: level 0 = `"THÁNG {m:02d}/2026"` (gom
    nhóm), level 1 = loại cột (Thực hiện/Kế hoạch/% hoàn thành/Tăng/Giảm) — TRƯỚC đó level 0 là loại
    cột (lặp lại xuyên suốt mọi tháng), level 1 là tháng, khiến `style_dvkd_summary_table()` tô màu
    header cấp 1 theo LOẠI CỘT thay vì theo TỪNG THÁNG như mong muốn thực tế (`Styler.map_index()`
    chỉ tô theo level 0). Cột "Cả năm" tách thành nhóm CỐ ĐỊNH riêng `"KH NĂM"` (2 sub-cột "Kế
    hoạch"/"% hoàn thành") — không còn nằm lẫn dưới nhóm "Kế hoạch"/"% hoàn thành" cũ (2 tên đó giờ
    là SUB-CỘT lặp lại theo từng tháng, không còn là tên group cố định nữa). 2 nhóm `"SO VỚI KH
    THÁNG"`/`"SO VỚI KH NĂM"` (v4.22) GIỮ NGUYÊN cấu trúc `(tên_nhóm, sub_cột)` — không đảo — nên
    MultiIndex cột cuối cùng là HỖN HỢP: level 0 vừa có thể là tên tháng ("THÁNG 03/2026"...) vừa có
    thể là tên nhóm cố định ("KH NĂM"/"SO VỚI KH THÁNG"/"SO VỚI KH NĂM") tuỳ cột — xem
    `_header_group_color()` trong `style_dvkd_summary_table()` xử lý cả 2 trường hợp.
    """

    def _fmt_num(v):
        if v is None:
            return "—"
        formatted = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"-{formatted}" if v < 0 else formatted

    def _fmt_pct(v):
        return "—" if v is None else f"{v * 100:.1f}%"

    rows = []
    section_row_positions = set()
    current_section = None
    for it in dvkd_items_scoped:
        sec = it.get("section") or ""
        if sec and sec != current_section:
            current_section = sec
            section_row_positions.add(len(rows))
            rows.append({("", "STT"): "", ("", "Chỉ tiêu"): sec, ("", "ĐVT"): "", ("", "Chu kỳ"): ""})
        row = {
            ("", "STT"): it.get("stt_display", ""),
            ("", "Chỉ tiêu"): it["label"],
            ("", "ĐVT"): it.get("dvt", ""),
            ("", "Chu kỳ"): it.get("chu_ky", ""),
        }
        for m in sel_months:
            # (v4.23) Đảo thứ tự tuple key: level 0 = THÁNG (gom nhóm), level 1 = loại cột — để
            # style_dvkd_summary_table() tô màu header cấp 1 theo TỪNG THÁNG (map_index bám level 0)
            # thay vì theo loại cột lặp lại xuyên suốt mọi tháng như trước (v4.20-v4.22).
            month_label = f"THÁNG {m:02d}/2026"
            th_val = it["thuc_hien"].get(m)
            th_prev = it["thuc_hien"].get(m - 1)
            row[(month_label, "Thực hiện")] = _fmt_num(th_val)
            row[(month_label, "Kế hoạch")] = _fmt_num(it["ke_hoach"].get(m))
            row[(month_label, "% hoàn thành")] = _fmt_pct(it["pct_ht"].get(m))
            row[(month_label, "Tăng/Giảm")] = _fmt_num(th_val - th_prev) if (th_val is not None and th_prev is not None) else "—"
        # (v4.23) Cả năm tách thành nhóm CỐ ĐỊNH riêng "KH NĂM" (không còn nằm dưới "Kế hoạch"/"%
        # hoàn thành" như trước, vì 2 tên đó giờ là SUB-CỘT theo tháng, không còn là group cố định).
        row[("KH NĂM", "Kế hoạch")] = _fmt_num(it.get("ke_hoach_nam"))
        row[("KH NĂM", "% hoàn thành")] = _fmt_pct(it.get("pct_ht_nam"))
        # (v4.22) 2 nhóm cột cuối cùng — So với Kế hoạch Tháng/Năm — dùng 4 giá trị ĐƠN (không lặp
        # theo sel_months, xem ghi chú field trong load_and_preprocess_dvkd_excel()).
        row[("SO VỚI KH THÁNG", "Todo tháng")] = _fmt_num(it.get("todo_thang"))
        row[("SO VỚI KH THÁNG", "% hoàn thành")] = _fmt_pct(it.get("pct_ht_thang"))
        row[("SO VỚI KH NĂM", "Todo năm")] = _fmt_num(it.get("todo_nam"))
        row[("SO VỚI KH NĂM", "% hoàn thành/năm")] = _fmt_pct(it.get("pct_ht_nam2"))
        rows.append(row)

    df = pd.DataFrame(rows)
    df.columns = pd.MultiIndex.from_tuples(df.columns)
    df = df.fillna("")
    return df, section_row_positions


def style_dvkd_summary_table(df: pd.DataFrame, section_row_positions: set):
    """
    (v4.20, mở rộng v4.21) Tô màu bảng thống kê ĐVKD theo NHÓM CỘT — Thực hiện = xanh dương nhạt
    (`#eff6ff`), Kế hoạch = xanh lá nhạt (`#ecfdf5`), Tăng/Giảm = tím nhạt (`#f3e8ff`, mới v4.21,
    cùng tông với màu header tím của nhóm này) — nhất quán với bảng màu 3 thẻ KPI Pill trước v4.20.
    Dòng tiêu đề section (v4.21: đổi sang nền xanh dương nhạt + viền trên đậm, rõ hơn bản v4.20 cũ).
    Theo đúng pattern `df.style.apply` hàng-theo-hàng đã dùng ở `format_and_style()`/
    `style_unit_detail_table()`.

    (v4.21) % hoàn thành đổi từ tô 1 màu cam cố định sang tô theo NGƯỠNG: xanh lá nếu ≥100%, vàng
    nếu ≥80%, đỏ nếu <80% (parse lại số từ chuỗi đã format, chấp nhận cả dấu `,` lẫn `.` thập phân —
    xem `_pct_cell_style()`); ô "—" (thiếu dữ liệu) fallback về cam nhạt như cũ. Header cấp 1 (nhóm
    cột: Thực hiện/Kế hoạch/% hoàn thành/Tăng/Giảm) tô đậm theo màu riêng từng nhóm qua
    `Styler.map_index()` (không dùng `nth-child` tĩnh vì số cột nhóm-tháng thay đổi theo `sel_months`
    đang chọn, khiến vị trí cột không cố định — `map_index` bám theo TÊN nhóm nên luôn đúng bất kể số
    tháng); header cấp 2 (tên tháng/"Cả năm") và ô dữ liệu style qua `set_table_styles()`.

    (v4.22) Thêm màu cho 2 nhóm cột mới `"SO VỚI KH THÁNG"` (cyan đậm `#0369a1` header, `#e0f2fe`
    nền ô cho sub-cột "Todo tháng") và `"SO VỚI KH NĂM"` (vàng/cam đậm `#b45309` header, `#fef3c7`
    nền ô cho sub-cột "Todo năm") — riêng sub-cột "% hoàn thành"/"% hoàn thành/năm" trong 2 nhóm này
    VẪN tô theo NGƯỠNG (dùng chung `_pct_cell_style()`), không dùng màu cố định của cả nhóm.

    (v4.23) `build_dvkd_summary_table()` đảo tuple key các cột theo tháng thành `(THÁNG, loại cột)`
    (trước là `(loại cột, THÁNG)`) để `map_index(level=0)` tô header theo TỪNG THÁNG thay vì theo
    loại cột — kéo theo `highlight_row()` PHẢI đổi cách nhận diện loại cột để tô màu Ô DỮ LIỆU (giờ
    đọc `col[1]` thay vì `col[0]` cho các cột theo tháng) và `_header_group_color()` PHẢI đổi từ tra
    dict cố định sang: (a) tra dict cố định CHỈ cho 3 nhóm không theo tháng (`"KH NĂM"`, `"SO VỚI KH
    THÁNG"`, `"SO VỚI KH NĂM"`, và `""` cho các cột đơn STT/Chỉ tiêu/ĐVT/Chu kỳ), (b) parse số tháng
    từ chuỗi `"THÁNG {mm}/2026"` rồi tô XEN KẼ theo 1 bảng 12 màu cố định (không còn theo TÊN loại
    cột như trước, vì level 0 giờ là tháng chứ không phải loại cột nữa).

    (v4.24) Header cấp 2 (loại cột: Thực hiện/Kế hoạch/% hoàn thành/Tăng-Giảm/Todo...) đổi từ
    `set_table_styles({"selector": "th.col_heading.level1"})` sang `Styler.map_index(level=1)`
    (`_header_sub_color()`) — selector CSS `th.col_heading.level1` ở một số phiên bản pandas/
    Streamlit khớp NHẦM luôn cả header cấp 1 (`th.col_heading.level0`), ghi đè mất màu vừa tô qua
    `map_index(level=0)` phía trên. `map_index` áp trực tiếp style theo TỪNG LABEL của đúng 1 level,
    không dùng CSS selector nên không có rủi ro khớp chéo giữa 2 cấp header.
    """

    def _pct_cell_style(raw_value) -> str:
        raw = str(raw_value).replace("%", "").replace(",", ".").replace("—", "").strip()
        try:
            pct = float(raw)
        except ValueError:
            return "background-color:#fff7ed; color:#9a3412;"
        if pct >= 100:
            return "background-color:#dcfce7; color:#166534; font-weight:700;"
        elif pct >= 80:
            return "background-color:#fef9c3; color:#854d0e; font-weight:700;"
        else:
            return "background-color:#fee2e2; color:#991b1b; font-weight:700;"

    def highlight_row(row):
        if row.name in section_row_positions:
            return ["background-color:#dbeafe; font-weight:800; color:#1e3a8a; border-top: 2px solid #1e40af;"] * len(row)
        styles = []
        for col in row.index:
            level0 = col[0] if isinstance(col, tuple) else ""
            level1 = col[1] if isinstance(col, tuple) and len(col) > 1 else ""
            # (v4.23) 3 nhóm KHÔNG theo tháng vẫn giữ cấu trúc (tên_nhóm, sub_cột) cũ — nhận diện
            # qua level0 (tên nhóm) TRƯỚC, y hệt logic v4.22, độc lập với việc đảo trục ở các cột
            # theo tháng bên dưới.
            if level0 == "SO VỚI KH THÁNG":
                styles.append(_pct_cell_style(row[col]) if level1 == "% hoàn thành" else "background-color:#e0f2fe; color:#0369a1;")
            elif level0 == "SO VỚI KH NĂM":
                styles.append(_pct_cell_style(row[col]) if level1 == "% hoàn thành/năm" else "background-color:#fef3c7; color:#b45309;")
            elif level0 == "KH NĂM":
                styles.append(_pct_cell_style(row[col]) if level1 == "% hoàn thành" else "background-color:#ecfdf5; color:#065f46;")
            # Cột theo tháng: (THÁNG, loại cột) — nhận diện qua level1 (loại cột) vì level0 giờ là tên tháng.
            elif level1 == "Thực hiện":
                styles.append("background-color:#eff6ff; color:#1e40af;")
            elif level1 == "Kế hoạch":
                styles.append("background-color:#ecfdf5; color:#065f46;")
            elif level1 == "Tăng/Giảm":
                styles.append("background-color:#f3e8ff; color:#6b21a8;")
            elif level1 == "% hoàn thành":
                styles.append(_pct_cell_style(row[col]))
            else:
                styles.append("background-color:#ffffff;")
        return styles

    # (v4.23) 12 màu xen kẽ theo số tháng (khác hẳn v4.20-v4.22 tô theo TÊN loại cột) — vì level 0
    # giờ là "THÁNG {mm}/2026", CÙNG 1 tháng phải ra CÙNG 1 màu cho mọi sub-cột của nó, và mỗi tháng
    # kế tiếp đổi màu để phân biệt trực quan giữa các khối tháng liền kề trên header.
    _MONTH_HEADER_COLORS = [
        "#1e40af", "#065f46", "#7c3aed", "#b45309", "#be185d",
        "#0e7490", "#166534", "#9a3412", "#1e3a8a", "#6b21a8", "#134e4a", "#7f1d1d",
    ]

    def _header_group_color(colname: str) -> str:
        fixed = {
            "KH NĂM": "#166534",
            "SO VỚI KH THÁNG": "#0369a1",
            "SO VỚI KH NĂM": "#b45309",
            "": "#334155",
        }
        if colname in fixed:
            bg = fixed[colname]
        else:
            try:
                m = int(colname.split(" ")[1].split("/")[0])  # "THÁNG 05/2026" -> 5
                bg = _MONTH_HEADER_COLORS[(m - 1) % len(_MONTH_HEADER_COLORS)]
            except (IndexError, ValueError):
                bg = "#1e40af"
        return f"background-color:{bg}; color:white; font-weight:700; text-align:center;"

    def _header_sub_color(colname: str) -> str:
        # (v4.24) Tô header cấp 2 (loại cột: Thực hiện/Kế hoạch/% hoàn thành/Tăng-Giảm/Todo...) qua
        # `map_index(level=1)` thay vì `set_table_styles({"selector": "th.col_heading.level1"})` —
        # selector đó ở một số phiên bản pandas/Streamlit khớp NHẦM luôn cả `th.col_heading.level0`
        # (do CSS class list `col_heading level0 col3` chứa cả "level1" nếu chọn sai combinator),
        # ghi đè mất màu vừa tô cho header cấp 1 qua `map_index(level=0)` ở trên.
        sub_map = {
            "Thực hiện": "background-color:#bfdbfe; color:#1e3a8a; font-weight:600; font-size:11px;",
            "Kế hoạch": "background-color:#bbf7d0; color:#14532d; font-weight:600; font-size:11px;",
            "% hoàn thành": "background-color:#fed7aa; color:#7c2d12; font-weight:600; font-size:11px;",
            "Tăng/Giảm": "background-color:#ddd6fe; color:#4c1d95; font-weight:600; font-size:11px;",
            "Todo tháng": "background-color:#bae6fd; color:#0c4a6e; font-weight:600; font-size:11px;",
            "Todo năm": "background-color:#fde68a; color:#78350f; font-weight:600; font-size:11px;",
            "% hoàn thành/năm": "background-color:#fed7aa; color:#7c2d12; font-weight:600; font-size:11px;",
        }
        return sub_map.get(colname, "background-color:#f1f5f9; color:#334155; font-size:11px;")

    styled = df.style.apply(highlight_row, axis=1)
    styled = styled.map_index(_header_group_color, level=0, axis=1)
    styled = styled.map_index(_header_sub_color, level=1, axis=1)
    styled = styled.set_table_styles(
        [
            {"selector": "th.row_heading", "props": [("display", "none")]},
            {"selector": "td", "props": [("font-size", "12px"), ("padding", "4px 8px")]},
        ]
    )
    return styled


def render_dvkd_table_html(df: pd.DataFrame, section_row_positions: set) -> str:
    """
    (v4.25) Render bảng thống kê ĐVKD thành HTML THUẦN (dùng với `st.markdown(..., unsafe_allow_html=True)`)
    — thay thế hoàn toàn việc gọi `st.dataframe(style_dvkd_summary_table(...))`.

    Lý do: `st.dataframe()` chỉ áp dụng phần tô màu Ô DỮ LIỆU của `Styler` (`df.style.apply(...)`),
    **hoàn toàn bỏ qua** mọi cách tô màu HEADER (`<th>`) của Styler — dù là `set_table_styles()` hay
    `map_index()` — đây là GIỚI HẠN CỨNG của `st.dataframe()`, không phải lỗi ở cách gọi API hay
    version pandas/Streamlit như 2 lần fix trước (v4.23, v4.24) từng phỏng đoán. Cả 2 lần đó đều tô
    ĐÚNG theo đúng API Styler chính thức nhưng vẫn không hiện màu header khi xem qua `st.dataframe()`
    vì bản thân `st.dataframe()` chưa từng render `<th>` style từ Styler ra DOM. Giải pháp triệt để
    duy nhất: tự dựng bảng `<table>` HTML thuần (2 dòng `<thead>`: dòng 1 = tên nhóm/tháng với
    `colspan` theo đúng số cột con, dòng 2 = tên loại cột) rồi `st.markdown(html, unsafe_allow_html=True)`.

    `style_dvkd_summary_table()`/`build_dvkd_summary_table()` GIỮ NGUYÊN, không xoá — hàm này chỉ
    thay thế bước RENDER CUỐI CÙNG, tự đọc trực tiếp DataFrame thô (`df`, không phải bản đã `.style`)
    và tự tô màu lại từ đầu bằng HTML/CSS inline (logic màu giữ đúng y hệt `style_dvkd_summary_table()`
    — 2 hàm này hiện trùng lặp logic màu, chấp nhận được vì `style_dvkd_summary_table()` được giữ lại
    có chủ đích cho khả năng tái dùng sau này, vd xuất Excel qua `Styler` — nơi ĐÓ `st.dataframe()`
    không phải lớp trung gian nên giới hạn trên không áp dụng).
    """
    MONTH_COLORS = [
        "#1e40af", "#065f46", "#7c3aed", "#b45309", "#be185d",
        "#0e7490", "#166534", "#9a3412", "#1e3a8a", "#6b21a8", "#134e4a", "#7f1d1d",
    ]
    FIXED_COLORS = {
        "KH NĂM": "#166534",
        "SO VỚI KH THÁNG": "#0369a1",
        "SO VỚI KH NĂM": "#b45309",
        "": "#334155",
    }
    SUB_COLORS = {
        "Thực hiện": "#bfdbfe",
        "Kế hoạch": "#bbf7d0",
        "% hoàn thành": "#fed7aa",
        "Tăng/Giảm": "#ddd6fe",
        "Todo tháng": "#bae6fd",
        "Todo năm": "#fde68a",
        "% hoàn thành/năm": "#fed7aa",
    }

    level0 = df.columns.get_level_values(0)
    level1 = df.columns.get_level_values(1)

    # Gom nhóm các cột liền kề cùng level0 thành (tên_nhóm, colspan, vị_trí_bắt_đầu)
    groups = []
    prev = None
    for i, g in enumerate(level0):
        if g != prev:
            groups.append([g, 1, i])
            prev = g
        else:
            groups[-1][1] += 1

    month_idx = 0
    group_colors = {}
    for gname, _span, _start in groups:
        if gname in FIXED_COLORS:
            group_colors[gname] = FIXED_COLORS[gname]
        elif gname.startswith("THÁNG"):
            group_colors[gname] = MONTH_COLORS[month_idx % len(MONTH_COLORS)]
            month_idx += 1
        else:
            group_colors[gname] = "#334155"

    html = ['<div style="overflow-x:auto; max-height:620px; overflow-y:auto;">']
    html.append('<table style="border-collapse:collapse; width:100%; font-size:12px; font-family:Arial,sans-serif;">')

    html.append("<thead>")
    html.append('<tr style="position:sticky;top:0;z-index:3;">')
    for gname, span, _start in groups:
        color = group_colors.get(gname, "#334155")
        html.append(
            f'<th colspan="{span}" style="background-color:{color};color:white;font-weight:700;'
            f'text-align:center;padding:6px 8px;border:1px solid rgba(255,255,255,0.3);">{gname}</th>'
        )
    html.append("</tr>")

    html.append('<tr style="position:sticky;top:33px;z-index:3;">')
    for col_l1 in level1:
        bg = SUB_COLORS.get(col_l1, "#f1f5f9")
        html.append(
            f'<th style="background-color:{bg};color:#334155;font-weight:600;font-size:11px;'
            f'text-align:center;padding:4px 6px;border:1px solid #e2e8f0;min-width:80px;white-space:nowrap;">{col_l1}</th>'
        )
    html.append("</tr></thead>")

    html.append("<tbody>")
    for row_idx, row in df.iterrows():
        is_section = row_idx in section_row_positions
        row_bg = "background-color:#dbeafe;" if is_section else ""
        row_fw = "font-weight:800;" if is_section else ""
        html.append(f'<tr style="{row_bg}{row_fw}">')
        for col, val in row.items():
            col_l1 = col[1] if isinstance(col, tuple) and len(col) > 1 else ""
            cell_style = "padding:3px 8px;border:1px solid #e8edf2;white-space:nowrap;"
            if "%" in col_l1 and "hoàn thành" in col_l1:
                raw = str(val).replace("%", "").replace(",", ".").replace("—", "").strip()
                try:
                    pct = float(raw)
                    if pct >= 100:
                        cell_style += "background-color:#dcfce7;color:#166534;font-weight:700;"
                    elif pct >= 80:
                        cell_style += "background-color:#fef9c3;color:#854d0e;font-weight:700;"
                    else:
                        cell_style += "background-color:#fee2e2;color:#991b1b;font-weight:700;"
                except ValueError:
                    pass
            elif col_l1 == "Thực hiện":
                cell_style += "color:#1e40af;"
            elif col_l1 == "Kế hoạch":
                cell_style += "color:#065f46;"
            elif col_l1 == "Tăng/Giảm":
                cell_style += "color:#6b21a8;font-style:italic;"
            html.append(f'<td style="{cell_style}">{val}</td>')
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "".join(html)


def _map_dvkd_chart_type_3(chart_type5: str) -> str:
    """
    (v4.20) Thu gọn 5 loại `chart_type` gốc (`_assign_chart_type()`, vẫn GIỮ NGUYÊN không đổi) về
    còn 3 dạng cho biểu đồ trang 'Dashboard số liệu ĐVKD' (combo/pie/lollipop) — chỉ ánh xạ tại thời
    điểm vẽ, không đụng vào field `chart_type` đã lưu trong data.
    """
    return {"line": "combo", "bar": "combo", "hbar": "lollipop", "area": "pie", "combo": "combo"}.get(chart_type5, "combo")


DVKD_CHART_BUILDERS_3 = {
    "combo": build_dvkd_combo3,
    "pie": build_dvkd_pie3,
    "lollipop": build_dvkd_lollipop3,
}


def build_dvkd_chart_by_type3(
    chart_type3: str,
    months_labels: list[str],
    thuc_hien_vals: list[float],
    ke_hoach_vals: list[float],
    pct_ht_vals: list,
    title: str,
    dvt: str = "",
):
    """(v4.20) Dispatcher: chọn 1 trong 3 hàm vẽ biểu đồ trang 'Dashboard số liệu ĐVKD' theo `chart_type3` (xem `_map_dvkd_chart_type_3()`)."""
    builder = DVKD_CHART_BUILDERS_3.get(chart_type3, build_dvkd_combo3)
    return builder(months_labels, thuc_hien_vals, ke_hoach_vals, pct_ht_vals, title, dvt)


# Ánh xạ lựa chọn Slicer KHỐI CHỈ TIÊU -> hàm dựng biểu đồ tương ứng (dùng cho biểu đồ tiêu điểm ở Tab 1)
CHI_TIEU_CHART_MAP = {
    "Tỷ lệ đạt KPIs": build_kv_kpi_bar,
    "Tỷ lệ thỏa LPT": build_kv_lpt_bar,
    "Phân bổ dải LPT thực nhận": build_lpt_distribution_chart,
    "KHDN mới/ĐVKD": build_khdn_moi_bar,
    "KHTD tăng ròng/ĐVKD": build_khtd_tang_bar,
    "KHDN tăng ròng/ĐVKD": build_khdn_tang_bar,
    "Chi tiết dư nợ": build_khoi_g_bar,
    "Chi tiết huy động": build_khoi_h_bar,
    "Chi tiết TOI": build_khoi_i_bar,
}

# ──────────────────────────────────────────────────────────────────────────────
# XUẤT BÁO CÁO EXCEL TRỌN BỘ 6 SHEET
# ──────────────────────────────────────────────────────────────────────────────


def export_blocks_to_excel(blocks: dict[str, pd.DataFrame]) -> bytes:
    """Xuất file Excel chuẩn hóa, mỗi sheet tương ứng 1 khối chỉ tiêu trong dict `blocks` truyền vào."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, block_df in blocks.items():
            flat_df = block_df.copy()
            if isinstance(flat_df.columns, pd.MultiIndex):
                flat_df.columns = [f"{c[0]} | {c[1]}" for c in flat_df.columns]
            flat_df.to_excel(writer, sheet_name=sheet_name[:31])
            ws = writer.sheets[sheet_name[:31]]
            for col_cells in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col_cells)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 4, 14), 40)
    return output.getvalue()


# ──────────────────────────────────────────────────────────────────────────────
# DRILL-DOWN: CHI TIẾT ĐƠN VỊ KINH DOANH (MA_DV, TEN_DV) THEO KHU VỰC ĐANG CHỌN
# ──────────────────────────────────────────────────────────────────────────────


# Ánh xạ lựa chọn Slicer KHỐI CHỈ TIÊU -> metadata (cột số liệu, tiêu đề, màu, có âm/dương hay không)
# trong bảng/biểu đồ drill-down Đơn vị kinh doanh — dùng để sắp xếp, highlight cột và dựng biểu đồ
# theo MA_DV khi kết hợp bộ lọc KHỐI CHỈ TIÊU với KHU VỰC.
# (v4.13) "chart_type" nay CỐ ĐỊNH theo từng khối chỉ tiêu — GIỐNG HỆT loại biểu đồ dùng ở cấp Khu
# vực (CHI_TIEU_CHART_MAP) và cấp theo tháng (CHI_TIEU_TREND_META), bất kể đang lọc Toàn Hàng/Khu
# vực/ĐVKD/Nhân sự/tháng nào: "hbar" (cột ngang — Khối A/B), "pie" (donut tỷ trọng — Khối C),
# "lollipop" (kẹo mút — Khối D/E/F), "vbar" + "diverging": True (cột dọc dương/âm — Khối G/H/I).
CHI_TIEU_UNIT_META = {
    "Tỷ lệ đạt KPIs": {"column": "% đạt KPIs", "title": "TỶ LỆ ĐẠT KPIs", "color": "#60a5fa", "diverging": False, "chart_type": "hbar"},
    "Tỷ lệ thỏa LPT": {"column": "% thỏa LPT", "title": "TỶ LỆ THỎA LƯƠNG PHỤ TRỘI", "color": "#34d399", "diverging": False, "chart_type": "hbar"},
    "Phân bổ dải LPT thực nhận": {"column": "Hệ số LPT TB (%)", "title": "HỆ SỐ LPT THỰC NHẬN BÌNH QUÂN", "color": "#fb923c", "diverging": False, "chart_type": "pie"},
    "KHDN mới/ĐVKD": {"column": "SL KHDN mới", "title": "SL KHDN MỚI", "color": "#f472b6", "diverging": False, "chart_type": "lollipop"},
    "KHTD tăng ròng/ĐVKD": {"column": "SL KHTD tăng ròng", "title": "SL KHTD TĂNG RÒNG", "color": "#60a5fa", "diverging": True, "chart_type": "lollipop"},
    "KHDN tăng ròng/ĐVKD": {"column": "SL KHDN tăng ròng", "title": "SL KHDN TĂNG RÒNG", "color": "#60a5fa", "diverging": True, "chart_type": "lollipop"},
    "Chi tiết dư nợ": {"column": "Tăng ròng CV", "title": "TĂNG/GIẢM RÒNG CHO VAY (ĐVKD)", "color": "#34d399", "diverging": True, "chart_type": "vbar"},
    "Chi tiết huy động": {"column": "Tăng ròng HĐ", "title": "TĂNG/GIẢM RÒNG HUY ĐỘNG (ĐVKD)", "color": "#fbbf24", "diverging": True, "chart_type": "vbar"},
    "Chi tiết TOI": {"column": "Tổng TOI", "title": "TỔNG TOI (ĐVKD)", "color": "#a78bfa", "diverging": True, "chart_type": "vbar"},
}
DEFAULT_UNIT_META = {"column": "% đạt KPIs", "title": "TỶ LỆ ĐẠT KPIs", "color": "#60a5fa", "diverging": False, "chart_type": "hbar"}


def get_unit_detail_table(
    df: pd.DataFrame,
    kv_filter,
    thang_filter,
    chi_tieu_filter: str = "",
    ma_dv_filter: Optional[list] = None,
    ns_filter: Optional[list] = None,
) -> pd.DataFrame:
    """
    Group theo (TEN_KV, MA_DV, TEN_DV) trong phạm vi (các) Khu vực & kỳ tháng đang chọn — luôn kèm
    cột Khu vực để xem được ĐVKD thuộc khu vực nào kể cả khi không chọn riêng Khu vực nào.
    Nếu có chọn KHỐI CHỈ TIÊU, sắp xếp ưu tiên theo cột chỉ tiêu đó (giảm dần) để "kết hợp" 2 bộ lọc.
    Nếu có `ma_dv_filter`/`ns_filter` (đa lựa chọn ĐVKD/Nhân sự), chỉ tính cho phạm vi đó.
    `kv_filter`/`thang_filter` là list (đa lựa chọn, v4.12).
    Trả về DataFrame phẳng kèm 1 dòng "Tổng" cuối bảng.
    """
    df_thang, _, _ = apply_thang_filter(df, thang_filter)
    df_kv = apply_nhan_su_filter(apply_ma_dv_filter(apply_kv_filter(df_thang, kv_filter), ma_dv_filter), ns_filter)

    if df_kv.empty:
        return pd.DataFrame()

    rows = []
    for (ten_kv, ma_dv, ten_dv), sub in df_kv.groupby(["TEN_KV", "MA_DV", "TEN_DV"], sort=True):
        tot = len(sub)
        dat = len(sub[sub["KET_QUA_KPIS"] == "DAT_KPIS"])
        thoa = len(sub[sub["KET_QUA_LPT"] == "THOA_LPT"])
        lpt_vals = sub["HE_SO_LPT_RAW"].dropna()
        lpt_tb = round(lpt_vals.mean() * 100, 1) if not lpt_vals.empty else None
        rows.append({
            "Khu vực": ten_kv,
            "Mã ĐVKD": ma_dv,
            "Tên ĐVKD": ten_dv,
            "Tổng NS đánh giá": tot,
            "SL đạt KPIs": dat,
            "% đạt KPIs": round(dat / tot * 100, 1) if tot > 0 else 0.0,
            "SL thỏa LPT": thoa,
            "% thỏa LPT": round(thoa / tot * 100, 1) if tot > 0 else 0.0,
            "Hệ số LPT TB (%)": lpt_tb,
            "SL KHDN mới": int(sub["SL_KHDN_MOI_N"].sum()),
            "SL KHDN tăng ròng": int(sub["SL_KHDN_TANG_RONG"].sum()),
            "SL KHTD tăng ròng": int(sub["SL_KH_TD_TANG_RONG"].sum()),
            # Khối G/H/I: giá trị cấp ĐVKD — dùng agg_unit_level_metric (dedup theo SYM_RUN_DATE)
            "Tổng CV": round(agg_unit_level_metric(sub, "TONG_CV_THANG_N"), 2),
            "Tổng CV N-1": round(agg_unit_level_metric(sub, "TONG_CV_THANG_N_1"), 2),
            "Tăng ròng CV": round(agg_unit_level_metric(sub, "TANG_RONG_CV"), 2),
            "Tổng HĐ": round(agg_unit_level_metric(sub, "TONG_HD_THANG_N"), 2),
            "Tổng HĐ N-1": round(agg_unit_level_metric(sub, "TONG_HD_THANG_N_1"), 2),
            "Tăng ròng HĐ": round(agg_unit_level_metric(sub, "TANG_RONG_HD"), 2),
            "Tổng TOI": round(sum(agg_unit_level_metric(sub, c) for c in TOI_SOURCE_COLS.values()), 2),
        })

    res = pd.DataFrame(rows)

    focus_col = CHI_TIEU_UNIT_META.get(chi_tieu_filter, {}).get("column")
    if focus_col and focus_col in res.columns:
        res = res.sort_values(focus_col, ascending=False, na_position="last").reset_index(drop=True)
    else:
        res = res.sort_values(["Khu vực", "Tên ĐVKD"]).reset_index(drop=True)

    tot_ns = int(res["Tổng NS đánh giá"].sum())
    tot_dat = int(res["SL đạt KPIs"].sum())
    tot_thoa = int(res["SL thỏa LPT"].sum())
    lpt_vals_all = df_kv["HE_SO_LPT_RAW"].dropna()
    total_row = {
        "Khu vực": "",
        "Mã ĐVKD": "",
        "Tên ĐVKD": "Tổng",
        "Tổng NS đánh giá": tot_ns,
        "SL đạt KPIs": tot_dat,
        "% đạt KPIs": round(tot_dat / tot_ns * 100, 1) if tot_ns > 0 else 0.0,
        "SL thỏa LPT": tot_thoa,
        "% thỏa LPT": round(tot_thoa / tot_ns * 100, 1) if tot_ns > 0 else 0.0,
        "Hệ số LPT TB (%)": round(lpt_vals_all.mean() * 100, 1) if not lpt_vals_all.empty else None,
        "SL KHDN mới": int(res["SL KHDN mới"].sum()),
        "SL KHDN tăng ròng": int(res["SL KHDN tăng ròng"].sum()),
        "SL KHTD tăng ròng": int(res["SL KHTD tăng ròng"].sum()),
        "Tổng CV": round(agg_unit_level_metric(df_kv, "TONG_CV_THANG_N"), 2),
        "Tổng CV N-1": round(agg_unit_level_metric(df_kv, "TONG_CV_THANG_N_1"), 2),
        "Tăng ròng CV": round(agg_unit_level_metric(df_kv, "TANG_RONG_CV"), 2),
        "Tổng HĐ": round(agg_unit_level_metric(df_kv, "TONG_HD_THANG_N"), 2),
        "Tổng HĐ N-1": round(agg_unit_level_metric(df_kv, "TONG_HD_THANG_N_1"), 2),
        "Tăng ròng HĐ": round(agg_unit_level_metric(df_kv, "TANG_RONG_HD"), 2),
        "Tổng TOI": round(sum(agg_unit_level_metric(df_kv, c) for c in TOI_SOURCE_COLS.values()), 2),
    }
    return pd.concat([res, pd.DataFrame([total_row])], ignore_index=True)


def build_unit_metric_bar(unit_df: pd.DataFrame, meta: dict, kv_label: str):
    """
    Bar ngang theo từng Đơn vị kinh doanh (MA_DV) trong 1 Khu vực, cho 1 chỉ tiêu cụ thể —
    Biểu đồ kết hợp cả 3 bộ lọc KHU VỰC + SỐ THÁNG/KỲ LŨY KẾ (đã áp vào unit_df) + KHỐI CHỈ TIÊU (meta).
    """
    df_plot = unit_df[unit_df["Tên ĐVKD"] != "Tổng"].copy()
    value_col = meta["column"]
    labels = [f"{r['Mã ĐVKD']} - {r['Tên ĐVKD']}" for _, r in df_plot.iterrows()]
    vals = [float(v) if pd.notna(v) else 0.0 for v in df_plot[value_col]]
    is_pct = "%" in value_col
    if meta.get("diverging"):
        bar_colors = [(PASTEL_POSITIVE if v >= 0 else PASTEL_NEGATIVE) for v in vals]
    else:
        bar_colors = [get_unit_color(i) for i in range(len(labels))]

    fig = go.Figure(
        go.Bar(
            x=vals,
            y=labels,
            orientation="h",
            marker=dict(color=bar_colors),
            text=[f"{v:,.1f}{'%' if is_pct else ''}".replace(",", ".") for v in vals],
            textposition="outside",
            textfont=dict(size=15, color="#111827", family="Segoe UI, Arial"),
            cliponaxis=False,
        )
    )
    fig.update_layout(
        title=dict(
            text=f"<b>{meta['title']} THEO ĐVKD — {kv_label}</b>",
            x=0.5,
            y=0.98,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        yaxis=dict(autorange="reversed", tickfont=dict(size=14.5, color="#374151")),
        xaxis=dict(
            gridcolor="#f0f2f5",
            tickfont=dict(size=14, color="#6b7280"),
            zeroline=meta.get("diverging", False),
            zerolinecolor="#9ca3af",
            zerolinewidth=1.5,
        ),
        margin=dict(l=10, r=40, t=40, b=25),
        height=max(220, 60 + 42 * max(len(labels), 1)),
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def build_unit_vertical_bar(unit_df: pd.DataFrame, meta: dict, kv_label: str):
    """Cột dọc theo từng Đơn vị kinh doanh (MA_DV) trong 1 Khu vực, cho 1 chỉ tiêu cụ thể."""
    df_plot = unit_df[unit_df["Tên ĐVKD"] != "Tổng"].copy()
    value_col = meta["column"]
    labels = [f"{r['Mã ĐVKD']} - {r['Tên ĐVKD']}" for _, r in df_plot.iterrows()]
    vals = [float(v) if pd.notna(v) else 0.0 for v in df_plot[value_col]]
    is_pct = "%" in value_col
    if meta.get("diverging"):
        bar_colors = [(PASTEL_POSITIVE if v >= 0 else PASTEL_NEGATIVE) for v in vals]
    else:
        bar_colors = [get_unit_color(i) for i in range(len(labels))]

    fig = go.Figure(
        go.Bar(
            x=labels,
            y=vals,
            marker=dict(color=bar_colors),
            text=[f"{v:,.1f}{'%' if is_pct else ''}".replace(",", ".") for v in vals],
            textposition="outside",
            textfont=dict(size=14, color="#111827", family="Segoe UI, Arial"),
        )
    )
    fig.update_layout(
        title=dict(
            text=f"<b>{meta['title']} THEO ĐVKD — {kv_label}</b>",
            x=0.5,
            y=0.98,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        xaxis=dict(tickfont=dict(size=13.5, color="#374151"), tickangle=-30),
        yaxis=dict(
            gridcolor="#f0f2f5",
            tickfont=dict(size=14, color="#6b7280"),
            zeroline=True,
            zerolinecolor="#9ca3af",
            zerolinewidth=1.5,
        ),
        margin=dict(l=10, r=10, t=40, b=90),
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def build_unit_lollipop(unit_df: pd.DataFrame, meta: dict, kv_label: str):
    """
    Biểu đồ Lollipop (kẹo mút) theo từng Đơn vị kinh doanh trong 1 Khu vực — que mảnh từ 0 tới giá
    trị + đầu tròn tại giá trị (v4.13, cố định cho Khối D/E/F bất kể bộ lọc đang chọn).
    """
    df_plot = unit_df[unit_df["Tên ĐVKD"] != "Tổng"].copy()
    value_col = meta["column"]
    labels = [f"{r['Mã ĐVKD']} - {r['Tên ĐVKD']}" for _, r in df_plot.iterrows()]
    vals = [float(v) if pd.notna(v) else 0.0 for v in df_plot[value_col]]
    is_pct = "%" in value_col
    if meta.get("diverging"):
        dot_colors = [(PASTEL_POSITIVE if v >= 0 else PASTEL_NEGATIVE) for v in vals]
    else:
        dot_colors = [get_unit_color(i) for i in range(len(labels))]

    stem_x, stem_y = [], []
    for lbl, v in zip(labels, vals):
        stem_x += [lbl, lbl, None]
        stem_y += [0, v, None]

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=stem_x, y=stem_y, mode="lines", line=dict(color="#cbd5e1", width=2), showlegend=False, hoverinfo="skip"))
    fig.add_trace(
        go.Scatter(
            x=labels,
            y=vals,
            mode="markers+text",
            marker=dict(color=dot_colors, size=16, line=dict(color="white", width=1.5)),
            text=[f"{v:,.1f}{'%' if is_pct else ''}".replace(",", ".") for v in vals],
            textposition="top center",
            textfont=dict(size=13.5, color="#111827", family="Segoe UI, Arial"),
            showlegend=False,
        )
    )
    fig.update_layout(
        title=dict(
            text=f"<b>{meta['title']} THEO ĐVKD — {kv_label}</b>",
            x=0.5,
            y=0.98,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        xaxis=dict(tickfont=dict(size=13.5, color="#374151"), tickangle=-30),
        yaxis=dict(
            gridcolor="#f0f2f5",
            tickfont=dict(size=14, color="#6b7280"),
            zeroline=True,
            zerolinecolor="#9ca3af",
            zerolinewidth=1.5,
        ),
        showlegend=False,
        margin=dict(l=10, r=25, t=40, b=90),
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def build_unit_pie(unit_df: pd.DataFrame, meta: dict, kv_label: str):
    """Donut Chart — tỷ trọng theo cột chỉ tiêu giữa các ĐVKD trong 1 Khu vực (v4.13, dùng cho Khối C)."""
    df_plot = unit_df[unit_df["Tên ĐVKD"] != "Tổng"].copy()
    value_col = meta["column"]
    pairs = [
        (f"{r['Mã ĐVKD']} - {r['Tên ĐVKD']}", float(r[value_col]))
        for _, r in df_plot.iterrows()
        if pd.notna(r[value_col]) and float(r[value_col]) > 0
    ]
    labels = [p[0] for p in pairs]
    vals = [p[1] for p in pairs]
    colors = [get_unit_color(i) for i in range(len(labels))]

    if not labels:
        fig = go.Figure()
        fig.update_layout(
            title=dict(text=f"<b>{meta['title']} — TỶ TRỌNG THEO ĐVKD — {kv_label}</b>", x=0.5, font=dict(size=16.5, color="#111827", family="Segoe UI, Arial")),
            annotations=[dict(text="Không có dữ liệu dương để tính tỷ trọng", showarrow=False, font=dict(size=15, color="#6b7280"))],
            height=340,
            plot_bgcolor="white",
            paper_bgcolor="white",
        )
        return fig

    single_group_annotations = []
    if len(labels) == 1:
        single_group_annotations = [dict(text=labels[0], x=0.5, y=0.5, showarrow=False, font=dict(size=15, color="#111827", family="Segoe UI, Arial"))]

    fig = go.Figure(
        data=[
            go.Pie(
                labels=labels,
                values=vals,
                hole=0.55,
                marker=dict(colors=colors, line=dict(color="#ffffff", width=2)),
                textinfo="percent",
                textposition="inside",
                textfont=dict(size=15, color="#ffffff", family="Segoe UI, Arial"),
            )
        ]
    )
    fig.update_layout(
        title=dict(
            text=f"<b>{meta['title']} — TỶ TRỌNG THEO ĐVKD — {kv_label}</b>",
            x=0.5,
            y=0.96,
            xanchor="center",
            yanchor="top",
            font=dict(size=16.5, color="#111827", family="Segoe UI, Arial"),
        ),
        annotations=single_group_annotations,
        legend=dict(orientation="h", yanchor="top", y=-0.05, xanchor="center", x=0.5, font=dict(size=13.5)),
        margin=dict(l=10, r=10, t=45, b=45),
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    return fig


def build_unit_chart(unit_df: pd.DataFrame, meta: dict, kv_label: str):
    """Điều phối loại biểu đồ theo ĐVKD dựa trên `meta['chart_type']` (hbar/vbar/lollipop/pie, v4.13)."""
    chart_type = meta.get("chart_type", "hbar")
    if chart_type == "vbar":
        return build_unit_vertical_bar(unit_df, meta, kv_label)
    if chart_type == "lollipop":
        return build_unit_lollipop(unit_df, meta, kv_label)
    if chart_type == "pie":
        return build_unit_pie(unit_df, meta, kv_label)
    return build_unit_metric_bar(unit_df, meta, kv_label)


def _ud_fmt_int(v):
    try:
        return f"{int(v):,}".replace(",", ".")
    except (ValueError, TypeError):
        return v


def _ud_fmt_pct(v):
    try:
        return f"{float(v):.1f}%"
    except (ValueError, TypeError):
        return v


def _ud_fmt_pct_or_blank(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return _ud_fmt_pct(v)


def style_unit_detail_table(unit_df: pd.DataFrame, chi_tieu_filter: str = ""):
    """Định dạng số + highlight dòng Tổng, và highlight cột đang khớp với KHỐI CHỈ TIÊU đang chọn."""
    df_fmt = unit_df.copy()
    for col in [
        "Tổng NS đánh giá", "SL đạt KPIs", "SL thỏa LPT", "SL KHDN mới", "SL KHDN tăng ròng", "SL KHTD tăng ròng",
        "Tổng CV", "Tổng CV N-1", "Tăng ròng CV", "Tổng HĐ", "Tổng HĐ N-1", "Tăng ròng HĐ", "Tổng TOI",
    ]:
        df_fmt[col] = df_fmt[col].apply(_ud_fmt_int)
    for col in ["% đạt KPIs", "% thỏa LPT"]:
        df_fmt[col] = df_fmt[col].apply(_ud_fmt_pct)
    df_fmt["Hệ số LPT TB (%)"] = df_fmt["Hệ số LPT TB (%)"].apply(_ud_fmt_pct_or_blank)

    focus_col = CHI_TIEU_UNIT_META.get(chi_tieu_filter, {}).get("column")

    def highlight_total(row):
        styles = []
        for col in row.index:
            style = ""
            if row["Tên ĐVKD"] == "Tổng":
                style = "background-color: #fff3cd; font-weight: 700; color: #856404;"
            elif focus_col and col == focus_col:
                style = "background-color: #e8f0fe; font-weight: 700; color: #0b3d91;"
            styles.append(style)
        return styles

    return df_fmt.style.apply(highlight_total, axis=1)


# ──────────────────────────────────────────────────────────────────────────────
# DỮ LIỆU MÔ PHỎNG ĐẦY ĐỦ CHO DEMO
# ──────────────────────────────────────────────────────────────────────────────


def generate_demo_df() -> pd.DataFrame:
    """Tạo bộ dữ liệu mô phỏng 7 kỳ (T01 -> T07/2026) cho 6 khu vực."""
    random.seed(42)
    months = [1, 2, 3, 4, 5, 6, 7]
    toi_cols = [c for c in UNIT_LEVEL_NUMERIC_COLUMNS if c not in (
        "TONG_CV_THANG_N", "TONG_CV_THANG_N_1", "TANG_RONG_CV",
        "TONG_HD_THANG_N", "TONG_HD_THANG_N_1", "TANG_RONG_HD",
    )]
    rows = []
    for m in months:
        for kv_idx, kv in enumerate(STANDARD_KV_ORDER, start=1):
            for dv in range(1, 4):
                # Giá trị cấp ĐVKD — sinh 1 LẦN cho cả (ĐVKD, tháng), rồi lặp lại y hệt trên mọi
                # dòng nhân sự bên dưới, mô phỏng đúng đặc điểm dữ liệu thật (xem cảnh báo dedup
                # ở UNIT_LEVEL_NUMERIC_COLUMNS) để kiểm thử được agg_unit_level_metric().
                tong_cv_n = round(random.uniform(1e9, 5e10), 2)
                tong_cv_n1 = round(tong_cv_n * random.uniform(0.90, 1.05), 2)
                tong_hd_n = round(random.uniform(1e9, 3e10), 2)
                tong_hd_n1 = round(tong_hd_n * random.uniform(0.90, 1.05), 2)
                toi_vals = {c: round(random.uniform(1e6, 5e8), 2) for c in toi_cols}

                for ns in range(12):
                    is_null = random.random() < 0.03
                    if is_null:
                        hs = None
                    else:
                        hs = round(random.choice([0.0, 0.15, 0.35, 0.60, 0.85, 1.0]), 2)

                    rows.append({
                        "SYM_RUN_DATE": pd.Timestamp(2026, m, 28 if m == 2 else 30),
                        "THANG": m,
                        "NAM": 2026,
                        "MA_KV": kv_idx,
                        "TEN_KV": kv,
                        "MA_DV": f"DV_{kv_idx}_{dv}",
                        "TEN_DV": f"Chi nhánh {kv_idx}.{dv}",
                        "TEN_CV": f"Nhân sự {kv_idx}.{dv}.{ns + 1:02d}",
                        "KET_QUA_KPIS": "DAT_KPIS" if random.random() > 0.42 else "KHONG_DAT_KPIS",
                        "KET_QUA_LPT": "THOA_LPT" if random.random() > 0.30 else "KHONG_THOA_LPT",
                        "SL_KHDN_MOI_N": random.randint(1, 14),
                        "SL_KHDN_TANG_RONG": random.randint(-4, 9),
                        "SL_KH_TD_TANG_RONG": random.randint(-6, 12),
                        "Hệ số LPT thực nhận": hs,
                        "HE_SO_LPT_RAW": hs,
                        "HE_SO_LPT": hs if hs is not None else 0.0,
                        "TONG_CV_THANG_N": tong_cv_n,
                        "TONG_CV_THANG_N_1": tong_cv_n1,
                        "TANG_RONG_CV": round(tong_cv_n - tong_cv_n1, 2),
                        "TONG_HD_THANG_N": tong_hd_n,
                        "TONG_HD_THANG_N_1": tong_hd_n1,
                        "TANG_RONG_HD": round(tong_hd_n - tong_hd_n1, 2),
                        **toi_vals,
                    })

    df = pd.DataFrame(rows)
    df["YEAR_MONTH"] = df["SYM_RUN_DATE"].dt.to_period("M")
    df["THANG_STR"] = df.apply(lambda r: f"T{r['THANG']:02d}/{r['NAM']}", axis=1)
    return df


def generate_demo_dvkd_items() -> list[dict]:
    """Bộ dữ liệu mô phỏng cho 'Dashboard số liệu ĐVKD' — cùng shape với
    load_and_preprocess_dvkd_excel(), dựng từ vài chỉ tiêu tiêu biểu (số liệu tham khảo theo đúng
    dữ liệu mẫu thật đã cung cấp), LẶP LẠI theo 4 Đơn vị kinh doanh (v4.15, giống cấu trúc nhiều
    block của file thật — mỗi ĐVKD 1 bộ giá trị co theo tỷ lệ quy mô riêng) để demo đầy đủ tính
    năng lọc "Chọn đơn vị" mới."""
    random.seed(7)
    # (v4.21) Thêm cột "chu_ky" (Hàng ngày/Hàng tháng, đọc từ cột I file thật, DVKD_COL_CHU_KY) làm
    # phần tử thứ 6 mỗi seed item — mô phỏng đúng chu kỳ cập nhật của từng nhóm chỉ tiêu.
    seed_items = [
        ("A. Thông tin chung", "1", "Số lượng nhân sự", "NS", 1102, "Hàng tháng"),
        ("A. Chỉ tiêu Chính", "1", "Huy động NHDN", "Tỷ đồng", 135497, "Hàng ngày"),
        ("A. Chỉ tiêu Chính", "2", "Tổng cho vay", "Tỷ đồng", 419973, "Hàng ngày"),
        ("A. Chỉ tiêu Chính", "5", "Tỷ lệ nợ xấu", "%", 2.1, "Hàng ngày"),
        ("A. Chỉ tiêu Chính", "9", "Tổng thu phí dịch vụ", "Tỷ đồng", 8342, "Hàng tháng"),
        ("B. Chỉ tiêu Phụ", "7", "Nợ nhóm 2", "Tỷ đồng", 850, "Hàng ngày"),
        ("B. Chỉ tiêu Phụ", "12", "Số lượng khách hàng", "KH", 66615, "Hàng tháng"),
        ("B. Chỉ tiêu Phụ", "14", "Số lượng khách hàng active", "KH", 31731, "Hàng tháng"),
        ("B. Chỉ tiêu Phụ", "15", "Số lượng KH mới", "KH", 4200, "Hàng tháng"),
    ]
    # 4 Đơn vị kinh doanh giống đúng cấu trúc file thật (v4.15) — "Toàn hàng" giữ quy mô gốc, 3 chi
    # nhánh co lại theo tỷ lệ (chi nhánh luôn nhỏ hơn toàn hàng nhiều lần); ĐVT "%" không co theo
    # quy mô (tỷ lệ không phụ thuộc kích thước đơn vị).
    dvkd_scales = [
        (DVKD_DEFAULT_DON_VI, 1.0),
        ("CN Phú Yên", 0.012),
        ("CN TTKD", 0.018),
        ("CN Hoàn Kiếm", 0.025),
    ]
    items = []
    row_idx = 0
    for dvkd_name, scale in dvkd_scales:
        for section, prefix, ten, dvt, base_full, chu_ky in seed_items:
            base = base_full if dvt == "%" else round(base_full * scale, 1)
            thuc_hien = {m: round(base * random.uniform(0.85, 1.10), 1) for m in range(1, 9)}
            ke_hoach = {m: round(base * random.uniform(0.95, 1.05), 1) for m in range(1, 12)}
            pct_ht = {m: round(thuc_hien.get(m, base) / ke_hoach[m], 4) for m in range(2, 12)}
            label = f"{prefix}. {ten}"
            items.append({
                "row_idx": row_idx,
                "dvkd": dvkd_name,
                "section": section,
                "label": label,
                "dvt": dvt,
                "stt_display": prefix,
                "chu_ky": chu_ky,
                "chart_type": _assign_chart_type(label, dvt, section),
                "thuc_hien": thuc_hien,
                "ke_hoach": ke_hoach,
                "ke_hoach_nam": round(base * 12 * random.uniform(0.95, 1.05), 1),
                "pct_ht": pct_ht,
                "pct_ht_nam": round(random.uniform(0.8, 1.1), 4),
                # (v4.22) 4 field mo phong "So voi KH Thang/Nam" - gia tri don, khong lap theo thang.
                "todo_thang": round(random.uniform(-5000, 5000), 2),
                "pct_ht_thang": round(random.uniform(0.6, 1.2), 4),
                "todo_nam": round(random.uniform(-10000, 10000), 2),
                "pct_ht_nam2": round(random.uniform(0.5, 1.3), 4),
            })
            row_idx += 1
    return items


def get_persistent_data_dir() -> Path:
    """
    Thư mục lưu dữ liệu BỀN VỮNG qua các lần khởi động lại (vd accounts.json).
    - (v4.33) Nếu biến môi trường `PERSISTENT_DATA_DIR` được set (vd trỏ vào 1 Volume gắn ngoài
      khi deploy lên Railway/Render...) thì LUÔN ưu tiên dùng đường dẫn đó — tự tạo thư mục nếu
      chưa tồn tại — bất kể đang chạy thường hay đã đóng gói. Dùng cho môi trường container có hệ
      thống file gốc ephemeral (mất dữ liệu mỗi lần redeploy/restart) nếu không gắn Volume riêng.
    - Khi chạy thường (`python app.py` / `streamlit run app.py`): cạnh app.py.
    - Khi đã đóng gói bằng PyInstaller (`sys.frozen == True`): PHẢI dùng thư mục chứa file .exe
      thực tế (`sys.executable`), KHÔNG dùng `Path(__file__).parent` — vì app.py lúc đó được
      giải nén vào 1 thư mục tạm khác nhau mỗi lần chạy (`sys._MEIPASS`), ghi file vào đó sẽ
      mất dữ liệu ngay khi đóng ứng dụng.
    """
    env_dir = os.environ.get("PERSISTENT_DATA_DIR")
    if env_dir:
        p = Path(env_dir)
        p.mkdir(parents=True, exist_ok=True)
        return p
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def get_logo_base64() -> str:
    logo_file = Path(__file__).parent / "hdbank_logo.png"
    if logo_file.exists():
        with open(logo_file, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    return ""


# ──────────────────────────────────────────────────────────────────────────────
# ĐĂNG NHẬP / ĐĂNG KÝ — TÀI KHOẢN QUẢN TRỊ (ADMIN) & KHÓA THIẾT BỊ THEO IP
# ──────────────────────────────────────────────────────────────────────────────
# Ghi chú kiến trúc (đã thống nhất với người dùng qua câu hỏi làm rõ):
# - Chỉ tài khoản "admin" (mật khẩu cố định bên dưới) được phép tạo tài khoản mới; admin KHÔNG bị
#   khóa thiết bị (đăng nhập được từ bất kỳ máy nào).
# - Tài khoản thường do admin tạo: lần đăng nhập ĐẦU TIÊN sẽ "khóa" tài khoản đó vào đúng địa chỉ IP
#   của máy đang đăng nhập (lấy IP mạng nội bộ của máy đang chạy app — vì mô hình triển khai là mỗi
#   người tự chạy 1 bản .exe đóng gói riêng trên máy mình, không phải 1 server dùng chung nhiều người
#   qua trình duyệt). Các lần đăng nhập sau chỉ thành công nếu trùng đúng IP đã khóa.
# - Admin có thể gỡ khóa (xóa IP đã lưu) cho từng tài khoản trong tab "Quản lý tài khoản".
# - Dữ liệu tài khoản lưu tại accounts.json — cạnh app.py khi chạy thường, cạnh file .exe khi đã
#   đóng gói (xem get_persistent_data_dir()) — bền vững qua các lần khởi động lại app.
ADMIN_USERNAME = "admin"
# (v4.33) Cho phép ghi đè qua biến môi trường `ADMIN_PASSWORD` — cần thiết khi deploy lên môi
# trường public (vd Railway) để không phải giữ mật khẩu cố định ngay trong mã nguồn; mặc định
# giữ nguyên giá trị cũ khi chạy local/.exe (không set biến môi trường này).
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin@1508")
ACCOUNTS_FILE = get_persistent_data_dir() / "accounts.json"


def _hash_password(password: str, salt: Optional[str] = None) -> str:
    """Băm mật khẩu theo PBKDF2-HMAC-SHA256 kèm salt ngẫu nhiên — không lưu mật khẩu dạng chuỗi trần."""
    if salt is None:
        salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 100_000)
    return f"{salt}${digest.hex()}"


def _verify_password(password: str, stored_hash: str) -> bool:
    try:
        salt, _ = stored_hash.split("$", 1)
    except ValueError:
        return False
    return _hash_password(password, salt) == stored_hash


def get_local_ip() -> str:
    """Lấy địa chỉ IP mạng nội bộ của máy đang chạy app (dùng để khóa thiết bị cho tài khoản thường)."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "unknown"


def load_accounts() -> dict:
    if not ACCOUNTS_FILE.exists():
        return {}
    try:
        return json.loads(ACCOUNTS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_accounts(accounts: dict) -> None:
    ACCOUNTS_FILE.write_text(json.dumps(accounts, ensure_ascii=False, indent=2), encoding="utf-8")


def create_account(username: str, password: str) -> tuple[bool, str]:
    """Chỉ admin được gọi hàm này (kiểm tra vai trò ở nơi gọi). Trả về (thành công, thông báo)."""
    username = username.strip()
    if not username or not password:
        return False, "Vui lòng nhập đầy đủ tên đăng nhập và mật khẩu."
    if username.lower() == ADMIN_USERNAME.lower():
        return False, "Tên đăng nhập này đã được dùng cho tài khoản quản trị."
    accounts = load_accounts()
    if username in accounts:
        return False, f"Tài khoản '{username}' đã tồn tại."
    accounts[username] = {
        "password_hash": _hash_password(password),
        "locked_ip": None,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    save_accounts(accounts)
    return True, f"Đã tạo tài khoản '{username}' thành công."


def unlock_account_device(username: str) -> None:
    """Admin gỡ khóa thiết bị cho 1 tài khoản — lần đăng nhập kế tiếp sẽ khóa lại theo IP mới."""
    accounts = load_accounts()
    if username in accounts:
        accounts[username]["locked_ip"] = None
        save_accounts(accounts)


def authenticate(username: str, password: str) -> tuple[bool, str, str]:
    """Trả về (thành công, vai trò 'admin'/'user', thông báo lỗi nếu thất bại)."""
    username = username.strip()
    if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
        return True, "admin", ""

    accounts = load_accounts()
    user = accounts.get(username)
    if not user or not _verify_password(password, user.get("password_hash", "")):
        return False, "", "Sai tên đăng nhập hoặc mật khẩu."

    current_ip = get_local_ip()
    locked_ip = user.get("locked_ip")
    if locked_ip is None:
        user["locked_ip"] = current_ip
        accounts[username] = user
        save_accounts(accounts)
        return True, "user", ""
    if locked_ip != current_ip:
        return False, "", f"Tài khoản này đã đăng nhập trên thiết bị khác (IP: {locked_ip}). Liên hệ Admin để gỡ khóa thiết bị."
    return True, "user", ""


def render_login_page(logo_img_html: str) -> None:
    """Màn hình đăng nhập — chặn toàn bộ nội dung Dashboard cho tới khi đăng nhập thành công."""
    st.markdown(
        f"""
        <div style="text-align: center; padding: 40px 0 10px 0;">
            <div style="margin-bottom: 12px;">{logo_img_html}</div>
            <h1 style="font-size: 22px; font-weight: 800; color: #111827; margin: 0 0 6px 0; text-transform: uppercase;">
                Đăng nhập hệ thống
            </h1>
            <p style="font-size: 13.5px; color: #4b5563; margin: 0 0 20px 0;">
                Kết Quả Tổng Quan KPIs và Lương Phụ Trội QHKHDN
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    col_l, col_m, col_r = st.columns([1.2, 1.6, 1.2])
    with col_m:
        with st.container(border=True):
            with st.form("login_form"):
                username = st.text_input("Tên đăng nhập")
                password = st.text_input("Mật khẩu", type="password")
                submitted = st.form_submit_button("🔐 Đăng nhập", width="stretch", type="primary")
            if submitted:
                ok, role, err = authenticate(username, password)
                if ok:
                    st.session_state.auth_user = username
                    st.session_state.auth_role = role
                    st.rerun()
                else:
                    st.error(f"❌ {err}")
            st.caption("Tài khoản mới chỉ do Admin cấp — liên hệ Admin nếu chưa có tài khoản hoặc cần gỡ khóa thiết bị.")
    st.stop()


def render_account_admin_panel() -> None:
    """Khối quản lý tài khoản trong Sidebar — chỉ hiển thị khi đăng nhập bằng vai trò admin."""
    st.markdown("### 🔐 Quản Lý Tài Khoản")
    st.caption("Chỉ Admin thấy mục này. Tài khoản Admin không bị khóa thiết bị.")

    with st.expander("➕ Tạo tài khoản mới"):
        with st.form("create_account_form", clear_on_submit=True):
            new_user = st.text_input("Tên đăng nhập mới")
            new_pass = st.text_input("Mật khẩu", type="password", key="new_pass")
            new_pass2 = st.text_input("Nhập lại mật khẩu", type="password", key="new_pass2")
            submitted = st.form_submit_button("Tạo tài khoản", type="primary", width="stretch")
        if submitted:
            if new_pass != new_pass2:
                st.error("❌ Mật khẩu nhập lại không khớp.")
            else:
                ok, msg = create_account(new_user, new_pass)
                if ok:
                    st.success(f"✅ {msg}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")

    accounts = load_accounts()
    with st.expander(f"📋 Danh sách tài khoản ({len(accounts)})"):
        if not accounts:
            st.caption("Chưa có tài khoản nào được tạo ngoài Admin.")
        else:
            for i, (uname, info) in enumerate(accounts.items()):
                locked_ip = info.get("locked_ip")
                st.markdown(f"**{uname}**")
                st.caption(f"📍 {locked_ip or 'Chưa đăng nhập máy nào'}")
                st.caption(f"🗓️ {info.get('created_at', '')}")
                if locked_ip:
                    if st.button("Gỡ khóa thiết bị", key=f"unlock_{uname}", width="stretch"):
                        unlock_account_device(uname)
                        st.rerun()
                if i < len(accounts) - 1:
                    st.markdown("---")


# ──────────────────────────────────────────────────────────────────────────────
# GIAO DIỆN CHÍNH STREAMLIT
# ──────────────────────────────────────────────────────────────────────────────


def main():
    logo_b64 = get_logo_base64()
    if logo_b64:
        logo_img_html = f'<img src="data:image/png;base64,{logo_b64}" alt="HDBank Logo" style="height:85px; max-height:90px; object-fit:contain;" />'
    else:
        logo_img_html = '<span style="font-family:\'Arial Black\', Impact, sans-serif; font-size:32px; font-weight:900; color:#ed1c24;">HDBank</span>'

    # ── CỔNG ĐĂNG NHẬP: CHẶN TOÀN BỘ NỘI DUNG CHO TỚI KHI ĐĂNG NHẬP THÀNH CÔNG ──
    if "auth_user" not in st.session_state:
        render_login_page(logo_img_html)

    # ── THÔNG TIN TÀI KHOẢN & ĐĂNG XUẤT — LUÔN HIỂN THỊ (kể cả khi chưa tải dữ liệu) ──
    with st.sidebar:
        role_display = "Quản trị (Admin)" if st.session_state.auth_role == "admin" else "Người dùng"
        st.markdown(f"👤 **{st.session_state.auth_user}** _( {role_display} )_")
        if st.button("🚪 Đăng xuất", width="stretch", key="btn_logout"):
            del st.session_state["auth_user"]
            del st.session_state["auth_role"]
            st.rerun()

        confirm_close = st.checkbox("Xác nhận đóng ứng dụng", key="chk_confirm_close")
        if st.button("🔴 Đóng ứng dụng", width="stretch", key="btn_close_app", disabled=not confirm_close):
            st.info("Đang đóng ứng dụng, có thể tắt cửa sổ trình duyệt này...")
            os._exit(0)

        if st.session_state.auth_role == "admin":
            render_account_admin_panel()
        st.markdown("---")

    has_data = (
        (st.session_state.uploaded_df is not None)
        or (st.session_state.uploaded_dvkd_items is not None)
        or st.session_state.use_demo
    )

    # ── MÀN HÌNH TẢI FILE BAN ĐẦU ─────────────────────────────────────────────
    if not has_data:
        st.markdown(
            f"""
            <div style="text-align: center; padding: 25px 0 10px 0;">
                <div style="margin-bottom: 12px;">{logo_img_html}</div>
                <h1 style="font-size: 24px; font-weight: 800; color: #111827; margin: 0 0 6px 0; text-transform: uppercase; letter-spacing: 0.3px;">
                    KẾT QUẢ TỔNG QUAN KPIs VÀ LƯƠNG PHỤ TRỘI QHKHDN
                </h1>
                <p style="font-size: 14px; color: #4b5563; margin: 0 0 20px 0;">
                    Hệ thống tự động tính toán 9 khối chỉ tiêu theo sheet Đặc tả công thức, theo từng tháng thực tế
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        col_l, col_m, col_r = st.columns([1, 2.6, 1])
        with col_m:
            with st.container(border=True):
                st.markdown(
                    """
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:12px;">
                        <span style="font-size:22px;">📁</span>
                        <span style="font-size:16px; font-weight:700; color:#1e3a8a;">TẢI LÊN FILE DỮ LIỆU EXCEL (.XLSX)</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                # (v4.26) 2 ô tải file RIÊNG BIỆT, không còn tự xử lý ngay khi chọn file — chỉ đọc
                # và nạp dữ liệu khi người dùng bấm nút "Tải Dashboard" bên dưới (trước đó: mỗi ô
                # `st.file_uploader` tự trigger xử lý + rerun ngay khi có file, không cho người dùng
                # cơ hội chọn lại/xem trước cả 2 file cùng lúc trước khi tải). Ô 1 = file chính (đủ
                # cả 2 sheet hoặc chỉ 1 trong 2); Ô 2 = file ĐVKD riêng (tùy chọn — dùng khi 2 sheet
                # nằm ở 2 file Excel khác nhau, giá trị của ô này LUÔN ghi đè dữ liệu ĐVKD đọc được
                # từ ô 1 nếu cả 2 cùng có, vì đây là lựa chọn tường minh của người dùng).
                landing_file = st.file_uploader(
                    "1️⃣ File Excel chính (chứa sheet 'Data import KPIs' và/hoặc 'Data import số ĐVKD'):",
                    type=["xlsx"],
                    key="landing_file_uploader",
                    help="File Excel cần chứa sheet 'Data import KPIs' và/hoặc 'Data import số ĐVKD'.",
                )
                landing_file_dvkd = st.file_uploader(
                    "2️⃣ File Excel số ĐVKD (tùy chọn — chỉ cần nếu sheet 'Data import số ĐVKD' nằm ở 1 file Excel khác):",
                    type=["xlsx"],
                    key="landing_file_uploader_dvkd",
                )

                st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
                load_clicked = st.button(
                    "🚀 Tải Dashboard",
                    width="stretch",
                    type="primary",
                    disabled=(landing_file is None and landing_file_dvkd is None),
                    key="btn_landing_load",
                )
                if load_clicked:
                    try:
                        msgs = []
                        has_kpi_sheet = False
                        has_dvkd_sheet = False
                        if landing_file is not None:
                            file_bytes = landing_file.read()
                            xls_probe = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
                            has_kpi_sheet = find_data_sheet(xls_probe) is not None
                            has_dvkd_sheet = find_data_sheet_dvkd(xls_probe) is not None
                            if not has_kpi_sheet and not has_dvkd_sheet and landing_file_dvkd is None:
                                st.error("❌ Không tìm thấy sheet 'Data import KPIs' hoặc 'Data import số ĐVKD' trong file này.")
                            else:
                                with st.spinner("Đang xử lý dữ liệu và tính toán chỉ tiêu..."):
                                    if has_kpi_sheet:
                                        df = load_and_preprocess_excel(file_bytes)
                                        st.session_state.uploaded_df = df
                                        st.session_state.uploaded_file_name = landing_file.name
                                        msgs.append(f"{len(df):,} dòng KPIs")
                                    if has_dvkd_sheet:
                                        items = load_and_preprocess_dvkd_excel(file_bytes)
                                        st.session_state.uploaded_dvkd_items = items
                                        st.session_state.uploaded_file_name_dvkd = landing_file.name
                                        msgs.append(f"{len(items):,} dòng dữ liệu ĐVKD")

                        if landing_file_dvkd is not None:
                            with st.spinner("Đang xử lý dữ liệu ĐVKD..."):
                                items2 = load_and_preprocess_dvkd_excel(landing_file_dvkd.read())
                                st.session_state.uploaded_dvkd_items = items2
                                st.session_state.uploaded_file_name_dvkd = landing_file_dvkd.name
                                msgs.append(f"{len(items2):,} dòng dữ liệu ĐVKD (file riêng)")
                                has_dvkd_sheet = True

                        if not msgs:
                            if landing_file is not None or landing_file_dvkd is not None:
                                st.error("❌ Không đọc được dữ liệu hợp lệ từ (các) file đã chọn.")
                        else:
                            st.session_state.use_demo = False
                            # (Fix UX) Chỉ có dữ liệu ĐVKD (không có KPIs) -> mở thẳng tab ĐVKD sau
                            # khi rerun, tránh người dùng tưởng nhầm là lỗi khi thấy tab KPIs (mặc
                            # định) báo "Chưa có dữ liệu" dù ĐVKD đã tải thành công.
                            st.session_state.default_top_tab = 1 if (has_dvkd_sheet and not has_kpi_sheet) else 0
                            st.success("✅ Đã tải thành công: " + ", ".join(msgs))
                            st.rerun()
                    except Exception as ex:
                        st.error(f"❌ Lỗi xử lý file Excel: {ex}")

                st.markdown(
                    """
                    <div style="background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 6px; padding: 10px 14px; margin-top: 14px; font-size: 12.5px; color: #475569; line-height: 1.6;">
                        <strong>📌 Đặc tả công thức tính toán:</strong><br>
                        • <strong>Theo từng tháng thực tế:</strong> Mỗi cột kết quả ứng với 1 tháng riêng lẻ có trong data (tối đa 12 tháng gần nhất), không cộng dồn nhiều tháng.<br>
                        • <strong>9 Khối chỉ tiêu:</strong> Bảng A (Đạt KPIs), B (Thỏa LPT), C (6 Dải LPT), D (KHDN mới), E (KHTD tăng ròng), F (KHDN tăng ròng), G (Chi tiết dư nợ), H (Chi tiết huy động), I (Chi tiết TOI).<br>
                        • <strong>Phân bổ LPT (Bảng C):</strong> Chuẩn hóa nửa khoảng đóng-trái, hệ số NULL không tính vào dải nào.
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
                if st.button("👉 Hoặc bấm vào đây để xem trước với Dữ liệu mẫu (Chuẩn 9 khối chỉ tiêu)", width="stretch", type="secondary"):
                    st.session_state.use_demo = True
                    st.rerun()

        st.stop()

    # ── MÀN HÌNH DASHBOARD ĐẦY ĐỦ KHI ĐÃ CÓ DỮ LIỆU ────────────────────────────

    has_kpi_data = (st.session_state.uploaded_df is not None) or st.session_state.use_demo

    # (v4.18) `st.tabs()` không tự lưu tab đang chọn phía server — mỗi rerun (kể cả rerun do 1
    # widget bên trong tab ĐVKD gây ra) khiến trình duyệt vẽ lại từ đầu và mặc định nhảy về tab đầu.
    # Khắc phục bằng cờ `active_top_tab`, được SET =1 CHỈ bởi đúng những widget bên trong tab ĐVKD
    # (xem `on_change`/nhánh `if st.button(...)` của chúng bên dưới — KHÔNG set unconditional ở đầu
    # mỗi `with top_tab_*:`, vì cả 2 khối `with` đều LUÔN thực thi mỗi lượt rerun bất kể tab nào đang
    # hiển thị phía client, nên set cứng ở đầu mỗi khối sẽ luôn cho kết quả = giá trị của khối chạy
    # SAU CÙNG trong code — tức luôn kẹt ở tab ĐVKD, khiến tab KPIs không thể tương tác được nữa).
    # `active_tab` (biến cục bộ) chốt QUYẾT ĐỊNH có tự bấm tab ĐVKD cho LƯỢT RENDER NÀY hay không,
    # đọc từ giá trị `active_top_tab` do lượt chạy TRƯỚC để lại (với widget dùng `on_change`, callback
    # đã chạy XONG trước khi thân script này bắt đầu nên đọc được giá trị MỚI ngay trong cùng lượt;
    # với widget gọi `st.rerun()` tường minh, giá trị mới chỉ có hiệu lực ở lượt chạy KẾ TIẾP — đúng
    # bằng lượt mà `st.rerun()` kích hoạt, nên người dùng vẫn thấy đúng tab ngay lập tức).
    active_tab = st.session_state.get("active_top_tab", 0)
    if st.session_state.get("default_top_tab") == 1:
        st.session_state.default_top_tab = 0
        active_tab = 1

    # Reset về mặc định (tab KPIs) làm nền cho lượt chạy NÀY — nếu lượt rerun này thực sự xuất phát
    # từ 1 tương tác trong tab ĐVKD, đúng đoạn code xử lý widget đó (chạy SAU dòng này) sẽ tự set lại
    # =1 để lượt render KẾ TIẾP tiếp tục ở đúng tab ĐVKD.
    st.session_state.active_top_tab = 0

    if active_tab == 1:
        st.markdown(
            """<script>
            (function tryClick() {
                const tabs = window.parent.document.querySelectorAll('[data-baseweb="tab"]');
                if (tabs.length > 1) { tabs[1].click(); }
                else { setTimeout(tryClick, 50); }
            })();
            </script>""",
            unsafe_allow_html=True,
        )

    # (v4.27) Marker ẩn cho CSS `:has()` — xem khối CSS "2 TAB CẤP CAO NHẤT" ở HDBANK_CSS phía trên,
    # đảm bảo CSS làm nổi bật tab chỉ áp dụng cho ĐÚNG cặp tab cấp cao nhất này, không lan sang các
    # `st.tabs()` con khác (Dashboard Tổng Quan/Báo Cáo/Xuất Excel...).
    st.markdown('<span class="top-nav-marker" style="display:none;"></span>', unsafe_allow_html=True)
    top_tab_kpis, top_tab_dvkd = st.tabs(["📊 Dashboard KPIs và LPT", "🏦 Dashboard số liệu ĐVKD"])

    with top_tab_kpis:
        if not has_kpi_data:
            if st.session_state.uploaded_dvkd_items is not None:
                st.success("✅ Đã có dữ liệu ĐVKD — bấm sang tab '🏦 Dashboard số liệu ĐVKD' để xem.")
            st.info("📂 Chưa có dữ liệu 'Data import KPIs'. Vui lòng tải lên tại đây để xem Dashboard KPIs và LPT:")
            inline_kpi_file = st.file_uploader(
                "Chọn file Excel chứa sheet 'Data import KPIs':",
                type=["xlsx"],
                key="inline_kpi_uploader",
            )
            if inline_kpi_file is not None:
                try:
                    with st.spinner("Đang xử lý dữ liệu và tính toán chỉ tiêu..."):
                        df = load_and_preprocess_excel(inline_kpi_file.read())
                        st.session_state.uploaded_df = df
                        st.session_state.uploaded_file_name = inline_kpi_file.name
                        st.success(f"✅ Đã tải thành công {len(df):,} dòng dữ liệu!")
                        st.rerun()
                except Exception as ex:
                    st.error(f"❌ Lỗi xử lý file Excel: {ex}")
        else:
            # Chuẩn bị DataFrame tính toán
            if st.session_state.uploaded_df is not None:
                calc_df = st.session_state.uploaded_df
            else:
                calc_df = generate_demo_df()

            # Sidebar quản lý file
            with st.sidebar:
                if logo_b64:
                    st.markdown(
                        f'<div style="text-align:center; margin-bottom:12px;"><img src="data:image/png;base64,{logo_b64}" style="width:170px; object-fit:contain;" /></div>',
                        unsafe_allow_html=True,
                    )
                st.markdown("### ⚙️ Quản lý Dữ liệu")

                if st.session_state.uploaded_df is not None:
                    st.success(f"📄 File: **{st.session_state.uploaded_file_name}**")
                    st.write(f"Tổng số dòng: **{len(calc_df):,}**")
                else:
                    st.info("🌟 Đang xem: **Dữ liệu mẫu HDBank** (7 kỳ T01-T07/2026)")

                re_upload = st.file_uploader(
                    "📁 Đổi file Excel khác:",
                    type=["xlsx"],
                    key="sidebar_reupload",
                    help="Tải file Excel khác để cập nhật Dashboard.",
                )
                if re_upload is not None:
                    try:
                        df = load_and_preprocess_excel(re_upload.read())
                        st.session_state.uploaded_df = df
                        st.session_state.uploaded_file_name = re_upload.name
                        st.session_state.use_demo = False
                        st.success("✅ Đã cập nhật file mới!")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"❌ Lỗi: {ex}")

                if st.button("🔄 Nhập file mới từ đầu", width="stretch"):
                    st.session_state.uploaded_df = None
                    st.session_state.uploaded_file_name = ""
                    st.session_state.use_demo = False
                    st.session_state.filter_kv = []
                    st.session_state.filter_thang = []
                    st.session_state.filter_chi_tieu = []
                    st.session_state.filter_ma_dv = []
                    st.session_state.filter_nhan_su = []
                    st.rerun()

                st.markdown("---")
                st.caption("Khối QHKHDN - HDBank Dashboard v2.0")

            # Tính toán các chỉ số Overview Metrics từ dữ liệu thật
            metrics = compute_overview_metrics(
                calc_df,
                st.session_state.filter_kv,
                st.session_state.filter_thang,
                st.session_state.filter_chi_tieu,
                st.session_state.filter_ma_dv,
                st.session_state.filter_nhan_su,
            )

            # ── HEADER SECTION ────────────────────────────────────────────────────────
            header_html = f"""
            <div class="hdbank-header">
                <div class="hdbank-brand">
                    {logo_img_html}
                </div>
                <div class="hdbank-title-center">
                    <h1 class="hdbank-main-title">KẾT QUẢ TỔNG QUAN KPIs VÀ LƯƠNG PHỤ TRỘI QHKHDN</h1>
                    <div class="hdbank-sub-title">Hệ thống tổng hợp hiệu quả KPIs, chi trả lương phụ trội và phát triển khách hàng</div>
                </div>
                <div class="hdbank-date-right">
                    <div><em>Dữ liệu tại ngày: <strong>{metrics['run_date']}</strong></em></div>
                    <div><em>So sánh với kỳ trước: <strong>{metrics['prev_date']}</strong></em></div>
                </div>
            </div>
            """
            st.markdown(header_html, unsafe_allow_html=True)

            # ── THANH TRẠNG THÁI FILE & NÚT ĐỔI FILE ──────────────────────────────────
            bar_col1, bar_col2 = st.columns([4, 1.2])
            with bar_col1:
                if st.session_state.uploaded_df is not None:
                    st.info(f"📂 **Nguồn dữ liệu:** File `{st.session_state.uploaded_file_name}` ({len(calc_df):,} nhân sự) • Dữ liệu chốt ngày: **{metrics['run_date']}**")
                else:
                    st.info(f"🌟 **Nguồn dữ liệu:** Dữ liệu mẫu chuẩn HDBank ({len(calc_df):,} nhân sự) • Dữ liệu chốt ngày: **{metrics['run_date']}**")
            with bar_col2:
                if st.button("🔄 Nhập file Excel khác", width="stretch", key="btn_top_change_file"):
                    st.session_state.uploaded_df = None
                    st.session_state.uploaded_file_name = ""
                    st.session_state.use_demo = False
                    st.session_state.filter_kv = []
                    st.session_state.filter_thang = []
                    st.session_state.filter_chi_tieu = []
                    st.session_state.filter_ma_dv = []
                    st.session_state.filter_nhan_su = []
                    st.rerun()

            # ── TÍNH TOÁN TRƯỚC 9 BẢNG CHỈ TIÊU CHO TOÀN BỘ DATA (theo từng tháng thực tế) ──
            periods = get_monthly_periods(calc_df)
            b1 = calc_block1(calc_df, periods)
            b2 = calc_block2(calc_df, periods)
            b3 = calc_block3(calc_df, periods)
            b4 = calc_block4(calc_df, periods)
            b5 = calc_block5(calc_df, periods)
            b6 = calc_block6(calc_df, periods)
            b7 = calc_block7(calc_df, periods)
            b8 = calc_block8(calc_df, periods)
            b9 = calc_block9(calc_df, periods)

            # ── NAVIGATION TABS ──────────────────────────────────────────────────────
            tab_overview, tab_tables, tab_export = st.tabs(
                [
                    "📊 Dashboard Tổng Quan",
                    "📋 Báo Cáo 9 Khối Chỉ Tiêu (Đặc tả công thức)",
                    "📥 Xuất Báo Cáo Excel",
                ]
            )

            with tab_overview:
                # ── 3 HỘP BỘ LỌC SLICER PHONG CÁCH POWER BI ─────────────────────────
                col_kv, col_thang, col_ct = st.columns([1.1, 2.4, 1.7])

                # Hộp 1: KHU VỰC (KV1 -> KV6)
                with col_kv:
                    with st.container(border=True):
                        st.markdown(
                            """
                            <div class="slicer-card-header slicer-card-header-kv">
                                <span>KHU VỰC</span>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                        # v4.12: filter_kv là LIST (đa lựa chọn). Nút "Toàn Hàng" luôn SET về rỗng
                        # (ẩn hẳn tầng ĐVKD/Nhân sự, reset cả 2 tầng đó — giữ đúng hành vi cũ). Mỗi
                        # nút KV con toggle add/remove khỏi list (không tự reset ĐVKD/Nhân sự khi
                        # toggle — dựa vào bước prune sẵn có ngay bên dưới để tự loại các ĐVKD không
                        # còn thuộc (các) Khu vực đang chọn, tương tự cách tầng ĐVKD/Nhân sự vẫn hoạt
                        # động khi đa lựa chọn).
                        is_all_kv = not st.session_state.filter_kv
                        if st.button("🏢 Toàn Hàng", type="primary" if is_all_kv else "secondary", key="btn_all_kv", width="stretch"):
                            st.session_state.filter_kv = []
                            st.session_state.filter_ma_dv = []
                            st.session_state.filter_nhan_su = []
                            st.rerun()

                        def _toggle_kv_button(kv_key: str, key: str):
                            is_sel = kv_key in st.session_state.filter_kv
                            if st.button(kv_key, type="primary" if is_sel else "secondary", key=key):
                                if is_sel:
                                    st.session_state.filter_kv = [k for k in st.session_state.filter_kv if k != kv_key]
                                else:
                                    st.session_state.filter_kv = st.session_state.filter_kv + [kv_key]
                                st.rerun()

                        r1_c1, r1_c2 = st.columns(2)
                        with r1_c1:
                            _toggle_kv_button("KV 1", "btn_kv1")
                            _toggle_kv_button("KV 3", "btn_kv3")
                            _toggle_kv_button("KV 5", "btn_kv5")
                        with r1_c2:
                            _toggle_kv_button("KV 2", "btn_kv2")
                            _toggle_kv_button("KV 4", "btn_kv4")
                            _toggle_kv_button("KV 6", "btn_kv6")

                        # Khi đã chọn 1 Khu vực cụ thể: xổ ra các nút ĐVKD (MA_DV) trong Khu vực đó để đa lựa chọn
                        if not is_all_kv:
                            kv_units = get_kv_units(calc_df, st.session_state.filter_kv)
                            valid_ma_dv = {u[0] for u in kv_units}
                            st.session_state.filter_ma_dv = [m for m in st.session_state.filter_ma_dv if m in valid_ma_dv]

                            if kv_units:
                                st.markdown(
                                    """
                                    <div class="slicer-divider"></div>
                                    <div class="slicer-sub-label">
                                        <span class="sub-label-indicator indicator-navy"></span>
                                        ĐƠN VỊ KINH DOANH (chọn 1 hoặc nhiều)
                                    </div>
                                    """,
                                    unsafe_allow_html=True,
                                )
                                for i in range(0, len(kv_units), 2):
                                    row_units = kv_units[i : i + 2]
                                    row_cols = st.columns(2)
                                    for j, (ma_dv, ten_dv) in enumerate(row_units):
                                        with row_cols[j]:
                                            is_sel = ma_dv in st.session_state.filter_ma_dv
                                            if st.button(
                                                ten_dv,
                                                type="primary" if is_sel else "secondary",
                                                key=f"btn_dv_{ma_dv}",
                                                width="stretch",
                                            ):
                                                if is_sel:
                                                    st.session_state.filter_ma_dv.remove(ma_dv)
                                                else:
                                                    st.session_state.filter_ma_dv.append(ma_dv)
                                                    st.session_state.filter_nhan_su = []
                                                st.rerun()

                            # Tầng lọc thứ 3 — khi đã chọn ít nhất 1 ĐVKD cụ thể: xổ ra nút chọn theo
                            # Họ tên nhân sự (TEN_CV) trong (các) ĐVKD đang chọn, cùng phong cách đa
                            # lựa chọn với nút ĐVKD phía trên.
                            if st.session_state.filter_ma_dv:
                                dv_staff = get_dv_staff(calc_df, st.session_state.filter_kv, st.session_state.filter_ma_dv)
                                valid_staff = set(dv_staff)
                                st.session_state.filter_nhan_su = [
                                    n for n in st.session_state.filter_nhan_su if n in valid_staff
                                ]

                                if dv_staff:
                                    st.markdown(
                                        """
                                        <div class="slicer-divider"></div>
                                        <div class="slicer-sub-label">
                                            <span class="sub-label-indicator indicator-navy"></span>
                                            NHÂN SỰ (chọn 1 hoặc nhiều)
                                        </div>
                                        """,
                                        unsafe_allow_html=True,
                                    )
                                    for i in range(0, len(dv_staff), 2):
                                        row_staff = dv_staff[i : i + 2]
                                        row_cols_ns = st.columns(2)
                                        for j, ten_cv in enumerate(row_staff):
                                            with row_cols_ns[j]:
                                                is_sel_ns = ten_cv in st.session_state.filter_nhan_su
                                                if st.button(
                                                    ten_cv,
                                                    type="primary" if is_sel_ns else "secondary",
                                                    key=f"btn_ns_{i + j}",
                                                    width="stretch",
                                                ):
                                                    if is_sel_ns:
                                                        st.session_state.filter_nhan_su.remove(ten_cv)
                                                    else:
                                                        st.session_state.filter_nhan_su.append(ten_cv)
                                                    st.rerun()
                            elif st.session_state.filter_nhan_su:
                                # Không còn ĐVKD nào được chọn -> tầng Nhân sự không còn hợp lệ, reset theo.
                                st.session_state.filter_nhan_su = []

                # Hộp 2: SỐ THÁNG (Từng tháng cụ thể)
                with col_thang:
                    with st.container(border=True):
                        unique_yms = sorted(calc_df["YEAR_MONTH"].dropna().unique())
                        total_months = len(unique_yms)
                        max_n = min(12, total_months)
                        avail_months = sorted([m for m in calc_df["THANG_STR"].dropna().unique() if m != "(blank)" and m != ""])

                        # Kiểm tra tính hợp lệ: prune list, chỉ giữ lại các tháng vẫn còn trong avail_months
                        st.session_state.filter_thang = [m for m in st.session_state.filter_thang if m in avail_months]

                        st.markdown(
                            """
                            <div class="slicer-card-header slicer-card-header-thang">
                                <span>SỐ THÁNG</span>
                                <span style="color:#94a3b8; cursor:pointer;" title="Bấm nút để lọc; bấm lại để xem tất cả">ℹ️</span>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                        # v4.12: filter_thang là LIST (đa lựa chọn) — cùng pattern loại trừ lẫn nhau
                        # với KHỐI CHỈ TIÊU: bấm "Tất cả các tháng" luôn SET lại thành đủ avail_months
                        # (hoặc về rỗng nếu đang bật); bấm 1 tháng lẻ trong lúc đang ở trạng thái "Tất
                        # cả" sẽ thoát trạng thái đó và chỉ còn đúng tháng vừa bấm.
                        is_all_months = bool(avail_months) and set(st.session_state.filter_thang) == set(avail_months)
                        if st.button(
                            "📅 Tất cả các tháng (So sánh)",
                            type="primary" if is_all_months else "secondary",
                            key="btn_all_months",
                            width="stretch",
                        ):
                            st.session_state.filter_thang = [] if is_all_months else list(avail_months)
                            st.rerun()

                        def _toggle_thang_button(m: str, key: str):
                            is_sel = m in st.session_state.filter_thang
                            if st.button(m, type="primary" if is_sel else "secondary", key=key):
                                if is_all_months:
                                    # Đang ở trạng thái "Tất cả các tháng" -> bấm 1 tháng lẻ nghĩa là
                                    # chuyển hẳn sang chọn riêng đúng tháng đó (loại trừ "Tất cả").
                                    st.session_state.filter_thang = [m]
                                elif is_sel:
                                    st.session_state.filter_thang = [x for x in st.session_state.filter_thang if x != m]
                                else:
                                    st.session_state.filter_thang = st.session_state.filter_thang + [m]
                                st.rerun()

                        # Hàng 1 Tháng cụ thể: tối đa 6 tháng đầu
                        r1_m = avail_months[:6]
                        cols_m1 = st.columns(6)
                        for i, m in enumerate(r1_m):
                            with cols_m1[i]:
                                _toggle_thang_button(m, f"btn_m_{i+1}")

                        # Hàng 2 Tháng cụ thể (nếu có trên 6 tháng): tháng 7 trở đi
                        if len(avail_months) > 6:
                            r2_m = avail_months[6:12]
                            cols_m2 = st.columns(6)
                            for i, m in enumerate(r2_m):
                                with cols_m2[i]:
                                    _toggle_thang_button(m, f"btn_m_{i+7}")

                # Hộp 3: KHỐI CHỈ TIÊU (9 lựa chọn: Bảng A-F + Khối G/H/I)
                with col_ct:
                    with st.container(border=True):
                        st.markdown(
                            """
                            <div class="slicer-card-header slicer-card-header-ct">
                                <span>KHỐI CHỈ TIÊU (BẢNG A - I)</span>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                        # v4.11: filter_chi_tieu là LIST (đa lựa chọn) — mỗi nút chỉ tiêu con toggle
                        # add/remove khỏi list (giống hệt pattern nút ĐVKD/Nhân sự). Nút "Tất cả chỉ
                        # tiêu" và các nút chỉ tiêu con LOẠI TRỪ LẪN NHAU theo đúng yêu cầu người
                        # dùng ("chọn toàn chỉ tiêu thì không được chọn mấy cái kia, chọn nhiều chỉ
                        # tiêu nhỏ thì không được chọn toàn chỉ tiêu"): bấm "Tất cả chỉ tiêu" luôn SET
                        # lại toàn bộ list thành đủ 9 (ghi đè mọi lựa chọn lẻ đang có, hoặc về rỗng
                        # nếu đang bật); bấm 1 chỉ tiêu lẻ khi đang ở trạng thái "Tất cả" sẽ tự động
                        # thoát trạng thái đó và chỉ còn đúng chỉ tiêu vừa bấm.
                        _ALL_9_CT = list(CHI_TIEU_TREND_META.keys())
                        is_all_ct = set(st.session_state.filter_chi_tieu) == set(_ALL_9_CT)
                        if st.button("📊 Tất cả chỉ tiêu (A - I)", type="primary" if is_all_ct else "secondary", key="btn_all_ct", width="stretch"):
                            st.session_state.filter_chi_tieu = [] if is_all_ct else list(_ALL_9_CT)
                            st.rerun()

                        def _toggle_ct_button(display_label: str, ct_key: str, key: str):
                            is_sel = ct_key in st.session_state.filter_chi_tieu
                            if st.button(display_label, type="primary" if is_sel else "secondary", key=key):
                                if is_all_ct:
                                    # Đang ở trạng thái "Tất cả chỉ tiêu" — MỌI nút chỉ tiêu con đều
                                    # đang hiển thị is_sel=True lúc này, nên phải kiểm is_all_ct TRƯỚC
                                    # nhánh "đã chọn -> bỏ chọn" bên dưới, nếu không bấm 1 nút bất kỳ
                                    # sẽ bị hiểu nhầm thành "bỏ chọn nút đó khỏi 9" (còn lại 8) thay vì
                                    # đúng ý nghĩa "chuyển sang chọn riêng đúng 1 chỉ tiêu đó".
                                    st.session_state.filter_chi_tieu = [ct_key]
                                elif is_sel:
                                    st.session_state.filter_chi_tieu = [c for c in st.session_state.filter_chi_tieu if c != ct_key]
                                else:
                                    st.session_state.filter_chi_tieu = st.session_state.filter_chi_tieu + [ct_key]
                                st.rerun()

                        ct_c1, ct_c2, ct_c3 = st.columns(3)
                        with ct_c1:
                            _toggle_ct_button("Tỷ lệ đạt KPIs", "Tỷ lệ đạt KPIs", "btn_ct_kpi")
                            _toggle_ct_button("Phân bổ dải LPT thực nhận", "Phân bổ dải LPT thực nhận", "btn_ct_pb")
                            _toggle_ct_button("KHTD tăng ròng/ĐVKD", "KHTD tăng ròng/ĐVKD", "btn_ct_khtd")

                        with ct_c2:
                            _toggle_ct_button("Tỷ lệ thỏa LPT", "Tỷ lệ thỏa LPT", "btn_ct_lpt")
                            _toggle_ct_button("KHDN mới/ĐVKD", "KHDN mới/ĐVKD", "btn_ct_khdn_moi")
                            _toggle_ct_button("KHDN tăng ròng/ĐVKD", "KHDN tăng ròng/ĐVKD", "btn_ct_khdn_tang")

                        with ct_c3:
                            _toggle_ct_button("Chi tiết dư nợ", "Chi tiết dư nợ", "btn_ct_khoi_g")
                            _toggle_ct_button("Chi tiết huy động", "Chi tiết huy động", "btn_ct_khoi_h")
                            _toggle_ct_button("Chi tiết TOI", "Chi tiết TOI", "btn_ct_khoi_i")

                st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)

                # ── 6 THẺ KPI HÌNH VIÊN THUỐC (ROUNDED PILL CARDS) ──────────────────
                k1, k2, k3, k4, k5, k6 = st.columns(6)

                def _get_delta_cls(d_str: str) -> str:
                    if "▲" in d_str:
                        return "delta-up"
                    if "▼" in d_str:
                        return "delta-down"
                    return "delta-neutral"

                # Thẻ 1: Tổng số ĐVKD (xanh dương pastel) — đẩy lên đầu hàng theo yêu cầu
                with k1:
                    st.markdown(
                        f"""
                        <div class="kpi-pill kpi-bg-blue">
                            <span class="kpi-icon-badge kpi-icon-blue">👥</span>
                            <div class="kpi-title">Tổng số ĐVKD</div>
                            <div class="kpi-value-container">
                                <span class="kpi-big-num color-blue">{metrics['total_dvkd']:,}</span>
                            </div>
                            <div class="kpi-delta {_get_delta_cls(metrics['delta_dvkd'])}">{metrics['delta_dvkd']}</div>
                        </div>
                        """.replace(",", "."),
                        unsafe_allow_html=True,
                    )

                # Thẻ 2: Tổng NS đánh giá KPIs (tím pastel)
                with k2:
                    st.markdown(
                        f"""
                        <div class="kpi-pill kpi-bg-purple">
                            <span class="kpi-icon-badge kpi-icon-purple">📊</span>
                            <div class="kpi-title">Tổng NS đánh giá KPIs</div>
                            <div class="kpi-value-container">
                                <span class="kpi-big-num color-navy">{metrics['total_personnel']:,}</span>
                            </div>
                            <div class="kpi-delta {_get_delta_cls(metrics['delta_personnel'])}">{metrics['delta_personnel']}</div>
                        </div>
                        """.replace(",", "."),
                        unsafe_allow_html=True,
                    )

                # Thẻ 3: Đạt KPIs, kèm % đạt (xanh lá pastel)
                with k3:
                    st.markdown(
                        f"""
                        <div class="kpi-pill kpi-bg-green">
                            <span class="kpi-icon-badge kpi-icon-green">🎯</span>
                            <div class="kpi-title">Đạt KPIs</div>
                            <div class="kpi-value-container">
                                <span class="kpi-big-num color-green">{metrics['kpi_dat']:,}</span>
                                <span class="kpi-sub-pct">{metrics['pct_kpi_dat']:.1f}%</span>
                            </div>
                            <div class="kpi-delta {_get_delta_cls(metrics['delta_kpi_dat'])}">{metrics['delta_kpi_dat']}</div>
                        </div>
                        """.replace(",", "."),
                        unsafe_allow_html=True,
                    )

                # Thẻ 4: Thỏa Lương Phụ Trội, kèm % thỏa (cam pastel)
                with k4:
                    st.markdown(
                        f"""
                        <div class="kpi-pill kpi-bg-orange">
                            <span class="kpi-icon-badge kpi-icon-orange">👤</span>
                            <div class="kpi-title">Thỏa Lương Phụ Trội</div>
                            <div class="kpi-value-container">
                                <span class="kpi-big-num color-orange">{metrics['thoa_lpt']:,}</span>
                                <span class="kpi-sub-pct">{metrics['pct_thoa_lpt']:.1f}%</span>
                            </div>
                            <div class="kpi-delta {_get_delta_cls(metrics['delta_thoa_lpt'])}">{metrics['delta_thoa_lpt']}</div>
                        </div>
                        """.replace(",", "."),
                        unsafe_allow_html=True,
                    )

                # Thẻ 5: KHDN tăng ròng, kèm bình quân/ĐVKD (hồng pastel) — (v4.28) đổi từ "KHDN mới"
                # (SL_KHDN_MOI_N) sang "KHDN tăng ròng" (SL_KHDN_TANG_RONG) theo yêu cầu người dùng.
                with k5:
                    st.markdown(
                        f"""
                        <div class="kpi-pill kpi-bg-pink">
                            <span class="kpi-icon-badge kpi-icon-pink">💼</span>
                            <div class="kpi-title">KHDN tăng ròng<br>(BQ/ĐVKD)</div>
                            <div class="kpi-value-container">
                                <span class="kpi-big-num color-magenta">{metrics['total_khdn_tang']:,}</span>
                                <span class="kpi-sub-pct">{metrics['avg_khdn_tang']:.1f}/ĐV</span>
                            </div>
                            <div class="kpi-delta {_get_delta_cls(metrics['delta_khdn_tang'])}">{metrics['delta_khdn_tang']}</div>
                        </div>
                        """.replace(",", "."),
                        unsafe_allow_html=True,
                    )

                # Thẻ 6: KHTD tăng ròng, kèm bình quân/ĐVKD (vàng hổ phách pastel)
                with k6:
                    st.markdown(
                        f"""
                        <div class="kpi-pill kpi-bg-amber">
                            <span class="kpi-icon-badge kpi-icon-amber">📈</span>
                            <div class="kpi-title">KHTD tăng ròng<br>(BQ/ĐVKD)</div>
                            <div class="kpi-value-container">
                                <span class="kpi-big-num color-red">{metrics['total_khtd_tang']:,}</span>
                                <span class="kpi-sub-pct">{metrics['avg_khtd_tang']:.1f}/ĐV</span>
                            </div>
                            <div class="kpi-delta {_get_delta_cls(metrics['delta_khtd_tang'])}">{metrics['delta_khtd_tang']}</div>
                        </div>
                        """.replace(",", "."),
                        unsafe_allow_html=True,
                    )

                st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)

                # ── KẾT HỢP 3 BỘ LỌC: KHU VỰC + SỐ THÁNG/KỲ LŨY KẾ + KHỐI CHỈ TIÊU ──────
                # - Không chọn Khu vực (hoặc bấm "Toàn Hàng") & không chọn Khối chỉ tiêu -> 3 biểu đồ toàn hàng mặc định.
                # - Có chọn 1 trong 2 (hoặc cả 2) -> ẨN 3 biểu đồ toàn hàng, thay bằng biểu đồ kết hợp:
                #     + Bấm "Tất cả chỉ tiêu (A-F)": hiện đủ 6 biểu đồ (mỗi Khối 1 biểu đồ riêng, dạng lưới).
                #     + Chỉ chọn 1 Khối chỉ tiêu (chưa chọn Khu vực): 1 biểu đồ theo 6 Khu vực, bấm vào 1 cột
                #       Khu vực sẽ tự lọc theo Khu vực đó (như bấm nút Slicer KHU VỰC).
                #     + Có chọn Khu vực: biểu đồ theo từng Đơn vị kinh doanh (MA_DV) trong Khu vực đó
                #       (1 biểu đồ nếu chọn đúng 1 Khối chỉ tiêu, hoặc lưới 6 biểu đồ nếu chọn "Tất cả chỉ tiêu").
                # v4.12: filter_kv là LIST (đa lựa chọn). `sel_kv` (chuỗi đơn, nối bằng ", ") chỉ
                # dùng để HIỂN THỊ (tiêu đề/caption) — mọi nơi truyền làm bộ lọc dữ liệu phải dùng
                # `sel_kv_list` (apply_kv_filter/get_unit_detail_table/calc_monthly_series_by_*... đều
                # đã tổng quát hóa nhận list).
                sel_kv_list = st.session_state.filter_kv
                has_kv = len(sel_kv_list) > 0
                sel_kv = ", ".join(sel_kv_list) if sel_kv_list else ""
                # v4.11: filter_chi_tieu là LIST (đa lựa chọn). `sel_ct` (chuỗi đơn) chỉ có ý nghĩa
                # khi CHỌN ĐÚNG 1 chỉ tiêu — giữ lại biến này để không phải sửa lại toàn bộ các
                # nhánh xử lý "1 chỉ tiêu"/"mặc định" phía dưới (n_ct 0 hoặc 1), vốn đã dùng `sel_ct`
                # xuyên suốt (tiêu đề, cột sắp xếp/tô sáng bảng ĐVKD...). Khi n_ct >= 2, các nhánh
                # MỚI (is_multi_ct) tự lặp qua `sel_ct_ordered` thay vì dùng `sel_ct`.
                sel_ct_list = st.session_state.filter_chi_tieu
                n_ct = len(sel_ct_list)
                _ALL_9_CT_ORDER = list(CHI_TIEU_TREND_META.keys())
                sel_ct_ordered = [c for c in _ALL_9_CT_ORDER if c in sel_ct_list]
                sel_ct = sel_ct_list[0] if n_ct == 1 else ""
                has_ct = n_ct > 0
                is_all_ct = n_ct == len(_ALL_9_CT_ORDER)
                is_multi_ct = has_ct and not is_all_ct and n_ct > 1
                # v4.12: filter_thang là LIST (đa lựa chọn). `is_all_months` (tên biến giữ nguyên để
                # không phải sửa lại hàng chục điều kiện nhánh phía dưới) giờ nghĩa rộng hơn: BẬT khi
                # chọn >= 2 tháng (dù là bấm nút "Tất cả các tháng" hay tự chọn tay nhiều tháng lẻ) —
                # cả 2 trường hợp đều hiển thị dạng biểu đồ xu hướng (trend) giống hệt nhau, chỉ khác
                # tập hợp tháng được lọc (`sel_thang_list`, truyền vào các hàm `calc_monthly_series_by_*`
                # qua tham số `months_filter`).
                sel_thang_list = st.session_state.filter_thang
                n_thang = len(sel_thang_list)
                is_all_months = n_thang >= 2
                unit_df = pd.DataFrame()

                def _ct_grid_cols(n: int) -> list:
                    """Dựng danh sách cột lưới (3/hàng) đủ cho n biểu đồ — dùng cho các nhánh đa lựa
                    chọn (2..8 chỉ tiêu, v4.11), khác với lưới cố định 3x3 khi chọn đủ 9."""
                    cols = []
                    for _ in range((n + 2) // 3):
                        cols.extend(st.columns(3))
                    return cols

                def _register_kv_click(click_result, current_kv_list: list) -> Optional[str]:
                    """
                    Đọc điểm bấm trên biểu đồ Plotly (nếu có) và trả về nhãn Khu vực tương ứng.
                    Bấm-để-drill-down luôn THAY THẾ (không cộng dồn) lựa chọn Khu vực hiện tại bằng
                    đúng 1 Khu vực vừa bấm — khác với các nút Slicer KHU VỰC (đa lựa chọn cộng dồn).
                    """
                    pts = click_result.get("selection", {}).get("points", []) if click_result is not None else []
                    if not pts:
                        return None
                    lbl = pts[0].get("y", "") or ""
                    kv_parsed = lbl.split(" (")[0].strip()
                    return kv_parsed if kv_parsed and kv_parsed not in current_kv_list else None

                if not has_kv and not has_ct and is_all_months:
                    # "Tất cả các tháng" (chưa chọn Khu vực/Khối chỉ tiêu): so sánh biến động 6 Khu vực qua từng tháng
                    st.markdown(
                        """
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">📅</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                So sánh biến động hàng tháng — 6 Khu vực
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    monthly_kv_df = calc_monthly_series_by_kv(calc_df, months_filter=sel_thang_list)
                    ch1, ch2 = st.columns(2)
                    with ch1:
                        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                        st.plotly_chart(
                            build_multi_bar_trend_chart(monthly_kv_df, "kv_label", "pct_dat", "% ĐẠT KPIs THEO THÁNG — 6 KHU VỰC", "%"),
                            width="stretch",
                            key="fig_trend_kv_dat",
                        )
                        st.markdown("</div>", unsafe_allow_html=True)
                    with ch2:
                        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                        st.plotly_chart(
                            build_multi_bar_trend_chart(monthly_kv_df, "kv_label", "pct_thoa", "% THỎA LPT THEO THÁNG — 6 KHU VỰC", "%"),
                            width="stretch",
                            key="fig_trend_kv_thoa",
                        )
                        st.markdown("</div>", unsafe_allow_html=True)

                elif not has_kv and not has_ct:
                    # Mặc định (chưa kết hợp bộ lọc nào): 3 biểu đồ toàn hàng (Donut + 2 Bar theo 6 Khu vực)
                    ch1, ch2, ch3 = st.columns([1.1, 1.45, 1.45])
                    with ch1:
                        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                        fig_donut = build_donut_chart(metrics["donut_kpi_dat"], metrics["donut_kpi_khong_dat"])
                        st.plotly_chart(fig_donut, width="stretch", key="fig_donut_overview")
                        st.markdown("</div>", unsafe_allow_html=True)
                    with ch2:
                        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                        fig_bar_kpi = build_kv_kpi_bar(metrics["kv_charts"])
                        st.plotly_chart(fig_bar_kpi, width="stretch", key="fig_bar_kpi")
                        st.markdown("</div>", unsafe_allow_html=True)
                    with ch3:
                        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                        fig_bar_lpt = build_kv_lpt_bar(metrics["kv_charts"])
                        st.plotly_chart(fig_bar_lpt, width="stretch", key="fig_bar_lpt")
                        st.markdown("</div>", unsafe_allow_html=True)

                elif is_all_ct and not has_kv and is_all_months:
                    # "Tất cả chỉ tiêu" + "Tất cả các tháng" — chưa chọn Khu vực: lưới 6 biểu đồ so sánh 6 Khu vực qua từng tháng
                    st.markdown(
                        """
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">📅</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                Tất cả 9 Khối chỉ tiêu (A - I) — So sánh biến động hàng tháng theo 6 Khu vực
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    monthly_kv_df = calc_monthly_series_by_kv(calc_df, months_filter=sel_thang_list)
                    _kv_grid_labels = list(dict.fromkeys(monthly_kv_df["kv_label"]))
                    render_shared_group_legend(_kv_grid_labels, [get_kv_color(lbl, i) for i, lbl in enumerate(_kv_grid_labels)])
                    grid_cols = st.columns(3) + st.columns(3) + st.columns(3)
                    _sparse_khoi_ghi = {"Chi tiết dư nợ", "Chi tiết huy động", "Chi tiết TOI"}
                    for idx, (label, tmeta) in enumerate(CHI_TIEU_TREND_META.items()):
                        with grid_cols[idx]:
                            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                            st.plotly_chart(
                                build_trend_chart_by_type(monthly_kv_df, "kv_label", tmeta, show_legend=False, value_col_key="kv_key"),
                                width="stretch",
                                key=f"fig_all_ct_trend_kv_{idx}",
                            )
                            if label in _sparse_khoi_ghi and (monthly_kv_df[tmeta["kv_key"]] != 0).mean() < 0.3:
                                st.caption("Dữ liệu hiện chỉ có ở 1 vài tháng gần nhất, các tháng khác sẽ được bổ sung sau")
                            st.markdown("</div>", unsafe_allow_html=True)

                elif is_multi_ct and not has_kv and is_all_months:
                    # v4.11: N chỉ tiêu đã chọn (2-8) + "Tất cả các tháng" — chưa chọn Khu vực: lưới N
                    # biểu đồ so sánh 6 Khu vực qua từng tháng (chỉ đúng các chỉ tiêu đã chọn).
                    st.markdown(
                        f"""
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">📅</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                {n_ct} Khối chỉ tiêu đã chọn — So sánh biến động hàng tháng theo 6 Khu vực
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    monthly_kv_df = calc_monthly_series_by_kv(calc_df, months_filter=sel_thang_list)
                    _kv_grid_labels = list(dict.fromkeys(monthly_kv_df["kv_label"]))
                    render_shared_group_legend(_kv_grid_labels, [get_kv_color(lbl, i) for i, lbl in enumerate(_kv_grid_labels)])
                    grid_cols = _ct_grid_cols(n_ct)
                    _sparse_khoi_ghi = {"Chi tiết dư nợ", "Chi tiết huy động", "Chi tiết TOI"}
                    for idx, label in enumerate(sel_ct_ordered):
                        tmeta = CHI_TIEU_TREND_META[label]
                        with grid_cols[idx]:
                            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                            st.plotly_chart(
                                build_trend_chart_by_type(monthly_kv_df, "kv_label", tmeta, show_legend=False, value_col_key="kv_key"),
                                width="stretch",
                                key=f"fig_multi_ct_trend_kv_{idx}",
                            )
                            if label in _sparse_khoi_ghi and (monthly_kv_df[tmeta["kv_key"]] != 0).mean() < 0.3:
                                st.caption("Dữ liệu hiện chỉ có ở 1 vài tháng gần nhất, các tháng khác sẽ được bổ sung sau")
                            st.markdown("</div>", unsafe_allow_html=True)

                elif is_all_ct and not has_kv:
                    # "Tất cả chỉ tiêu" — chưa chọn Khu vực: lưới 6 biểu đồ theo 6 Khu vực, bấm 1 cột để drill-down
                    st.markdown(
                        """
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">🎯</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                Tất cả 9 Khối chỉ tiêu (A - I) — theo 6 Khu vực (bấm vào 1 cột Khu vực để xem chi tiết theo Đơn vị kinh doanh)
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    clicked_kv_grid = None

                    # Hàng 1: Bảng A + Bảng B (2 cột rộng)
                    row_ab = st.columns(2)
                    for i, label in enumerate(["Tỷ lệ đạt KPIs", "Tỷ lệ thỏa LPT"]):
                        with row_ab[i]:
                            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                            click_result = st.plotly_chart(
                                CHI_TIEU_CHART_MAP[label](metrics["kv_charts"]),
                                width="stretch",
                                key=f"fig_all_ct_kv_{i}",
                                on_select="rerun",
                                selection_mode="points",
                            )
                            st.markdown("</div>", unsafe_allow_html=True)
                            if clicked_kv_grid is None:
                                clicked_kv_grid = _register_kv_click(click_result, sel_kv_list)

                    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

                    # Hàng 2: Bảng C — Stacked Bar + Pie cạnh nhau (không cần cuộn xuống mới thấy Pie)
                    col_stack, col_pie = st.columns([1.6, 1])
                    with col_stack:
                        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                        click_result = st.plotly_chart(
                            CHI_TIEU_CHART_MAP["Phân bổ dải LPT thực nhận"](metrics["kv_charts"]),
                            width="stretch",
                            key="fig_all_ct_kv_2",
                            on_select="rerun",
                            selection_mode="points",
                        )
                        st.markdown("</div>", unsafe_allow_html=True)
                        if clicked_kv_grid is None:
                            clicked_kv_grid = _register_kv_click(click_result, sel_kv_list)
                    with col_pie:
                        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                        st.plotly_chart(
                            build_lpt_distribution_pie(calc_df, st.session_state.filter_thang),
                            width="stretch",
                            key="fig_all_ct_lpt_pie",
                        )
                        st.markdown("</div>", unsafe_allow_html=True)

                    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

                    # Hàng 3: Bảng D + Bảng E + Bảng F
                    row_def = st.columns(3)
                    for i, label in enumerate(["KHDN mới/ĐVKD", "KHTD tăng ròng/ĐVKD", "KHDN tăng ròng/ĐVKD"]):
                        with row_def[i]:
                            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                            click_result = st.plotly_chart(
                                CHI_TIEU_CHART_MAP[label](metrics["kv_charts"]),
                                width="stretch",
                                key=f"fig_all_ct_kv_{3 + i}",
                                on_select="rerun",
                                selection_mode="points",
                            )
                            st.markdown("</div>", unsafe_allow_html=True)
                            if clicked_kv_grid is None:
                                clicked_kv_grid = _register_kv_click(click_result, sel_kv_list)

                    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

                    # Hàng 4: Khối G + Khối H + Khối I
                    row_ghi = st.columns(3)
                    for i, label in enumerate(["Chi tiết dư nợ", "Chi tiết huy động", "Chi tiết TOI"]):
                        with row_ghi[i]:
                            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                            click_result = st.plotly_chart(
                                CHI_TIEU_CHART_MAP[label](metrics["kv_charts"]),
                                width="stretch",
                                key=f"fig_all_ct_kv_{6 + i}",
                                on_select="rerun",
                                selection_mode="points",
                            )
                            st.markdown("</div>", unsafe_allow_html=True)
                            if clicked_kv_grid is None:
                                clicked_kv_grid = _register_kv_click(click_result, sel_kv_list)

                    if clicked_kv_grid:
                        st.session_state.filter_kv = [clicked_kv_grid]
                        st.rerun()

                elif is_multi_ct and not has_kv:
                    # v4.11: N chỉ tiêu đã chọn (2-8) — chưa chọn Khu vực: lưới N biểu đồ theo 6 Khu
                    # vực (chỉ đúng các chỉ tiêu đã chọn), bấm 1 cột để drill-down. Dùng lưới đơn giản
                    # 3 cột/hàng thay vì bố cục đặc biệt 4-hàng của "Tất cả chỉ tiêu" (A-B rộng, C+Pie
                    # cạnh nhau...) — bố cục đặc biệt đó chỉ áp dụng khi chọn ĐỦ 9 chỉ tiêu.
                    st.markdown(
                        f"""
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">🎯</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                {n_ct} Khối chỉ tiêu đã chọn — theo 6 Khu vực (bấm vào 1 cột Khu vực để xem chi tiết theo Đơn vị kinh doanh)
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    clicked_kv_grid = None
                    grid_cols = _ct_grid_cols(n_ct)
                    for idx, label in enumerate(sel_ct_ordered):
                        with grid_cols[idx]:
                            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                            click_result = st.plotly_chart(
                                CHI_TIEU_CHART_MAP[label](metrics["kv_charts"]),
                                width="stretch",
                                key=f"fig_multi_ct_kv_{idx}",
                                on_select="rerun",
                                selection_mode="points",
                            )
                            st.markdown("</div>", unsafe_allow_html=True)
                            if clicked_kv_grid is None:
                                clicked_kv_grid = _register_kv_click(click_result, sel_kv_list)

                    if clicked_kv_grid:
                        st.session_state.filter_kv = [clicked_kv_grid]
                        st.rerun()

                elif is_all_ct and is_all_months:
                    # "Tất cả chỉ tiêu" + "Tất cả các tháng" — đã chọn Khu vực: lưới 9 biểu đồ so sánh ĐVKD (hoặc Nhân sự nếu có chọn) qua từng tháng
                    has_ns = bool(st.session_state.filter_nhan_su)
                    trend_scope_label = "Nhân sự" if has_ns else "ĐVKD"
                    st.markdown(
                        f"""
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">📅</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                Tất cả 9 Khối chỉ tiêu (A - I) — {sel_kv} — So sánh biến động hàng tháng theo {trend_scope_label}
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    if has_ns:
                        monthly_unit_df = calc_monthly_series_by_staff(
                            calc_df, sel_kv_list, st.session_state.filter_ma_dv, st.session_state.filter_nhan_su, months_filter=sel_thang_list
                        )
                        group_col_trend = "staff_label"
                    else:
                        monthly_unit_df = calc_monthly_series_by_unit(
                            calc_df, sel_kv_list, st.session_state.filter_ma_dv, st.session_state.filter_nhan_su, months_filter=sel_thang_list
                        )
                        group_col_trend = "unit_label"
                    if monthly_unit_df.empty:
                        st.info("Không có dữ liệu Đơn vị kinh doanh/Nhân sự cho Khu vực và kỳ đang chọn.")
                    else:
                        _unit_grid_labels = list(dict.fromkeys(monthly_unit_df[group_col_trend]))
                        render_shared_group_legend(_unit_grid_labels, [get_unit_color(i) for i in range(len(_unit_grid_labels))])
                        grid_cols = st.columns(3) + st.columns(3) + st.columns(3)
                        _sparse_khoi_ghi = {"Chi tiết dư nợ", "Chi tiết huy động", "Chi tiết TOI"}
                        for idx, (label, tmeta) in enumerate(CHI_TIEU_TREND_META.items()):
                            with grid_cols[idx]:
                                st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                                st.plotly_chart(
                                    build_trend_chart_by_type(monthly_unit_df, group_col_trend, tmeta, show_legend=False),
                                    width="stretch",
                                    key=f"fig_all_ct_trend_unit_{idx}",
                                )
                                if label in _sparse_khoi_ghi and (monthly_unit_df[tmeta["unit_key"]] != 0).mean() < 0.3:
                                    st.caption("Dữ liệu hiện chỉ có ở 1 vài tháng gần nhất, các tháng khác sẽ được bổ sung sau")
                                st.markdown("</div>", unsafe_allow_html=True)

                elif is_multi_ct and is_all_months:
                    # v4.11: N chỉ tiêu đã chọn (2-8) + "Tất cả các tháng" — đã chọn Khu vực: lưới N
                    # biểu đồ so sánh ĐVKD (hoặc Nhân sự nếu có chọn) qua từng tháng.
                    has_ns = bool(st.session_state.filter_nhan_su)
                    trend_scope_label = "Nhân sự" if has_ns else "ĐVKD"
                    st.markdown(
                        f"""
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">📅</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                {n_ct} Khối chỉ tiêu đã chọn — {sel_kv} — So sánh biến động hàng tháng theo {trend_scope_label}
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    if has_ns:
                        monthly_unit_df = calc_monthly_series_by_staff(
                            calc_df, sel_kv_list, st.session_state.filter_ma_dv, st.session_state.filter_nhan_su, months_filter=sel_thang_list
                        )
                        group_col_trend = "staff_label"
                    else:
                        monthly_unit_df = calc_monthly_series_by_unit(
                            calc_df, sel_kv_list, st.session_state.filter_ma_dv, st.session_state.filter_nhan_su, months_filter=sel_thang_list
                        )
                        group_col_trend = "unit_label"
                    if monthly_unit_df.empty:
                        st.info("Không có dữ liệu Đơn vị kinh doanh/Nhân sự cho Khu vực và kỳ đang chọn.")
                    else:
                        _unit_grid_labels = list(dict.fromkeys(monthly_unit_df[group_col_trend]))
                        render_shared_group_legend(_unit_grid_labels, [get_unit_color(i) for i in range(len(_unit_grid_labels))])
                        grid_cols = _ct_grid_cols(n_ct)
                        _sparse_khoi_ghi = {"Chi tiết dư nợ", "Chi tiết huy động", "Chi tiết TOI"}
                        for idx, label in enumerate(sel_ct_ordered):
                            tmeta = CHI_TIEU_TREND_META[label]
                            with grid_cols[idx]:
                                st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                                st.plotly_chart(
                                    build_trend_chart_by_type(monthly_unit_df, group_col_trend, tmeta, show_legend=False),
                                    width="stretch",
                                    key=f"fig_multi_ct_trend_unit_{idx}",
                                )
                                if label in _sparse_khoi_ghi and (monthly_unit_df[tmeta["unit_key"]] != 0).mean() < 0.3:
                                    st.caption("Dữ liệu hiện chỉ có ở 1 vài tháng gần nhất, các tháng khác sẽ được bổ sung sau")
                                st.markdown("</div>", unsafe_allow_html=True)

                elif is_all_ct:
                    # "Tất cả chỉ tiêu" — đã chọn Khu vực: lưới 6 biểu đồ theo từng Đơn vị kinh doanh (MA_DV)
                    unit_df = get_unit_detail_table(
                        calc_df, sel_kv_list, st.session_state.filter_thang, sel_ct,
                        st.session_state.filter_ma_dv, st.session_state.filter_nhan_su,
                    )
                    st.markdown(
                        f"""
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">🎯</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                Tất cả 9 Khối chỉ tiêu (A - I) — {sel_kv} (theo từng Đơn vị kinh doanh)
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    if unit_df.empty:
                        st.info("Không có dữ liệu Đơn vị kinh doanh cho Khu vực và kỳ đang chọn.")
                    else:
                        grid_cols = st.columns(3) + st.columns(3) + st.columns(3)
                        for idx, (label, meta) in enumerate(CHI_TIEU_UNIT_META.items()):
                            with grid_cols[idx]:
                                st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                                st.plotly_chart(build_unit_chart(unit_df, meta, sel_kv), width="stretch", key=f"fig_all_ct_unit_{idx}")
                                st.markdown("</div>", unsafe_allow_html=True)

                elif is_multi_ct:
                    # v4.11: N chỉ tiêu đã chọn (2-8) — đã chọn Khu vực: lưới N biểu đồ theo từng
                    # Đơn vị kinh doanh (MA_DV), chỉ đúng các chỉ tiêu đã chọn.
                    unit_df = get_unit_detail_table(
                        calc_df, sel_kv_list, st.session_state.filter_thang, sel_ct,
                        st.session_state.filter_ma_dv, st.session_state.filter_nhan_su,
                    )
                    st.markdown(
                        f"""
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">🎯</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                {n_ct} Khối chỉ tiêu đã chọn — {sel_kv} (theo từng Đơn vị kinh doanh)
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    if unit_df.empty:
                        st.info("Không có dữ liệu Đơn vị kinh doanh cho Khu vực và kỳ đang chọn.")
                    else:
                        grid_cols = _ct_grid_cols(n_ct)
                        for idx, label in enumerate(sel_ct_ordered):
                            meta = CHI_TIEU_UNIT_META[label]
                            with grid_cols[idx]:
                                st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                                st.plotly_chart(build_unit_chart(unit_df, meta, sel_kv), width="stretch", key=f"fig_multi_ct_unit_{idx}")
                                st.markdown("</div>", unsafe_allow_html=True)

                elif not has_kv and is_all_months:
                    # Chỉ chọn 1 KHỐI CHỈ TIÊU + "Tất cả các tháng": so sánh biến động 6 Khu vực qua từng tháng
                    tmeta = CHI_TIEU_TREND_META.get(sel_ct, DEFAULT_TREND_META)
                    st.markdown(
                        f"""
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">📅</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                Biểu đồ kết hợp — {sel_ct} — So sánh biến động hàng tháng theo 6 Khu vực
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    monthly_kv_df = calc_monthly_series_by_kv(calc_df, months_filter=sel_thang_list)
                    st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                    st.plotly_chart(
                        build_trend_chart_by_type(monthly_kv_df, "kv_label", tmeta, title_suffix="THEO THÁNG — 6 KHU VỰC", value_col_key="kv_key"),
                        width="stretch",
                        key="fig_focus_trend_kv",
                    )
                    st.markdown("</div>", unsafe_allow_html=True)

                elif not has_kv:
                    # Chỉ chọn 1 KHỐI CHỈ TIÊU: biểu đồ kết hợp theo 6 Khu vực — bấm cột để drill-down
                    st.markdown(
                        f"""
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">🎯</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                Biểu đồ kết hợp — {sel_ct} (bấm vào 1 cột Khu vực để xem chi tiết theo Đơn vị kinh doanh)
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    focus_chart_fn = CHI_TIEU_CHART_MAP.get(sel_ct)
                    if sel_ct == "Phân bổ dải LPT thực nhận":
                        col_stack, col_pie = st.columns([1.6, 1])
                        with col_stack:
                            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                            click_result = st.plotly_chart(
                                focus_chart_fn(metrics["kv_charts"]),
                                width="stretch",
                                key="fig_focus_chi_tieu",
                                on_select="rerun",
                                selection_mode="points",
                            )
                            st.markdown("</div>", unsafe_allow_html=True)
                        with col_pie:
                            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                            st.plotly_chart(
                                build_lpt_distribution_pie(calc_df, st.session_state.filter_thang),
                                width="stretch",
                                key="fig_focus_lpt_pie",
                            )
                            st.markdown("</div>", unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                        click_result = st.plotly_chart(
                            focus_chart_fn(metrics["kv_charts"]),
                            width="stretch",
                            key="fig_focus_chi_tieu",
                            on_select="rerun",
                            selection_mode="points",
                        )
                        st.markdown("</div>", unsafe_allow_html=True)

                    clicked_kv = _register_kv_click(click_result, sel_kv_list)
                    if clicked_kv:
                        st.session_state.filter_kv = [clicked_kv]
                        st.rerun()

                elif is_all_months:
                    # Có chọn KHU VỰC (+ có thể 1 Khối chỉ tiêu) + "Tất cả các tháng": so sánh biến động ĐVKD (hoặc Nhân sự nếu có chọn) qua từng tháng
                    tmeta = CHI_TIEU_TREND_META.get(sel_ct, DEFAULT_TREND_META)
                    ct_label = sel_ct if has_ct else f"{tmeta['title']} (mặc định)"
                    has_ns = bool(st.session_state.filter_nhan_su)
                    trend_scope_label = "Nhân sự" if has_ns else "ĐVKD"
                    st.markdown(
                        f"""
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">📅</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                Biểu đồ kết hợp — {sel_kv} • {ct_label} — So sánh biến động hàng tháng theo {trend_scope_label}
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    if has_ns:
                        monthly_unit_df = calc_monthly_series_by_staff(
                            calc_df, sel_kv_list, st.session_state.filter_ma_dv, st.session_state.filter_nhan_su, months_filter=sel_thang_list
                        )
                        group_col_trend = "staff_label"
                    else:
                        monthly_unit_df = calc_monthly_series_by_unit(
                            calc_df, sel_kv_list, st.session_state.filter_ma_dv, st.session_state.filter_nhan_su, months_filter=sel_thang_list
                        )
                        group_col_trend = "unit_label"
                    st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                    if monthly_unit_df.empty:
                        st.info("Không có dữ liệu Đơn vị kinh doanh/Nhân sự cho Khu vực và kỳ đang chọn.")
                    else:
                        st.plotly_chart(
                            build_trend_chart_by_type(monthly_unit_df, group_col_trend, tmeta, title_suffix=f"THEO THÁNG — {sel_kv}"),
                            width="stretch",
                            key="fig_unit_trend_combined",
                        )
                    st.markdown("</div>", unsafe_allow_html=True)

                else:
                    # Có chọn KHU VỰC + 1 Khối chỉ tiêu cụ thể: biểu đồ theo từng Đơn vị kinh doanh (MA_DV)
                    meta = CHI_TIEU_UNIT_META.get(sel_ct, DEFAULT_UNIT_META)
                    unit_df = get_unit_detail_table(
                        calc_df, sel_kv_list, st.session_state.filter_thang, sel_ct,
                        st.session_state.filter_ma_dv, st.session_state.filter_nhan_su,
                    )
                    ct_label = sel_ct if has_ct else f"{meta['title']} (mặc định)"
                    st.markdown(
                        f"""
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:2px;">
                            <span style="font-size:13px;">🎯</span>
                            <span style="font-size:12.5px; font-weight:800; color:#111827; text-transform:uppercase; letter-spacing:0.3px;">
                                Biểu đồ kết hợp — {sel_kv} • {ct_label} (theo từng Đơn vị kinh doanh)
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                    if unit_df.empty:
                        st.info("Không có dữ liệu Đơn vị kinh doanh cho Khu vực và kỳ đang chọn.")
                    else:
                        st.plotly_chart(build_unit_chart(unit_df, meta, sel_kv), width="stretch", key="fig_unit_combined")
                    st.markdown("</div>", unsafe_allow_html=True)

                # ── BIỂU ĐỒ ĐƯỜNG (LINE CHART): XU HƯỚNG THEO TỪNG THÁNG THỰC TẾ ────────
                # Luôn hiển thị (không phụ thuộc Slicer SỐ THÁNG vì đây là góc nhìn xuyên suốt thời gian),
                # tôn trọng Slicer KHU VỰC/ĐVKD/Nhân sự nếu có chọn — bổ sung loại biểu đồ khác (đường)
                # bên cạnh cột/donut.
                st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)
                trend_df = calc_monthly_trend(
                    calc_df, sel_kv_list, st.session_state.filter_ma_dv, st.session_state.filter_nhan_su
                )
                if not trend_df.empty and len(trend_df) >= 2:
                    st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                    trend_title = sel_kv if has_kv else "Toàn hàng"
                    st.plotly_chart(build_trend_line_chart(trend_df, trend_title), width="stretch", key="fig_trend_line")
                    st.markdown("</div>", unsafe_allow_html=True)

                # ── BẢNG CHI TIẾT ĐƠN VỊ KINH DOANH — SỐ LIỆU ĐẦY ĐỦ (KẾT HỢP KHU VỰC + KHỐI CHỈ TIÊU) ──
                if has_kv or has_ct:
                    st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)
                    if unit_df.empty:
                        unit_df = get_unit_detail_table(
                            calc_df, sel_kv_list, st.session_state.filter_thang, sel_ct,
                            st.session_state.filter_ma_dv, st.session_state.filter_nhan_su,
                        )
                    thang_display = ", ".join(sel_thang_list) if sel_thang_list else "Toàn bộ dữ liệu"
                    kv_display = sel_kv if has_kv else "Toàn hàng (6 Khu vực)"
                    if is_all_ct:
                        ct_display = "Tất cả 9 chỉ tiêu"
                    elif is_multi_ct:
                        ct_display = f"{n_ct} chỉ tiêu đã chọn"
                    elif has_ct:
                        ct_display = sel_ct
                    else:
                        ct_display = "(Chưa chọn Khối chỉ tiêu)"
                    n_dv_display = max(len(unit_df) - 1, 0) if not unit_df.empty else 0
                    n_dv_selected = len(st.session_state.filter_ma_dv)
                    dv_display = f" • {n_dv_selected} ĐVKD đã chọn" if n_dv_selected > 0 else ""
                    with st.expander(
                        f"🏢 **Chi tiết Đơn vị kinh doanh — {kv_display} • {ct_display}** "
                        f"({n_dv_display} ĐVKD • Kỳ: {thang_display}{dv_display})",
                        expanded=True,
                    ):
                        if n_dv_selected > 0:
                            st.caption(f"📌 Đang lọc riêng {n_dv_selected} Đơn vị kinh doanh đã chọn trong Slicer KHU VỰC.")
                        if n_ct == 1 and sel_ct in CHI_TIEU_UNIT_META:
                            focus_meta = CHI_TIEU_UNIT_META[sel_ct]
                            st.caption(f"📌 Đã sắp xếp và tô sáng theo cột **{focus_meta['column']}** tương ứng khối chỉ tiêu đang chọn.")
                        elif is_all_ct:
                            st.caption("📌 Đang xem Tất cả chỉ tiêu — bảng sắp xếp mặc định theo Khu vực / Tên ĐVKD (xem đủ 9 chỉ tiêu ở các cột bên phải).")
                        elif is_multi_ct:
                            st.caption(f"📌 Đang xem {n_ct} chỉ tiêu đã chọn — bảng sắp xếp mặc định theo Khu vực / Tên ĐVKD (không tô sáng riêng cột nào khi chọn nhiều hơn 1 chỉ tiêu).")
                        if unit_df.empty:
                            st.info("Không có dữ liệu Đơn vị kinh doanh cho Khu vực và kỳ đang chọn.")
                        else:
                            st.dataframe(style_unit_detail_table(unit_df, sel_ct), width="stretch", hide_index=True)

            # ── TAB 2: BÁO CÁO 9 KHỐI CHỈ TIÊU (Theo đặc tả công thức) ─────────────
            with tab_tables:
                st.subheader("📋 Báo Cáo Chi Tiết 9 Khối Chỉ Tiêu (Đặc tả công thức)")

                if st.session_state.uploaded_df is None:
                    st.info("💡 Đang hiển thị bảng mẫu với dữ liệu mô phỏng 7 kỳ dữ liệu. Tải lên file Excel để tính toán số liệu thực tế.")

                def _fmt_int(v):
                    try:
                        return f"{int(v):,}".replace(",", ".")
                    except (ValueError, TypeError):
                        return v

                def _fmt_pct(v):
                    try:
                        return f"{float(v):.1f}%"
                    except (ValueError, TypeError):
                        return v

                def _fmt_float(v):
                    try:
                        return f"{float(v):,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    except (ValueError, TypeError):
                        return v

                def format_and_style(df_block: pd.DataFrame, pct_keys: list[str]):
                    df_fmt = df_block.copy()
                    for col in df_fmt.columns:
                        gl = str(col[0]) if isinstance(col, tuple) else str(col)
                        if any(pk in gl for pk in pct_keys):
                            df_fmt[col] = df_fmt[col].apply(_fmt_pct)
                        elif "/ ĐVKD" in gl:
                            df_fmt[col] = df_fmt[col].apply(_fmt_float)
                        else:
                            df_fmt[col] = df_fmt[col].apply(_fmt_int)

                    def highlight_total(row):
                        if row.name == "Tổng":
                            return ["background-color: #fff3cd; font-weight: 700; color: #856404;"] * len(row)
                        return [""] * len(row)

                    return df_fmt.style.apply(highlight_total, axis=1)

                # Mở rộng khối dựa theo Slicer KHỐI CHỈ TIÊU — v4.11: filter_chi_tieu là list (đa lựa
                # chọn); chưa chọn gì (list rỗng) = mở rộng tất cả (hành vi cũ), có chọn thì chỉ mở
                # rộng đúng (các) khối đã chọn.
                sel_ct_list_tab2 = st.session_state.filter_chi_tieu
                exp_k1 = not sel_ct_list_tab2 or "Tỷ lệ đạt KPIs" in sel_ct_list_tab2
                exp_k2 = not sel_ct_list_tab2 or "Tỷ lệ thỏa LPT" in sel_ct_list_tab2
                exp_k3 = not sel_ct_list_tab2 or "Phân bổ dải LPT thực nhận" in sel_ct_list_tab2
                exp_k4 = not sel_ct_list_tab2 or "KHDN mới/ĐVKD" in sel_ct_list_tab2
                exp_k5 = not sel_ct_list_tab2 or "KHTD tăng ròng/ĐVKD" in sel_ct_list_tab2
                exp_k6 = not sel_ct_list_tab2 or "KHDN tăng ròng/ĐVKD" in sel_ct_list_tab2
                exp_k7 = not sel_ct_list_tab2 or "Chi tiết dư nợ" in sel_ct_list_tab2
                exp_k8 = not sel_ct_list_tab2 or "Chi tiết huy động" in sel_ct_list_tab2
                exp_k9 = not sel_ct_list_tab2 or "Chi tiết TOI" in sel_ct_list_tab2

                _thang_display_tab2 = ", ".join(st.session_state.filter_thang) if st.session_state.filter_thang else "Toàn bộ dữ liệu"
                chart_caption = (
                    f"📊 Biểu đồ theo bộ lọc SỐ THÁNG đang chọn: **{_thang_display_tab2}** "
                    "(luôn hiển thị đủ 6 Khu vực, không lọc theo Slicer Khu vực)."
                )

                # Render 9 Khối có expander, mỗi khối kèm 1 biểu đồ phù hợp với bản chất chỉ tiêu
                with st.expander("📌 **Khối 1: Tỷ lệ đạt KPIs** (Bảng A - Theo từng tháng)", expanded=exp_k1):
                    st.caption(chart_caption)
                    st.plotly_chart(build_kv_kpi_bar(metrics["kv_charts"]), width="stretch", key="fig_block1_kv")
                    st.dataframe(format_and_style(b1, ["% nhân sự đạt KPIs", "% Đạt"]), width="stretch")

                with st.expander("📌 **Khối 2: Tỷ lệ thỏa LPT** (Bảng B - Theo từng tháng)", expanded=exp_k2):
                    st.caption(chart_caption)
                    st.plotly_chart(build_kv_lpt_bar(metrics["kv_charts"]), width="stretch", key="fig_block2_kv")
                    st.dataframe(format_and_style(b2, ["% nhân sự thỏa LPT", "% Thỏa"]), width="stretch")

                with st.expander("📌 **Khối 3: Phân bổ dải Hệ số LPT thực nhận** (Bảng C - 6 dải chuẩn hóa)", expanded=exp_k3):
                    st.caption(chart_caption)
                    st.caption("Dòng có Hệ số LPT thực nhận rỗng/NULL không tính vào bất kỳ dải nào (gộp vào 'Chưa xác định' trên biểu đồ).")
                    col_stack, col_pie = st.columns([1.6, 1])
                    with col_stack:
                        st.plotly_chart(build_lpt_distribution_chart(metrics["kv_charts"]), width="stretch", key="fig_block3_kv")
                    with col_pie:
                        st.plotly_chart(build_lpt_distribution_pie(calc_df, st.session_state.filter_thang), width="stretch", key="fig_block3_pie")
                    st.dataframe(format_and_style(b3, []), width="stretch")

                with st.expander("📌 **Khối 4: KHDN mới bình quân / ĐVKD** (Bảng D)", expanded=exp_k4):
                    st.caption(chart_caption)
                    st.plotly_chart(build_khdn_moi_bar(metrics["kv_charts"]), width="stretch", key="fig_block4_kv")
                    st.dataframe(format_and_style(b4, []), width="stretch")

                with st.expander("📌 **Khối 5: KHTD tăng ròng bình quân / ĐVKD** (Bảng E)", expanded=exp_k5):
                    st.caption(chart_caption)
                    st.plotly_chart(build_khtd_tang_bar(metrics["kv_charts"]), width="stretch", key="fig_block5_kv")
                    st.dataframe(format_and_style(b5, []), width="stretch")

                with st.expander("📌 **Khối 6: KHDN tăng ròng bình quân / ĐVKD** (Bảng F)", expanded=exp_k6):
                    st.caption(chart_caption)
                    st.plotly_chart(build_khdn_tang_bar(metrics["kv_charts"]), width="stretch", key="fig_block6_kv")
                    st.dataframe(format_and_style(b6, []), width="stretch")

                with st.expander("📌 **Khối 7: Chi tiết dư nợ** (Bảng G - Theo từng tháng)", expanded=exp_k7):
                    st.caption(chart_caption)
                    st.caption("⚠️ Số liệu cấp ĐVKD, đã khử trùng lặp theo (Mã ĐVKD, Ngày chốt) trước khi cộng — không nhân sai theo số nhân sự.")
                    st.plotly_chart(build_khoi_g_bar(metrics["kv_charts"]), width="stretch", key="fig_block7_kv")
                    st.dataframe(format_and_style(b7, []), width="stretch")

                with st.expander("📌 **Khối 8: Chi tiết huy động** (Bảng H - Theo từng tháng)", expanded=exp_k8):
                    st.caption(chart_caption)
                    st.caption("⚠️ Số liệu cấp ĐVKD, đã khử trùng lặp theo (Mã ĐVKD, Ngày chốt) trước khi cộng — không nhân sai theo số nhân sự.")
                    st.plotly_chart(build_khoi_h_bar(metrics["kv_charts"]), width="stretch", key="fig_block8_kv")
                    st.dataframe(format_and_style(b8, []), width="stretch")

                with st.expander("📌 **Khối 9: Chi tiết TOI** (Bảng I - 10 nguồn thu + Tổng TOI, Theo từng tháng)", expanded=exp_k9):
                    st.caption(chart_caption)
                    st.caption("⚠️ Số liệu cấp ĐVKD, đã khử trùng lặp theo (Mã ĐVKD, Ngày chốt) trước khi cộng — không nhân sai theo số nhân sự.")
                    st.plotly_chart(build_khoi_i_bar(metrics["kv_charts"]), width="stretch", key="fig_block9_kv")
                    st.dataframe(format_and_style(b9, []), width="stretch")

            # ── TAB 3: XUẤT FILE EXCEL BÁO CÁO TRỌN BỘ ──────────────────────────────
            with tab_export:
                st.subheader("📥 Xuất Báo Cáo Excel Trọn Bộ 9 Khối")
                st.write("Nhấn nút bên dưới để tải file Excel đã tính toán và định dạng tự động gồm 9 sheet tương ứng với 9 khối chỉ tiêu theo đúng đặc tả công thức.")

                blocks_to_export = {
                    "1-Ty le dat KPIs": b1,
                    "2-Ty le thoa LPT": b2,
                    "3-Phan bo dai LPT": b3,
                    "4-KHDN moi BQ-DVKD": b4,
                    "5-KHTD tang rong BQ": b5,
                    "6-KHDN tang rong BQ": b6,
                    "7-Chi tiet du no": b7,
                    "8-Chi tiet huy dong": b8,
                    "9-Chi tiet TOI": b9,
                }

                excel_data = export_blocks_to_excel(blocks_to_export)

                st.download_button(
                    label="📊 Tải Xuống File Báo Cáo Excel (.xlsx)",
                    data=excel_data,
                    file_name="HDBank_Bao_Cao_KPIs_LPT_9_Khoi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.document",
                )

    with top_tab_dvkd:
        has_dvkd_data = (st.session_state.uploaded_dvkd_items is not None) or st.session_state.use_demo
        if not has_dvkd_data:
            st.info("📂 Chưa có dữ liệu 'Data import số ĐVKD'. Vui lòng tải lên tại đây để xem Dashboard số liệu ĐVKD:")
            inline_dvkd_file = st.file_uploader(
                "Chọn file Excel chứa sheet 'Data import số ĐVKD':",
                type=["xlsx"],
                key="inline_dvkd_uploader",
            )
            if inline_dvkd_file is not None:
                try:
                    with st.spinner("Đang xử lý dữ liệu số liệu ĐVKD..."):
                        items = load_and_preprocess_dvkd_excel(inline_dvkd_file.read())
                        st.session_state.uploaded_dvkd_items = items
                        st.session_state.uploaded_file_name_dvkd = inline_dvkd_file.name
                        st.success(f"✅ Đã tải thành công {len(items):,} dòng dữ liệu ĐVKD!")
                        st.rerun()
                except Exception as ex:
                    st.error(f"❌ Lỗi xử lý file Excel: {ex}")
        else:
            dvkd_items = (
                st.session_state.uploaded_dvkd_items
                if st.session_state.uploaded_dvkd_items is not None
                else generate_demo_dvkd_items()
            )

            dvkd_header_html = f"""
            <div class="hdbank-header">
                <div class="hdbank-brand">{logo_img_html}</div>
                <div class="hdbank-title-center">
                    <h1 class="hdbank-main-title">DASHBOARD SỐ LIỆU ĐVKD</h1>
                    <div class="hdbank-sub-title">Thực hiện / Kế hoạch / % hoàn thành theo từng chỉ tiêu, xem theo Toàn hàng hoặc từng Đơn vị kinh doanh</div>
                </div>
            </div>
            """
            st.markdown(dvkd_header_html, unsafe_allow_html=True)
            st.info(f"📂 **Nguồn dữ liệu:** {st.session_state.uploaded_file_name_dvkd or 'Dữ liệu mẫu'} • {len(dvkd_items):,} dòng dữ liệu")

            # (v4.15) Danh sách Đơn vị kinh doanh duy nhất, giữ đúng thứ tự xuất hiện trong file —
            # validate/chốt lựa chọn đang lưu TRƯỚC khi lọc dvkd_items, để hộp CHỈ TIÊU bên dưới chỉ
            # liệt kê đúng các chỉ tiêu thuộc Đơn vị đang chọn.
            dvkd_don_vi_list = list(dict.fromkeys(it.get("dvkd", "Toàn hàng") for it in dvkd_items)) or ["Toàn hàng"]
            if st.session_state.filter_dvkd_don_vi not in dvkd_don_vi_list:
                st.session_state.filter_dvkd_don_vi = (
                    "Toàn hàng" if "Toàn hàng" in dvkd_don_vi_list else dvkd_don_vi_list[0]
                )

            # ── 3 FILTER: ĐƠN VỊ + THÁNG (đa lựa chọn) + CHỈ TIÊU ─────────────────
            col_dv_dvkd, col_thang_dvkd, col_ct_dvkd = st.columns([1.2, 1.4, 2])
            with col_dv_dvkd:
                with st.container(border=True):
                    st.markdown(
                        """
                        <div class="slicer-card-header slicer-card-header-kv">
                            <span>ĐƠN VỊ</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    dv_default_idx = dvkd_don_vi_list.index(st.session_state.filter_dvkd_don_vi)
                    sel_dv = st.selectbox(
                        "Chọn đơn vị:",
                        options=dvkd_don_vi_list,
                        index=dv_default_idx,
                        key="dvkd_donvi_selectbox",
                        on_change=lambda: st.session_state.update(
                            filter_dvkd_don_vi=st.session_state["dvkd_donvi_selectbox"],
                            active_top_tab=1,
                        ),
                    )

            # Lọc theo Đơn vị kinh doanh đang chọn TRƯỚC khi liệt kê chỉ tiêu — mỗi Đơn vị có 1 tập
            # row_idx RIÊNG (khóa toàn cục duy nhất, xem load_and_preprocess_dvkd_excel()) dù cùng
            # nhãn chỉ tiêu lặp lại ở mọi Đơn vị.
            dvkd_items_scoped = [it for it in dvkd_items if it.get("dvkd", "Toàn hàng") == sel_dv]
            item_by_idx = {it["row_idx"]: it for it in dvkd_items_scoped}

            with col_thang_dvkd:
                with st.container(border=True):
                    st.markdown(
                        """
                        <div class="slicer-card-header slicer-card-header-thang">
                            <span>THÁNG (chọn 1 hoặc nhiều — bỏ trống = cả năm)</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    row1_m, row2_m = st.columns(6), st.columns(6)
                    for i, cols in enumerate([row1_m, row2_m]):
                        for j in range(6):
                            m = i * 6 + j + 1
                            with cols[j]:
                                is_sel_m = m in st.session_state.filter_dvkd_thang
                                if st.button(
                                    f"T{m:02d}",
                                    type="primary" if is_sel_m else "secondary",
                                    key=f"btn_dvkd_m_{m}",
                                    width="stretch",
                                ):
                                    if is_sel_m:
                                        st.session_state.filter_dvkd_thang.remove(m)
                                    else:
                                        st.session_state.filter_dvkd_thang.append(m)
                                    st.session_state.active_top_tab = 1
                                    st.rerun()

            with col_ct_dvkd:
                with st.container(border=True):
                    st.markdown(
                        """
                        <div class="slicer-card-header slicer-card-header-ct">
                            <span>CHỈ TIÊU (đa lựa chọn — bỏ trống = hiển thị tất cả)</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    options = [it["row_idx"] for it in dvkd_items_scoped]

                    def _fmt_dvkd_item(idx):
                        it = item_by_idx[idx]
                        sec = it["section"].replace("A. ", "").replace("B. ", "")
                        return f"[{sec}] {it['label']}" if sec else it["label"]

                    if not options:
                        st.warning("Không có chỉ tiêu nào trong dữ liệu.")
                        sel_indices = []
                    else:
                        # (v4.18) Đổi Đơn vị làm 1 số row_idx đang chọn không còn hợp lệ (row_idx là
                        # khóa TOÀN CỤC riêng cho từng Đơn vị) — đa lựa chọn không có "1 vị trí" duy
                        # nhất để cố map lại theo nhãn như bản selectbox-đơn cũ, nên chỉ PRUNE (loại
                        # bỏ) các row_idx không còn hợp lệ, giữ lại đúng phần còn valid — cùng cách
                        # filter_ma_dv/filter_nhan_su tự prune khi đổi phạm vi Khu vực ở tab KPIs.
                        # `dict.fromkeys()` khử trùng lặp (giữ đúng thứ tự) — phòng lỗi StreamlitDuplicateElementKey
                        # (key biểu đồ/bảng dùng chính row_idx) nếu session_state["filter_dvkd_chi_tieu"]
                        # từng bị lưu trùng row_idx do bất kỳ nguyên nhân nào (vd đổi file dữ liệu mới
                        # trong lúc đang có lựa chọn cũ, hoặc tương tác đồng thời với nhiều widget).
                        valid_selected = list(dict.fromkeys(i for i in st.session_state.filter_dvkd_chi_tieu if i in item_by_idx))
                        if valid_selected != st.session_state.filter_dvkd_chi_tieu:
                            st.session_state.filter_dvkd_chi_tieu = valid_selected
                            # Đồng bộ TRỰC TIẾP vào key riêng của widget TRƯỚC khi widget được tạo ở
                            # dưới (an toàn — khác việc gán SAU khi widget đã tồn tại, xem Lesson #16) —
                            # nếu không, giá trị cũ (đã invalid) vẫn còn kẹt trong session_state[key] từ
                            # lần render trước và ĐÈ mất tham số `default=` bên dưới.
                            st.session_state["dvkd_ct_multiselect"] = valid_selected

                        if st.button(
                            "📋 Tất cả chỉ tiêu",
                            key="dvkd_ct_all_btn",
                            width="stretch",
                            type="secondary",
                        ):
                            options_dedup = list(dict.fromkeys(options))
                            st.session_state.filter_dvkd_chi_tieu = options_dedup
                            st.session_state["dvkd_ct_multiselect"] = options_dedup
                            st.session_state.active_top_tab = 1
                            st.rerun()

                        # (v4.18) Streamlit CẢNH BÁO khi 1 widget vừa được gán qua Session State API
                        # (nhánh pruning/nút "Tất cả" phía trên) vừa nhận `default=` trong CÙNG 1 lượt
                        # render — nên chỉ truyền `default=` khi widget CHƯA TỪNG được tạo (key chưa
                        # tồn tại trong session_state, vd lần render đầu tiên của cả phiên, hoặc khi
                        # test tự set `filter_dvkd_chi_tieu` trước khi widget kịp render lần nào) —
                        # những lượt SAU, widget tự đọc đúng giá trị đã lưu trong session_state[key],
                        # không cần `default=` nữa (khác hẳn cách `st.selectbox` dùng `index=`, không
                        # bị policy này áp dụng — xem Lesson #16).
                        multiselect_kwargs = dict(
                            options=options,
                            format_func=_fmt_dvkd_item,
                            key="dvkd_ct_multiselect",
                            placeholder="Bỏ trống = hiển thị tất cả",
                            on_change=lambda: st.session_state.update(
                                filter_dvkd_chi_tieu=list(dict.fromkeys(st.session_state["dvkd_ct_multiselect"])),
                                active_top_tab=1,
                            ),
                        )
                        if "dvkd_ct_multiselect" not in st.session_state:
                            multiselect_kwargs["default"] = valid_selected
                        sel_indices_raw = st.multiselect("Chọn chỉ tiêu (gõ để tìm kiếm):", **multiselect_kwargs)
                        sel_indices = sel_indices_raw if sel_indices_raw else options

            st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)

            # (v4.18) CHỈ TIÊU nay là đa lựa chọn — lặp qua từng chỉ tiêu đã chọn (bỏ trống =
            # `sel_indices` đã fallback về TOÀN BỘ `options` ở trên), mỗi chỉ tiêu 1 khối bảng/thẻ +
            # biểu đồ riêng, phân cách bằng tiêu đề + đường kẻ. `dict.fromkeys()` khử trùng lặp lần
            # cuối TRƯỚC khi render — key của mỗi bảng/biểu đồ dùng chính row_idx
            # (`f"fig_dvkd_{row_idx}"`), row_idx trùng nhau trong `sel_items` sẽ ném
            # `StreamlitDuplicateElementKey` (đã gặp thực tế — xem Mục 6 mục 22), nên đây là lớp
            # phòng thủ CUỐI CÙNG, độc lập với việc dedup ở các bước phía trên có bắt hết hay không.
            sel_indices_dedup = list(dict.fromkeys(sel_indices))
            sel_items = [item_by_idx[i] for i in sel_indices_dedup if i in item_by_idx]
            sel_months = sorted(st.session_state.filter_dvkd_thang) or list(range(1, 13))

            # (v4.20) BẢNG THỐNG KÊ toàn bộ chỉ tiêu của Đơn vị đang chọn — ĐỘC LẬP với filter CHỈ
            # TIÊU (filter đó giờ chỉ còn điều khiển phần biểu đồ bên dưới, không lọc bảng nữa),
            # mirror đúng layout hàng/cột của file Excel gốc, giữ nguyên số liệu, tô màu theo nhóm
            # cột kiểu bảng chuyên nghiệp (xem build_dvkd_summary_table/style_dvkd_summary_table).
            summary_df, section_row_positions = build_dvkd_summary_table(dvkd_items_scoped, sel_months)
            st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
            # (v4.25) st.dataframe() không render được màu <th> header của Styler dù dùng
            # set_table_styles/map_index (giới hạn cứng của Streamlit) — render HTML thuần thay thế,
            # xem render_dvkd_table_html().
            st.markdown(render_dvkd_table_html(summary_df, section_row_positions), unsafe_allow_html=True)

            st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)

            if not sel_items:
                st.info("Chưa chọn chỉ tiêu nào để hiển thị biểu đồ.")
            else:
                for sel_item in sel_items:
                    st.markdown(
                        f"<div style='font-weight:700; font-size:14px; color:#1e3a8a; "
                        f"margin: 12px 0 4px 0; padding-left:4px; "
                        f"border-left: 3px solid #6366f1;'>{sel_item['label']}</div>",
                        unsafe_allow_html=True,
                    )
                    months_labels = [f"T{m:02d}" for m in sel_months]
                    thuc_hien_vals = [sel_item["thuc_hien"].get(m) or 0 for m in sel_months]
                    ke_hoach_vals = [sel_item["ke_hoach"].get(m) or 0 for m in sel_months]
                    pct_ht_vals = [sel_item["pct_ht"].get(m) for m in sel_months]
                    chart_type3 = _map_dvkd_chart_type_3(sel_item.get("chart_type", "combo"))
                    st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                    st.plotly_chart(
                        build_dvkd_chart_by_type3(
                            chart_type3, months_labels, thuc_hien_vals, ke_hoach_vals, pct_ht_vals, sel_item["label"], sel_item.get("dvt", "")
                        ),
                        width="stretch",
                        key=f"fig_dvkd_{sel_item['row_idx']}",
                    )
                    st.markdown("</div>", unsafe_allow_html=True)
                    st.markdown("<hr style='margin: 8px 0; border-color:#e2e8f0;'>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()