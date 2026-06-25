import streamlit as st
import pandas as pd
import pickle
import time
import io
import base64

# Tambahkan ini
import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

from docx import Document as DocxDocument
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ==========================================
# 1. KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(
    page_title="Portal Evaluasi Studi Akademik",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# 2. ICON SET
# ==========================================
BARCHART_ICON  = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="20" x2="12" y2="10"/><line x1="18" y1="20" x2="18" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>"""
IDCARD_ICON    = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="5" width="20" height="14" rx="2"/><circle cx="8" cy="11" r="2"/><line x1="14" y1="9" x2="19" y2="9"/><line x1="14" y1="13" x2="19" y2="13"/><line x1="6" y1="16" x2="10" y2="16"/></svg>"""
CLIPBOARD_ICON = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 2h6a1 1 0 0 1 1 1v2H8V3a1 1 0 0 1 1-1Z"/><line x1="8" y1="11" x2="16" y2="11"/><line x1="8" y1="15" x2="16" y2="15"/></svg>"""
TABLE_ICON     = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/><line x1="9" y1="4" x2="9" y2="20"/></svg>"""
TARGET_ICON    = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="4"/><circle cx="12" cy="12" r="0.6" fill="{color}"/></svg>"""
CHECK_ICON     = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>"""
ALERT_ICON     = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0Z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>"""
CALENDAR_ICON  = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>"""
LOGO_SVG       = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3 2 8l10 5 10-5-10-5Z"/><path d="M6 10.5V16c0 1.2 2.7 2.5 6 2.5s6-1.3 6-2.5v-5.5"/><path d="M22 8v6"/></svg>"""
ZAPPER_ICON    = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>"""
BULB_ICON      = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="9" y1="18" x2="15" y2="18"/><line x1="10" y1="22" x2="14" y2="22"/><path d="M15.09 14c.18-.98.65-1.74 1.41-2.5A4.65 4.65 0 0 0 18 8 6 6 0 0 0 6 8c0 1 .23 2.23 1.5 3.5A4.61 4.61 0 0 1 8.91 14"/></svg>"""
USER_ICON      = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>"""
DOWNLOAD_ICON  = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>"""
SHIELD_ICON    = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>"""
BOOK_ICON      = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/></svg>"""
COIN_ICON      = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M14.8 9A2 2 0 0 0 13 8h-2a2 2 0 0 0 0 4h2a2 2 0 0 1 0 4h-2a2 2 0 0 1-1.8-1"/><line x1="12" y1="6" x2="12" y2="8"/><line x1="12" y1="16" x2="12" y2="18"/></svg>"""
PRINT_ICON     = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>"""
WORD_ICON      = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>"""
BACK_ICON      = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 12H5"/><polyline points="12 19 5 12 12 5"/></svg>"""

# ==========================================
# 3. CSS
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,400&family=Space+Grotesk:wght@300;400;500;600;700&family=Fira+Code:wght@400;500;600&display=swap');

:root {
    --ink:#0D1117; --ink-2:#1C2333; --muted:#64748B; --muted-2:#94A3B8;
    --paper:#F0F4FA; --surface:#FFFFFF;
    --blue:#3B5BDB; --blue-dark:#2F4AC0; --blue-bg:#EEF2FF;
    --violet:#7048E8; --violet-bg:#F3F0FF;
    --safe:#0D9488; --safe-bg:#CCFBF1;
    --risk:#DC2626; --risk-bg:#FEE2E2;
    --amber:#D97706; --amber-bg:#FEF3C7;
    --hairline:#E2E8F0; --radius:16px;
    --font-head:'Space Grotesk',sans-serif;
    --font-body:'Plus Jakarta Sans',sans-serif;
    --font-code:'Fira Code',monospace;
}
html,body,[class*="css"]{font-family:var(--font-body)!important;-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility;}
.stApp{background-color:var(--paper);}
header[data-testid="stHeader"]{height:0;visibility:hidden;}
div[data-testid="stDecoration"]{display:none;}
footer{visibility:hidden;}
#MainMenu{visibility:hidden;}
.block-container{padding-top:1rem!important;padding-bottom:3rem!important;padding-left:2.4rem!important;padding-right:2.4rem!important;max-width:100%!important;}

/* ══════════════════════════════════════════════
   GLOBAL FIX — sembunyikan "Press Enter to submit
   form" di SEMUA input & number input di seluruh app
   ══════════════════════════════════════════════ */
[data-testid="InputInstructions"] {
    display: none !important;
    visibility: hidden !important;
    height: 0 !important;
    overflow: hidden !important;
}
/* Pastikan wrapper input tidak punya padding bawah
   yang nampung teks tersebut */
div[data-testid="stTextInput"] > div,
div[data-testid="stNumberInput"] > div > div {
    overflow: visible !important;
}

.st-key-topbar{position:fixed!important;top:0;left:0;right:0;z-index:999999;background:rgba(255,255,255,0.94);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);border-bottom:1.5px solid var(--hairline);box-shadow:0 1px 16px rgba(15,23,42,0.05);padding:10px 2.4rem!important;}
div[data-testid="stPopover"] button{background:var(--blue-bg)!important;border:1.5px solid #c5d0fa!important;border-radius:10px!important;color:var(--blue)!important;font-weight:600!important;font-family:var(--font-body)!important;}
div[data-testid="stPopover"] button:hover{background:#dde4ff!important;}
div[data-testid="stPopover"] button p{color:var(--blue)!important;font-family:var(--font-body)!important;}

/* ── LOGIN CARD ── */
.st-key-login_card_wrapper .stTextInput input{
    padding:11px 14px!important;
    font-size:14px!important;
    font-family:var(--font-body)!important;
    border-radius:12px!important;
    border:1.5px solid var(--hairline)!important;
    background:#F8FAFC!important;
    color:var(--ink)!important;
    transition:border-color .18s,box-shadow .18s;
}
.st-key-login_card_wrapper .stTextInput input:focus{
    border-color:var(--blue)!important;
    background:#fff!important;
    box-shadow:0 0 0 3px rgba(59,91,219,0.10)!important;
}
.st-key-login_card_wrapper .stTextInput label{
    font-size:13px!important;
    font-weight:600!important;
    font-family:var(--font-body)!important;
    color:var(--ink)!important;
}
/* Ikon mata show/hide password — rapikan posisinya */
.st-key-login_card_wrapper .stTextInput [data-testid="textInputRootElement"] input {
    padding-right: 44px !important;
}
.st-key-login_card_wrapper .stTextInput button[data-testid="textInput-ShowHideButton"] {
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    color: #94A3B8 !important;
}

/* Tombol submit login — gradient biru-ungu */
.st-key-login_card_wrapper .stFormSubmitButton button,
.st-key-login_card_wrapper .stFormSubmitButton button:hover,
.st-key-login_card_wrapper .stFormSubmitButton button:focus,
.st-key-login_card_wrapper .stFormSubmitButton button:active {
    background: linear-gradient(135deg, #3B5BDB 0%, #7048E8 100%) !important;
    border: none !important;
    border-radius: 50px !important;
    color: #fff !important;
    font-weight: 700 !important;
    font-size: 15px !important;
    font-family: var(--font-body) !important;
    padding: 12px 0 !important;
    width: 100% !important;
    box-shadow: 0 4px 20px rgba(59,91,219,0.30) !important;
    margin-top: 8px;
    transition: opacity .18s !important;
}
.st-key-login_card_wrapper .stFormSubmitButton button:hover { opacity: .88 !important; }

.section-label{font-family:var(--font-code)!important;font-size:10.5px;font-weight:600;color:var(--blue);text-transform:uppercase;letter-spacing:1.4px;border-left:3px solid var(--blue);padding-left:10px;margin:4px 0 18px 0;display:flex;align-items:center;gap:8px;}

div[data-testid="stVerticalBlockBorderWrapper"]{background:#ffffff!important;border-radius:20px!important;border:1.5px solid var(--hairline)!important;box-shadow:0 2px 16px rgba(15,23,42,0.05)!important;}

.stTextInput input,.stNumberInput input{padding:10px 14px!important;font-size:14px!important;font-family:var(--font-body)!important;border-radius:10px!important;border:1.5px solid var(--hairline)!important;background:#F8FAFC!important;transition:border-color .18s,box-shadow .18s;}
.stTextInput input:focus,.stNumberInput input:focus{border-color:var(--blue)!important;box-shadow:0 0 0 3px rgba(59,91,219,0.09)!important;background:#fff!important;}
.stSelectbox div[data-baseweb="select"]>div{border-radius:10px!important;border:1.5px solid var(--hairline)!important;min-height:42px!important;background:#F8FAFC!important;font-family:var(--font-body)!important;}
.stTextInput label,.stNumberInput label,.stSelectbox label{font-size:13px!important;font-weight:600!important;font-family:var(--font-body)!important;color:var(--ink)!important;}

/* Semua tombol primary — gradient biru-ungu (BUKAN merah) */
.stButton button[kind="primary"],
.stFormSubmitButton button[kind="primary"],
.stButton button[kind="primary"]:hover,
.stFormSubmitButton button[kind="primary"]:hover {
    background: linear-gradient(135deg, #3B5BDB 0%, #7048E8 100%) !important;
    border: none !important;
    border-radius: 50px !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    font-family: var(--font-body) !important;
    color: #fff !important;
    padding: 10px 0 !important;
    box-shadow: 0 3px 14px rgba(59,91,219,0.25) !important;
    transition: opacity .18s !important;
}
.stButton button[kind="primary"]:hover,
.stFormSubmitButton button[kind="primary"]:hover { opacity: .88 !important; }

.stButton button[kind="secondary"]{border-radius:50px!important;border:1.5px solid var(--hairline)!important;color:var(--ink)!important;font-weight:600!important;font-family:var(--font-body)!important;background:#fff!important;transition:border-color .18s,color .18s;}
.stButton button[kind="secondary"]:hover{border-color:var(--blue)!important;color:var(--blue)!important;}

div[data-testid="stMetric"]{background:#ffffff!important;border:1.5px solid var(--hairline)!important;border-radius:16px!important;padding:18px 20px!important;}
div[data-testid="stMetricLabel"]{font-family:var(--font-code)!important;font-size:10px!important;letter-spacing:0.8px;color:var(--muted)!important;text-transform:uppercase;}
div[data-testid="stMetricValue"]{font-family:var(--font-head)!important;font-size:26px!important;font-weight:700!important;color:var(--ink)!important;}

.stDownloadButton button{border-radius:50px!important;font-weight:600!important;font-family:var(--font-body)!important;border:1.5px solid var(--hairline)!important;background:#fff!important;color:var(--ink)!important;transition:all .18s;}
.stDownloadButton button:hover{border-color:var(--blue)!important;color:var(--blue)!important;}

.stamp{display:inline-flex;align-items:center;gap:8px;font-family:var(--font-code)!important;font-weight:600;font-size:11.5px;letter-spacing:0.6px;text-transform:uppercase;padding:10px 20px;border-radius:50px;border:none;}
.stamp-safe{color:var(--safe);background:var(--safe-bg);}
.stamp-risk{color:var(--risk);background:var(--risk-bg);}
.stRadio label{font-size:13px!important;font-weight:500!important;font-family:var(--font-body)!important;}
div[data-testid="stDataFrame"]{border-radius:14px;overflow:hidden;}

/* Back button custom */
.st-key-back_btn button {
    background: linear-gradient(135deg, #3B5BDB 0%, #7048E8 100%) !important;
    border: none !important;
    border-radius: 50px !important;
    color: #ffffff !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    font-family: var(--font-body) !important;
    padding: 10px 0 !important;
    box-shadow: 0 3px 14px rgba(59,91,219,0.30) !important;
    transition: opacity .18s !important;
}
.st-key-back_btn button:hover { opacity: .88 !important; }

/* Export panel */
.export-panel{background:linear-gradient(135deg,#0D1117 0%,#1E1B4B 100%);border-radius:20px;padding:24px 28px;border:1px solid rgba(255,255,255,0.07);box-shadow:0 8px 32px rgba(15,23,42,0.2);}

/* ── MENU POPOVER — sembunyikan label teks tombol navigasi,
       biarkan hanya HTML card custom yang terlihat ── */
div[data-testid="stPopover"] [data-testid="stVerticalBlock"] .stButton button {
    opacity: 0 !important;
    height: 0 !important;
    padding: 0 !important;
    margin: 0 !important;
    border: none !important;
    min-height: 0 !important;
    overflow: hidden !important;
    pointer-events: auto !important;
    position: absolute !important;
    width: 100% !important;
}
div[data-testid="stPopover"] [data-testid="stVerticalBlock"] .stButton {
    position: relative !important;
    margin-top: -52px !important;
    z-index: 10 !important;
}
/* Popover container ukuran & shadow */
div[data-testid="stPopover"] > div[data-testid="stPopoverBody"],
div[data-testid="stPopoverBody"] {
    min-width: 260px !important;
    border-radius: 18px !important;
    border: 1.5px solid #E2E8F0 !important;
    box-shadow: 0 8px 32px rgba(15,23,42,0.12) !important;
    padding: 16px !important;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 4. LOAD MODEL
# ==========================================
@st.cache_resource
def load_model():
    with open("model_dropout_terbaik.pkl", "rb") as f:
        return pickle.load(f)

try:
    model = load_model()
except FileNotFoundError:
    st.error("File model `model_dropout_terbaik.pkl` tidak ditemukan.")
    st.stop()

# ==========================================
# 5. SESSION STATE
# ==========================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "history" not in st.session_state:
    st.session_state.history = []
if "page" not in st.session_state:
    st.session_state.page = "input"
if "selected_student_uid" not in st.session_state:
    st.session_state.selected_student_uid = ""

def switch_page(p):
    st.session_state.page = p
    st.rerun()

# ==========================================
# HELPER — Rekomendasi
# ==========================================
def generate_recommendations(data):
    recs = []
    is_safe = data["SkorAman"] >= 75
    if data["SPP"] == "Menunggak":
        recs.append({"title":"Pembayaran SPP Bermasalah","desc":"Mahasiswa memiliki tunggakan SPP. Rekomendasikan konsultasi keuangan atau pengajuan keringanan biaya segera.","level":"KRITIS"})
    if data["IPK1"] < 2.0 or data["IPK2"] < 2.0:
        recs.append({"title":"IPK di Bawah Standar Minimum","desc":"IPK semester berada di bawah 2.0. Wajibkan sesi konseling akademik dan bimbingan intensif dengan dosen wali.","level":"KRITIS"})
    if data["IPK2"] < data["IPK1"] and abs(data["IPK2"] - data["IPK1"]) > 0.3:
        recs.append({"title":"Tren IPK Menurun Signifikan","desc":f"IPK turun {round(data['IPK1']-data['IPK2'],2)} poin dari semester 1 ke 2. Monitor beban studi dan kondisi psikososial mahasiswa.","level":"WASPADA"})
    if data["SKS2"] < data["SKS1"] and (data["SKS1"] - data["SKS2"]) >= 4:
        recs.append({"title":"Pengambilan SKS Menurun Drastis","desc":"Jumlah SKS semester 2 turun drastis. Identifikasi hambatan akademik atau pribadi yang menyebabkan penurunan.","level":"WASPADA"})
    if is_safe and data["SkorRisiko"] < 20:
        recs.append({"title":"Profil Retensi Sangat Kuat","desc":"Mahasiswa menunjukkan performa optimal. Pertimbangkan program akselerasi atau beasiswa prestasi.","level":"BAIK"})
    if not recs:
        lbl = "BAIK" if is_safe else "WASPADA"
        desc = "Tidak ada indikator risiko spesifik. Pertahankan pemantauan rutin setiap akhir semester." if is_safe else "Model mendeteksi risiko tinggi. Jadwalkan wawancara personal dengan mahasiswa."
        recs.append({"title":"Evaluasi Lanjutan Diperlukan","desc":desc,"level":lbl})
    return recs

def factor_score(data):
    ipk_avg = (data["IPK1"] + data["IPK2"]) / 2
    return {
        "IPK Rata-rata": min(round(ipk_avg / 4.0 * 100), 100),
        "Beban SKS": min(round((data["TotalSKS"] / 48) * 100), 100),
        "Kelancaran SPP": 100 if data["SPP"] == "Lancar" else 20,
        "Tren Nilai": min(max(round(((data["IPK2"] - data["IPK1"]) + 1) / 2 * 100), 0), 100),
    }

# ==========================================
# HELPER — Chart generator (returns PNG bytes)
# ==========================================
def make_ipk_chart(ipk1, ipk2):
    fig, ax = plt.subplots(figsize=(4.5, 2.6))
    fig.patch.set_facecolor("#F8FAFF")
    ax.set_facecolor("#F8FAFF")
    sems = ["Semester 1", "Semester 2"]
    vals = [ipk1, ipk2]
    color = "#3B5BDB"
    ax.plot(sems, vals, color=color, linewidth=2.5, marker="o", markersize=8, markerfacecolor="#fff", markeredgewidth=2.5, markeredgecolor=color)
    ax.fill_between(sems, vals, alpha=0.10, color=color)
    for i, (s, v) in enumerate(zip(sems, vals)):
        ax.annotate(f"{v:.2f}", (s, v), textcoords="offset points", xytext=(0, 10), ha="center", fontsize=10, fontweight="bold", color=color)
    ax.set_ylim(0, 4.2)
    ax.set_ylabel("IPK", fontsize=9, color="#64748B")
    ax.tick_params(colors="#64748B", labelsize=9)
    for spine in ax.spines.values():
        spine.set_edgecolor("#E2E8F0")
    ax.grid(axis="y", color="#E2E8F0", linestyle="--", linewidth=0.8)
    plt.tight_layout(pad=0.8)
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()

def make_sks_chart(sks1, sks2):
    fig, ax = plt.subplots(figsize=(4.5, 2.6))
    fig.patch.set_facecolor("#F8F7FF")
    ax.set_facecolor("#F8F7FF")
    sems = ["Semester 1", "Semester 2"]
    vals = [sks1, sks2]
    colors = ["#7048E8", "#A78BFA"]
    bars = ax.bar(sems, vals, color=colors, width=0.45, edgecolor="white", linewidth=1.5, zorder=3)
    for bar, v in zip(bars, vals):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3, str(v), ha="center", va="bottom", fontsize=10, fontweight="bold", color="#7048E8")
    ax.set_ylim(0, max(vals) * 1.3 + 2)
    ax.set_ylabel("SKS", fontsize=9, color="#64748B")
    ax.tick_params(colors="#64748B", labelsize=9)
    for spine in ax.spines.values():
        spine.set_edgecolor("#E2E8F0")
    ax.grid(axis="y", color="#E2E8F0", linestyle="--", linewidth=0.8, zorder=0)
    plt.tight_layout(pad=0.8)
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()

def make_risiko_gauge(risk_pct, safe_pct):
    fig, ax = plt.subplots(figsize=(4.5, 2.4))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    cats = ["Risiko Dropout", "Peluang Lulus"]
    vals = [risk_pct, safe_pct]
    colors = ["#DC2626", "#0D9488"]
    bars = ax.barh(cats, vals, color=colors, height=0.4, edgecolor="white", linewidth=1)
    for bar, v in zip(bars, vals):
        ax.text(min(v + 1, 97), bar.get_y() + bar.get_height()/2, f"{v:.1f}%", va="center", fontsize=10, fontweight="bold", color="#0D1117")
    ax.set_xlim(0, 100)
    ax.set_xlabel("Persentase (%)", fontsize=9, color="#64748B")
    ax.tick_params(colors="#64748B", labelsize=9)
    for spine in ax.spines.values():
        spine.set_edgecolor("#E2E8F0")
    ax.grid(axis="x", color="#E2E8F0", linestyle="--", linewidth=0.8)
    plt.tight_layout(pad=0.8)
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()

def make_factor_chart(scores):
    fig, ax = plt.subplots(figsize=(5.5, 2.8))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    labels = list(scores.keys())
    vals = list(scores.values())
    colors = ["#0D9488" if v >= 70 else "#D97706" if v >= 40 else "#DC2626" for v in vals]
    bars = ax.barh(labels, vals, color=colors, height=0.42, edgecolor="white", linewidth=1)
    for bar, v in zip(bars, vals):
        ax.text(min(v + 1, 97), bar.get_y() + bar.get_height()/2, f"{v}%", va="center", fontsize=9, fontweight="bold", color="#0D1117")
    ax.set_xlim(0, 100)
    ax.set_xlabel("Skor (%)", fontsize=9, color="#64748B")
    ax.tick_params(colors="#64748B", labelsize=9)
    for spine in ax.spines.values():
        spine.set_edgecolor("#E2E8F0")
    ax.grid(axis="x", color="#E2E8F0", linestyle="--", linewidth=0.8)
    plt.tight_layout(pad=0.8)
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()

# ==========================================
# HELPER — Set cell background color
# ==========================================
def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.replace("#",""))
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top","left","bottom","right"):
        if edge in kwargs:
            el = OxmlElement(f"w:{edge}")
            for k,v in kwargs[edge].items():
                el.set(qn(f"w:{k}"),v)
            tcBorders.append(el)
    tcPr.append(tcBorders)

# ==========================================
# HELPER — Generate HTML for print/PDF
# ==========================================
def generate_html_report(active_data):
    recs = generate_recommendations(active_data)
    scores = factor_score(active_data)
    _sa = active_data["SkorAman"]
    if _sa >= 75:
        status_color = "#0D9488"; status_bg = "#CCFBF1"
    elif _sa >= 40:
        status_color = "#D97706"; status_bg = "#FEF3C7"
    else:
        status_color = "#DC2626"; status_bg = "#FEE2E2"

    ipk_b64  = base64.b64encode(make_ipk_chart(active_data["IPK1"],   active_data["IPK2"])).decode()
    sks_b64  = base64.b64encode(make_sks_chart(active_data["SKS1"],   active_data["SKS2"])).decode()
    risk_b64 = base64.b64encode(make_risiko_gauge(active_data["SkorRisiko"], active_data["SkorAman"])).decode()
    fact_b64 = base64.b64encode(make_factor_chart(scores)).decode()

    level_style = {
        "KRITIS":  ("DC2626","FEE2E2","FECACA"),
        "WASPADA": ("D97706","FEF3C7","FDE68A"),
        "BAIK":    ("0D9488","CCFBF1","99F6E4"),
    }

    recs_html = ""
    for rec in recs:
        fg, bg, border = level_style.get(rec["level"], ("3B5BDB","EEF2FF","C5D0FA"))
        recs_html += f"""
        <div style="display:flex;gap:14px;align-items:flex-start;background:#{bg};border:1.5px solid #{border};
                    border-radius:10px;padding:12px 16px;margin-bottom:10px;">
            <div style="background:#{bg};border:1px solid #{border};border-radius:8px;padding:4px 10px;
                        color:#{fg};font-weight:700;font-size:9px;white-space:nowrap;letter-spacing:0.6px;">{rec["level"]}</div>
            <div>
                <div style="font-weight:700;font-size:12px;color:#0D1117;margin-bottom:3px;">{rec["title"]}</div>
                <div style="font-size:11px;color:#475569;line-height:1.55;">{rec["desc"]}</div>
            </div>
        </div>"""

    html = f"""<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8"/>
<title>Laporan Evaluasi — {active_data['Nama']}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;700&family=Fira+Code:wght@500&display=swap');
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{font-family:'Plus Jakarta Sans',sans-serif;background:#F0F4FA;color:#0D1117;font-size:12px;}}
  @media print{{
    body{{background:white;}}
    .no-print{{display:none!important;}}
    .page-break{{page-break-before:always;}}
    @page{{margin:1.5cm;size:A4;}}
  }}
  .container{{max-width:900px;margin:0 auto;padding:24px;}}
  .print-btn{{position:fixed;top:20px;right:20px;background:linear-gradient(135deg,#3B5BDB,#7048E8);color:#fff;
              border:none;border-radius:50px;padding:12px 24px;font-family:'Plus Jakarta Sans',sans-serif;
              font-size:13px;font-weight:700;cursor:pointer;box-shadow:0 4px 16px rgba(59,91,219,0.3);z-index:999;}}
  .print-btn:hover{{opacity:.9;}}
  .header{{background:linear-gradient(135deg,#3B5BDB 0%,#1E1B4B 100%);border-radius:16px;
           padding:28px 32px;margin-bottom:24px;display:flex;justify-content:space-between;align-items:center;}}
  .header-left h1{{font-family:'Space Grotesk',sans-serif;font-size:22px;font-weight:700;color:#fff;letter-spacing:-0.5px;}}
  .header-left p{{font-size:11px;color:#C5D0FA;margin-top:4px;font-family:'Fira Code',monospace;letter-spacing:0.5px;}}
  .header-right{{text-align:right;}}
  .header-right .date{{font-family:'Fira Code',monospace;font-size:11px;color:#A5B4FC;}}
  .section{{background:#fff;border-radius:14px;padding:24px;margin-bottom:18px;border:1.5px solid #E2E8F0;}}
  .section-title{{font-family:'Space Grotesk',sans-serif;font-size:13px;font-weight:700;color:#3B5BDB;
                  text-transform:uppercase;letter-spacing:1px;border-left:3px solid #3B5BDB;padding-left:10px;
                  margin-bottom:16px;font-family:'Fira Code',monospace;}}
  .id-grid{{display:grid;grid-template-columns:1fr 1fr;gap:10px;}}
  .id-item{{background:#F8FAFC;border:1.5px solid #E2E8F0;border-radius:10px;padding:10px 14px;}}
  .id-item .lbl{{font-size:9px;color:#94A3B8;text-transform:uppercase;letter-spacing:0.8px;font-family:'Fira Code',monospace;margin-bottom:2px;}}
  .id-item .val{{font-size:14px;font-weight:700;color:#0D1117;font-family:'Space Grotesk',sans-serif;}}
  .status-badge{{display:inline-flex;align-items:center;gap:6px;background:{status_bg};color:{status_color};
                 font-family:'Fira Code',monospace;font-size:11px;font-weight:600;padding:8px 18px;border-radius:50px;
                 letter-spacing:0.5px;text-transform:uppercase;}}
  .chart-row{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:4px;}}
  .chart-box{{background:#F8FAFC;border:1.5px solid #E2E8F0;border-radius:12px;padding:14px;text-align:center;}}
  .chart-box .chart-title{{font-family:'Fira Code',monospace;font-size:9.5px;font-weight:600;color:#3B5BDB;
                            text-transform:uppercase;letter-spacing:0.7px;margin-bottom:8px;}}
  .chart-box img{{max-width:100%;border-radius:8px;}}
  .factor-bar{{margin-bottom:10px;}}
  .factor-bar .fb-label{{display:flex;justify-content:space-between;margin-bottom:4px;font-size:11px;font-weight:600;}}
  .factor-bar .fb-track{{height:8px;background:#F1F5F9;border-radius:99px;overflow:hidden;}}
  .sem-table{{width:100%;border-collapse:collapse;margin-top:4px;}}
  .sem-table th{{background:#3B5BDB;color:#fff;padding:8px 12px;text-align:left;font-size:10px;font-family:'Fira Code',monospace;letter-spacing:0.5px;}}
  .sem-table td{{padding:8px 12px;font-size:11px;border-bottom:1px solid #F1F5F9;}}
  .sem-table tr:last-child td{{border-bottom:none;}}
  .all-table{{width:100%;border-collapse:collapse;font-size:10px;}}
  .all-table th{{background:#1E1B4B;color:#C5D0FA;padding:7px 10px;text-align:left;font-family:'Fira Code',monospace;font-size:9px;letter-spacing:0.5px;}}
  .all-table td{{padding:6px 10px;border-bottom:1px solid #F1F5F9;}}
  .footer{{text-align:center;padding:20px;font-size:10px;color:#94A3B8;font-family:'Fira Code',monospace;}}
</style>
</head>
<body>
<button class="print-btn no-print" onclick="window.print()">🖨 Cetak / Simpan PDF</button>
<div class="container">

  <!-- HEADER -->
  <div class="header">
    <div class="header-left">
      <h1>PortalAkademik</h1>
      <p>LAPORAN EVALUASI STUDI MAHASISWA · PREDIKSI RISIKO DROPOUT</p>
    </div>
    <div class="header-right">
      <div class="date">{time.strftime('%d %B %Y')}</div>
      <div class="date">{time.strftime('%H:%M:%S')} WIB</div>
      <div style="margin-top:10px;"><span class="status-badge">{active_data['Status']}</span></div>
    </div>
  </div>

  <!-- IDENTITAS -->
  <div class="section">
    <div class="section-title">1 · Identitas Mahasiswa</div>
    <div class="id-grid">
      <div class="id-item"><div class="lbl">Nama Lengkap</div><div class="val">{active_data['Nama']}</div></div>
      <div class="id-item"><div class="lbl">NIM</div><div class="val">{active_data['NIM']}</div></div>
      <div class="id-item"><div class="lbl">Status SPP</div>
        <div class="val" style="color:{'#0D9488' if active_data['SPP']=='Lancar' else '#DC2626'};">{active_data['SPP']}</div></div>
      <div class="id-item"><div class="lbl">Rerata IPK</div><div class="val">{active_data['RerataIPK']:.2f}</div></div>
      <div class="id-item"><div class="lbl">Total SKS</div><div class="val">{active_data['TotalSKS']} SKS</div></div>
      <div class="id-item"><div class="lbl">Peluang Lulus</div>
        <div class="val" style="color:#3B5BDB;">{active_data['SkorAman']}%</div></div>
      <div class="id-item"><div class="lbl">Faktor Risiko Dropout</div>
        <div class="val" style="color:#DC2626;">{active_data['SkorRisiko']}%</div></div>
      <div class="id-item"><div class="lbl">Prediksi Model</div>
        <div class="val" style="color:{status_color};">{active_data['Status']}</div></div>
    </div>
  </div>

  <!-- DATA SEMESTER -->
  <div class="section">
    <div class="section-title">2 · Data Akademik per Semester</div>
    <table class="sem-table">
      <thead><tr><th>Indikator</th><th>Semester 1</th><th>Semester 2</th><th>Keterangan</th></tr></thead>
      <tbody>
        <tr style="background:#F8FAFC;">
          <td style="font-weight:600;">IPK (Indeks Prestasi)</td>
          <td style="font-weight:700;color:#3B5BDB;">{active_data['IPK1']:.2f}</td>
          <td style="font-weight:700;color:{'#0D9488' if active_data['IPK2']>=active_data['IPK1'] else '#DC2626'};">{active_data['IPK2']:.2f}</td>
          <td style="color:{'#0D9488' if active_data['IPK2']>=active_data['IPK1'] else '#DC2626'};">
            {'▲ Meningkat' if active_data['IPK2']>active_data['IPK1'] else ('▼ Menurun' if active_data['IPK2']<active_data['IPK1'] else '→ Stabil')}
            {f" ({abs(active_data['IPK2']-active_data['IPK1']):.2f} poin)" if active_data['IPK2']!=active_data['IPK1'] else ""}
          </td>
        </tr>
        <tr>
          <td style="font-weight:600;">Jumlah SKS Diambil</td>
          <td style="font-weight:700;color:#7048E8;">{active_data['SKS1']} SKS</td>
          <td style="font-weight:700;color:{'#0D9488' if active_data['SKS2']>=active_data['SKS1'] else '#DC2626'};">{active_data['SKS2']} SKS</td>
          <td style="color:{'#0D9488' if active_data['SKS2']>=active_data['SKS1'] else '#D97706'};">
            {'▲ Bertambah' if active_data['SKS2']>active_data['SKS1'] else ('▼ Berkurang' if active_data['SKS2']<active_data['SKS1'] else '→ Sama')}
          </td>
        </tr>
      </tbody>
    </table>
  </div>

  <!-- GRAFIK -->
  <div class="section">
    <div class="section-title">3 · Visualisasi & Grafik</div>
    <div class="chart-row">
      <div class="chart-box">
        <div class="chart-title">Fluktuasi IPK per Semester</div>
        <img src="data:image/png;base64,{ipk_b64}" alt="IPK Chart"/>
      </div>
      <div class="chart-box">
        <div class="chart-title">Volume SKS per Semester</div>
        <img src="data:image/png;base64,{sks_b64}" alt="SKS Chart"/>
      </div>
    </div>
    <div class="chart-row" style="margin-top:14px;">
      <div class="chart-box">
        <div class="chart-title">Distribusi Skor Risiko vs Kelulusan</div>
        <img src="data:image/png;base64,{risk_b64}" alt="Risk Chart"/>
      </div>
      <div class="chart-box">
        <div class="chart-title">Indikator Faktor Keputusan Model</div>
        <img src="data:image/png;base64,{fact_b64}" alt="Factor Chart"/>
      </div>
    </div>
  </div>

  <!-- REKOMENDASI -->
  <div class="section">
    <div class="section-title">4 · Rekomendasi Intervensi Akademik</div>
    {recs_html}
  </div>

  <div class="footer">Dokumen ini digenerate otomatis oleh PortalAkademik &mdash; {time.strftime('%d %B %Y, %H:%M:%S')} WIB</div>
</div>
</body>
</html>"""
    return html

# ==========================================
# 6. LOGIN
# ==========================================
if not st.session_state.logged_in:
    _, center_col, _ = st.columns([1, 1.2, 1])
    with center_col:
        with st.container(key="login_card_wrapper"):
            st.markdown(f"""
                <div style="display:flex;flex-direction:column;align-items:center;margin:44px 0 32px 0;">
                    <div style="width:68px;height:68px;border-radius:20px;
                                background:linear-gradient(135deg,#EEF2FF 0%,#F3F0FF 100%);
                                border:1.5px solid #c5d0fa;
                                display:flex;align-items:center;justify-content:center;margin-bottom:16px;
                                box-shadow:0 4px 16px rgba(59,91,219,0.12);">
                        {LOGO_SVG.format(size=32, color="#3B5BDB")}
                    </div>
                    <div style="font-family:'Space Grotesk',sans-serif;font-size:28px;font-weight:700;
                                letter-spacing:-0.8px;line-height:1;margin-bottom:7px;text-align:center;">
                        <span style="color:#3B5BDB;">Portal</span><span style="color:#0D1117;">Akademik</span>
                    </div>
                    <div style="font-size:12.5px;color:#64748B;font-weight:400;letter-spacing:0.2px;font-family:'Plus Jakarta Sans',sans-serif;">
                        Evaluasi Studi &nbsp;·&nbsp; Deteksi Risiko Dropout
                    </div>
                </div>
            """, unsafe_allow_html=True)
            with st.form("login_form"):
                username = st.text_input("Username", placeholder="Masukkan ID pengguna")
                password = st.text_input("Password", type="password", placeholder="Masukkan kata sandi", help="")
                submitted = st.form_submit_button("→  Masuk ke Dashboard", type="primary", use_container_width=True)
            if submitted:
                if username == "admin" and password == "admin123":
                    st.session_state.logged_in = True
                    st.rerun()
                else:
                    st.error("Kredensial salah, periksa kembali username & password.")
            st.markdown("<div style='height:36px;'></div>", unsafe_allow_html=True)

# ==========================================
# 7. PASCA LOGIN
# ==========================================
else:
    with st.container(key="topbar"):
        tb1, tb2, tb3 = st.columns([6, 2, 1])
        with tb1:
            st.markdown(f"""
                <div style="display:flex;align-items:center;gap:12px;">
                    <div style="width:40px;height:40px;border-radius:12px;background:linear-gradient(135deg,#EEF2FF,#F3F0FF);
                                border:1.5px solid #c5d0fa;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
                        {LOGO_SVG.format(size=20, color="#3B5BDB")}
                    </div>
                    <div>
                        <p style="font-family:'Space Grotesk',sans-serif;color:#0D1117;font-size:15.5px;font-weight:700;margin:0;letter-spacing:-0.5px;">
                            <span style="color:#3B5BDB;">Portal</span>Akademik</p>
                        <p style="font-family:'Fira Code',monospace;color:#64748B;font-size:9.5px;text-transform:uppercase;letter-spacing:0.7px;margin:1px 0 0 0;">
                            Prediksi Risiko Dropout &middot; Random Forest</p>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with tb2:
            st.markdown(f"""<div style="display:flex;align-items:center;justify-content:flex-end;gap:6px;height:100%;
                font-family:'Fira Code',monospace;font-size:10.5px;color:#64748B;padding-top:10px;">
                {CALENDAR_ICON.format(size=13, color="#64748B")} {time.strftime('%d %B %Y')}</div>""", unsafe_allow_html=True)
        with tb3:
            with st.popover("☰  Menu", use_container_width=False):
                # Header menu
                st.markdown(f"""
                <div style="padding:4px 2px 14px 2px;border-bottom:1.5px solid #F1F5F9;margin-bottom:14px;">
                    <div style="display:flex;align-items:center;gap:10px;">
                        <div style="width:36px;height:36px;border-radius:10px;
                                    background:linear-gradient(135deg,#EEF2FF,#F3F0FF);
                                    border:1.5px solid #C5D0FA;
                                    display:flex;align-items:center;justify-content:center;flex-shrink:0;">
                            {LOGO_SVG.format(size=18, color="#3B5BDB")}
                        </div>
                        <div>
                            <div style="font-family:'Space Grotesk',sans-serif;font-size:13px;font-weight:700;color:#0D1117;line-height:1.2;">
                                <span style="color:#3B5BDB;">Portal</span>Akademik
                            </div>
                            <div style="font-family:'Fira Code',monospace;font-size:9px;color:#94A3B8;text-transform:uppercase;letter-spacing:0.6px;">
                                v1.0 · Random Forest
                            </div>
                        </div>
                    </div>
                </div>
                <div style="font-family:'Fira Code',monospace;font-size:9px;font-weight:600;color:#94A3B8;
                            text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;padding-left:2px;">
                    Navigasi
                </div>
                """, unsafe_allow_html=True)

                # Tombol Form Input
                _active_input = st.session_state.page == "input"
                _bg_input  = "linear-gradient(135deg,#EEF2FF,#F3F0FF)" if _active_input else "#F8FAFC"
                _bd_input  = "#C5D0FA" if _active_input else "#E2E8F0"
                _col_input = "#3B5BDB" if _active_input else "#64748B"
                _fw_input  = "700" if _active_input else "600"
                _tc_input  = "#3B5BDB" if _active_input else "#0D1117"
                st.markdown(f"""
                <div style="display:flex;align-items:center;gap:10px;width:100%;padding:10px 12px;
                            border-radius:10px;margin-bottom:6px;
                            background:{_bg_input};border:1.5px solid {_bd_input};">
                    {CLIPBOARD_ICON.format(size=15, color=_col_input)}
                    <div>
                        <div style="font-family:'Plus Jakarta Sans',sans-serif;font-size:12.5px;font-weight:{_fw_input};color:{_tc_input};">Form Input Parameter</div>
                        <div style="font-family:'Plus Jakarta Sans',sans-serif;font-size:10px;color:#94A3B8;">Input data mahasiswa baru</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                if st.button("Form Input Parameter", use_container_width=True, key="nav_input"):
                    switch_page("input")

                # Tombol Log Aktivitas
                if len(st.session_state.history) > 0:
                    _active_log = st.session_state.page == "log"
                    _bg_log  = "linear-gradient(135deg,#EEF2FF,#F3F0FF)" if _active_log else "#F8FAFC"
                    _bd_log  = "#C5D0FA" if _active_log else "#E2E8F0"
                    _col_log = "#3B5BDB" if _active_log else "#64748B"
                    _fw_log  = "700" if _active_log else "600"
                    _tc_log  = "#3B5BDB" if _active_log else "#0D1117"
                    _count   = len(st.session_state.history)
                    st.markdown(f"""
                    <div style="display:flex;align-items:center;gap:10px;width:100%;padding:10px 12px;
                                border-radius:10px;margin-bottom:6px;
                                background:{_bg_log};border:1.5px solid {_bd_log};">
                        {TABLE_ICON.format(size=15, color=_col_log)}
                        <div style="flex:1;">
                            <div style="display:flex;align-items:center;justify-content:space-between;">
                                <div style="font-family:'Plus Jakarta Sans',sans-serif;font-size:12.5px;font-weight:{_fw_log};color:{_tc_log};">Log Aktivitas Global</div>
                                <span style="font-family:'Fira Code',monospace;font-size:9px;font-weight:700;
                                             background:#3B5BDB;color:#fff;padding:2px 7px;border-radius:20px;">{_count}</span>
                            </div>
                            <div style="font-family:'Plus Jakarta Sans',sans-serif;font-size:10px;color:#94A3B8;">Riwayat evaluasi mahasiswa</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    if st.button("Log Aktivitas Global", use_container_width=True, key="nav_log"):
                        switch_page("log")

                # Info user aktif
                st.markdown(f"""
                <div style="border-top:1.5px solid #F1F5F9;margin:10px 0 12px 0;"></div>
                <div style="display:flex;align-items:center;gap:8px;padding:8px 10px;
                            background:#F8FAFC;border:1.5px solid #E2E8F0;border-radius:10px;margin-bottom:8px;">
                    <div style="width:30px;height:30px;border-radius:8px;
                                background:linear-gradient(135deg,#3B5BDB,#7048E8);
                                display:flex;align-items:center;justify-content:center;flex-shrink:0;">
                        {USER_ICON.format(size=14, color="#fff")}
                    </div>
                    <div>
                        <div style="font-family:'Plus Jakarta Sans',sans-serif;font-size:12px;font-weight:700;color:#0D1117;">Administrator</div>
                        <div style="font-family:'Fira Code',monospace;font-size:9px;color:#94A3B8;">admin · Sesi aktif</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # Tombol Logout
                st.markdown(f"""
                <div style="display:flex;align-items:center;gap:10px;width:100%;padding:10px 12px;
                            border-radius:10px;background:#FFF5F5;border:1.5px solid #FECACA;">
                    {SHIELD_ICON.format(size=15, color="#DC2626")}
                    <div style="font-family:'Plus Jakarta Sans',sans-serif;font-size:12.5px;font-weight:600;color:#DC2626;">
                        Keluar dari Aplikasi
                    </div>
                </div>
                """, unsafe_allow_html=True)
                if st.button("Keluar dari Aplikasi", use_container_width=True, key="nav_logout"):
                    st.session_state.clear()
                    st.rerun()

    st.markdown("<div style='height:78px;'></div>", unsafe_allow_html=True)

    # ── FORM INPUT ──
    if st.session_state.page == "input":
        st.markdown("""<div style="margin-bottom:22px;">
            <h2 style="font-family:'Space Grotesk',sans-serif;font-size:22px;font-weight:700;color:#0D1117;margin:0 0 4px 0;letter-spacing:-0.4px;">Form Evaluasi Mahasiswa</h2>
            <p style="font-family:'Plus Jakarta Sans',sans-serif;font-size:13px;color:#64748B;margin:0;">Isi data di bawah untuk memproses prediksi risiko dropout mahasiswa.</p>
        </div>""", unsafe_allow_html=True)

        with st.form("input_form"):
            with st.container(border=True):
                st.markdown(f'<div class="section-label">{IDCARD_ICON.format(size=14, color="#3B5BDB")} Identitas Pokok Mahasiswa</div>', unsafe_allow_html=True)
                col_id1, col_id2 = st.columns(2, gap="large")
                with col_id1:
                    nama = st.text_input("Nama Lengkap Mahasiswa", placeholder="Contoh: Muhammad Akbar")
                with col_id2:
                    nim = st.text_input("Nomor Induk Mahasiswa (NIM)", placeholder="Contoh: I00129031")
            st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown(f'<div class="section-label">{CLIPBOARD_ICON.format(size=14, color="#3B5BDB")} Nilai Atribut & Capaian Akademik Semester</div>', unsafe_allow_html=True)
                g1, g2, g3, g4 = st.columns(4, gap="large")
                with g1:
                    sem1_app      = st.number_input("SKS Diambil Sem 1", 0, 24, 12)
                    sem1_grd_indo = st.number_input("IPK Perolehan Sem 1", 0.0, 4.0, 3.0, step=0.01, format="%.2f")
                with g2:
                    sem2_app      = st.number_input("SKS Diambil Sem 2", 0, 24, 12)
                    sem2_grd_indo = st.number_input("IPK Perolehan Sem 2", 0.0, 4.0, 3.0, step=0.01, format="%.2f")
                with g3:
                    tuit_fees = st.selectbox("Pembayaran SPP Kuliah", ["Lancar", "Menunggak"])
                    debtor    = st.selectbox("Tunggakan Finansial Lain", ["Tidak Ada", "Ada"])
                with g4:
                    scholar   = st.selectbox("Klasifikasi Beasiswa", ["Bukan Penerima", "Penerima"])
                    age_enrol = st.number_input("Usia Saat Registrasi", 15, 60, 19)
            st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
            submitted = st.form_submit_button("→  Proses Evaluasi Data", type="primary", use_container_width=True)

        if submitted:
            if not nama or not nim:
                st.warning("Data parameter Nama dan NIM harus diisi lengkap.")
            else:
                input_feed = pd.DataFrame([[
                    sem2_app, sem2_grd_indo*5, sem1_app, sem1_grd_indo*5,
                    1 if tuit_fees=="Lancar" else 0,
                    1 if debtor=="Ada" else 0,
                    1 if scholar=="Penerima" else 0,
                    age_enrol
                ]])
                prediction  = model.predict(input_feed)
                probability = model.predict_proba(input_feed)
                risk = probability[0][0] * 100
                safe = probability[0][1] * 100
                st.session_state.history = [h for h in st.session_state.history if h["NIM"] != nim]
                st.session_state.history.append({
                    "NIM": nim, "Nama": nama.upper(),
                    "SKS1": sem1_app, "IPK1": sem1_grd_indo,
                    "SKS2": sem2_app, "IPK2": sem2_grd_indo,
                    "TotalSKS": sem1_app + sem2_app,
                    "RerataIPK": round((sem1_grd_indo + sem2_grd_indo)/2, 2),
                    "SPP": tuit_fees,
                    "SkorAman": round(safe, 2), "SkorRisiko": round(risk, 2),
                    "Class": prediction[0],
                    "Status": "AMAN (RETENSI TINGGI)" if safe >= 75 else ("WASPADA (RISIKO SEDANG)" if safe >= 40 else "KERENTANAN DROPOUT TINGGI")
                })
                st.session_state.selected_student_uid = nim
                switch_page("log")

    # ── LOG AKTIVITAS ──
    elif st.session_state.page == "log":
        if len(st.session_state.history) == 0:
            st.info("Belum ada data mahasiswa yang dievaluasi.")
            if st.button("Mulai Evaluasi Pertama", type="primary", use_container_width=True):
                switch_page("input")
            st.stop()

        df_full       = pd.DataFrame(st.session_state.history)
        total_mhs     = len(df_full)
        total_aman    = len(df_full[df_full['SkorAman'] >= 75])
        total_waspada = len(df_full[(df_full['SkorAman'] >= 40) & (df_full['SkorAman'] < 75)])
        total_risiko  = len(df_full[df_full['SkorAman'] < 40])

        st.markdown("""<div style="margin-bottom:22px;">
            <h2 style="font-family:'Space Grotesk',sans-serif;font-size:22px;font-weight:700;color:#0D1117;margin:0 0 4px 0;letter-spacing:-0.4px;">Log Aktivitas Evaluasi</h2>
            <p style="font-family:'Plus Jakarta Sans',sans-serif;font-size:13px;color:#64748B;margin:0;">Ringkasan seluruh mahasiswa yang telah dievaluasi pada sesi ini.</p>
        </div>""", unsafe_allow_html=True)

        total_aman    = sum(1 for h in st.session_state.history if h['SkorAman'] >= 75)
        total_waspada = sum(1 for h in st.session_state.history if 40 <= h['SkorAman'] < 75)
        total_risiko  = sum(1 for h in st.session_state.history if h['SkorAman'] < 40)

        s1, s2, s3, s4 = st.columns(4, gap="medium")
        with s1:
            st.markdown(f"""<div style="background:#fff;border:1.5px solid #E2E8F0;border-radius:16px;padding:18px 20px;border-top:4px solid #3B5BDB;">
                <div style="font-family:'Fira Code',monospace;font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:6px;">Total Dievaluasi</div>
                <div style="font-family:'Space Grotesk',sans-serif;font-size:30px;font-weight:700;color:#0D1117;">{total_mhs}</div>
                <div style="font-size:12px;color:#64748B;">Orang</div></div>""", unsafe_allow_html=True)
        with s2:
            st.markdown(f"""<div style="background:#fff;border:1.5px solid #E2E8F0;border-radius:16px;padding:18px 20px;border-top:4px solid #0D9488;">
                <div style="font-family:'Fira Code',monospace;font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:6px;">Retensi Tinggi (≥75%)</div>
                <div style="font-family:'Space Grotesk',sans-serif;font-size:30px;font-weight:700;color:#0D9488;">{total_aman}</div>
                <div style="font-size:12px;color:#64748B;">Mahasiswa aman</div></div>""", unsafe_allow_html=True)
        with s3:
            st.markdown(f"""<div style="background:#fff;border:1.5px solid #E2E8F0;border-radius:16px;padding:18px 20px;border-top:4px solid #D97706;">
                <div style="font-family:'Fira Code',monospace;font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:6px;">Waspada (40–74%)</div>
                <div style="font-family:'Space Grotesk',sans-serif;font-size:30px;font-weight:700;color:#D97706;">{total_waspada}</div>
                <div style="font-size:12px;color:#64748B;">Perlu monitoring</div></div>""", unsafe_allow_html=True)
        with s4:
            st.markdown(f"""<div style="background:#fff;border:1.5px solid #E2E8F0;border-radius:16px;padding:18px 20px;border-top:4px solid #DC2626;">
                <div style="font-family:'Fira Code',monospace;font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:6px;">Risiko Tinggi (&lt;40%)</div>
                <div style="font-family:'Space Grotesk',sans-serif;font-size:30px;font-weight:700;color:#DC2626;">{total_risiko}</div>
                <div style="font-size:12px;color:#64748B;">Kasus terdeteksi</div></div>""", unsafe_allow_html=True)

        st.markdown("<div style='height:22px;'></div>", unsafe_allow_html=True)
        st.markdown(f'<div class="section-label">{TABLE_ICON.format(size=14, color="#3B5BDB")} Tabel Riwayat Aktivitas Evaluasi</div>', unsafe_allow_html=True)
        st.dataframe(df_full[["NIM","Nama","TotalSKS","RerataIPK","SPP","Status"]], use_container_width=True, height=min(40+len(df_full)*35,240))

        st.markdown("<div style='height:22px;'></div>", unsafe_allow_html=True)
        st.markdown(f'<div class="section-label">{TARGET_ICON.format(size=14, color="#3B5BDB")} Pilih Record Mahasiswa untuk Visualisasi Performa</div>', unsafe_allow_html=True)
        options_list = [f"[{item['NIM']}] {item['Nama']}" for item in st.session_state.history]
        default_index = 0
        if st.session_state.selected_student_uid:
            for idx, item in enumerate(st.session_state.history):
                if item["NIM"] == st.session_state.selected_student_uid:
                    default_index = idx
        selection = st.radio("", options=options_list, index=default_index, horizontal=True, label_visibility="collapsed")
        current_nim = selection.split("]")[0].replace("[","").strip()
        if current_nim != st.session_state.selected_student_uid:
            st.session_state.selected_student_uid = current_nim
            st.rerun()

        active_data = next(item for item in st.session_state.history if item["NIM"] == st.session_state.selected_student_uid)
        st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

        _skor_aman  = active_data["SkorAman"]
        if _skor_aman >= 75:
            badge_color = "#0D9488"; badge_bg = "#CCFBF1"; badge_icon = CHECK_ICON.format(size=14, color="#0D9488")
        elif _skor_aman >= 40:
            badge_color = "#D97706"; badge_bg = "#FEF3C7"; badge_icon = ALERT_ICON.format(size=14, color="#D97706")
        else:
            badge_color = "#DC2626"; badge_bg = "#FEE2E2"; badge_icon = ALERT_ICON.format(size=14, color="#DC2626")

        with st.container(border=True):
            st.markdown(f"""
                <div style="display:flex;align-items:center;justify-content:space-between;padding:4px 4px 16px 4px;border-bottom:1.5px solid #F1F5F9;margin-bottom:18px;">
                    <div>
                        <div style="font-size:10.5px;font-family:'Fira Code',monospace;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:3px;">Mahasiswa Terpilih</div>
                        <div style="font-family:'Space Grotesk',sans-serif;font-size:17px;font-weight:700;color:#0D1117;letter-spacing:-0.3px;">
                            {active_data['Nama']}
                            <span style="font-family:'Plus Jakarta Sans',sans-serif;font-size:13px;font-weight:400;color:#64748B;margin-left:8px;">NIM {active_data['NIM']}</span>
                        </div>
                    </div>
                    <span style="display:inline-flex;align-items:center;gap:8px;font-family:'Fira Code',monospace;font-weight:600;
                                 font-size:11px;letter-spacing:0.6px;text-transform:uppercase;padding:10px 22px;border-radius:50px;
                                 color:{badge_color};background:{badge_bg};">
                        {badge_icon}&nbsp;{active_data['Status']}
                    </span>
                </div>
            """, unsafe_allow_html=True)
            pl, pm, pr = st.columns([1,2,2], gap="large")
            with pl:
                _sa = active_data['SkorAman']
                if _sa >= 75:
                    _card_color = "#3B5BDB"; _card_bg = "#F8FAFF"; _card_border = "#c5d0fa"; _card_label = "AMAN · Retensi Tinggi"
                elif _sa >= 40:
                    _card_color = "#D97706"; _card_bg = "#FFFBEB"; _card_border = "#FDE68A"; _card_label = "WASPADA · Risiko Sedang"
                else:
                    _card_color = "#DC2626"; _card_bg = "#FFF5F5"; _card_border = "#fecaca"; _card_label = "DROPOUT · Risiko Tinggi"
                st.markdown(f"""
                    <div style="display:flex;flex-direction:column;gap:12px;">
                        <div style="background:{_card_bg};border:1.5px solid {_card_border};border-radius:14px;padding:16px 18px;">
                            <div style="font-family:'Fira Code',monospace;font-size:9.5px;color:{_card_color};text-transform:uppercase;letter-spacing:0.7px;margin-bottom:4px;">Persentase Kelulusan</div>
                            <div style="font-family:'Space Grotesk',sans-serif;font-size:30px;font-weight:700;color:{_card_color};line-height:1;">{active_data['SkorAman']}%</div>
                            <div style="font-family:'Fira Code',monospace;font-size:9px;color:{_card_color};margin-top:5px;opacity:0.8;">{_card_label}</div>
                        </div>
                        <div style="background:#FFF5F5;border:1.5px solid #fecaca;border-radius:14px;padding:16px 18px;">
                            <div style="font-family:'Fira Code',monospace;font-size:9.5px;color:#DC2626;text-transform:uppercase;letter-spacing:0.7px;margin-bottom:4px;">Faktor Risiko Dropout</div>
                            <div style="font-family:'Space Grotesk',sans-serif;font-size:30px;font-weight:700;color:#DC2626;line-height:1;">{active_data['SkorRisiko']}%</div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
            with pm:
                st.markdown(f'<div style="font-family:\'Fira Code\',monospace;font-size:11.5px;font-weight:600;color:#0D1117;margin-bottom:8px;">{BARCHART_ICON.format(size=13,color="#3B5BDB")} &nbsp;Fluktuasi IPK per Semester</div>', unsafe_allow_html=True)
                st.line_chart(pd.DataFrame({"IPK":[active_data["IPK1"],active_data["IPK2"]]},index=["Semester 1","Semester 2"]), height=165, use_container_width=True)
            with pr:
                st.markdown(f'<div style="font-family:\'Fira Code\',monospace;font-size:11.5px;font-weight:600;color:#0D1117;margin-bottom:8px;">{BARCHART_ICON.format(size=13,color="#7048E8")} &nbsp;Volume SKS per Semester</div>', unsafe_allow_html=True)
                st.bar_chart(pd.DataFrame({"Volume SKS":[active_data["SKS1"],active_data["SKS2"]]},index=["Semester 1","Semester 2"]), height=165, use_container_width=True)

        st.markdown("<div style='height:24px;'></div>", unsafe_allow_html=True)

        # ── FITUR: Faktor & Rekomendasi ──
        fa, fb = st.columns([1,1], gap="large")
        with fa:
            with st.container(border=True):
                st.markdown(f'<div class="section-label">{ZAPPER_ICON.format(size=13,color="#3B5BDB")} Indikator Faktor Keputusan Model</div>', unsafe_allow_html=True)
                scores = factor_score(active_data)
                for label, score in scores.items():
                    bar_color = "#0D9488" if score>=70 else "#D97706" if score>=40 else "#DC2626"
                    txt_bg    = "#CCFBF1" if score>=70 else "#FEF3C7" if score>=40 else "#FEE2E2"
                    st.markdown(f"""
                        <div style="margin-bottom:14px;">
                            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
                                <span style="font-family:'Plus Jakarta Sans',sans-serif;font-size:13px;font-weight:600;color:#0D1117;">{label}</span>
                                <span style="font-family:'Fira Code',monospace;font-size:12px;font-weight:600;color:{bar_color};background:{txt_bg};padding:2px 10px;border-radius:20px;">{score}%</span>
                            </div>
                            <div style="width:100%;height:8px;background:#F1F5F9;border-radius:99px;overflow:hidden;">
                                <div style="width:{score}%;height:100%;background:{bar_color};border-radius:99px;"></div>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
        with fb:
            with st.container(border=True):
                st.markdown(f'<div class="section-label">{BULB_ICON.format(size=13,color="#3B5BDB")} Rekomendasi Intervensi Akademik</div>', unsafe_allow_html=True)
                recs = generate_recommendations(active_data)
                level_colors = {"KRITIS":("DC2626","FEE2E2","FECACA"),"WASPADA":("D97706","FEF3C7","FDE68A"),"BAIK":("0D9488","CCFBF1","99F6E4")}
                for rec in recs:
                    fg, bg, border = level_colors.get(rec["level"],("3B5BDB","EEF2FF","C5D0FA"))
                    st.markdown(f"""
                        <div style="display:flex;gap:14px;align-items:flex-start;background:#{bg};border:1.5px solid #{border};border-radius:14px;padding:14px 16px;margin-bottom:12px;">
                            <div style="background:rgba(255,255,255,0.7);border-radius:8px;padding:4px 10px;color:#{fg};font-family:'Fira Code',monospace;font-weight:700;font-size:9px;white-space:nowrap;">{rec['level']}</div>
                            <div>
                                <div style="font-family:'Space Grotesk',sans-serif;font-size:13.5px;font-weight:700;color:#0D1117;margin-bottom:3px;">{rec['title']}</div>
                                <div style="font-family:'Plus Jakarta Sans',sans-serif;font-size:12px;color:#475569;line-height:1.55;">{rec['desc']}</div>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)

        # ══════════════════════════════════════════════════════════
        # EXPORT PANEL — PDF Print + Kembali
        # ══════════════════════════════════════════════════════════
        st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown(f"""
                <div style="display:flex;align-items:center;gap:16px;padding:6px 4px 14px 4px;border-bottom:1.5px solid #F1F5F9;margin-bottom:18px;">
                    <div style="width:44px;height:44px;border-radius:13px;background:linear-gradient(135deg,#EEF2FF,#F3F0FF);
                                border:1.5px solid #c5d0fa;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
                        {DOWNLOAD_ICON.format(size=20, color="#3B5BDB")}
                    </div>
                    <div>
                        <div style="font-family:'Space Grotesk',sans-serif;font-size:15px;font-weight:700;color:#0D1117;">Ekspor Laporan Lengkap</div>
                        <div style="font-family:'Plus Jakarta Sans',sans-serif;font-size:12px;color:#94A3B8;">
                            Laporan berisi identitas, data semester, semua grafik, faktor risiko &amp; rekomendasi intervensi.
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

            html_report = generate_html_report(active_data)
            st.markdown(f"""
                <div style="background:#EEF2FF;border:1.5px solid #C5D0FA;border-radius:14px;padding:16px 18px;margin-bottom:10px;">
                    <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px;">
                        {PRINT_ICON.format(size=18, color="#3B5BDB")}
                        <span style="font-family:'Space Grotesk',sans-serif;font-size:13px;font-weight:700;color:#0D1117;">Cetak / PDF</span>
                    </div>
                    <div style="font-family:'Plus Jakarta Sans',sans-serif;font-size:11.5px;color:#64748B;line-height:1.5;">
                        Buka laporan HTML interaktif di tab baru. Klik tombol <b>Cetak / Simpan PDF</b> untuk print atau save as PDF langsung dari browser.
                    </div>
                </div>
            """, unsafe_allow_html=True)
            st.download_button(
                label="🖨 Buka Laporan untuk Print/PDF",
                data=html_report.encode("utf-8"),
                file_name=f"Laporan_{active_data['NIM']}_{active_data['Nama']}.html",
                mime="text/html",
                use_container_width=True,
                key="dl_html"
            )

            st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
            if st.button("Evaluasi Mahasiswa Baru", use_container_width=True, key="back_btn"):
                switch_page("input")