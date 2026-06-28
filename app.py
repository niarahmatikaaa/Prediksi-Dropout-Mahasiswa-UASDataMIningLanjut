import streamlit as st
import pandas as pd
import pickle
import time
import io

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, Image, HRFlowable, KeepTogether
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import Flowable

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
LOGO_SVG       = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3 2 8l10 5 10-5-10-5Z"/><path d="M6 10.5V16c0 1.2 2.7 2.5 6 2.5s6-1.3 6-2.5v-5.5"/><path d="M22 8v6"/></svg>"""
ZAPPER_ICON    = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>"""
BULB_ICON      = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="9" y1="18" x2="15" y2="18"/><line x1="10" y1="22" x2="14" y2="22"/><path d="M15.09 14c.18-.98.65-1.74 1.41-2.5A4.65 4.65 0 0 0 18 8 6 6 0 0 0 6 8c0 1 .23 2.23 1.5 3.5A4.61 4.61 0 0 1 8.91 14"/></svg>"""
DOWNLOAD_ICON  = """<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>"""

# ==========================================
# 3. CSS
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=Space+Grotesk:wght@300;400;500;600;700&family=Fira+Code:wght@400;500;600&display=swap');
:root {
    --ink:#0D1117; --muted:#64748B;
    --paper:#F4F7FB; --surface:#FFFFFF;
    --blue:#3B5BDB; --blue-bg:#EEF2FF;
    --violet:#7048E8;
    --safe:#0D9488; --safe-bg:#CCFBF1;
    --risk:#DC2626; --risk-bg:#FEE2E2;
    --amber:#D97706; --amber-bg:#FEF3C7;
    --hairline:#E2E8F0;
    --font-head:'Space Grotesk',sans-serif;
    --font-body:'Plus Jakarta Sans',sans-serif;
    --font-code:'Fira Code',monospace;
}
html,body,[class*="css"]{font-family:var(--font-body)!important;-webkit-font-smoothing:antialiased;}
.stApp{background-color:var(--paper);}
header[data-testid="stHeader"]{display:none!important;}
div[data-testid="stDecoration"]{display:none!important;}
button[title="View sidebar"]{display:none!important;}
div[data-testid="InputInstructions"]{display:none!important;}
.block-container{padding-top:0!important;padding-bottom:3rem!important;max-width:100%!important;}
.st-key-topbar{
    position:fixed!important;top:0!important;left:0!important;right:0!important;z-index:999999!important;
    background:rgba(255,255,255,0.97)!important;backdrop-filter:blur(14px);
    border-bottom:1.5px solid var(--hairline)!important;box-shadow:0 4px 20px rgba(15,23,42,0.04)!important;
    padding:14px 2.4rem!important;
}
.st-key-topbar > div[data-testid="stVerticalBlock"]{position:static!important;}
.st-key-topbar div[data-testid="stVerticalBlockBorderWrapper"]{
    position:static!important;background:transparent!important;border:none!important;box-shadow:none!important;border-radius:0!important;
}
.st-key-topbar_spacer > div{height:74px!important;min-height:74px!important;padding:0!important;margin:0!important;}
.section-label{
    font-family:var(--font-code)!important;font-size:11px;font-weight:600;color:var(--blue);text-transform:uppercase;
    letter-spacing:1.4px;border-left:3px solid var(--blue);padding-left:12px;margin:12px 0 20px 0;display:flex;align-items:center;gap:8px;
}
div[data-testid="stVerticalBlockBorderWrapper"]{
    background:var(--surface)!important;border-radius:18px!important;border:1px solid var(--hairline)!important;
    box-shadow:0 4px 16px rgba(15,23,42,0.03)!important;
}
div[data-baseweb="input"],div[data-baseweb="select"]{background-color:#F8FAFC!important;border-radius:10px!important;border:1px solid var(--hairline)!important;}
div[data-baseweb="input"]:focus-within,div[data-baseweb="select"]:focus-within{border-color:var(--blue)!important;box-shadow:0 0 0 1.5px var(--blue)!important;}
.stTextInput input,.stNumberInput input{padding:12px 16px!important;font-size:14px!important;}
.stTextInput label,.stNumberInput label,.stSelectbox label{font-size:13px!important;font-weight:600!important;color:var(--ink)!important;}
.stButton button[kind="primary"],div[data-testid="stFormSubmitButton"] button,.stDownloadButton button{
    background:linear-gradient(135deg,var(--blue) 0%,var(--violet) 100%)!important;
    border:none!important;border-radius:12px!important;color:#fff!important;
    font-weight:700!important;font-size:14px!important;padding:12px 0!important;
    box-shadow:0 4px 16px rgba(59,91,219,0.25)!important;transition:all .2s!important;
}
div[data-testid="stFormSubmitButton"] button p{color:#ffffff!important;}
.stButton button[kind="primary"]:hover,div[data-testid="stFormSubmitButton"] button:hover,.stDownloadButton button:hover{opacity:0.9!important;transform:translateY(-1px);}
.stButton button[kind="secondary"]{border-radius:12px!important;border:1.5px solid var(--hairline)!important;color:var(--ink)!important;font-weight:600!important;background:#fff!important;transition:all .2s!important;padding:12px 0!important;}
.stButton button[kind="secondary"]:hover{border-color:var(--blue)!important;color:var(--blue)!important;background:#F8FAFF!important;}
div[data-testid="stDataFrame"]{border-radius:14px;overflow:hidden;border:1px solid var(--hairline);}
.stat-cards-row{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:24px;}
@media(max-width:992px){.stat-cards-row{grid-template-columns:repeat(2,1fr);}}
@media(max-width:576px){.stat-cards-row{grid-template-columns:1fr;}}
.stat-card{background:#fff;border:1px solid var(--hairline);border-radius:16px;padding:24px;display:flex;flex-direction:column;justify-content:space-between;box-shadow:0 4px 16px rgba(15,23,42,0.02);}
.stat-card .sc-label{font-family:var(--font-code);font-size:11px;color:#64748B;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;font-weight:600;}
.stat-card .sc-value{font-family:var(--font-head);font-size:38px;font-weight:700;line-height:1;margin-bottom:6px;}
.stat-card .sc-sub{font-size:13px;color:#64748B;font-weight:500;}
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
# 5. SESSION STATE & ROUTING
# ==========================================
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
# HELPER
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
        "IPK Rata-rata":  min(round(ipk_avg / 4.0 * 100), 100),
        "Beban SKS":      min(round((data["TotalSKS"] / 48) * 100), 100),
        "Kelancaran SPP": 100 if data["SPP"] == "Lancar" else 20,
        "Tren Nilai":     min(max(round(((data["IPK2"]-data["IPK1"])+1)/2*100), 0), 100),
    }

# ==========================================
# CHART HELPERS
# ==========================================
def make_ipk_chart(ipk1, ipk2, w_inch=2.8, h_inch=1.5):
    fig, ax = plt.subplots(figsize=(w_inch, h_inch))
    fig.patch.set_facecolor("#F8FAFF"); ax.set_facecolor("#F8FAFF")
    sems = ["Sem 1","Sem 2"]; vals = [ipk1, ipk2]; color = "#3B5BDB"
    ax.plot(sems, vals, color=color, linewidth=2.0, marker="o", markersize=7,
            markerfacecolor="#fff", markeredgewidth=2.0, markeredgecolor=color)
    ax.fill_between(sems, vals, alpha=0.12, color=color)
    for s, v in zip(sems, vals):
        ax.annotate(f"{v:.2f}", (s, v), textcoords="offset points", xytext=(0,7),
                    ha="center", fontsize=7.5, fontweight="bold", color=color)
    ax.set_ylim(0, 4.3); ax.set_ylabel("IPK", fontsize=7, color="#64748B")
    ax.tick_params(colors="#64748B", labelsize=7)
    for sp in ax.spines.values(): sp.set_edgecolor("#E2E8F0")
    ax.grid(axis="y", color="#E2E8F0", linestyle="--", linewidth=0.6)
    plt.tight_layout(pad=0.4)
    buf = io.BytesIO(); plt.savefig(buf, format="png", dpi=150, bbox_inches="tight"); plt.close(fig); buf.seek(0)
    return buf

def make_sks_chart(sks1, sks2, w_inch=2.8, h_inch=1.5):
    fig, ax = plt.subplots(figsize=(w_inch, h_inch))
    fig.patch.set_facecolor("#F8F7FF"); ax.set_facecolor("#F8F7FF")
    sems = ["Sem 1","Sem 2"]; vals = [sks1, sks2]
    bars = ax.bar(sems, vals, color=["#7048E8","#A78BFA"], width=0.35, edgecolor="white", linewidth=1.0, zorder=3)
    for bar, v in zip(bars, vals):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.1, str(v),
                ha="center", va="bottom", fontsize=7.5, fontweight="bold", color="#7048E8")
    ax.set_ylim(0, max(vals)*1.45+1); ax.set_ylabel("SKS", fontsize=7, color="#64748B")
    ax.tick_params(colors="#64748B", labelsize=7)
    for sp in ax.spines.values(): sp.set_edgecolor("#E2E8F0")
    ax.grid(axis="y", color="#E2E8F0", linestyle="--", linewidth=0.6, zorder=0)
    plt.tight_layout(pad=0.4)
    buf = io.BytesIO(); plt.savefig(buf, format="png", dpi=150, bbox_inches="tight"); plt.close(fig); buf.seek(0)
    return buf

def make_risiko_gauge(risk_pct, safe_pct, w_inch=2.8, h_inch=1.2):
    fig, ax = plt.subplots(figsize=(w_inch, h_inch))
    fig.patch.set_facecolor("white"); ax.set_facecolor("white")
    cats = ["Risiko Dropout","Peluang Lulus"]; vals = [risk_pct, safe_pct]
    bars = ax.barh(cats, vals, color=["#DC2626","#0D9488"], height=0.32, edgecolor="white")
    for bar, v in zip(bars, vals):
        ax.text(min(v+1,93), bar.get_y()+bar.get_height()/2,
                f"{v:.1f}%", va="center", fontsize=7.5, fontweight="bold", color="#0D1117")
    ax.set_xlim(0,100); ax.set_xlabel("Persentase (%)", fontsize=7, color="#64748B")
    ax.tick_params(colors="#64748B", labelsize=7)
    for sp in ax.spines.values(): sp.set_edgecolor("#E2E8F0")
    ax.grid(axis="x", color="#E2E8F0", linestyle="--", linewidth=0.6)
    plt.tight_layout(pad=0.4)
    buf = io.BytesIO(); plt.savefig(buf, format="png", dpi=150, bbox_inches="tight"); plt.close(fig); buf.seek(0)
    return buf

def make_factor_chart(scores, w_inch=2.8, h_inch=1.5):
    fig, ax = plt.subplots(figsize=(w_inch, h_inch))
    fig.patch.set_facecolor("white"); ax.set_facecolor("white")
    labels = list(scores.keys()); vals = list(scores.values())
    bar_colors = ["#0D9488" if v>=70 else "#D97706" if v>=40 else "#DC2626" for v in vals]
    bars = ax.barh(labels, vals, color=bar_colors, height=0.32, edgecolor="white")
    for bar, v in zip(bars, vals):
        ax.text(min(v+1,93), bar.get_y()+bar.get_height()/2,
                f"{v}%", va="center", fontsize=7, fontweight="bold", color="#0D1117")
    ax.set_xlim(0,100); ax.set_xlabel("Skor (%)", fontsize=7, color="#64748B")
    ax.tick_params(colors="#64748B", labelsize=7)
    for sp in ax.spines.values(): sp.set_edgecolor("#E2E8F0")
    ax.grid(axis="x", color="#E2E8F0", linestyle="--", linewidth=0.6)
    plt.tight_layout(pad=0.4)
    buf = io.BytesIO(); plt.savefig(buf, format="png", dpi=150, bbox_inches="tight"); plt.close(fig); buf.seek(0)
    return buf

# ==========================================
# REPORTLAB: Progress Bar Flowable
# ==========================================
class ProgressBar(Flowable):
    def __init__(self, score, bar_width):
        super().__init__()
        self.score     = score
        self.bar_width = bar_width
        self.bar_height = 7
        if score >= 70:   self.color = colors.HexColor("#0D9488")
        elif score >= 40: self.color = colors.HexColor("#D97706")
        else:             self.color = colors.HexColor("#DC2626")

    def wrap(self, aw, ah): return (self.bar_width, self.bar_height)

    def draw(self):
        self.canv.saveState()
        r = self.bar_height / 2
        self.canv.setFillColor(colors.HexColor("#F1F5F9"))
        self.canv.roundRect(0, 0, self.bar_width, self.bar_height, r, fill=1, stroke=0)
        self.canv.setFillColor(self.color)
        filled = max(self.bar_width * self.score / 100, r * 2)
        self.canv.roundRect(0, 0, filled, self.bar_height, r, fill=1, stroke=0)
        self.canv.restoreState()

# ==========================================
# PDF GENERATOR  — 1 halaman penuh, tidak terpotong
# ==========================================
def generate_pdf_bytes(active_data):
    buf = io.BytesIO()
    W, H = A4
    M    = 9 * mm

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=M, rightMargin=M,
        topMargin=M, bottomMargin=M
    )

    PW = W - 2 * M

    C_BLUE   = colors.HexColor("#3B5BDB")
    C_VIOLET = colors.HexColor("#7048E8")
    C_SAFE   = colors.HexColor("#0D9488")
    C_RISK   = colors.HexColor("#DC2626")
    C_AMBER  = colors.HexColor("#D97706")
    C_INK    = colors.HexColor("#0D1117")
    C_MUTED  = colors.HexColor("#64748B")
    C_HAIR   = colors.HexColor("#E2E8F0")
    C_BG     = colors.HexColor("#F8FAFC")
    C_WHITE  = colors.white
    C_SAFEBG = colors.HexColor("#CCFBF1")
    C_RISKBG = colors.HexColor("#FEE2E2")
    C_AMBBG  = colors.HexColor("#FEF3C7")
    C_BLUEBG = colors.HexColor("#EEF2FF")

    def st_(name, **kw):
        base = ParagraphStyle(name, fontName="Helvetica", fontSize=7.5,
                              leading=10, textColor=C_INK)
        for k, v in kw.items(): setattr(base, k, v)
        return base

    S_HDR_T = st_("ht", fontName="Helvetica-Bold", fontSize=12, textColor=C_WHITE, leading=15)
    S_HDR_S = st_("hs", fontSize=6,   textColor=colors.HexColor("#C5D0FA"), leading=8)
    S_HDR_D = st_("hd", fontSize=6.5, textColor=colors.HexColor("#C5D0FA"), alignment=TA_RIGHT)
    S_SEC   = st_("sc", fontName="Helvetica-Bold", fontSize=7.5, textColor=C_BLUE, leading=10)
    S_LBL   = st_("lb", fontName="Helvetica-Bold", fontSize=5.5, textColor=C_MUTED, leading=7)
    S_TH    = st_("th", fontName="Helvetica-Bold", fontSize=7,   textColor=C_WHITE, leading=9)
    S_TD    = st_("td", fontName="Helvetica",      fontSize=7.5, textColor=C_INK,   leading=10)
    S_TDB   = st_("tb", fontName="Helvetica-Bold", fontSize=7.5, textColor=C_INK,   leading=10)
    S_CTIT  = st_("ct", fontName="Helvetica-Bold", fontSize=6,   textColor=C_BLUE,  alignment=TA_CENTER, leading=8)
    S_FOOT  = st_("ft", fontSize=6, textColor=C_MUTED, alignment=TA_CENTER)
    S_RTT   = st_("rt", fontName="Helvetica-Bold", fontSize=7.5, textColor=C_INK,   leading=10)
    S_RDD   = st_("rd", fontSize=6.5, textColor=colors.HexColor("#475569"),          leading=9)

    def p(text, style): return Paragraph(text, style)

    def zpad(t):
        t.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),  ("BOTTOMPADDING",(0,0),(-1,-1),0),
        ]))
        return t

    # ── derived ──
    sa = active_data["SkorAman"]
    if sa >= 75:   sc, sb, sl = C_SAFE,  C_SAFEBG, "AMAN (RETENSI TINGGI)"
    elif sa >= 40: sc, sb, sl = C_AMBER, C_AMBBG,  "WASPADA (RISIKO SEDANG)"
    else:          sc, sb, sl = C_RISK,  C_RISKBG, "RISIKO DROPOUT TINGGI"

    spp_c = C_SAFE if active_data["SPP"] == "Lancar" else C_RISK
    di = active_data["IPK2"] - active_data["IPK1"]
    if di > 0:   it, ic = f"Naik {abs(di):.2f} poin",   C_SAFE
    elif di < 0: it, ic = f"Turun {abs(di):.2f} poin",  C_RISK
    else:        it, ic = "Stabil",                       C_MUTED
    ds = active_data["SKS2"] - active_data["SKS1"]
    if ds > 0:   st2, sc2 = "Bertambah", C_SAFE
    elif ds < 0: st2, sc2 = "Berkurang", C_AMBER
    else:        st2, sc2 = "Sama",      C_MUTED

    scores = factor_score(active_data)
    recs   = generate_recommendations(active_data)
    story  = []

    # ── HEADER ──
    from reportlab.pdfbase.pdfmetrics import stringWidth
    LH = round(PW * 0.60)
    RH = PW - LH

    badge_st = ParagraphStyle("bs", fontName="Helvetica-Bold", fontSize=7,
                               textColor=sc, alignment=TA_CENTER, leading=9)
    bw = min(stringWidth(sl, "Helvetica-Bold", 7) + 26, RH - 12)
    bdg = Table([[p(sl, badge_st)]], colWidths=[bw])
    bdg.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1), C_WHITE),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
    ]))
    bdg.cornerRadii = [8,8,8,8]

    hl = zpad(Table([[p("PortalAkademik", S_HDR_T)],
                     [p("LAPORAN EVALUASI STUDI MAHASISWA \u00b7 PREDIKSI RISIKO DROPOUT", S_HDR_S)]],
                    colWidths=[LH]))
    hr = Table([[p(time.strftime('%d %B %Y'), S_HDR_D)],
                [p(f"{time.strftime('%H:%M:%S')} WIB", S_HDR_D)],
                [Spacer(1,2)],[bdg]],
               colWidths=[RH-12])
    hr.setStyle(TableStyle([
        ("ALIGN",(0,0),(-1,-1),"RIGHT"),
        ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0), ("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))

    hdr = Table([[hl, hr]], colWidths=[LH, RH])
    hdr.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1), C_BLUE),
        ("TOPPADDING",(0,0),(-1,-1),10),("BOTTOMPADDING",(0,0),(-1,-1),10),
        ("LEFTPADDING",(0,0),(0,-1),14), ("RIGHTPADDING",(1,0),(1,-1),12),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    hdr.cornerRadii = [12,12,12,12]
    story.append(hdr)
    story.append(Spacer(1, 2.5*mm))

    def sec(num, text):
        t = Table([[p(f"{num}  \u00b7  {text}", S_SEC)]], colWidths=[PW])
        t.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),10), ("RIGHTPADDING",(0,0),(-1,-1),6),
            ("TOPPADDING",(0,0),(-1,-1),3),   ("BOTTOMPADDING",(0,0),(-1,-1),3),
            ("LINEBEFORE",(0,0),(0,-1),3, C_BLUE),
        ]))
        return t

    def id_cell(label, value, val_color=C_INK, cell_w=None):
        vs = ParagraphStyle("vs", fontName="Helvetica-Bold", fontSize=8.5,
                             textColor=val_color, leading=11)
        t = Table([[p(label, S_LBL)],[p(str(value), vs)]], colWidths=[cell_w])
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), C_BG),
            ("BOX",(0,0),(-1,-1),0.5, C_HAIR),
            ("LEFTPADDING",(0,0),(-1,-1),8), ("RIGHTPADDING",(0,0),(-1,-1),8),
            ("TOPPADDING",(0,0),(-1,-1),5),  ("BOTTOMPADDING",(0,0),(-1,-1),5),
        ]))
        t.cornerRadii = [8,8,8,8]
        return t

    GAP  = 3
    CW2  = round((PW - GAP) / 2)

    def id_row_2(la, va, ca, lb, vb, cb):
        return Table([[id_cell(la, va, ca, CW2), id_cell(lb, vb, cb, CW2)]],
                     colWidths=[CW2, CW2])

    id_grid = [
        id_row_2("NAMA LENGKAP",          active_data["Nama"],                      C_INK,
                 "NIM",                   active_data["NIM"],                       C_INK),
        id_row_2("STATUS SPP",            active_data["SPP"],                       spp_c,
                 "IPK RATA-RATA",         f"{active_data['RerataIPK']:.2f}",        C_INK),
        id_row_2("TOTAL SKS DITEMPUH",    f"{active_data['TotalSKS']} SKS",         C_INK,
                 "PELUANG LULUS",         f"{active_data['SkorAman']}%",            C_BLUE),
        id_row_2("FAKTOR RISIKO DROPOUT", f"{active_data['SkorRisiko']}%",          C_RISK,
                 "PREDIKSI MODEL",        active_data["Status"],                    sc),
    ]
    for t in id_grid:
        zpad(t)

    id_stack = Table([[r] for r in id_grid], colWidths=[PW])
    id_stack.setStyle(TableStyle([
        ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),1),  ("BOTTOMPADDING",(0,0),(-1,-1),1),
    ]))

    story.append(sec("1", "IDENTITAS MAHASISWA"))
    story.append(Spacer(1, 1.5*mm))
    story.append(id_stack)
    story.append(Spacer(1, 2.5*mm))

    # ── SECTION 2 ──
    def cv(c): return ParagraphStyle("cv", fontName="Helvetica-Bold", fontSize=7.5, textColor=c, leading=10)

    c1 = round(PW*0.38); c2 = round(PW*0.18); c3 = round(PW*0.18); c4 = PW - c1 - c2 - c3
    sem = Table([
        [p("Indikator",S_TH), p("Semester 1",S_TH), p("Semester 2",S_TH), p("Keterangan",S_TH)],
        [p("IPK (Indeks Prestasi Kumulatif)", S_TDB),
         p(f"{active_data['IPK1']:.2f}", cv(C_BLUE)),
         p(f"{active_data['IPK2']:.2f}", cv(ic)),
         p(it, cv(ic))],
        [p("Jumlah SKS Diambil", S_TDB),
         p(f"{active_data['SKS1']} SKS", cv(C_VIOLET)),
         p(f"{active_data['SKS2']} SKS", cv(sc2)),
         p(st2, cv(sc2))],
    ], colWidths=[c1, c2, c3, c4])
    sem.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0), C_BLUE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[C_BG, C_WHITE]),
        ("LINEBELOW",(0,0),(-1,-1),0.4, C_HAIR),
        ("LEFTPADDING",(0,0),(-1,-1),8), ("RIGHTPADDING",(0,0),(-1,-1),8),
        ("TOPPADDING",(0,0),(-1,-1),5),  ("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    sem.cornerRadii = [10,10,10,10]

    story.append(sec("2", "DATA AKADEMIK PER SEMESTER"))
    story.append(Spacer(1, 1.5*mm))
    story.append(sem)
    story.append(Spacer(1, 2.5*mm))

    # ── SECTION 3 — Grafik ──
    CG  = 4
    CCW = round((PW - CG) / 2)

    def chart_cell(title, buf):
        # FIX: rasio tinggi dikurangi dari 0.52 -> 0.48 agar ada ruang untuk Section 4 & 5
        img = Image(buf, width=CCW - 8, height=(CCW - 8) * 0.48)
        t = Table([[p(title, S_CTIT)], [img]], colWidths=[CCW])
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), C_BG),
            ("BOX",(0,0),(-1,-1),0.5, C_HAIR),
            ("TOPPADDING",(0,0),(-1,-1),4),   ("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("LEFTPADDING",(0,0),(-1,-1),4),  ("RIGHTPADDING",(0,0),(-1,-1),4),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ]))
        t.cornerRadii = [8,8,8,8]
        return t

    def chart_row(a, b):
        t = Table([[a, b]], colWidths=[CCW, CCW])
        zpad(t)
        return t

    ipk_b  = make_ipk_chart(active_data["IPK1"], active_data["IPK2"])
    sks_b  = make_sks_chart(active_data["SKS1"], active_data["SKS2"])
    risk_b = make_risiko_gauge(active_data["SkorRisiko"], active_data["SkorAman"])
    fact_b = make_factor_chart(scores)

    story.append(sec("3", "VISUALISASI & GRAFIK"))
    story.append(Spacer(1, 1.5*mm))
    story.append(chart_row(
        chart_cell("FLUKTUASI IPK PER SEMESTER", ipk_b),
        chart_cell("VOLUME SKS PER SEMESTER",    sks_b)
    ))
    story.append(Spacer(1, 2*mm))
    story.append(chart_row(
        chart_cell("SKOR RISIKO VS KELULUSAN",         risk_b),
        chart_cell("INDIKATOR FAKTOR KEPUTUSAN MODEL", fact_b)
    ))
    story.append(Spacer(1, 2.5*mm))

    # ── SECTION 4 — Faktor Skor ──
    P4  = 10
    IW  = PW - 2 * P4
    LW4 = round(IW * 0.30)
    BW4 = round(IW * 0.52)
    PW4 = IW - LW4 - BW4

    factor_rows = []
    for label, score in scores.items():
        if score >= 70:   bc = colors.HexColor("#0D9488")
        elif score >= 40: bc = colors.HexColor("#D97706")
        else:             bc = colors.HexColor("#DC2626")
        s_l = ParagraphStyle("sl", fontName="Helvetica-Bold", fontSize=7, textColor=C_INK)
        s_p = ParagraphStyle("sp", fontName="Helvetica-Bold", fontSize=7, textColor=bc, alignment=TA_RIGHT)
        pb  = ProgressBar(score, BW4)
        row = Table([[p(label, s_l), pb, p(f"{score}%", s_p)]],
                    colWidths=[LW4, BW4, PW4])
        row.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",(0,0),(-1,-1),1), ("RIGHTPADDING",(0,0),(-1,-1),1),
            ("TOPPADDING",(0,0),(-1,-1),5),  ("BOTTOMPADDING",(0,0),(-1,-1),5),
        ]))
        factor_rows.append(row)

    f4_card = Table([[r] for r in factor_rows], colWidths=[IW])
    f4_card.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1), C_WHITE),
        ("BOX",(0,0),(-1,-1),0.5, C_HAIR),
        ("LEFTPADDING",(0,0),(-1,-1),P4), ("RIGHTPADDING",(0,0),(-1,-1),P4),
        ("TOPPADDING",(0,0),(-1,-1),1),   ("BOTTOMPADDING",(0,0),(-1,-1),1),
        # FIX: cegah tabel ini terpotong antar baris progress bar
        ("NOSPLIT",(0,0),(-1,-1)),
    ]))
    f4_card.cornerRadii = [10,10,10,10]

    # FIX: KeepTogether memastikan judul section 4 dan isinya selalu satu blok
    story.append(KeepTogether([
        sec("4", "INDIKATOR FAKTOR KEPUTUSAN MODEL"),
        Spacer(1, 1.5*mm),
        f4_card,
    ]))
    story.append(Spacer(1, 2.5*mm))

    # ── SECTION 5 — Rekomendasi ──
    level_map = {
        "KRITIS":  (colors.HexColor("#DC2626"), colors.HexColor("#FEE2E2"), colors.HexColor("#FECACA")),
        "WASPADA": (colors.HexColor("#D97706"), colors.HexColor("#FEF3C7"), colors.HexColor("#FDE68A")),
        "BAIK":    (colors.HexColor("#0D9488"), colors.HexColor("#CCFBF1"), colors.HexColor("#99F6E4")),
    }

    story.append(sec("5", "REKOMENDASI INTERVENSI AKADEMIK"))
    story.append(Spacer(1, 1.5*mm))

    n = len(recs)
    RG = 3
    RW = round((PW - RG * (n - 1)) / n) if n > 1 else PW

    rec_cells  = []
    rec_widths = []
    for i, rec in enumerate(recs):
        fg, bg, border = level_map.get(rec["level"], (C_BLUE, C_BLUEBG, colors.HexColor("#C5D0FA")))
        RP  = 8
        RIW = RW - 2 * RP

        lv_t = Table([[p(rec["level"], ParagraphStyle("lv", fontName="Helvetica-Bold",
                          fontSize=6, textColor=fg, alignment=TA_CENTER, leading=8))]],
                     colWidths=[RIW])
        lv_t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), C_WHITE),
            ("BOX",(0,0),(-1,-1),0.5, border),
            ("TOPPADDING",(0,0),(-1,-1),3),  ("BOTTOMPADDING",(0,0),(-1,-1),3),
            ("LEFTPADDING",(0,0),(-1,-1),4), ("RIGHTPADDING",(0,0),(-1,-1),4),
        ]))
        lv_t.cornerRadii = [6,6,6,6]

        inner = Table([[lv_t],[p(rec["title"], S_RTT)],[p(rec["desc"], S_RDD)]],
                      colWidths=[RIW])
        inner.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),2),  ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ]))

        card = Table([[inner]], colWidths=[RW])
        card.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), bg),
            ("BOX",(0,0),(-1,-1),0.5, border),
            ("TOPPADDING",(0,0),(-1,-1),7),  ("BOTTOMPADDING",(0,0),(-1,-1),7),
            ("LEFTPADDING",(0,0),(-1,-1),RP),("RIGHTPADDING",(0,0),(-1,-1),RP),
        ]))
        card.cornerRadii = [10,10,10,10]
        rec_cells.append(card)
        rec_widths.append(RW)
        if i < n - 1:
            rec_cells.append(Spacer(RG, 1))
            rec_widths.append(RG)

    if rec_cells:
        rec_row = Table([rec_cells], colWidths=rec_widths)
        zpad(rec_row)
        rec_row.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),  ("BOTTOMPADDING",(0,0),(-1,-1),0),
        ]))
        story.append(rec_row)

    story.append(Spacer(1, 2.5*mm))
    story.append(HRFlowable(width=PW, thickness=0.4, color=C_HAIR))
    story.append(Spacer(1, 1.5*mm))
    story.append(p(
        f"Dokumen ini digenerate otomatis oleh PortalAkademik \u2014 "
        f"{time.strftime('%d %B %Y, %H:%M:%S')} WIB",
        S_FOOT
    ))

    doc.build(story)
    return buf.getvalue()


# ==========================================
# 6. TOPBAR
# ==========================================
with st.container(key="topbar"):
    tb1, tb2 = st.columns([5, 5], vertical_alignment="center")
    with tb1:
        st.markdown(f"""
            <div style="display:flex;align-items:center;gap:12px;">
                <div style="width:46px;height:46px;border-radius:12px;background:linear-gradient(135deg,#EEF2FF,#F3F0FF);
                            border:1.5px solid #c5d0fa;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
                    {LOGO_SVG.format(size=24, color="#3B5BDB")}
                </div>
                <div>
                    <p style="font-family:'Space Grotesk',sans-serif;color:#0D1117;font-size:18px;font-weight:700;margin:0;letter-spacing:-0.5px;">
                        <span style="color:#3B5BDB;">Portal</span>Akademik</p>
                    <p style="font-family:'Fira Code',monospace;color:#64748B;font-size:10px;text-transform:uppercase;letter-spacing:0.7px;margin:2px 0 0 0;">
                        AI Evaluasi Studi</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
    with tb2:
        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
        n1, n2, n3, n4 = st.columns([2, 2, 4, 4])
        with n3:
            if st.button("Form Evaluasi", use_container_width=True,
                         type="primary" if st.session_state.page=="input" else "secondary"):
                switch_page("input")
        with n4:
            if st.button("Dashboard Hasil", use_container_width=True,
                         type="primary" if st.session_state.page=="log" else "secondary"):
                switch_page("log")

with st.container(key="topbar_spacer"):
    st.markdown("<div style='height:74px;'></div>", unsafe_allow_html=True)

# ==========================================
# 7. MAIN CONTENT
# ==========================================

# ── FORM INPUT ──
if st.session_state.page == "input":
    st.markdown("""<div style="margin-bottom:22px;">
        <h2 style="font-family:'Space Grotesk',sans-serif;font-size:26px;font-weight:700;color:#0D1117;margin:0 0 6px 0;letter-spacing:-0.5px;">Form Evaluasi Mahasiswa</h2>
        <p style="font-family:'Plus Jakarta Sans',sans-serif;font-size:14px;color:#64748B;margin:0;">Isi data di bawah untuk memproses prediksi risiko dropout mahasiswa.</p>
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
        submitted = st.form_submit_button("Proses Evaluasi Data", type="primary", use_container_width=True)

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
                "SKS1": sem1_app,  "IPK1": sem1_grd_indo,
                "SKS2": sem2_app,  "IPK2": sem2_grd_indo,
                "TotalSKS":  sem1_app + sem2_app,
                "RerataIPK": round((sem1_grd_indo + sem2_grd_indo)/2, 2),
                "SPP": tuit_fees,
                "SkorAman":   round(safe, 2),
                "SkorRisiko": round(risk, 2),
                "Class":  prediction[0],
                "Status": "AMAN (RETENSI TINGGI)" if safe >= 75
                          else ("WASPADA (RISIKO SEDANG)" if safe >= 40
                          else "KERENTANAN DROPOUT TINGGI")
            })
            st.session_state.selected_student_uid = nim
            switch_page("log")


# ── DASHBOARD ──
elif st.session_state.page == "log":
    if len(st.session_state.history) == 0:
        st.info("Belum ada data mahasiswa yang dievaluasi.")
        if st.button("Mulai Evaluasi Pertama", type="primary", use_container_width=True):
            switch_page("input")
        st.stop()

    df_full       = pd.DataFrame(st.session_state.history)
    total_mhs     = len(df_full)
    total_aman    = sum(1 for h in st.session_state.history if h['SkorAman'] >= 75)
    total_waspada = sum(1 for h in st.session_state.history if 40 <= h['SkorAman'] < 75)
    total_risiko  = sum(1 for h in st.session_state.history if h['SkorAman'] < 40)

    st.markdown("""<div style="margin:0 0 20px 0;">
        <h2 style="font-family:'Space Grotesk',sans-serif;font-size:26px;font-weight:700;color:#0D1117;margin:0 0 6px 0;letter-spacing:-0.5px;">Dashboard Hasil Evaluasi</h2>
        <p style="font-family:'Plus Jakarta Sans',sans-serif;font-size:14px;color:#64748B;margin:0;">Ringkasan seluruh mahasiswa yang telah dievaluasi pada sesi ini.</p>
    </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <div class="stat-cards-row">
        <div class="stat-card" style="border-top:4px solid #3B5BDB;">
            <div class="sc-label">Total Dievaluasi</div>
            <div class="sc-value" style="color:#0D1117;">{total_mhs}</div>
            <div class="sc-sub">Orang</div>
        </div>
        <div class="stat-card" style="border-top:4px solid #0D9488;">
            <div class="sc-label">Retensi Tinggi (&ge;75%)</div>
            <div class="sc-value" style="color:#0D9488;">{total_aman}</div>
            <div class="sc-sub">Mahasiswa aman</div>
        </div>
        <div class="stat-card" style="border-top:4px solid #D97706;">
            <div class="sc-label">Waspada (40&#8211;74%)</div>
            <div class="sc-value" style="color:#D97706;">{total_waspada}</div>
            <div class="sc-sub">Perlu monitoring</div>
        </div>
        <div class="stat-card" style="border-top:4px solid #DC2626;">
            <div class="sc-label">Risiko Tinggi (&lt;40%)</div>
            <div class="sc-value" style="color:#DC2626;">{total_risiko}</div>
            <div class="sc-sub">Kasus terdeteksi</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f'<div class="section-label">{TABLE_ICON.format(size=14, color="#3B5BDB")} Tabel Riwayat Aktivitas Evaluasi</div>', unsafe_allow_html=True)
    st.dataframe(df_full[["NIM","Nama","TotalSKS","RerataIPK","SPP","Status"]],
                 use_container_width=True, height=min(40+len(df_full)*35, 240))

    st.markdown("<div style='height:24px;'></div>", unsafe_allow_html=True)
    st.markdown(f'<div class="section-label">{TARGET_ICON.format(size=14, color="#3B5BDB")} Detail Performa Mahasiswa Terpilih</div>', unsafe_allow_html=True)

    options_list = [f"[{item['NIM']}] {item['Nama']}" for item in st.session_state.history]
    default_index = 0
    if st.session_state.selected_student_uid:
        for idx, item in enumerate(st.session_state.history):
            if item["NIM"] == st.session_state.selected_student_uid:
                default_index = idx
    selection = st.radio("Pilih Mahasiswa:", options=options_list, index=default_index, horizontal=True)
    current_nim = selection.split("]")[0].replace("[","").strip()
    if current_nim != st.session_state.selected_student_uid:
        st.session_state.selected_student_uid = current_nim
        st.rerun()

    active_data = next(item for item in st.session_state.history
                       if item["NIM"] == st.session_state.selected_student_uid)
    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

    _sa = active_data["SkorAman"]
    if _sa >= 75:
        badge_color = "#0D9488"; badge_bg = "#CCFBF1"; badge_icon = CHECK_ICON.format(size=14, color="#0D9488")
    elif _sa >= 40:
        badge_color = "#D97706"; badge_bg = "#FEF3C7"; badge_icon = ALERT_ICON.format(size=14, color="#D97706")
    else:
        badge_color = "#DC2626"; badge_bg = "#FEE2E2"; badge_icon = ALERT_ICON.format(size=14, color="#DC2626")

    with st.container(border=True):
        st.markdown(f"""
            <div style="display:flex;align-items:center;justify-content:space-between;padding:4px 4px 16px 4px;border-bottom:1px solid #F1F5F9;margin-bottom:20px;">
                <div style="font-family:'Space Grotesk',sans-serif;font-size:20px;font-weight:700;color:#0D1117;letter-spacing:-0.3px;">
                    {active_data['Nama']}
                    <span style="font-family:'Plus Jakarta Sans',sans-serif;font-size:14px;font-weight:500;color:#64748B;margin-left:8px;">NIM {active_data['NIM']}</span>
                </div>
                <span style="display:inline-flex;align-items:center;gap:8px;font-family:'Fira Code',monospace;font-weight:600;
                             font-size:12px;letter-spacing:0.6px;text-transform:uppercase;padding:10px 22px;border-radius:50px;
                             color:{badge_color};background:{badge_bg};">
                    {badge_icon}&nbsp;{active_data['Status']}
                </span>
            </div>
        """, unsafe_allow_html=True)
        pl, pm, pr = st.columns([1,2,2], gap="large")
        with pl:
            if _sa >= 75:   _cc="#3B5BDB";_cb="#F8FAFF";_cbr="#c5d0fa";_cl="AMAN · Retensi Tinggi"
            elif _sa >= 40: _cc="#D97706";_cb="#FFFBEB";_cbr="#FDE68A";_cl="WASPADA · Risiko Sedang"
            else:           _cc="#DC2626";_cb="#FFF5F5";_cbr="#fecaca";_cl="DROPOUT · Risiko Tinggi"
            st.markdown(f"""
                <div style="display:flex;flex-direction:column;gap:12px;">
                    <div style="background:{_cb};border:1.5px solid {_cbr};border-radius:14px;padding:16px 18px;">
                        <div style="font-family:'Fira Code',monospace;font-size:10px;color:{_cc};text-transform:uppercase;letter-spacing:0.7px;margin-bottom:4px;">Peluang Kelulusan</div>
                        <div style="font-family:'Space Grotesk',sans-serif;font-size:32px;font-weight:700;color:{_cc};line-height:1;">{active_data['SkorAman']}%</div>
                        <div style="font-family:'Fira Code',monospace;font-size:9.5px;color:{_cc};margin-top:6px;opacity:0.8;">{_cl}</div>
                    </div>
                    <div style="background:#FFF5F5;border:1.5px solid #fecaca;border-radius:14px;padding:16px 18px;">
                        <div style="font-family:'Fira Code',monospace;font-size:10px;color:#DC2626;text-transform:uppercase;letter-spacing:0.7px;margin-bottom:4px;">Faktor Risiko Dropout</div>
                        <div style="font-family:'Space Grotesk',sans-serif;font-size:32px;font-weight:700;color:#DC2626;line-height:1;">{active_data['SkorRisiko']}%</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with pm:
            st.markdown(f'<div style="font-family:\'Fira Code\',monospace;font-size:11.5px;font-weight:600;color:#0D1117;margin-bottom:8px;">{BARCHART_ICON.format(size=14,color="#3B5BDB")} &nbsp;Fluktuasi IPK per Semester</div>', unsafe_allow_html=True)
            st.line_chart(pd.DataFrame({"IPK":[active_data["IPK1"],active_data["IPK2"]]},
                                       index=["Semester 1","Semester 2"]), height=165, use_container_width=True)
        with pr:
            st.markdown(f'<div style="font-family:\'Fira Code\',monospace;font-size:11.5px;font-weight:600;color:#0D1117;margin-bottom:8px;">{BARCHART_ICON.format(size=14,color="#7048E8")} &nbsp;Volume SKS per Semester</div>', unsafe_allow_html=True)
            st.bar_chart(pd.DataFrame({"Volume SKS":[active_data["SKS1"],active_data["SKS2"]]},
                                      index=["Semester 1","Semester 2"]), height=165, use_container_width=True)

    st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)

    fa, fb = st.columns([1,1], gap="large")
    with fa:
        with st.container(border=True):
            st.markdown(f'<div class="section-label">{ZAPPER_ICON.format(size=14,color="#3B5BDB")} Indikator Faktor Keputusan Model</div>', unsafe_allow_html=True)
            scores = factor_score(active_data)
            for label, score in scores.items():
                bar_color = "#0D9488" if score>=70 else "#D97706" if score>=40 else "#DC2626"
                txt_bg    = "#CCFBF1" if score>=70 else "#FEF3C7" if score>=40 else "#FEE2E2"
                st.markdown(f"""
                    <div style="margin-bottom:16px;">
                        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
                            <span style="font-family:'Plus Jakarta Sans',sans-serif;font-size:13.5px;font-weight:600;color:#0D1117;">{label}</span>
                            <span style="font-family:'Fira Code',monospace;font-size:12px;font-weight:600;color:{bar_color};background:{txt_bg};padding:4px 12px;border-radius:20px;">{score}%</span>
                        </div>
                        <div style="width:100%;height:10px;background:#F1F5F9;border-radius:99px;overflow:hidden;">
                            <div style="width:{score}%;height:100%;background:{bar_color};border-radius:99px;"></div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
    with fb:
        with st.container(border=True):
            st.markdown(f'<div class="section-label">{BULB_ICON.format(size=14,color="#3B5BDB")} Rekomendasi Intervensi Akademik</div>', unsafe_allow_html=True)
            recs = generate_recommendations(active_data)
            level_colors = {
                "KRITIS":  ("DC2626","FEE2E2","FECACA"),
                "WASPADA": ("D97706","FEF3C7","FDE68A"),
                "BAIK":    ("0D9488","CCFBF1","99F6E4")
            }
            for rec in recs:
                fg, bg, border = level_colors.get(rec["level"],("3B5BDB","EEF2FF","C5D0FA"))
                st.markdown(f"""
                    <div style="display:flex;gap:14px;align-items:flex-start;background:#{bg};border:1px solid #{border};border-radius:14px;padding:16px;margin-bottom:14px;">
                        <div style="background:rgba(255,255,255,0.7);border-radius:8px;padding:4px 10px;color:#{fg};font-family:'Fira Code',monospace;font-weight:700;font-size:9.5px;white-space:nowrap;">{rec['level']}</div>
                        <div>
                            <div style="font-family:'Space Grotesk',sans-serif;font-size:14.5px;font-weight:700;color:#0D1117;margin-bottom:4px;">{rec['title']}</div>
                            <div style="font-family:'Plus Jakarta Sans',sans-serif;font-size:12.5px;color:#475569;line-height:1.55;">{rec['desc']}</div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)

    st.markdown("<div style='height:24px;'></div>", unsafe_allow_html=True)

    # ── EKSPOR LAPORAN PDF ──
    with st.container(border=True):
        col_export, col_btn = st.columns([2, 1], gap="large", vertical_alignment="center")
        with col_export:
            st.markdown(f"""
                <div style="display:flex;align-items:center;gap:16px;">
                    <div style="width:52px;height:52px;border-radius:14px;background:linear-gradient(135deg,#EEF2FF,#F3F0FF);
                                border:1.5px solid #c5d0fa;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
                        {DOWNLOAD_ICON.format(size=24, color="#3B5BDB")}
                    </div>
                    <div>
                        <div style="font-family:'Space Grotesk',sans-serif;font-size:18px;font-weight:700;color:#0D1117;">Download Laporan PDF</div>
                        <div style="font-family:'Plus Jakarta Sans',sans-serif;font-size:13px;color:#64748B;margin-top:2px;">
                            Laporan PDF 1 halaman: identitas, data akademik, grafik, faktor keputusan &amp; rekomendasi intervensi.
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with col_btn:
            with st.spinner("Menyiapkan PDF..."):
                pdf_data = generate_pdf_bytes(active_data)
            st.download_button(
                label="Unduh Laporan PDF",
                data=pdf_data,
                file_name=f"Laporan_{active_data['NIM']}_{active_data['Nama']}.pdf",
                mime="application/pdf",
                use_container_width=True,
                type="primary",
                key="dl_pdf_direct"
            )