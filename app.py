# -*- coding: utf-8 -*-
"""
Dashboard Pemeringkatan Universitas Halu Oleo (UHO)
Gaya akademis-formal, kontras tinggi, dibangun dengan Streamlit.
"""

import os
import io
import base64
import textwrap

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from PIL import Image

# Konfigurasi halaman (page_title & page_icon) dipasang lebih ke bawah,
# tepat setelah logo resmi UHO didefinisikan, supaya favicon tab browser
# memakai logo asli UHO alih-alih emoji generik.

def block(html_str: str):
    """Render blok HTML dengan aman.
    Semua string HTML multi-baris DI-DEDENT terlebih dahulu, karena Markdown/
    CommonMark akan membaca teks yang menjorok >=4 spasi sebagai *code block*
    mentah (inilah sebab tampilan identitas sebelumnya muncul sebagai teks
    kode, bukan halaman yang dirender)."""
    st.markdown(textwrap.dedent(html_str), unsafe_allow_html=True)

# ==========================================================================
# 2. PALET WARNA RESMI UHO
#    Warna dasar almamater UHO adalah KUNING (sumber: profil resmi & infobox
#    Wikipedia "Colors: Yellow"), dipadukan navy gelap untuk kontras teks dan
#    hijau untuk aksen "Kampus Hijau Bumi Tridharma".
# ==========================================================================
UHO_NAVY = "#12213D"
UHO_INK = "#101828"        # warna teks utama, kontras tinggi di atas putih
UHO_GOLD = "#F2B705"
UHO_GOLD_DEEP = "#8A6400"  # emas gelap — dipakai untuk TEKS agar tetap terbaca di atas putih
UHO_GREEN = "#155D34"
UHO_RED = "#9B2C22"
UHO_GREY = "#4B5563"       # abu-abu gelap (bukan abu-abu terang) demi kontras
UHO_BORDER = "#D8DCE3"
UHO_BG = "#F5F6F8"
UHO_CARD = "#FFFFFF"

IDENTITAS_UHO = [
    ("Nama Resmi", "Universitas Halu Oleo (UHO)"),
    ("Status", "Perguruan Tinggi Negeri Badan Layanan Umum (PTN-BLU)"),
    ("Didirikan", "19 Agustus 1981 (Keppres RI No. 37 Tahun 1981)"),
    ("Urutan PTN", "PTN ke-42 di Indonesia; PTN pertama di Sulawesi Tenggara"),
    ("Lokasi", "Kampus Hijau Bumi Tridharma, Anduonohu, Kota Kendari, Sulawesi Tenggara"),
    ("Plt. Rektor", "Prof. Dr. Khairul Munadi, S.T., M.Eng. (menjabat sejak 15 Juni 2026, SK Mendiktisaintek No. 0065/M/KP.10.00/2026)"),
    ("Akreditasi Institusi", "Unggul — Nilai 363 (SK BAN-PT No. 146/SK/BAN-PT/Ak/PT/V/2026, 5 Mei 2026, berlaku 5 tahun hingga 2031)"),
    ("Situs Resmi", "https://uho.ac.id"),
]

# ==========================================================================
# 3. DATA PEMERINGKATAN
#    Sumber data DIPISAH ke berkas "data_ranking.csv" (format tabel biasa,
#    bisa dibuka & diedit lewat Excel/Google Sheets) supaya pembaruan data
#    di tahun-tahun berikutnya TIDAK PERLU mengedit kode ini sama sekali.
#    Format kolom: Lembaga, Tahun, Nasional, Dunia (kosongkan sel bila
#    lembaga tsb belum memeringkat UHO pada tahun itu).
#    Untuk menambah data tahun baru (mis. 2027): cukup tambahkan baris baru
#    di data_ranking.csv untuk tiap lembaga dengan Tahun=2027, lalu commit
#    ke GitHub — dashboard otomatis menampilkannya tanpa perlu ubah app.py.
#    Bila berkas CSV tidak ditemukan, dashboard tetap berjalan memakai data
#    bawaan (DEFAULT_RAW) di bawah ini sebagai cadangan (fallback).
# ==========================================================================
DEFAULT_RAW = {
    "THE WUR": {
        2021: (None, None), 2022: (None, None), 2023: (None, None),
        2024: (None, None), 2025: (11, 1501), 2026: (35, 1501),
    },
    "QS WUR": {
        2021: (None, None), 2022: (None, None), 2023: (None, None),
        2024: (None, None), 2025: (None, None), 2026: (None, None),
    },
    "THE Interdisciplinary Science Rank (Physics)": {
        2021: (None, None), 2022: (None, None), 2023: (None, None),
        2024: (None, None), 2025: (28, 801), 2026: (27, 801),
    },
    "THE Sustainability Ranking": {
        2021: (None, None), 2022: (None, None), 2023: (None, None),
        2024: (None, None), 2025: (33, None), 2026: (35, 801),
    },
    "THE ASIA Ranking": {
        2021: (None, None), 2022: (None, None), 2023: (None, None),
        2024: (None, None), 2025: (None, None), 2026: (23, 801),
    },
    "UI GreenMetric": {
        2021: (26, 236), 2022: (30, 263), 2023: (31, 240),
        2024: (32, 243), 2025: (28, 193), 2026: (28, 193),
    },
    "Webometrics": {
        2021: (None, None), 2022: (56, 4149), 2023: (32, 2678),
        2024: (33, 2392), 2025: (49, 2894), 2026: (30, 2481),
    },
    "UniRank": {
        2021: (None, None), 2022: (None, None), 2023: (69, 4124),
        2024: (77, 3477), 2025: (54, 3097), 2026: (63, 2901),
    },
    "EduRank": {
        2021: (None, None), 2022: (None, None), 2023: (None, None),
        2024: (54, 3142), 2025: (48, 2616), 2026: (42, 2200),
    },
    "Scimago Institutions Ranking": {
        2021: (24, 4935), 2022: (34, 5311), 2023: (40, 5481),
        2024: (47, 5207), 2025: (52, 6191), 2026: (47, 5979),
    },
}

_DATA_CSV_CANDIDATES = ["data_ranking.csv", "data/data_ranking.csv"]

def _find_data_csv_path():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for rel in _DATA_CSV_CANDIDATES:
        path = os.path.join(base_dir, rel)
        if os.path.exists(path):
            return path
    return None

def load_ranking_data():
    """Muat data peringkat dari data_ranking.csv bila tersedia; kalau tidak,
    pakai DEFAULT_RAW bawaan. Mengembalikan (raw_dict, lembaga_order, is_csv)."""
    csv_path = _find_data_csv_path()
    if csv_path is None:
        return DEFAULT_RAW, list(DEFAULT_RAW.keys()), False
    try:
        df_csv = pd.read_csv(csv_path)
        df_csv.columns = [c.strip() for c in df_csv.columns]
        required = {"Lembaga", "Tahun", "Nasional", "Dunia"}
        if not required.issubset(set(df_csv.columns)):
            return DEFAULT_RAW, list(DEFAULT_RAW.keys()), False
        raw = {}
        order = []
        for _, row in df_csv.iterrows():
            lembaga = str(row["Lembaga"]).strip()
            if lembaga not in raw:
                raw[lembaga] = {}
                order.append(lembaga)
            tahun = int(row["Tahun"])
            nas = row["Nasional"]
            dun = row["Dunia"]
            nas = None if pd.isna(nas) else int(nas)
            dun = None if pd.isna(dun) else int(dun)
            raw[lembaga][tahun] = (nas, dun)
        if not raw:
            return DEFAULT_RAW, list(DEFAULT_RAW.keys()), False
        return raw, order, True
    except Exception:
        # Kalau CSV rusak/format salah, tetap jalan dengan data bawaan
        # daripada membuat seluruh dashboard error.
        return DEFAULT_RAW, list(DEFAULT_RAW.keys()), False

RAW, LEMBAGA_LIST, DATA_FROM_CSV = load_ranking_data()

# Tahun dihitung otomatis dari data yang ada — begitu tahun baru (mis. 2027)
# ditambahkan di CSV, opsi tahun di filter & grafik ikut bertambah otomatis.
YEARS = sorted({year for data in RAW.values() for year in data.keys()})

DESKRIPSI_LEMBAGA = {
    "THE WUR": "Times Higher Education World University Rankings — pengajaran, riset, sitasi, kolaborasi internasional, dan income industri.",
    "QS WUR": "Quacquarelli Symonds World University Rankings — reputasi akademik, reputasi pemberi kerja, dan rasio dosen-mahasiswa.",
    "THE Interdisciplinary Science Rank (Physics)": "THE Interdisciplinary Science Rankings bidang Fisika — kualitas riset lintas-disiplin ilmu fisika.",
    "THE Sustainability Ranking": "THE Impact Rankings — kontribusi universitas terhadap Sustainable Development Goals (SDGs) PBB.",
    "THE ASIA Ranking": "THE Asia University Rankings — pemeringkatan khusus perguruan tinggi se-kawasan Asia.",
    "UI GreenMetric": "UI GreenMetric World University Ranking — pemeringkatan kampus berkelanjutan (green campus), digagas Universitas Indonesia.",
    "Webometrics": "Ranking Web of Universities — visibilitas dan dampak akademik berbasis kehadiran web (CSIC, Spanyol).",
    "UniRank": "4ICU UniRank — indeks popularitas web universitas dunia berbasis lalu lintas dan keterhubungan domain.",
    "EduRank": "EduRank.org — reputasi akademik, kekuatan alumni, dan sitasi riset.",
    "Scimago Institutions Ranking": "SCImago Institutions Rankings — output riset, inovasi (paten), dan visibilitas web (data Scopus).",
}
DESKRIPSI_DEFAULT = "Deskripsi lembaga ini belum ditambahkan — lengkapi di kamus DESKRIPSI_LEMBAGA pada kode bila diperlukan."

# ==========================================================================
# 4. TRANSFORMASI DATA
#    Rata-rata per tahun kini DIHITUNG OTOMATIS (bukan angka tetap) dari
#    seluruh nilai yang tersedia pada RAW — sehingga ikut ter-update begitu
#    data tahun baru ditambahkan di data_ranking.csv.
# ==========================================================================
def _hitung_rata_rata():
    hasil = {}
    for year in YEARS:
        nas_vals = [RAW[l][year][0] for l in LEMBAGA_LIST if year in RAW[l] and RAW[l][year][0] is not None]
        dun_vals = [RAW[l][year][1] for l in LEMBAGA_LIST if year in RAW[l] and RAW[l][year][1] is not None]
        hasil[year] = {
            "Nasional": (sum(nas_vals) / len(nas_vals)) if nas_vals else None,
            "Dunia": (sum(dun_vals) / len(dun_vals)) if dun_vals else None,
        }
    return hasil

RATA_RATA = _hitung_rata_rata()


# ==========================================================================
# 5. METODOLOGI PENSKORAN (A/B/C) — KLASIFIKASI INTERNAL DASHBOARD
# ==========================================================================
GRADE_TIERS_NASIONAL = [
    (10, "A+", UHO_GREEN, "Sangat Unggul"),
    (25, "A", UHO_GREEN, "Unggul"),
    (50, "B+", UHO_GOLD_DEEP, "Sangat Baik"),
    (100, "B", UHO_GOLD_DEEP, "Baik"),
    (250, "C+", UHO_RED, "Cukup"),
    (float("inf"), "C", UHO_RED, "Perlu Peningkatan"),
]

GRADE_TIERS_DUNIA = [
    (500, "A+", UHO_GREEN, "Sangat Unggul"),
    (1000, "A", UHO_GREEN, "Unggul"),
    (2000, "B+", UHO_GOLD_DEEP, "Sangat Baik"),
    (4000, "B", UHO_GOLD_DEEP, "Baik"),
    (8000, "C+", UHO_RED, "Cukup"),
    (float("inf"), "C", UHO_RED, "Perlu Peningkatan"),
]

def rank_to_grade(rank, scope="Nasional"):
    if rank is None or pd.isna(rank):
        return ("N/A", UHO_GREY, "Belum Terpetakan")
    tiers = GRADE_TIERS_NASIONAL if scope == "Nasional" else GRADE_TIERS_DUNIA
    for threshold, letter, color, label in tiers:
        if rank <= threshold:
            return (letter, color, label)
    return ("C", UHO_RED, "Perlu Peningkatan")

def latest_value(lembaga, scope):
    idx = 0 if scope == "Nasional" else 1
    for year in reversed(YEARS):
        val = RAW.get(lembaga, {}).get(year, (None, None))[idx]
        if val is not None:
            return year, val
    return None, None

def previous_value(lembaga, scope, before_year):
    idx = 0 if scope == "Nasional" else 1
    for year in reversed(YEARS):
        if year >= before_year:
            continue
        val = RAW.get(lembaga, {}).get(year, (None, None))[idx]
        if val is not None:
            return year, val
    return None, None

# ==========================================================================
# 6. LOGO RESMI UHO
#    Prioritas 1: jika berkas logo resmi (mis. "logo_uho.png") ditaruh di
#    folder yang sama dengan app.py, dashboard OTOMATIS memakai berkas asli
#    tersebut — ini cara paling akurat untuk memastikan logo 100% sesuai
#    identitas resmi kampus.
#    Logo resmi berikut SUDAH DITANAMKAN LANGSUNG (base64) di dalam kode ini,
#    sehingga app.py bersifat mandiri/self-contained — tidak perlu berkas
#    logo terpisah untuk berjalan atau untuk diserahkan ke pihak lain.
#    Jika suatu saat logo resmi berganti, cukup taruh berkas baru bernama
#    "logo_uho.png" (atau .jpg/.svg) di folder yang sama dengan app.py; berkas
#    tersebut akan otomatis dipakai TANPA perlu mengedit base64 di bawah ini.
# ==========================================================================
_LOGO_CANDIDATES = [
    "logo_uho.png", "logo_uho.jpg", "logo_uho.jpeg", "logo_uho.svg",
    "assets/logo_uho.png", "assets/logo_uho.jpg",
]

# --- Logo resmi Universitas Halu Oleo, ditanam sebagai base64 (PNG) ---
UHO_LOGO_BASE64_PNG = "iVBORw0KGgoAAAANSUhEUgAAA+gAAAPoCAMAAAENehRAAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAYBQTFRF/oUQICAg/f39/u7uNEGVr6+v/t6//pUw/l4Af39//s3N/s6fv7+//u7e6+lW/iAg/r5//tavvLpFzMpLCQoY/p+e3NlRn5+fsrfWnZs62Nrq/q1f/saP/rVv/t7eQUybAwQJcHBwgIe8fnwuj4+P/qVPjJPCTVihEBAQQEBA/n9/jYw0Xl0jWmSoc3u1ras/Tk4dXy8Au1wA/t3Hbm0ov8PdT09P/jQDHx8MpqvQ/o+P+/hc/gAA/hQA/m9v7e3t+/hb7u7uMDAw+vhbEBAGPz4XX19fzs7O8fL33t7eZ3CvzM/kLy8RDRAlmZ/JMTyL/vbub29vJzBvGiBK5ebx/ubO7+/v/p1AIShdFBg4LjiCKjR4+vhc/o0gHSRT39/fJCxmYGBgEBQuz8/PFxxB/r+/UFBQ/kBA/hAQ/l9fmaDK/jAw/q+v/k9O/vXyPz4+/g8P3d3dZ3Cw/oUP8fL4/ggAgICA/ndf/opqb0cg/sav/n0A/v7+AAAANECU////kpHQWwAAAIB0Uk5T/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////wA4BUtnAAEm+UlEQVR42uyb61MUSRLAGw/W671Z1gBPFokLvR1OB0TXD+wGbM2HngeigAoeIqIXt3q3ERdxcX9BVw7/+tX7mdVTM8OuXyidB9Pd9avMysyqrsouLr9kKa7p1/Rr+gTlDODL0UGUL0NnYMrKLPxiBrGpLL3pFVBMrXDqlGm7oJgGDB1FLW0j1JGN34rej2QGGI3cH17IM9avnh5pe46zeQEoo164WjqrsB8iRqaERzu5/CKT3TNVt8BHmwa4qmnl6b/Id2ypbYDPo0S5zY+WRjufroR+ptSKiRyVC3bSiTr9SugmpI0yC8BcJn48fQG0c+UX1eAroAtB4CKon/0m3+WX2ASyhM+g83oiNKUH4sv8/Lxph3vKOeQI30Rf18HNVfsr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"

def _find_override_logo_tag(width_px: int):
    """Cek apakah ada berkas logo custom di folder aplikasi yang sengaja
    dipakai untuk MENGGANTIKAN logo bawaan (mis. saat logo resmi diperbarui)."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for rel in _LOGO_CANDIDATES:
        path = os.path.join(base_dir, rel)
        if os.path.exists(path):
            ext = rel.rsplit(".", 1)[-1].lower()
            mime = "svg+xml" if ext == "svg" else ("jpeg" if ext in ("jpg", "jpeg") else "png")
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            return f'<img src="data:image/{mime};base64,{b64}" style="width:{width_px}px; height:auto;" alt="Logo Universitas Halu Oleo"/>'
    return None

def logo_markup(width_px: int = 84) -> str:
    override = _find_override_logo_tag(width_px)
    if override:
        return override
    if UHO_LOGO_BASE64_PNG and not UHO_LOGO_BASE64_PNG.startswith("__"):
        return (f'<img src="data:image/png;base64,{UHO_LOGO_BASE64_PNG}" '
                f'style="width:{width_px}px; height:auto;" alt="Logo Universitas Halu Oleo"/>')
    # --- fallback terakhir: emblem skematik (dipakai hanya jika base64 kosong) ---
    return f"""
    <svg viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg" width="{width_px}" height="{width_px}"
         role="img" aria-label="Emblem Universitas Halu Oleo">
      <g transform="translate(100,100)">
        {''.join([
            f'<path transform="rotate({i*72})" d="M0,-94 C26,-72 32,-28 0,-4 '
            f'C-32,-28 -26,-72 0,-94 Z" fill="{UHO_GOLD}" stroke="{UHO_NAVY}" stroke-width="2.5"/>'
            for i in range(5)
        ])}
        <circle r="44" fill="{UHO_NAVY}" stroke="{UHO_GOLD}" stroke-width="3"/>
        <path d="M-28,6 C-15,-4 -5,-4 0,2 C5,-4 15,-4 28,6 L28,15 C15,7 5,7 0,13 C-5,7 -15,7 -28,15 Z"
              fill="#FFFFFF"/>
        <line x1="0" y1="2" x2="0" y2="13" stroke="{UHO_NAVY}" stroke-width="1.3"/>
        <rect x="-2.3" y="-28" width="4.6" height="22" rx="2" fill="#E9E9EE"/>
        <path d="M0,-44 C7,-37 7,-27 0,-21 C-7,-27 -7,-37 0,-44 Z" fill="{UHO_GOLD}"/>
      </g>
    </svg>
    """

LOGO_IS_OFFICIAL = True

# ==========================================================================
# 6b. KONFIGURASI HALAMAN (page_title & favicon)
#     Favicon tab browser memakai logo asli UHO (bukan emoji generik).
# ==========================================================================
try:
    _favicon = Image.open(io.BytesIO(base64.b64decode(UHO_LOGO_BASE64_PNG)))
except Exception:
    _favicon = None

st.set_page_config(
    page_title="Dashboard Pemeringkatan UHO",
    page_icon=_favicon,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==========================================================================
# 7. GAYA / CSS — LETTERHEAD AKADEMIS, KONTRAS DIPAKSA (ANTI-TEMA-BENTROK)
#    Catatan teknis: banyak masalah "teks tidak kelihatan" sebelumnya terjadi
#    karena CSS internal tema Streamlit (yang mengikuti preferensi terang/
#    gelap perangkat pengguna) memakai !important pada warna teks, sehingga
#    MENGALAHKAN warna custom yang tidak diberi !important. Solusinya: setiap
#    aturan warna di bawah ini diberi !important DAN specificity yang lebih
#    tinggi daripada aturan bawaan tema, sehingga warna selalu konsisten
#    apa pun tema/perangkat pengguna.
# ==========================================================================
NUM_FONT = "Georgia, 'Times New Roman', 'Cambria', serif"
UI_FONT = "'Segoe UI', 'Helvetica Neue', Arial, sans-serif"

block(f"""
<style>
    .stApp {{ background-color: {UHO_BG} !important; }}
    html, body {{ font-family: {UI_FONT}; }}

    /* --- Paksa warna teks gelap di SELURUH konten utama, apa pun tema --- */
    .stApp [data-testid="stAppViewContainer"],
    .stApp [data-testid="stAppViewContainer"] p,
    .stApp [data-testid="stAppViewContainer"] span,
    .stApp [data-testid="stAppViewContainer"] div,
    .stApp [data-testid="stAppViewContainer"] label,
    .stApp [data-testid="stAppViewContainer"] li,
    .stApp [data-testid="stAppViewContainer"] h1,
    .stApp [data-testid="stAppViewContainer"] h2,
    .stApp [data-testid="stAppViewContainer"] h3,
    .stApp [data-testid="stAppViewContainer"] h4,
    .stApp [data-testid="stAppViewContainer"] strong,
    .stApp [data-testid="stAppViewContainer"] b {{
        color: {UHO_INK} !important;
    }}

    /* --- Font angka akademis (serif) pada semua metrik & nilai numerik --- */
    [data-testid="stMetricValue"] {{
        font-family: {NUM_FONT} !important;
        font-weight: 700 !important;
        color: {UHO_NAVY} !important;
    }}
    [data-testid="stMetricLabel"] {{
        font-family: {UI_FONT} !important;
        color: {UHO_GREY} !important;
    }}
    [data-testid="stMetricDelta"] {{ font-family: {UI_FONT} !important; }}

    /* --- Sidebar (latar navy → teks terang), specificity dipaksa maksimal
           dengan trik pengulangan kelas ".stApp" agar mengalahkan aturan
           warna internal Streamlit apa pun (mis. pada <strong>, label
           widget, teks bantuan/help) yang sebelumnya bisa "kebobolan". --- */
    .stApp section[data-testid="stSidebar"] {{
        background-color: {UHO_NAVY} !important;
        border-right: 3px solid {UHO_GOLD};
    }}
    .stApp.stApp.stApp section[data-testid="stSidebar"],
    .stApp.stApp.stApp section[data-testid="stSidebar"] * {{
        color: #F2F4F8 !important;
    }}
    .stApp section[data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,0.18); }}

    /* --- Letterhead / kop dashboard (latar navy → teks terang) --- */
    .stApp .letterhead {{
        background: {UHO_NAVY};
        border-bottom: 5px solid {UHO_GOLD};
        padding: 18px 28px;
        margin: -1rem -1rem 22px -1rem;
        display: flex;
        align-items: center;
        gap: 20px;
    }}
    .stApp .letterhead .lh-title,
    .stApp .letterhead .lh-title * {{
        color: #FFFFFF !important;
        font-size: 22px;
        font-weight: 700;
        letter-spacing: 0.2px;
        margin: 0;
        font-family: {NUM_FONT};
    }}
    .stApp .letterhead .lh-sub, .stApp .letterhead .lh-sub * {{
        color: #D6DCE8 !important;
        font-size: 13.5px;
        margin: 2px 0 0 0;
    }}
    .stApp .letterhead .lh-meta, .stApp .letterhead .lh-meta * {{
        color: #AFC0DA !important;
        font-size: 12px;
        margin-top: 6px;
    }}

    .stApp .section-title, .stApp .section-title * {{
        color: {UHO_NAVY} !important;
        font-weight: 700;
        border-bottom: 3px solid {UHO_GOLD};
        padding-bottom: 6px;
        margin-top: 10px;
        margin-bottom: 14px;
        font-size: 19px;
        font-family: {NUM_FONT};
    }}
    .stApp .note-box, .stApp .note-box p, .stApp .note-box li, .stApp .note-box b,
    .stApp .note-box span, .stApp .note-box div {{
        color: {UHO_INK} !important;
    }}
    .note-box {{
        background: {UHO_CARD};
        border: 1px solid {UHO_BORDER};
        border-left: 5px solid {UHO_GOLD_DEEP};
        border-radius: 6px;
        padding: 14px 18px;
        margin-bottom: 14px;
        color: {UHO_INK};
        font-size: 14.5px;
        line-height: 1.55;
    }}
    .note-box b {{ color: {UHO_NAVY} !important; }}
    .stApp .grade-chip, .stApp .grade-chip * {{
        color: #FFFFFF !important;
        font-family: {NUM_FONT} !important;
    }}
    .grade-chip {{
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-width: 46px;
        height: 46px;
        border-radius: 8px;
        font-weight: 800;
        font-size: 17px;
        border: 1px solid rgba(0,0,0,0.15);
    }}
    .stApp .footer-note, .stApp .footer-note * {{ color: {UHO_GREY} !important; }}
    .footer-note {{
        font-size: 12px;
        text-align: center;
        margin-top: 36px;
        border-top: 1px solid {UHO_BORDER};
        padding-top: 12px;
    }}
    thead tr th {{ background-color: {UHO_NAVY} !important; color: #FFFFFF !important; }}

    /* --- Dropdown pilihan (Selectbox/MultiSelect di Streamlit versi baru
           memakai komponen "React-Aria ComboBox", BUKAN BaseWeb). Panelnya
           dikenali lewat data-testid="stSelectboxVirtualDropdown", dan tiap
           barisnya adalah <div role="option"> (bukan <li>). Diverifikasi
           langsung dari HTML sungguhan, bukan tebakan, supaya presisi. Ia
           dirender sebagai portal position:fixed menempel di document.body,
           jadi TIDAK cukup dijangkau lewat selector di dalam ".stApp" atau
           "section[data-testid='stSidebar']" — harus ditarget langsung
           sebagai elemen tingkat atas. --- */
    div[data-testid="stSelectboxVirtualDropdown"] {{
        background-color: #FFFFFF !important;
    }}
    div[data-testid="stSelectboxVirtualDropdown"] *,
    [role="listbox"],
    [role="listbox"] *,
    [role="option"],
    [role="option"] * {{
        color: {UHO_INK} !important;
    }}
    [role="option"] {{
        background-color: #FFFFFF !important;
    }}
    [role="option"]:hover,
    [role="option"][aria-selected="true"],
    [role="option"][data-focused="true"],
    [role="option"][data-selected="true"] {{
        background-color: #EEF0F3 !important;
    }}
    /* --- Kotak ISIAN di dalam sidebar (kotak Select box, kotak pencarian
           MultiSelect) memakai warna latar TERANG bawaan tema Streamlit
           (secara default, bukan navy), padahal warna teksnya ikut aturan
           umum sidebar (terang) — hasilnya teks nyaris tak terlihat di atas
           latar terang tersebut. Diverifikasi langsung lewat inspeksi HTML
           sungguhan (bukan tebakan): elemen sesungguhnya adalah <input>
           polos di dalam [data-testid="stSelectbox"] / "stMultiSelect",
           BUKAN struktur BaseWeb seperti pada versi Streamlit lama. --- */
    .stApp.stApp.stApp section[data-testid="stSidebar"] [data-testid="stSelectbox"] input,
    .stApp.stApp.stApp section[data-testid="stSidebar"] [data-testid="stMultiSelect"] input,
    .stApp.stApp.stApp section[data-testid="stSidebar"] [data-testid="stTextInput"] input,
    .stApp.stApp.stApp section[data-testid="stSidebar"] [data-testid="stNumberInput"] input {{
        color: {UHO_INK} !important;
    }}
    .stApp.stApp.stApp section[data-testid="stSidebar"] [data-testid="stSelectbox"] input::placeholder,
    .stApp.stApp.stApp section[data-testid="stSidebar"] [data-testid="stMultiSelect"] input::placeholder {{
        color: {UHO_GREY} !important;
        opacity: 1 !important;
    }}
</style>
""")

# ==========================================================================
# 8. SIDEBAR — NAVIGASI & FILTER
# ==========================================================================
with st.sidebar:
    block(f"""
    <div style="text-align:center; margin-bottom:4px;">{logo_markup(80)}</div>
    """)
    st.markdown("**UHO Ranking Dashboard**")
    st.caption("Universitas Halu Oleo — Kendari, Sulawesi Tenggara")
    st.markdown("---")
    halaman = st.radio(
        "Navigasi",
        ["Ringkasan Umum", "Detail per Lembaga", "Metodologi Skor", "Profil Perguruan Tinggi"],
        index=0,
    )
    st.markdown("---")
    tahun_terpilih = st.select_slider(
        "Rentang tahun ditampilkan",
        options=YEARS,
        value=(YEARS[0], YEARS[-1]),
    )
    lembaga_terpilih = st.multiselect(
        "Filter lembaga pemeringkat",
        LEMBAGA_LIST,
        default=LEMBAGA_LIST,
    )
    frekuensi_update = st.selectbox(
        "Frekuensi pembaruan data (informasi)",
        ["Tahunan", "Per 6 Bulan"],
        index=0,
        help="Sesuai arahan: data dapat dimutakhirkan setiap tahun atau setiap semester (6 bulan).",
    )
    st.markdown("---")
    _sumber_txt = "data_ranking.csv" if DATA_FROM_CSV else "data bawaan kode (data_ranking.csv tidak ditemukan)"
    st.caption(f"Sumber data: *{_sumber_txt}*, cakupan tahun {YEARS[0]}–{YEARS[-1]}. Skor huruf A–C adalah klasifikasi internal — lihat menu Metodologi Skor.")

years_range = [y for y in YEARS if tahun_terpilih[0] <= y <= tahun_terpilih[1]]
if not lembaga_terpilih:
    lembaga_terpilih = LEMBAGA_LIST

# ==========================================================================
# 9. LETTERHEAD / KOP DASHBOARD
# ==========================================================================
block(f"""
<div class="letterhead">
    <div>{logo_markup(72)}</div>
    <div>
        <p class="lh-title">DASHBOARD PEMERINGKATAN UNIVERSITAS HALU OLEO</p>
        <p class="lh-sub">Pemantauan capaian pemeringkatan nasional &amp; dunia periode {YEARS[0]}–{YEARS[-1]}</p>
        <p class="lh-meta">PTN-BLU &nbsp;·&nbsp; Kendari, Sulawesi Tenggara &nbsp;·&nbsp; Berdiri 19 Agustus 1981
        &nbsp;·&nbsp; Pembaruan data: {frekuensi_update}</p>
    </div>
</div>
""")

# ==========================================================================
# HALAMAN 1 — RINGKASAN UMUM
# ==========================================================================
if halaman == "Ringkasan Umum":

    tahun_acuan = years_range[-1] if years_range else YEARS[-1]
    nas_now = RATA_RATA.get(tahun_acuan, {}).get("Nasional")
    dunia_now = RATA_RATA.get(tahun_acuan, {}).get("Dunia")
    tahun_awal = years_range[0] if years_range else YEARS[0]
    nas_awal = RATA_RATA.get(tahun_awal, {}).get("Nasional")
    dunia_awal = RATA_RATA.get(tahun_awal, {}).get("Dunia")
    jumlah_terpetakan = sum(1 for l in lembaga_terpilih if latest_value(l, "Nasional")[1] is not None)
    grade_avg_nas, grade_color_nas, grade_label_nas = rank_to_grade(nas_now, "Nasional")
    grade_avg_dunia, grade_color_dunia, grade_label_dunia = rank_to_grade(dunia_now, "Dunia")

    col1, col2, col3 = st.columns(3)
    with col1:
        with st.container(border=True):
            st.metric(f"Rata-rata Peringkat Nasional ({tahun_acuan})", f"#{nas_now:.1f}" if nas_now is not None else "Belum ada data")
            if nas_awal and nas_now:
                arah = "membaik" if nas_now < nas_awal else ("menurun" if nas_now > nas_awal else "stabil")
                st.caption(f"{arah.capitalize()} sejak {tahun_awal}")
    with col2:
        with st.container(border=True):
            st.metric(f"Rata-rata Peringkat Dunia ({tahun_acuan})", f"#{dunia_now:,.0f}" if dunia_now is not None else "Belum ada data")
            if dunia_awal and dunia_now:
                arah = "membaik" if dunia_now < dunia_awal else ("menurun" if dunia_now > dunia_awal else "stabil")
                st.caption(f"{arah.capitalize()} sejak {tahun_awal}")
    with col3:
        with st.container(border=True):
            st.metric("Lembaga Aktif Memetakan UHO", f"{jumlah_terpetakan} / {len(lembaga_terpilih)}")
            st.caption(f"Memiliki data peringkat pada {tahun_acuan}")

    st.write("")
    col4, col5 = st.columns(2)
    with col4:
        with st.container(border=True):
            st.markdown(f"<div style='font-size:13px; color:{UHO_GREY};'>SKOR KLASIFIKASI RATA-RATA — NASIONAL</div>", unsafe_allow_html=True)
            cA, cB = st.columns([1, 3])
            with cA:
                st.markdown(f"<div class='grade-chip' style='background:{grade_color_nas}; font-size:22px; min-width:56px; height:56px;'>{grade_avg_nas}</div>", unsafe_allow_html=True)
            with cB:
                st.markdown(f"<div style='padding-top:8px; color:{UHO_INK};'>{grade_label_nas}</div>", unsafe_allow_html=True)
                st.caption("Berbasis rata-rata peringkat nasional")
    with col5:
        with st.container(border=True):
            st.markdown(f"<div style='font-size:13px; color:{UHO_GREY};'>SKOR KLASIFIKASI RATA-RATA — DUNIA</div>", unsafe_allow_html=True)
            cA, cB = st.columns([1, 3])
            with cA:
                st.markdown(f"<div class='grade-chip' style='background:{grade_color_dunia}; font-size:22px; min-width:56px; height:56px;'>{grade_avg_dunia}</div>", unsafe_allow_html=True)
            with cB:
                st.markdown(f"<div style='padding-top:8px; color:{UHO_INK};'>{grade_label_dunia}</div>", unsafe_allow_html=True)
                st.caption("Berbasis rata-rata peringkat dunia")

    st.write("")
    block(f'<div class="section-title">Tren Rata-rata Peringkat UHO ({YEARS[0]}–{YEARS[-1]})</div>')

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    yrs = years_range
    nas_series = [RATA_RATA[y]["Nasional"] for y in yrs]
    dunia_series = [RATA_RATA[y]["Dunia"] for y in yrs]

    fig.add_trace(go.Scatter(
        x=yrs, y=nas_series, name="Rata-rata Peringkat Nasional",
        mode="lines+markers", line=dict(color=UHO_GOLD_DEEP, width=3), marker=dict(size=9),
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=yrs, y=dunia_series, name="Rata-rata Peringkat Dunia",
        mode="lines+markers", line=dict(color=UHO_NAVY, width=3, dash="dot"), marker=dict(size=9),
    ), secondary_y=True)

    fig.update_yaxes(title_text="Peringkat Nasional (kecil = baik)", autorange="reversed", secondary_y=False)
    fig.update_yaxes(title_text="Peringkat Dunia (kecil = baik)", autorange="reversed", secondary_y=True)
    fig.update_xaxes(dtick=1, tickformat="d")
    fig.update_layout(
        height=420, plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=30, b=10, l=10, r=10),
        font=dict(family="Georgia, Times New Roman, serif", color=UHO_INK),
    )
    st.plotly_chart(fig, width="stretch")
    st.caption("Sumbu peringkat dibalik: posisi lebih tinggi pada grafik berarti peringkat yang lebih baik.")

    block('<div class="section-title">Papan Skor Seluruh Lembaga</div>')

    def df_table_height(n_rows: int) -> int:
        """Tinggi tabel dipaskan dengan jumlah baris (header + baris data),
        supaya tidak menyisakan baris kosong di bawah seperti sebelumnya."""
        return 38 + 35 * max(n_rows, 1) + 3

    tab_nas, tab_dunia, tab_snapshot = st.tabs([
        "Peringkat & Skor Nasional per Tahun",
        "Peringkat & Skor Dunia per Tahun",
        "Cuplikan Tahun Terbaru",
    ])

    def build_year_matrix(scope):
        """Bangun matriks Lembaga x Tahun berisi 'peringkat (skor)' + peta warna skor,
        mencakup SELURUH tahun yang sedang difilter (bukan hanya tahun terbaru)."""
        idx = 0 if scope == "Nasional" else 1
        text_rows, color_rows = {}, {}
        for lembaga in lembaga_terpilih:
            trow, crow = {}, {}
            for y in years_range:
                v = RAW.get(lembaga, {}).get(y, (None, None))[idx]
                if v is None:
                    trow[y] = "–"
                    crow[y] = None
                else:
                    g_letter, g_color, _ = rank_to_grade(v, scope)
                    trow[y] = f"#{v:,} ({g_letter})"
                    crow[y] = g_color
            text_rows[lembaga] = trow
            color_rows[lembaga] = crow
        df_text = pd.DataFrame(text_rows).T
        df_color = pd.DataFrame(color_rows).T
        df_text.columns = [str(c) for c in df_text.columns]
        df_color.columns = df_text.columns
        return df_text, df_color

    def render_year_matrix(scope):
        df_text, df_color = build_year_matrix(scope)
        if df_text.empty:
            st.info("Tidak ada lembaga yang dipilih pada filter saat ini.")
            return

        def styler(_):
            out = pd.DataFrame("", index=df_text.index, columns=df_text.columns)
            for r in df_text.index:
                for c in df_text.columns:
                    color = df_color.loc[r, c]
                    if pd.notna(color):
                        out.loc[r, c] = f"background-color:{color}; color:#FFFFFF; font-weight:700; text-align:center; border-radius:4px;"
                    else:
                        out.loc[r, c] = "background-color:#EEF0F3; color:#4B5563; text-align:center;"
            return out

        styled = df_text.style.apply(styler, axis=None)
        st.dataframe(styled, width="stretch", height=df_table_height(len(df_text)))

    with tab_nas:
        st.caption(f"Menampilkan peringkat & skor nasional untuk tahun {years_range[0]}–{years_range[-1]} sesuai filter pada sidebar.")
        render_year_matrix("Nasional")

    with tab_dunia:
        st.caption(f"Menampilkan peringkat & skor dunia untuk tahun {years_range[0]}–{years_range[-1]} sesuai filter pada sidebar.")
        render_year_matrix("Dunia")

    with tab_snapshot:
        board_rows = []
        for lembaga in lembaga_terpilih:
            y_nas, v_nas = latest_value(lembaga, "Nasional")
            y_dun, v_dun = latest_value(lembaga, "Dunia")
            g_letter, g_color, g_label = rank_to_grade(v_nas, "Nasional")
            board_rows.append({
                "Lembaga Pemeringkat": lembaga,
                "Peringkat Nasional Terbaru": f"#{int(v_nas)} ({y_nas})" if v_nas else "Belum ada data",
                "Peringkat Dunia Terbaru": f"#{int(v_dun):,} ({y_dun})" if v_dun else "Belum ada data",
                "Skor": g_letter,
                "Predikat": g_label,
            })
        df_board = pd.DataFrame(board_rows)

        def color_grade(val):
            colors = {"A+": UHO_GREEN, "A": UHO_GREEN, "B+": UHO_GOLD_DEEP,
                      "B": UHO_GOLD_DEEP, "C+": UHO_RED, "C": UHO_RED, "N/A": UHO_GREY}
            c = colors.get(val, UHO_NAVY)
            return f"background-color:{c}; color:#FFFFFF; font-weight:700; text-align:center; border-radius:4px;"

        def _style_grade_col(styler_obj, func, subset):
            """Kompatibel lintas versi pandas: Styler.applymap() dihapus total
            di pandas 3.x (diganti Styler.map()). Ini pernah membuat app
            error di Streamlit Cloud walau berjalan mulus di lokal, karena
            versi pandas yang ter-install bisa berbeda antara kedua tempat."""
            if hasattr(styler_obj, "map"):
                return styler_obj.map(func, subset=subset)
            return styler_obj.applymap(func, subset=subset)

        styled = _style_grade_col(df_board.style, color_grade, ["Skor"])
        st.dataframe(styled, width="stretch", hide_index=True, height=df_table_height(len(df_board)))

    block(f"""
    <div style="font-size:13px; color:{UHO_GREY};">
    Legenda skor:
    <b style="color:{UHO_GREEN}">A+/A</b> Sangat Unggul–Unggul ·
    <b style="color:{UHO_GOLD_DEEP}">B+/B</b> Sangat Baik–Baik ·
    <b style="color:{UHO_RED}">C+/C</b> Cukup–Perlu Peningkatan ·
    <b style="color:{UHO_GREY}">N/A / –</b> Belum masuk daftar peringkat lembaga tersebut pada tahun itu.
    </div>
    """)

# ==========================================================================
# HALAMAN 2 — DETAIL PER LEMBAGA
# ==========================================================================
elif halaman == "Detail per Lembaga":
    block('<div class="section-title">Rincian Capaian per Lembaga Pemeringkat</div>')
    st.caption("Setiap kartu menampilkan peringkat terbaru, tren, skor klasifikasi, dan grafik historis lembaga tersebut.")

    for lembaga in lembaga_terpilih:
        y_nas, v_nas = latest_value(lembaga, "Nasional")
        y_dun, v_dun = latest_value(lembaga, "Dunia")
        g_letter, g_color, g_label = rank_to_grade(v_nas, "Nasional")

        py_nas, pv_nas = previous_value(lembaga, "Nasional", y_nas) if y_nas else (None, None)
        if v_nas is None:
            trend_txt = None
        elif pv_nas is None:
            trend_txt = f"Data pertama tercatat pada {y_nas}"
        elif v_nas < pv_nas:
            trend_txt = f"Membaik dari #{int(pv_nas)} ({py_nas})"
        elif v_nas > pv_nas:
            trend_txt = f"Menurun dari #{int(pv_nas)} ({py_nas})"
        else:
            trend_txt = f"Stabil sejak {py_nas}"

        with st.container(border=True):
            c1, c2, c3, c4 = st.columns([3, 1.6, 1.6, 1])
            with c1:
                st.markdown(f"**{lembaga}**")
                st.caption(DESKRIPSI_LEMBAGA.get(lembaga, DESKRIPSI_DEFAULT))
            with c2:
                st.metric(
                    "Peringkat Nasional",
                    f"#{int(v_nas)}" if v_nas is not None else "Belum ada data",
                    trend_txt,
                    delta_color="off",  # dimatikan agar tidak muncul warna hijau/merah yang keliru arah
                )
            with c3:
                st.metric("Peringkat Dunia", f"#{int(v_dun):,}" if v_dun is not None else "Belum ada data")
            with c4:
                block(f"""
                <div style="text-align:center;">
                    <div class="grade-chip" style="background:{g_color};">{g_letter}</div>
                    <div style="font-size:11px; color:{UHO_GREY}; margin-top:4px;">{g_label}</div>
                </div>
                """)

            data = RAW.get(lembaga, {})
            nas_vals = [data.get(y, (None, None))[0] for y in years_range]
            dun_vals = [data.get(y, (None, None))[1] for y in years_range]
            ada_nas = any(v is not None for v in nas_vals)
            ada_dun = any(v is not None for v in dun_vals)

            if ada_nas or ada_dun:
                fig2 = make_subplots(specs=[[{"secondary_y": True}]])
                if ada_nas:
                    fig2.add_trace(go.Scatter(
                        x=years_range, y=nas_vals, name="Nasional", mode="lines+markers",
                        line=dict(color=UHO_GOLD_DEEP, width=2.5), marker=dict(size=8), connectgaps=True,
                    ), secondary_y=False)
                if ada_dun:
                    fig2.add_trace(go.Scatter(
                        x=years_range, y=dun_vals, name="Dunia", mode="lines+markers",
                        line=dict(color=UHO_NAVY, width=2.5, dash="dot"), marker=dict(size=8), connectgaps=True,
                    ), secondary_y=True)
                fig2.update_xaxes(dtick=1, tickformat="d")
                fig2.update_yaxes(autorange="reversed", secondary_y=False,
                                   title_text="Nasional" if ada_nas else "", title_font=dict(size=10),
                                   showgrid=True, gridcolor="#EEF1F5")
                fig2.update_yaxes(autorange="reversed", secondary_y=True,
                                   title_text="Dunia" if ada_dun else "", title_font=dict(size=10), showgrid=False)
                fig2.update_layout(
                    height=230, margin=dict(t=10, b=0, l=0, r=0),
                    plot_bgcolor="white", paper_bgcolor="white",
                    legend=dict(orientation="h", yanchor="bottom", y=1.0, font=dict(size=10, color=UHO_INK)),
                    font=dict(family="Georgia, Times New Roman, serif", size=11, color=UHO_INK),
                )
                st.plotly_chart(fig2, width="stretch", key=f"chart_{lembaga}")
            else:
                st.info(
                    f"Universitas Halu Oleo belum tercatat dalam pemeringkatan **{lembaga}** "
                    f"pada rentang tahun {years_range[0]}–{years_range[-1]} yang dipilih."
                )

# ==========================================================================
# HALAMAN 3 — METODOLOGI SKOR
# ==========================================================================
elif halaman == "Metodologi Skor":
    block('<div class="section-title">Metodologi Penyusunan Skor A–C</div>')

    block(f"""
    <div class="note-box">
    Skor huruf pada dashboard ini merupakan <b>klasifikasi internal</b> yang dibuat khusus untuk
    memudahkan pembacaan posisi relatif Universitas Halu Oleo pada setiap lembaga pemeringkat.
    Skor ini <b>bukan nilai resmi</b> yang diterbitkan oleh THE, QS, Webometrics, atau lembaga
    pemeringkat lainnya, melainkan hasil pengelompokan (tiering) atas peringkat numerik yang
    mereka terbitkan, dengan prinsip dasar: semakin kecil angka peringkat, semakin baik posisi
    universitas.
    </div>
    """)

    colA, colB = st.columns(2)
    with colA:
        st.markdown("**Tier Peringkat Nasional**")
        df_nas_tier = pd.DataFrame([
            {"Rentang Peringkat Nasional": "1 – 10", "Skor": "A+", "Predikat": "Sangat Unggul"},
            {"Rentang Peringkat Nasional": "11 – 25", "Skor": "A", "Predikat": "Unggul"},
            {"Rentang Peringkat Nasional": "26 – 50", "Skor": "B+", "Predikat": "Sangat Baik"},
            {"Rentang Peringkat Nasional": "51 – 100", "Skor": "B", "Predikat": "Baik"},
            {"Rentang Peringkat Nasional": "101 – 250", "Skor": "C+", "Predikat": "Cukup"},
            {"Rentang Peringkat Nasional": "> 250", "Skor": "C", "Predikat": "Perlu Peningkatan"},
            {"Rentang Peringkat Nasional": "Tidak masuk daftar", "Skor": "N/A", "Predikat": "Belum Terpetakan"},
        ])
        st.dataframe(df_nas_tier, hide_index=True, width="stretch")
    with colB:
        st.markdown("**Tier Peringkat Dunia**")
        df_dunia_tier = pd.DataFrame([
            {"Rentang Peringkat Dunia": "1 – 500", "Skor": "A+", "Predikat": "Sangat Unggul"},
            {"Rentang Peringkat Dunia": "501 – 1.000", "Skor": "A", "Predikat": "Unggul"},
            {"Rentang Peringkat Dunia": "1.001 – 2.000", "Skor": "B+", "Predikat": "Sangat Baik"},
            {"Rentang Peringkat Dunia": "2.001 – 4.000", "Skor": "B", "Predikat": "Baik"},
            {"Rentang Peringkat Dunia": "4.001 – 8.000", "Skor": "C+", "Predikat": "Cukup"},
            {"Rentang Peringkat Dunia": "> 8.000", "Skor": "C", "Predikat": "Perlu Peningkatan"},
            {"Rentang Peringkat Dunia": "Tidak masuk daftar", "Skor": "N/A", "Predikat": "Belum Terpetakan"},
        ])
        st.dataframe(df_dunia_tier, hide_index=True, width="stretch")

    block("""
    <div class="note-box">
    <b>Dasar penentuan skor pada dashboard:</b>
    <ol style="margin-top:6px; margin-bottom:0;">
        <li>Skor utama tiap lembaga dihitung dari peringkat nasional terbaru yang tersedia,
        karena posisi nasional lebih mudah dibandingkan antar-lembaga dan antar-tahun.</li>
        <li>Peringkat dunia ditampilkan sebagai data pendukung (bukan penentu skor huruf),
        karena skala jumlah universitas yang dinilai berbeda-beda antar lembaga.</li>
        <li>Jika suatu lembaga belum pernah memetakan UHO pada tahun tertentu, status ditandai
        "N/A - Belum Terpetakan", bukan otomatis dinilai buruk.</li>
        <li>Tren dihitung dengan membandingkan capaian tahun terbaru terhadap capaian tahun
        sebelumnya yang tersedia datanya, karena beberapa lembaga tidak memeringkat UHO
        setiap tahun.</li>
    </ol>
    </div>
    """)

    block("""
    <div class="note-box">
    <b>Catatan pengembangan lanjutan:</b> apabila diperlukan skema skor berbasis nilai total
    0–100 dengan ambang batas tertentu per kriteria (mis. ambang A/B/C/D), skema tersebut dapat
    ditambahkan sebagai lapisan penilaian tambahan setelah ambang batas dan bobot masing-masing
    kriteria ditetapkan secara resmi oleh tim penyusun.
    </div>
    """)

# ==========================================================================
# HALAMAN 4 — PROFIL PERGURUAN TINGGI
# ==========================================================================
elif halaman == "Profil Perguruan Tinggi":
    block('<div class="section-title">Identitas Perguruan Tinggi</div>')

    col1, col2 = st.columns([1, 2.2])
    with col1:
        with st.container(border=True):
            block(f'<div style="text-align:center;">{logo_markup(120)}</div>')
            st.markdown(
                f"<p style='text-align:center; font-weight:700; color:{UHO_NAVY}; margin:10px 0 0 0;'>"
                f"Universitas Halu Oleo</p>",
                unsafe_allow_html=True,
            )
            st.markdown(
                f"<p style='text-align:center; color:{UHO_GOLD_DEEP}; font-weight:700; letter-spacing:1px;'>UHO</p>",
                unsafe_allow_html=True,
            )

    with col2:
        with st.container(border=True):
            # Ditulis dengan komponen native Streamlit (bukan tabel HTML mentah)
            # agar dijamin selalu ter-render sebagai tampilan, bukan teks kode.
            for label, value in IDENTITAS_UHO:
                lc, vc = st.columns([1, 2.3])
                with lc:
                    st.markdown(f"**{label}**")
                with vc:
                    st.write(value)
                st.markdown(
                    f"<hr style='margin:4px 0; border:none; border-top:1px solid {UHO_BORDER};'>",
                    unsafe_allow_html=True,
                )

    block('<div class="section-title">Lembaga Pemeringkat yang Dipantau</div>')
    cols = st.columns(2)
    for i, lembaga in enumerate(LEMBAGA_LIST):
        with cols[i % 2]:
            with st.container(border=True):
                st.markdown(f"**{lembaga}**")
                st.caption(DESKRIPSI_LEMBAGA.get(lembaga, DESKRIPSI_DEFAULT))

# ==========================================================================
# FOOTER
# ==========================================================================
block(f"""
<div class="footer-note">
    Dashboard Pemeringkatan Universitas Halu Oleo · Data dikelola melalui <i>data_ranking.csv</i> (cakupan {YEARS[0]}–{YEARS[-1]})<br>
    Skor A–C adalah klasifikasi internal untuk kebutuhan monitoring, bukan nilai resmi lembaga pemeringkat.
</div>
""")
