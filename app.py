# -*- coding: utf-8 -*-
"""
LMS SMP Negeri 1 Cijambe
Sistem pembelajaran online terpadu dan portal akademik sekolah.
Peran: guru, siswa, kepala sekolah, orang tua.
Stack: Streamlit, SQLite, Google Gemini.
"""
import hashlib
import hmac
import html as _html
import inspect
import io
import json
import os
import random
import re
import sqlite3
from contextlib import closing
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from PIL import Image

try:
    import google.generativeai as genai
except Exception:  # paket belum terpasang
    genai = None
try:
    import streamlit.components.v1 as components
    from streamlit_geolocation import streamlit_geolocation
except Exception:  # paket lokasi belum terpasang
    streamlit_geolocation = None
    import streamlit.components.v1 as components

st.set_page_config(page_title="LMS SMP Negeri 1 Cijambe", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")
try:  # cadangan bila .streamlit/config.toml belum dipasang
    for _k, _v in {"base": "light", "primaryColor": "#1B4F9C", "backgroundColor": "#F3F5F9",
                   "secondaryBackgroundColor": "#FFFFFF", "textColor": "#1B2A41"}.items():
        st._config.set_option(f"theme.{_k}", _v)
except Exception:
    pass

# ══════════════════════════ KONFIGURASI ══════════════════════════
NAMA_SEKOLAH = "SMP Negeri 1 Cijambe"
DB_PATH = os.environ.get("SPOT_DB", "spot_v2.db")  # nama tidak diubah agar data lama tetap terbaca
WIB = timezone(timedelta(hours=7))
BATAS_MASUK = "07:15"  # presensi guru setelah jam ini dicatat Terlambat
KELAS = ["VIII-A", "VIII-B"]
MP = [("Matematika", "guru"), ("IPA", "guru3"), ("Bahasa Indonesia", "guru2"), ("Bahasa Inggris", "guru4")]
SLOT = ["07:30-09:00", "09:15-10:45", "11:00-12:30", "13:00-14:30"]
WALI = {"VIII-A": "Ibu Sari Wulandari, S.Pd.", "VIII-B": "Bpk. Hendra Gunawan, S.Si."}
IKON = {"Matematika": "📐", "IPA": "🔬", "Bahasa Indonesia": "📖", "Bahasa Inggris": "🌐"}
ROLE = {"guru": "Pengajar", "siswa": "Peserta didik", "kepsek": "Kepala sekolah", "ortu": "Orang tua / wali"}
HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
BLN = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
STATUS_HADIR = ["Hadir", "Sakit", "Izin", "Alpa"]
KKM = 75  # batas ketuntasan nilai akhir
RADIUS_DEFAULT = 100  # meter, dipakai bila kepala sekolah belum mengatur radius
MAX_AKURASI = 200  # meter, akurasi GPS terburuk yang diterima
DOKUMEN_ADMIN = ["Kalender Pendidikan & Hari Efektif", "Capaian Pembelajaran (CP) & Alur Tujuan Pembelajaran (ATP)",
                 "Program Tahunan (Prota)", "Program Semester (Promes)", "Modul Ajar", "KKTP / Kriteria Ketercapaian",
                 "Jadwal Pelajaran", "Bahan Ajar / LKPD", "Soal & Kisi-kisi Asesmen", "Daftar Hadir Siswa", "Daftar Nilai",
                 "Analisis Hasil Penilaian", "Program Remedial & Pengayaan", "Jurnal Mengajar", "Laporan Refleksi / PKG"]
GENERATOR = {
    "Program Tahunan (Prota)": "Tabel: semester, bab/lingkup materi, tujuan pembelajaran, alokasi jam pelajaran, dan keterangan. Total JP sesuai minggu efektif.",
    "Program Semester (Promes)": "Tabel: no, materi, alokasi JP, dan kolom minggu ke-1 sampai ke-20 per bulan (tandai minggu pelaksanaan), plus minggu tidak efektif dan asesmen.",
    "Capaian Pembelajaran (CP) & Alur Tujuan Pembelajaran (ATP)": "Uraikan CP fase, elemen, tujuan pembelajaran berurutan, alokasi waktu, dan dimensi Profil Lulusan/Pelajar Pancasila.",
    "KKTP / Kriteria Ketercapaian": "Tabel: tujuan pembelajaran, kriteria ketercapaian (rubrik 4 level), indikator, teknik asesmen.",
    "Program Remedial & Pengayaan": "Berisi identifikasi kesulitan, strategi remedial, kegiatan pengayaan, jadwal, instrumen, dan tabel tindak lanjut siswa.",
    "Soal & Kisi-kisi Asesmen": "Berisi kisi-kisi (tujuan, indikator, level kognitif, bentuk, nomor), 10 soal pilihan ganda, 3 soal uraian, kunci, dan pedoman skor.",
    "Bahan Ajar / LKPD": "Berisi ringkasan materi, contoh, LKPD bertahap, latihan, dan refleksi, dengan bahasa sesuai jenjang.",
}
BD = {"Hadir": "ok", "Tepat Waktu": "ok", "Sudah Dinilai": "ok", "Baik": "ok", "A": "ok", "B": "info", "C": "warn",
      "D": "bad", "Terlambat": "warn", "Sakit": "warn", "Menunggu Penilaian": "info", "Izin": "info",
      "Belum Dikumpulkan": "info", "Cukup": "info", "Alpa": "bad", "Melewati Deadline": "bad",
      "Perlu Perhatian": "bad", "Belum Presensi": "gray", "Lengkap": "ok", "Draf": "warn", "Belum Ada": "bad",
      "Tuntas": "ok", "Remedial": "bad", "Cocok": "ok", "Perlu Verifikasi": "warn", "Tidak Cocok": "bad", "Terdaftar": "ok"}


def secret(nama, default=""):
    """Baca dari Streamlit Secrets, lalu environment variable."""
    try:
        if nama in st.secrets:
            return str(st.secrets[nama])
    except Exception:
        pass
    return os.environ.get(nama, default)


API_KEY = secret("GEMINI_API_KEY")
SHOW_DEMO = secret("SHOW_DEMO", "1") != "0"  # isi SHOW_DEMO = "0" di Secrets untuk menyembunyikan akun demo
MODEL_KANDIDAT = [m.strip() for m in secret("GEMINI_MODEL").split(",") if m.strip()] + [
    "gemini-3.8-flash", "gemini-3.5-flash", "gemini-2.5-flash", "gemini-1.5-flash"]
_PAKAI_UCW = "use_container_width" in inspect.signature(st.data_editor).parameters


# ══════════════════════════ UTILITAS ══════════════════════════
def now():
    return datetime.now(WIB)


def today():
    return now().date()


def fmt(d, hari=True):
    if d is None or d != d:
        return "–"
    if isinstance(d, str):
        d = date.fromisoformat(d[:10])
    return (HARI[d.weekday()][:3] + ", " if hari else "") + f"{d.day} {BLN[d.month - 1]} {d.year}"


def hari_kerja(n, mulai=None):
    """n hari kerja terakhir (Senin-Jumat), terbaru lebih dulu."""
    d, out = mulai or today(), []
    while len(out) < n:
        if d.weekday() < 5:
            out.append(d)
        d -= timedelta(days=1)
    return out


def sapa():
    j = now().hour
    return "Selamat pagi" if j < 11 else "Selamat siang" if j < 15 else "Selamat sore" if j < 18 else "Selamat malam"


def pct(a, b):
    a = 0 if a is None or a != a else a
    return round(100 * float(a) / float(b)) if b else 0


def huruf(n):
    return "A" if n >= 85 else "B" if n >= 75 else "C" if n >= 65 else "D"


def initials(nama):
    p = [w for w in re.split(r"[ ,]+", nama) if w and "." not in w]
    return "".join(w[0] for w in p[:2]).upper() or "U"


def hpw(username, pw):
    # salt jangan diubah, agar password lama tetap valid
    return hashlib.pbkdf2_hmac("sha256", pw.encode(), f"spot:{username}".encode(), 60000).hex()


# ══════════════════════════ GAYA (CSS) ══════════════════════════
CSS = """
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');
:root{--navy:#0B2A55;--blue:#1B4F9C;--gold:#F2B705;--bg:#F3F5F9;--line:#D5DEEB;--tx:#1B2A41;--mu:#5B6B82}
html,body,.stApp,.stApp p,.stApp label,.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp button,.stApp input,.stApp textarea,.stApp th,.stApp td,.stApp li{font-family:'IBM Plex Sans','Segoe UI',system-ui,sans-serif}
.stApp{background:var(--bg);color:var(--tx)}
#MainMenu,footer,[data-testid="stToolbar"],[data-testid="stDecoration"]{display:none!important}
header[data-testid="stHeader"]{background:transparent}
.block-container{padding:3rem 2rem 3rem;max-width:1280px}
.stApp h1,.stApp h2,.stApp h3,.stApp h4{color:var(--navy);font-weight:600}
/* sidebar: menu ala SIAK */
section[data-testid="stSidebar"]{background:var(--navy);border-right:0}
section[data-testid="stSidebar"] *{color:#DCE6F5}
section[data-testid="stSidebar"] hr{border-color:rgba(255,255,255,.16)}
.sb-brand{display:flex;gap:10px;align-items:center;padding:4px 2px 14px;border-bottom:1px solid rgba(255,255,255,.16);margin-bottom:12px}
.sb-brand .logo{width:40px;height:40px;background:var(--gold);border-radius:4px;display:flex;align-items:center;justify-content:center;font-size:1.3rem}
.sb-brand b{display:block;color:#fff!important;font-size:1.08rem;line-height:1.2}
.sb-brand span{font-size:.78rem;opacity:.8}
.sb-user{display:flex;gap:10px;align-items:center;background:rgba(255,255,255,.07);border-radius:4px;padding:10px;margin-bottom:14px}
.sb-user b{display:block;color:#fff!important;font-size:.88rem;line-height:1.25}
.sb-user span{font-size:.76rem;opacity:.8}
.av{width:40px;height:40px;border-radius:50%;background:var(--gold);color:var(--navy)!important;font-weight:700;display:flex;align-items:center;justify-content:center;flex:none}
.av.big{width:64px;height:64px;font-size:1.4rem}
section[data-testid="stSidebar"] .stButton>button{width:100%;background:transparent;border:0;border-left:3px solid transparent;border-radius:0;justify-content:flex-start;padding:.6rem .8rem;font-weight:500;color:#DCE6F5}
section[data-testid="stSidebar"] .stButton>button *{text-align:left;justify-content:flex-start;color:inherit}
section[data-testid="stSidebar"] .stButton>button:hover{background:rgba(255,255,255,.08)}
section[data-testid="stSidebar"] .stButton>button[kind="primary"],section[data-testid="stSidebar"] .stButton>button[data-testid="stBaseButton-primary"]{background:rgba(242,183,5,.16);border-left-color:var(--gold);font-weight:600;color:#fff}
.st-key-logout button{border:1px solid rgba(255,255,255,.35)!important;border-radius:4px!important}
.st-key-logout button *{text-align:center!important;justify-content:center!important}
/* judul halaman */
.hero{background:var(--navy);border-bottom:3px solid var(--gold);border-radius:4px;padding:16px 22px;margin-bottom:18px;display:flex;justify-content:space-between;align-items:flex-end;gap:12px;flex-wrap:wrap}
.hero .crumb{font-size:.8rem;color:var(--gold)}
.hero h1{color:#fff!important;font-size:1.55rem;font-weight:600;margin:2px 0 4px;padding:0!important;line-height:1.25}
.hero p{margin:0;color:#C9D7EE;font-size:.9rem}
.hero .rt{text-align:right;color:#C9D7EE;font-size:.85rem}
.hero .rt b{display:block;color:#fff;font-weight:600}
/* kotak ringkasan */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:12px;margin-bottom:18px}
.kpi{background:var(--blue);color:#fff;border-radius:4px;padding:13px 16px}
.kpi span{display:block;font-size:.85rem;opacity:.92}
.kpi b{display:block;font-size:1.9rem;font-weight:700;line-height:1.25;font-variant-numeric:tabular-nums}
.kpi small{font-size:.78rem;opacity:.88}
.kpi.g{background:#15803D}.kpi.o{background:#B45309}.kpi.r{background:#B91C1C}.kpi.t{background:#0F766E}
/* panel dan tabel */
.pn{background:#fff;border:1px solid var(--line);border-radius:4px;margin-bottom:16px}
.pn-h{background:#E8EEF7;border-bottom:1px solid var(--line);padding:9px 14px;font-weight:600;color:var(--navy)}
.pn-b{padding:14px}.pn-b.flush{padding:0}
.sec{background:#E8EEF7;border:1px solid var(--line);border-left:4px solid var(--gold);border-radius:4px;padding:8px 14px;font-weight:600;color:var(--navy);margin:14px 0 10px}
.tw{overflow-x:auto;border:1px solid var(--line);border-radius:4px;background:#fff}
.pn .tw{border:0;border-radius:0}
.tb{width:100%;border-collapse:collapse;font-size:.87rem}
.tb th{background:var(--navy);color:#fff;text-align:left;padding:9px 12px;font-weight:600;white-space:nowrap}
.tb td{padding:8px 12px;border-bottom:1px solid #E6ECF5;color:var(--tx);font-variant-numeric:tabular-nums}
.tb tbody tr:nth-child(even){background:#F6F9FD}
.bd{display:inline-block;padding:2px 9px;border-radius:3px;font-size:.76rem;font-weight:600;white-space:nowrap}
.bd.ok{background:#DCFCE7;color:#166534}.bd.warn{background:#FEF3C7;color:#92400E}.bd.info{background:#DBEAFE;color:#1E40AF}.bd.bad{background:#FEE2E2;color:#991B1B}.bd.gray{background:#E2E8F0;color:#334155}
.ann{border-left:3px solid var(--gold);padding:2px 0 2px 12px;margin-bottom:12px}
.ann small{color:var(--mu)}.ann b{display:block;color:var(--navy)}.ann p{margin:2px 0 0;font-size:.86rem;color:#475569}
/* kartu mata pelajaran */
.courses{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:14px;margin-bottom:18px}
.course{background:#fff;border:1px solid var(--line);border-radius:4px;overflow:hidden}
.cband{height:62px;display:flex;align-items:center;padding:0 14px;font-size:1.6rem;background:var(--blue)}
.c1{background:#0F766E}.c2{background:#B45309}.c3{background:#6D28D9}
.cbody{padding:12px 14px}.cbody b{display:block;color:var(--navy)}.cbody small{color:var(--mu);display:block}
.bar{height:6px;background:#E3EAF4;border-radius:3px;margin:10px 0 6px;overflow:hidden}.bar i{display:block;height:100%;background:var(--blue)}
.note{border:1px solid;border-radius:4px;padding:9px 13px;margin:8px 0;font-size:.88rem}
.note.ok{background:#F0FDF4;border-color:#BBF7D0;color:#166534}.note.warn{background:#FFFBEB;border-color:#FDE68A;color:#92400E}.note.bad{background:#FEF2F2;border-color:#FECACA;color:#991B1B}.note.info{background:#EFF6FF;border-color:#BFDBFE;color:#1E40AF}
.empty{padding:22px;text-align:center;color:var(--mu);background:#fff;border:1px dashed var(--line);border-radius:4px}
.profile{display:flex;gap:16px;align-items:center;background:#fff;border:1px solid var(--line);border-left:4px solid var(--gold);border-radius:4px;padding:14px 18px;margin-bottom:16px}
.profile b{display:block;font-size:1.15rem;color:var(--navy)}.profile span{display:block;color:var(--mu);font-size:.88rem}
.tugas-h{display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap}.tugas-h b{color:var(--navy);font-size:1.02rem}
.tugas-m{color:var(--mu);font-size:.85rem;margin:2px 0 6px}
/* halaman masuk */
.lg{background:var(--navy);border-bottom:4px solid var(--gold);border-radius:4px;padding:32px 28px;color:#fff;min-height:430px}
.lg .logo{width:52px;height:52px;background:var(--gold);border-radius:4px;display:flex;align-items:center;justify-content:center;font-size:1.6rem}
.lg h2{color:#fff!important;font-size:1.8rem;margin:16px 0 6px;padding:0!important}
.lg p{color:#C9D7EE;margin:0 0 18px}
.lg ul{margin:0;padding:0}.lg li{list-style:none;padding:9px 12px;border-left:3px solid var(--gold);margin-bottom:8px;background:rgba(255,255,255,.07);color:#E4ECF9}
/* komponen bawaan Streamlit */
[data-testid="stForm"]{background:#fff;border:1px solid var(--line);border-radius:4px;padding:16px}
[data-testid="stFormSubmitButton"]>button{width:100%}
div[class*="st-key-card"]{background:#fff;border:1px solid var(--line);border-radius:4px;padding:14px 16px;margin-bottom:6px}
.stButton>button,.stDownloadButton>button,[data-testid="stFormSubmitButton"]>button{border-radius:4px;font-weight:600}
.stTabs [data-baseweb="tab-list"]{gap:2px;border-bottom:1px solid var(--line)}
.stTabs [data-baseweb="tab"]{height:44px;padding:0 16px;font-weight:600;color:var(--mu)}
.stTabs [aria-selected="true"]{color:var(--navy)}
/* paksa tema terang agar teks tidak hilang saat HP memakai mode gelap */
:root,.stApp{color-scheme:light}
.stApp [data-testid="stMarkdownContainer"] p,.stApp [data-testid="stMarkdownContainer"] li,.stApp [data-testid="stMarkdownContainer"] td{color:var(--tx)}
.stApp .hero p,.stApp .lg p{color:#C9D7EE}
.stApp .lg li{color:#E4ECF9}
.stApp .ann p{color:#475569}
.stApp .ann small,.stApp .tugas-m,.stApp .cbody small,.stApp .profile span{color:var(--mu)}
.stApp .ann b,.stApp .tugas-h b,.stApp .cbody b,.stApp .profile b{color:var(--navy)}
.stApp label,.stApp [data-testid="stWidgetLabel"] p,.stApp [data-testid="stCaptionContainer"] *{color:var(--tx)}
.stApp input,.stApp textarea,.stApp [data-baseweb="select"] *{color:var(--tx)!important}
.stApp input,.stApp textarea,.stApp [data-baseweb="select"]>div,.stApp [data-baseweb="input"],.stApp [data-baseweb="textarea"]{background:#fff!important}
[data-baseweb="popover"] *{color:var(--tx)}
.block-container .stButton>button[kind="secondary"],.block-container .stDownloadButton>button{background:#fff;color:var(--navy);border:1px solid var(--line)}
.block-container .stButton>button[kind="primary"],.block-container .stDownloadButton>button[kind="primary"],.block-container [data-testid="stFormSubmitButton"]>button{background:var(--blue);color:#fff;border:1px solid var(--blue)}
.block-container .stButton>button[kind="primary"] *,.block-container [data-testid="stFormSubmitButton"]>button *{color:#fff}
[data-testid="stExpander"] summary,[data-testid="stExpander"] summary *{color:var(--navy)}
/* menu atas: tampil di HP, disembunyikan di layar lebar karena sudah ada sidebar */
.st-key-topnav{background:#fff;border:1px solid var(--line);border-left:4px solid var(--gold);border-radius:4px;padding:8px 12px;margin-bottom:14px}
.st-key-topnav [role="radiogroup"]{gap:6px 14px;flex-wrap:wrap}
/* kotak isian, unggah berkas, dan tombol: paksa terang */
[data-baseweb="input"],[data-baseweb="base-input"],[data-baseweb="textarea"],[data-baseweb="select"]>div,[data-testid="stNumberInput"] div[data-baseweb="input"],[data-testid="stDateInput"] div[data-baseweb="input"]{background:#fff!important;border-color:var(--line)!important}
.stApp input,.stApp textarea{background:#fff!important;color:var(--tx)!important;-webkit-text-fill-color:var(--tx)!important;caret-color:var(--tx)}
.stApp input::placeholder,.stApp textarea::placeholder{color:#8493A8!important;-webkit-text-fill-color:#8493A8!important}
[data-testid="stNumberInput"] button,[data-testid="stNumberInputStepUp"],[data-testid="stNumberInputStepDown"]{background:#EEF3FA!important;color:var(--navy)!important}
[data-testid="stNumberInput"] button *{color:var(--navy)!important;fill:var(--navy)!important}
[data-testid="stFileUploaderDropzone"]{background:#F6F9FD!important;border:1px dashed #9DB0CB!important}
[data-testid="stFileUploaderDropzone"] *{color:var(--tx)!important}
[data-testid="stFileUploaderDropzone"] button{background:#fff!important;border:1px solid var(--blue)!important}
[data-testid="stFileUploaderDropzone"] button *{color:var(--blue)!important}
[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"],[data-testid="stFileUploader"] section+div{background:#fff!important}
[data-testid="stFileUploader"] small,[data-testid="stFileUploader"] span{color:var(--tx)!important}
.block-container [data-testid="stBaseButton-secondary"],.block-container [data-testid="stBaseButton-secondaryFormSubmit"],.block-container .stDownloadButton>button{background:#fff!important;color:var(--navy)!important;border:1px solid #9DB0CB!important}
.block-container [data-testid="stBaseButton-secondary"] *,.block-container [data-testid="stBaseButton-secondaryFormSubmit"] *,.block-container .stDownloadButton>button *{color:var(--navy)!important}
.block-container [data-testid="stBaseButton-primary"],.block-container [data-testid="stBaseButton-primaryFormSubmit"],.block-container .stDownloadButton>button[kind="primary"],.block-container [data-testid="stFormSubmitButton"]>button[kind="primary"]{background:var(--blue)!important;color:#fff!important;border:1px solid var(--blue)!important}
.block-container [data-testid="stBaseButton-primary"] *,.block-container [data-testid="stBaseButton-primaryFormSubmit"] *,.block-container .stDownloadButton>button[kind="primary"] *,.block-container [data-testid="stFormSubmitButton"]>button[kind="primary"] *{color:#fff!important}
.block-container button:hover{filter:brightness(.96)}
[data-baseweb="calendar"],[data-baseweb="calendar"] *,[data-baseweb="popover"] ul,[data-baseweb="popover"] li,[role="listbox"],[role="listbox"] li{background-color:#fff!important;color:var(--tx)!important}
[role="option"][aria-selected="true"],[data-baseweb="popover"] li:hover{background-color:#E8EEF7!important}
[data-baseweb="radio"] div,[data-baseweb="checkbox"] div,.stRadio label,.stRadio p{color:var(--tx)!important}
[data-testid="stAlert"] *{color:inherit}
[data-testid="stChatInput"],[data-testid="stChatInput"] textarea{background:#fff!important;color:var(--tx)!important}
[data-testid="stChatMessage"]{background:#fff;border:1px solid var(--line);border-radius:4px}
[data-testid="stChatMessage"] *{color:var(--tx)}
[data-testid="stDataFrame"],[data-testid="stDataEditor"]{color-scheme:light;background:#fff}
/* kotak isian punya garis tepi yang jelas (username, password, dan lainnya) */
[data-baseweb="input"]{border:1px solid #9DB0CB!important;border-radius:4px!important;background:#fff!important}
[data-baseweb="base-input"]{border:0!important;background:#fff!important}
[data-baseweb="select"]>div{border:1px solid #9DB0CB!important}
[data-testid="stTextInput"] input{min-height:42px;font-size:1rem}
[data-testid="stTextInput"] button{background:transparent!important}
[data-testid="stTextInput"] button *{color:var(--mu)!important;fill:var(--mu)!important}
[data-testid="stForm"] [data-testid="stWidgetLabel"] p{font-weight:600;color:var(--navy)!important}
/* tombol presensi besar dan tombol nonaktif tetap terbaca */
.st-key-btn_masuk button,.st-key-btn_pulang button{width:100%;min-height:56px;font-size:1.05rem}
.block-container .st-key-btn_masuk button:not(:disabled){background:#15803D!important;border-color:#15803D!important}
.block-container .st-key-btn_pulang button:not(:disabled){background:#B45309!important;border-color:#B45309!important}
.block-container button:disabled,.block-container button[disabled]{background:#E2E8F0!important;border:1px solid #CBD5E1!important;opacity:1!important}
.block-container button:disabled *,.block-container button[disabled] *{color:#64748B!important}
@media(min-width:0px){.st-key-topnav{display:block}}
section[data-testid="stSidebar"] .stButton>button,section[data-testid="stSidebar"] .stButton>button *{color:#DCE6F5!important}
section[data-testid="stSidebar"] .stButton>button[kind="primary"],section[data-testid="stSidebar"] .stButton>button[kind="primary"] *{color:#fff!important}
section[data-testid="stSidebar"] .sb-brand span,section[data-testid="stSidebar"] .sb-user span{color:#DCE6F5!important}
[data-testid="stExpandSidebarButton"],[data-testid="stSidebarCollapsedControl"],[data-testid="collapsedControl"]{display:flex!important;visibility:visible!important;background:var(--navy);border-radius:4px}
[data-testid="stExpandSidebarButton"] *,[data-testid="stSidebarCollapsedControl"] *,[data-testid="collapsedControl"] *{color:#fff!important}
@media(max-width:640px){.block-container{padding:3rem 1rem 2rem}.hero .rt{text-align:left}}
"""


def pasang_css():
    st.markdown("<style>" + " ".join(CSS.split()) + "</style>", unsafe_allow_html=True)


# ══════════════════════════ DATABASE ══════════════════════════
SCHEMA = """
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, username TEXT UNIQUE, pw TEXT, role TEXT, nama TEXT, siswa_id INTEGER);
CREATE TABLE IF NOT EXISTS siswa(id INTEGER PRIMARY KEY, nis TEXT, nama TEXT, kelas TEXT);
CREATE TABLE IF NOT EXISTS presensi_guru(id INTEGER PRIMARY KEY, guru_id INTEGER, tanggal TEXT, jam_masuk TEXT, jam_pulang TEXT, status TEXT, keterangan TEXT, UNIQUE(guru_id, tanggal));
CREATE TABLE IF NOT EXISTS sesi(id INTEGER PRIMARY KEY, tanggal TEXT, kelas TEXT, mapel TEXT, materi TEXT, catatan TEXT, guru_id INTEGER);
CREATE TABLE IF NOT EXISTS presensi_siswa(id INTEGER PRIMARY KEY, sesi_id INTEGER, siswa_id INTEGER, status TEXT);
CREATE TABLE IF NOT EXISTS tugas(id INTEGER PRIMARY KEY, judul TEXT, deskripsi TEXT, mapel TEXT, kelas TEXT, deadline TEXT, guru_id INTEGER);
CREATE TABLE IF NOT EXISTS pengumpulan(id INTEGER PRIMARY KEY, tugas_id INTEGER, siswa_id INTEGER, waktu TEXT, nama_file TEXT, berkas BLOB, nilai INTEGER, umpan_balik TEXT, UNIQUE(tugas_id, siswa_id));
CREATE TABLE IF NOT EXISTS nilai(siswa_id INTEGER, mapel TEXT, harian INTEGER, ujian INTEGER, PRIMARY KEY(siswa_id, mapel));
CREATE TABLE IF NOT EXISTS pengumuman(id INTEGER PRIMARY KEY, tanggal TEXT, judul TEXT, isi TEXT);
CREATE TABLE IF NOT EXISTS admin_guru(id INTEGER PRIMARY KEY, guru_id INTEGER, jenis TEXT, nama_file TEXT, berkas BLOB, status TEXT, catatan TEXT, diperbarui TEXT, UNIQUE(guru_id, jenis));
CREATE TABLE IF NOT EXISTS pengaturan(kunci TEXT PRIMARY KEY, nilai TEXT);
CREATE TABLE IF NOT EXISTS wajah_guru(guru_id INTEGER PRIMARY KEY, foto BLOB, setuju_pada TEXT, diperbarui TEXT);
CREATE TABLE IF NOT EXISTS bukti_presensi(id INTEGER PRIMARY KEY, guru_id INTEGER, tanggal TEXT, jenis TEXT, waktu TEXT, lat REAL, lon REAL, akurasi REAL, jarak INTEGER, foto BLOB, wajah TEXT, catatan TEXT, UNIQUE(guru_id, tanggal, jenis));
"""


def db():
    return sqlite3.connect(DB_PATH, timeout=15)


def qdf(sql, p=()):
    with closing(db()) as c:
        return pd.read_sql_query(sql, c, params=p)


def run(sql, p=(), many=False):
    with closing(db()) as c:
        cur = c.executemany(sql, p) if many else c.execute(sql, p)
        c.commit()
        return cur.lastrowid


def one(sql, p=()):
    with closing(db()) as c:
        return c.execute(sql, p).fetchone()


def pastikan_skema():
    """Jalankan setiap sesi: tabel baru dibuat walau init_db sudah tersimpan di cache server."""
    with closing(db()) as c:
        c.executescript(SCHEMA)


@st.cache_resource
def init_db():
    """Buat tabel dan isi data contoh saat database masih kosong."""
    with closing(db()) as c:
        c.executescript(SCHEMA)
        if c.execute("SELECT COUNT(*) FROM users").fetchone()[0]:
            return True
        rnd, t0 = random.Random(2026), today()
        c.executemany("INSERT INTO siswa(nis,nama,kelas) VALUES (?,?,?)", [
            ("2401", "Budi Santoso", "VIII-A"), ("2402", "Siti Rahma", "VIII-A"), ("2403", "Andi Wijaya", "VIII-A"),
            ("2404", "Dewi Lestari", "VIII-A"), ("2405", "Rizky Pratama", "VIII-A"), ("2406", "Nadia Putri", "VIII-A"),
            ("2407", "Fajar Ramadhan", "VIII-A"), ("2408", "Aulia Zahra", "VIII-A"), ("2411", "Kevin Mahendra", "VIII-B"),
            ("2412", "Salsabila Nur", "VIII-B"), ("2413", "Dimas Prasetyo", "VIII-B"), ("2414", "Putri Anjani", "VIII-B"),
            ("2415", "Reza Fauzi", "VIII-B"), ("2416", "Intan Permata", "VIII-B")])
        nama = dict(c.execute("SELECT id, nama FROM siswa"))
        budi = [i for i, n in nama.items() if n == "Budi Santoso"][0]
        akun = [("guru", "guru123", "guru", "Bpk. Ramdani, M.Pd.", None),
                ("guru2", "guru123", "guru", "Ibu Sari Wulandari, S.Pd.", None),
                ("guru3", "guru123", "guru", "Bpk. Hendra Gunawan, S.Si.", None),
                ("guru4", "guru123", "guru", "Ibu Rina Marlina, S.Pd.", None),
                ("siswa", "siswa123", "siswa", "Budi Santoso", budi),
                ("ortu", "ortu123", "ortu", "Orang Tua Budi Santoso", budi),
                ("kepsek", "kepsek123", "kepsek", "Dra. Hj. Nurhayati", None)]
        c.executemany("INSERT INTO users(username,pw,role,nama,siswa_id) VALUES (?,?,?,?,?)",
                      [(u, hpw(u, p), r, n, s) for u, p, r, n, s in akun])
        gid = {u: i for i, u in c.execute("SELECT id, username FROM users WHERE role='guru'")}

        # presensi guru 20 hari kerja terakhir
        for u, g in gid.items():
            for d in hari_kerja(20):
                if d == t0 and u in ("guru", "guru4"):
                    continue  # supaya demo bisa mencoba presensi masuk
                r = rnd.random()
                if r < .90:
                    sts, jm = "Hadir", f"06:{rnd.randint(30, 59)}"
                elif r < .95:
                    sts, jm = "Terlambat", f"07:{rnd.randint(16, 45)}"
                elif r < .975:
                    sts, jm = "Sakit", None
                else:
                    sts, jm = "Izin", None
                jp = f"{rnd.randint(14, 15)}:{rnd.randint(0, 59):02d}" if jm and d != t0 else None
                ket = {"Sakit": "Surat dokter", "Izin": "Keperluan keluarga"}.get(sts, "")
                c.execute("INSERT INTO presensi_guru(guru_id,tanggal,jam_masuk,jam_pulang,status,keterangan) VALUES (?,?,?,?,?,?)",
                          (g, d.isoformat(), jm, jp, sts, ket))

        # jurnal mengajar dan presensi siswa
        materi = {"Matematika": ["Teorema Pythagoras", "Tripel Pythagoras", "Bilangan Berpangkat", "Pola Bilangan"],
                  "IPA": ["Fotosintesis", "Sistem Pencernaan", "Getaran dan Gelombang"],
                  "Bahasa Indonesia": ["Teks Eksposisi", "Kalimat Efektif", "Teks Prosedur"],
                  "Bahasa Inggris": ["Descriptive Text", "Simple Present Tense", "Procedure Text"]}
        kelas_siswa = {}
        for i, k in c.execute("SELECT id, kelas FROM siswa"):
            kelas_siswa.setdefault(k, []).append(i)
        for mp, u in MP:
            for k, ids in kelas_siswa.items():
                for d in hari_kerja(12)[1:]:
                    if rnd.random() > .5:
                        continue
                    sid = c.execute("INSERT INTO sesi(tanggal,kelas,mapel,materi,catatan,guru_id) VALUES (?,?,?,?,?,?)",
                                    (d.isoformat(), k, mp, rnd.choice(materi[mp]), "", gid[u])).lastrowid
                    c.executemany("INSERT INTO presensi_siswa(sesi_id,siswa_id,status) VALUES (?,?,?)",
                                  [(sid, i, rnd.choices(STATUS_HADIR, [91, 4, 3, 2])[0]) for i in ids])

        # tugas dan pengumpulan contoh
        A, B = kelas_siswa["VIII-A"], kelas_siswa["VIII-B"]
        spek = [("Latihan Soal Teorema Pythagoras", "Kerjakan soal nomor 1-10 pada LKPD, lalu unggah foto atau scan jawaban.", "Matematika", "VIII-A", 2, "guru"),
                ("Kuis Bilangan Berpangkat", "Selesaikan kuis dan unggah bukti pengerjaan.", "Matematika", "VIII-A", -4, "guru"),
                ("Laporan Praktikum Fotosintesis", "Susun laporan praktikum minimal 2 halaman (PDF atau DOCX).", "IPA", "VIII-A", 5, "guru3"),
                ("Menulis Teks Eksposisi", "Tulis teks eksposisi 3 paragraf bertema lingkungan sekolah.", "Bahasa Indonesia", "VIII-A", 1, "guru2"),
                ("Latihan Soal Teorema Pythagoras", "Kerjakan soal nomor 1-10 pada LKPD.", "Matematika", "VIII-B", 3, "guru"),
                ("Descriptive Text: My Favourite Place", "Write a descriptive text (minimum 100 words).", "Bahasa Inggris", "VIII-B", 4, "guru4")]
        tid = [c.execute("INSERT INTO tugas(judul,deskripsi,mapel,kelas,deadline,guru_id) VALUES (?,?,?,?,?,?)",
                         (j, ds, mp, k, (t0 + timedelta(days=dl)).isoformat(), gid[g])).lastrowid
               for j, ds, mp, k, dl, g in spek]

        def kirim(t, s, lalu, nilai=None, fb=None):
            waktu = (t0 - timedelta(days=lalu)).isoformat() + (" 19:30" if lalu else " 07:45")
            c.execute("INSERT INTO pengumpulan(tugas_id,siswa_id,waktu,nama_file,berkas,nilai,umpan_balik) VALUES (?,?,?,?,?,?,?)",
                      (tid[t], s, waktu, f"tugas_{nama[s].split()[0].lower()}.txt", b"Berkas contoh (data demo)", nilai, fb))

        for n, s in enumerate(A):
            if n != 2:  # Andi belum mengumpulkan kuis
                kirim(1, s, 5 if n != 4 else 3, rnd.randint(72, 98), "Pekerjaan rapi, perhatikan ketelitian hitung.")
        kirim(0, A[1], 1)
        kirim(0, A[3], 1)
        kirim(3, A[1], 0)
        kirim(3, A[5], 0)
        kirim(4, B[0], 1)
        kirim(5, B[1], 2)

        c.executemany("INSERT INTO nilai(siswa_id,mapel,harian,ujian) VALUES (?,?,?,?)",
                      [(i, mp, rnd.randint(68, 98), rnd.randint(62, 98)) for ids in kelas_siswa.values() for i in ids for mp, _ in MP])
        c.executemany("INSERT INTO pengumuman(tanggal,judul,isi) VALUES (?,?,?)", [
            ((t0 - timedelta(days=1)).isoformat(), "Jadwal Penilaian Tengah Semester", "Penilaian tengah semester dilaksanakan pekan depan. Siswa mohon mempersiapkan diri dan membawa kartu ujian."),
            ((t0 - timedelta(days=3)).isoformat(), "Rapat Wali Murid Semester Ganjil", "Rapat wali murid diadakan hari Sabtu pukul 09.00 WIB di aula sekolah."),
            ((t0 - timedelta(days=6)).isoformat(), "Pengumpulan Modul Ajar", "Bapak/Ibu guru mengunggah modul ajar semester berjalan paling lambat akhir bulan.")])
        c.commit()
    return True


# ══════════════════════════ KOMPONEN HTML ══════════════════════════
def esc(v):
    if v is None or v is pd.NA or (isinstance(v, float) and v != v):
        return "–"
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return _html.escape(str(v))


def H(s):
    """Render HTML dalam satu baris agar tidak dianggap blok kode oleh markdown."""
    st.markdown(" ".join(s.split()), unsafe_allow_html=True)


def badge(v):
    return f'<span class="bd {BD.get(str(v), "gray")}">{esc(v)}</span>'


def tb(df, badges=(), kosong="Belum ada data."):
    if df is None or len(df) == 0:
        return f'<div class="empty">{esc(kosong)}</div>'
    cols = list(df.columns)
    th = "".join(f"<th>{esc(c)}</th>" for c in cols)
    baris = "".join(
        "<tr>" + "".join(f"<td>{badge(v) if c in badges else esc(v)}</td>" for c, v in zip(cols, row)) + "</tr>"
        for row in df.itertuples(index=False))
    return f'<div class="tw"><table class="tb"><thead><tr>{th}</tr></thead><tbody>{baris}</tbody></table></div>'


def kpis(items):
    H('<div class="kpis">' + "".join(
        f'<div class="kpi {w}"><span>{esc(l)}</span><b>{esc(v)}</b><small>{esc(h)}</small></div>'
        for l, v, h, w in items) + "</div>")


def hero(judul, sub=""):
    n = now()
    H(f'<div class="hero"><div><div class="crumb">{esc(NAMA_SEKOLAH)} / {esc(judul)}</div>'
      f'<h1>{esc(judul)}</h1><p>{esc(sub)}</p></div>'
      f'<div class="rt"><b>{HARI[n.weekday()]}, {n.day} {BLN[n.month - 1]} {n.year}</b>{n:%H:%M} WIB</div></div>')


def sec(judul):
    H(f'<div class="sec">{esc(judul)}</div>')


def panel(judul, isi, flush=False):
    return f'<div class="pn"><div class="pn-h">{esc(judul)}</div><div class="pn-b{" flush" if flush else ""}">{isi}</div></div>'


def note(kind, isi):
    H(f'<div class="note {kind}">{isi}</div>')


def grid(df, **kw):
    """Tabel yang bisa diedit, kompatibel dengan versi Streamlit lama dan baru."""
    kw["use_container_width" if _PAKAI_UCW else "width"] = True if _PAKAI_UCW else "stretch"
    return st.data_editor(df, hide_index=True, **kw)


def flash(pesan):
    st.session_state["_flash"] = pesan


def tampil_flash():
    m = st.session_state.pop("_flash", None)
    if m:
        st.success(m)


def pengumuman_html(n=3):
    d = qdf("SELECT tanggal, judul, isi FROM pengumuman ORDER BY tanggal DESC, id DESC LIMIT ?", (n,))
    if d.empty:
        return '<div class="empty">Belum ada pengumuman.</div>'
    return "".join(f'<div class="ann"><small>{fmt(r.tanggal)}</small><b>{esc(r.judul)}</b><p>{esc(r.isi)}</p></div>'
                   for r in d.itertuples())


def jadwal(kelas, hari):
    """Jadwal contoh: 3 jam pelajaran per hari, tidak bentrok antar kelas."""
    off = 0 if kelas.endswith("A") else 2
    return [(SLOT[i], *MP[(i + hari + off) % 4]) for i in range(3)]


def jadwal_html(rows, kolom):
    if not rows:
        return '<div class="empty">Tidak ada jadwal hari ini.</div>'
    return tb(pd.DataFrame(rows, columns=kolom))


def peta_guru():
    u = dict(qdf("SELECT username, nama FROM users WHERE role='guru'").values.tolist())
    return {m: u.get(g, "-") for m, g in MP}


# ══════════════════════════ DATA BERSAMA (siswa, orang tua) ══════════════════════════
def status_kumpul(deadline, waktu):
    if waktu is None or waktu != waktu:
        return "Belum Dikumpulkan" if deadline >= today().isoformat() else "Melewati Deadline"
    return "Tepat Waktu" if str(waktu)[:10] <= deadline else "Terlambat"


def daftar_tugas_siswa(sid):
    d = qdf("""SELECT t.id, t.judul, t.mapel, t.deskripsi, t.deadline, p.waktu, p.nama_file, p.nilai, p.umpan_balik
               FROM siswa s JOIN tugas t ON t.kelas = s.kelas
               LEFT JOIN pengumpulan p ON p.tugas_id = t.id AND p.siswa_id = s.id
               WHERE s.id = ? ORDER BY t.deadline DESC, t.id DESC""", (sid,))
    d["status"] = [status_kumpul(a, b) for a, b in zip(d.deadline, d.waktu)]
    return d


def rekap_presensi(sid):
    return qdf("""SELECT s.tanggal, s.mapel, s.materi, p.status FROM presensi_siswa p
                  JOIN sesi s ON s.id = p.sesi_id WHERE p.siswa_id = ?
                  ORDER BY s.tanggal DESC, s.id DESC""", (sid,))


def blok_presensi(sid):
    d = rekap_presensi(sid)
    if d.empty:
        H('<div class="empty">Belum ada data presensi.</div>')
        return
    c, tot = d.status.value_counts(), len(d)
    kpis([("Kehadiran", f"{pct(c.get('Hadir', 0), tot)}%", f"{c.get('Hadir', 0)} dari {tot} pertemuan", "g"),
          ("Sakit", c.get("Sakit", 0), "pertemuan", "o"), ("Izin", c.get("Izin", 0), "pertemuan", ""),
          ("Alpa", c.get("Alpa", 0), "tanpa keterangan", "r")])
    t1, t2 = st.tabs(["Rekap per mata pelajaran", "Riwayat kehadiran"])
    with t1:
        r = d.pivot_table(index="mapel", columns="status", values="tanggal", aggfunc="count", fill_value=0)
        r = r.reindex(columns=STATUS_HADIR, fill_value=0)
        r["Total"] = r.sum(axis=1)
        r["% Hadir"] = (100 * r["Hadir"] / r["Total"]).round().astype(int).astype(str) + "%"
        r.columns.name, r.index.name = None, "Mata pelajaran"
        H(panel("Rekap kehadiran", tb(r.reset_index()), True))
    with t2:
        v = pd.DataFrame({"Tanggal": d.tanggal.map(fmt), "Mata pelajaran": d.mapel, "Materi": d.materi, "Status": d.status})
        H(panel("Riwayat 60 pertemuan terakhir", tb(v.head(60), ("Status",)), True))


def blok_nilai(sid):
    d = qdf("SELECT mapel, harian, ujian FROM nilai WHERE siswa_id=? ORDER BY mapel", (sid,))
    if d.empty:
        H('<div class="empty">Belum ada nilai.</div>')
        return
    d["akhir"] = (d.harian * .4 + d.ujian * .6).round(1)
    tinggi, rendah = d.loc[d.akhir.idxmax()], d.loc[d.akhir.idxmin()]
    kpis([("Rata-rata nilai akhir", f"{d.akhir.mean():.1f}", "seluruh mata pelajaran", ""),
          ("Nilai tertinggi", f"{tinggi.akhir:.1f}", tinggi.mapel, "g"),
          ("Perlu ditingkatkan", f"{rendah.akhir:.1f}", rendah.mapel, "o")])
    pg = peta_guru()
    v = pd.DataFrame({"Mata pelajaran": d.mapel, "Pengajar": d.mapel.map(pg), "Harian (40%)": d.harian,
                      "Ujian (60%)": d.ujian, "Nilai akhir": d.akhir, "Huruf": d.akhir.map(huruf)})
    H(panel("Nilai semester berjalan", tb(v, ("Huruf",)), True))
    rt = qdf("""SELECT t.mapel AS 'Mata pelajaran', COUNT(*) AS 'Tugas dinilai', ROUND(AVG(p.nilai),1) AS 'Rata-rata nilai tugas'
                FROM pengumpulan p JOIN tugas t ON t.id=p.tugas_id WHERE p.siswa_id=? AND p.nilai IS NOT NULL
                GROUP BY t.mapel""", (sid,))
    if not rt.empty:
        H(panel("Nilai tugas", tb(rt), True))


def tabel_tugas_anak(sid):
    d = daftar_tugas_siswa(sid)
    if d.empty:
        H('<div class="empty">Belum ada tugas untuk kelas ini.</div>')
        return
    sudah = int(d.waktu.notna().sum())
    belum = int(d.status.isin(["Belum Dikumpulkan", "Melewati Deadline"]).sum())
    kpis([("Total tugas", len(d), "untuk kelas ini", ""), ("Sudah dikumpulkan", sudah, f"{pct(sudah, len(d))}% dari total", "g"),
          ("Belum dikumpulkan", belum, "perlu ditindaklanjuti", "o" if belum else "g"),
          ("Terlambat dikumpulkan", int((d.status == "Terlambat").sum()), "dikirim setelah batas", "r")])
    v = pd.DataFrame({"Tugas": d.judul, "Mata pelajaran": d.mapel, "Batas kumpul": d.deadline.map(fmt), "Status": d.status,
                      "Waktu dikumpulkan": d.waktu, "Berkas": d.nama_file,
                      "Nilai": d.nilai.map(lambda x: "–" if pd.isna(x) else int(x)), "Catatan guru": d.umpan_balik})
    H(panel("Tugas dan pengumpulan", tb(v, ("Status",)), True))


# ══════════════════════════ AI (GEMINI) & DOKUMEN WORD ══════════════════════════
def ai(prompt, gambar=None):
    if genai is None or not API_KEY:
        st.error("GEMINI_API_KEY belum diatur. Tambahkan di Streamlit Cloud, menu Settings lalu Secrets.")
        return None
    genai.configure(api_key=API_KEY)
    if gambar is None:
        isi = prompt
    else:
        isi = [prompt] + (list(gambar) if isinstance(gambar, (list, tuple)) else [gambar])
    awal = [st.session_state["_model"]] if st.session_state.get("_model") else []
    for nama in dict.fromkeys(awal + MODEL_KANDIDAT):
        try:
            teks = genai.GenerativeModel(nama).generate_content(isi).text
            st.session_state["_model"] = nama
            return teks
        except Exception:
            continue
    try:
        for m in genai.list_models():
            if "generateContent" in m.supported_generation_methods:
                try:
                    teks = genai.GenerativeModel(m.name).generate_content(isi).text
                    st.session_state["_model"] = m.name
                    return teks
                except Exception:
                    continue
    except Exception as e:
        st.error(f"Gagal terhubung ke AI: {e}")
        return None
    st.error("Koneksi AI gagal. Periksa API key dan kuota Anda.")
    return None


def buat_docx(teks, judul):
    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(1)
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_heading(judul, 0)
    tabel = []

    def runs(p, t):
        for i, bagian in enumerate(t.split("**")):
            if bagian:
                p.add_run(bagian).bold = i % 2 == 1

    def tutup():
        if tabel:
            t = doc.add_table(rows=len(tabel), cols=max(len(r) for r in tabel))
            t.style = "Table Grid"
            for i, r in enumerate(tabel):
                for j, sel in enumerate(r):
                    runs(t.cell(i, j).paragraphs[0], sel)
            doc.add_paragraph()
            tabel.clear()

    for baris in teks.splitlines():
        b = baris.strip()
        if b.startswith("|"):
            sel = [x.strip() for x in b.strip("|").split("|")]
            if not set("".join(sel)) <= set("-: "):
                tabel.append(sel)
            continue
        tutup()
        if not b:
            continue
        if b.startswith("#"):
            doc.add_heading(b.lstrip("#").strip().replace("**", ""), min(len(b) - len(b.lstrip("#")), 3))
        elif b.startswith(("- ", "* ")):
            runs(doc.add_paragraph(style="List Bullet"), b[2:])
        else:
            runs(doc.add_paragraph(), b)
    tutup()
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ══════════════════════════ LOKASI & WAJAH (PRESENSI GURU) ══════════════════════════
def get_set(kunci, default=""):
    r = one("SELECT nilai FROM pengaturan WHERE kunci=?", (kunci,))
    return r[0] if r else default


def set_set(kunci, nilai):
    run("INSERT OR REPLACE INTO pengaturan(kunci,nilai) VALUES (?,?)", (kunci, str(nilai)))


def titik_sekolah():
    """(lintang, bujur, radius meter) atau None bila belum diatur kepala sekolah."""
    try:
        return float(get_set("lat")), float(get_set("lon")), int(float(get_set("radius", RADIUS_DEFAULT)))
    except ValueError:
        return None


def jarak_m(lat1, lon1, lat2, lon2):
    from math import asin, cos, radians, sin, sqrt
    p1, p2 = radians(lat1), radians(lat2)
    a = sin((p2 - p1) / 2) ** 2 + cos(p1) * cos(p2) * sin(radians(lon2 - lon1) / 2) ** 2
    return 2 * 6371000 * asin(sqrt(a))


def link_maps(lat, lon):
    return f"https://www.google.com/maps?q={lat},{lon}"


def kecilkan(data, maks=640):
    im = Image.open(io.BytesIO(data)).convert("RGB")
    im.thumbnail((maks, maks))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=82)
    return buf.getvalue()


def bandingkan_wajah(ref, baru):
    """Bandingkan foto acuan dan foto baru dengan AI. Hasil: (status, catatan)."""
    a = Image.open(io.BytesIO(ref)).convert("RGB")
    b = Image.open(io.BytesIO(baru)).convert("RGB")
    teks = ai("Kamu petugas verifikasi identitas. Gambar pertama adalah foto acuan, gambar kedua adalah foto selfie baru. "
              "Tentukan apakah keduanya orang yang sama. Jika wajah tidak terlihat jelas, atau foto kedua tampak diambil dari layar "
              "atau foto cetak, jawab TIDAK_JELAS. Jawab HANYA satu baris JSON tanpa teks lain: "
              '{"hasil":"SAMA|BEDA|TIDAK_JELAS","keyakinan":0-100,"alasan":"satu kalimat singkat"}', [a, b])
    if not teks:
        return "Perlu Verifikasi", "AI tidak dapat dihubungi, perlu dicek manual."
    try:
        d = json.loads(re.search(r"\{.*\}", teks, re.S).group(0))
        hasil, yakin, alasan = str(d.get("hasil", "")).upper(), float(d.get("keyakinan", 0)), str(d.get("alasan", ""))
    except Exception:
        return "Perlu Verifikasi", "Hasil AI tidak terbaca, perlu dicek manual."
    if hasil == "SAMA" and yakin >= 70:
        return "Cocok", alasan
    if hasil == "BEDA" and yakin >= 70:
        return "Tidak Cocok", alasan
    return "Perlu Verifikasi", alasan


def _geo(kunci):
    try:
        return streamlit_geolocation(key=f"geo_{kunci}")
    except TypeError:
        return streamlit_geolocation()


def ambil_bukti(kunci, titik):
    """Ambil lokasi GPS dan foto selfie. Kembalikan ((lat, lon, akurasi) atau None, bytes foto atau None)."""
    if streamlit_geolocation is None:
        st.error("Komponen lokasi belum terpasang. Tambahkan streamlit-geolocation ke requirements.txt lalu deploy ulang.")
        return None, None
    st.markdown("**1. Lokasi**")
    st.caption("Tekan ikon lokasi di bawah, lalu izinkan akses lokasi di browser. Gunakan HTTPS.")
    loc = _geo(kunci)
    lat = loc.get("latitude") if isinstance(loc, dict) else None
    lon = loc.get("longitude") if isinstance(loc, dict) else None
    akur = loc.get("accuracy") if isinstance(loc, dict) else None
    lokasi = None
    if lat is None or lon is None:
        note("warn", "Lokasi belum diambil.")
    else:
        lokasi = (float(lat), float(lon), float(akur) if akur is not None else None)
        info = f"Koordinat {lat:.5f}, {lon:.5f}" + (f", akurasi sekitar {akur:.0f} m" if akur is not None else "")
        if titik:
            j = int(jarak_m(lokasi[0], lokasi[1], titik[0], titik[1]))
            if j <= titik[2]:
                note("ok", f"{info}. Anda berada {j} m dari sekolah (dalam radius {titik[2]} m).")
            else:
                note("bad", f"{info}. Anda berada {j} m dari sekolah, di luar radius {titik[2]} m.")
        else:
            note("info", f"{info}. Titik sekolah belum diatur kepala sekolah, jadi radius belum diperiksa.")
    st.markdown("**2. Foto selfie**")
    foto = st.camera_input("Ambil foto wajah (menghadap lurus, cahaya cukup)", key=f"cam_{kunci}")
    return lokasi, (foto.getvalue() if foto else None)


def proses_bukti(g, tgl, jenis, lokasi, foto, titik):
    """Periksa lokasi dan wajah, simpan bukti. Kembalikan (berhasil, pesan, status wajah)."""
    if lokasi is None:
        return False, "Lokasi belum diambil. Tekan ikon lokasi dan izinkan akses lokasi.", None
    if not foto:
        return False, "Foto selfie belum diambil.", None
    lat, lon, akur = lokasi
    if akur is not None and akur > MAX_AKURASI:
        return False, f"Akurasi lokasi rendah (sekitar {akur:.0f} m). Pindah ke tempat terbuka lalu ambil lokasi lagi.", None
    jarak = None
    if titik:
        jarak = int(jarak_m(lat, lon, titik[0], titik[1]))
        if jarak > titik[2]:
            return False, f"Anda berada {jarak} m dari sekolah, di luar radius {titik[2]} m. Presensi hanya dapat dilakukan di area sekolah.", None
    ref = one("SELECT foto FROM wajah_guru WHERE guru_id=?", (g,))
    kecil = kecilkan(foto)
    with st.spinner("Mencocokkan wajah..."):
        wajah, cat = bandingkan_wajah(ref[0], kecil)
    if wajah == "Tidak Cocok":
        return False, f"Wajah tidak cocok dengan foto terdaftar. Ambil ulang foto dengan cahaya cukup. ({cat})", None
    run("""INSERT OR REPLACE INTO bukti_presensi(guru_id,tanggal,jenis,waktu,lat,lon,akurasi,jarak,foto,wajah,catatan)
           VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
        (g, tgl, jenis, now().strftime("%H:%M"), lat, lon, akur, jarak, kecil, wajah, cat))
    return True, "", wajah


def daftar_wajah(g):
    note("warn", "Wajah Anda belum didaftarkan. Daftarkan satu kali sebagai foto acuan presensi. "
                 "Perubahan selanjutnya hanya dapat dilakukan oleh kepala sekolah.")
    setuju = st.checkbox("Saya menyetujui foto wajah dan lokasi saya dipakai untuk presensi sesuai kebijakan sekolah.", key="setuju_wajah")
    foto = st.camera_input("Foto wajah acuan (menghadap lurus, cahaya cukup, tanpa masker)", key="cam_daftar")
    if st.button("Daftarkan wajah", type="primary", key="btn_daftar_wajah"):
        if not setuju:
            st.warning("Centang persetujuan terlebih dahulu.")
        elif not foto:
            st.warning("Ambil foto wajah terlebih dahulu.")
        else:
            ts = now().strftime("%Y-%m-%d %H:%M")
            run("INSERT OR REPLACE INTO wajah_guru(guru_id,foto,setuju_pada,diperbarui) VALUES (?,?,?,?)",
                (g, kecilkan(foto.getvalue()), ts, ts))
            flash("Wajah berhasil didaftarkan. Sekarang Anda dapat presensi.")
            st.rerun()


# ══════════════════════════ AUTENTIKASI & NAVIGASI ══════════════════════════
def login(username, pw):
    if st.session_state.get("gagal", 0) >= 5:
        st.error("Terlalu banyak percobaan. Muat ulang halaman untuk mencoba lagi.")
        return False
    usr = username.strip()
    r = one("SELECT id, role, nama, siswa_id, pw FROM users WHERE username=?", (usr,))
    if r and hmac.compare_digest(r[4], hpw(usr, pw)):
        st.session_state["user"] = {"id": r[0], "role": r[1], "nama": r[2], "siswa_id": r[3], "username": usr}
        st.session_state.pop("menu", None)
        st.session_state["gagal"] = 0
        return True
    st.session_state["gagal"] = st.session_state.get("gagal", 0) + 1
    st.error("Username atau password salah.")
    return False


def sidebar(u):
    menu = list(PAGES[u["role"]])
    aktif = st.session_state.get("menu")
    if aktif not in menu:
        aktif = st.session_state["menu"] = menu[0]
    with st.sidebar:
        H(f'<div class="sb-brand"><div class="logo">🎓</div><div><b>LMS</b><span>{esc(NAMA_SEKOLAH)}</span></div></div>')
        H(f'<div class="sb-user"><div class="av">{esc(initials(u["nama"]))}</div>'
          f'<div><b>{esc(u["nama"])}</b><span>{ROLE[u["role"]]}</span></div></div>')
        for m in menu:
            if st.button(m, key=f"nav_{m}", type="primary" if m == aktif else "secondary"):
                st.session_state["menu"] = m
                st.rerun()
        st.divider()
        st.button("Keluar", key="logout", on_click=lambda: st.session_state.clear())
    return aktif


def _form_login():
    a, b = st.columns([1.15, 1], gap="large")
    with a:
        H(f'<div class="lg"><div class="logo">🎓</div><h2>LMS {esc(NAMA_SEKOLAH)}</h2>'
          f'<p>Sistem pembelajaran online terpadu {esc(NAMA_SEKOLAH)}.</p><ul>'
          '<li>Presensi guru dan siswa tercatat langsung di sistem</li>'
          '<li>Tugas diunggah, dinilai, dan diberi umpan balik di satu tempat</li>'
          '<li>Modul ajar dan koreksi lembar jawaban dibantu AI</li>'
          '<li>Kepala sekolah dan orang tua memantau perkembangan tanpa menunggu rapor</li></ul></div>')
    with b:
        st.markdown("### Masuk ke portal")
        st.caption("Gunakan akun yang diberikan sekolah.")
        with st.form("form_login"):
            usr = st.text_input("Username", placeholder="contoh: guru")
            pw = st.text_input("Password", type="password")
            if st.form_submit_button("Masuk", type="primary"):
                if login(usr, pw):
                    st.rerun()
        if SHOW_DEMO:
            with st.expander("Akun demo"):
                H(tb(pd.DataFrame([("Guru", "guru", "guru123"), ("Siswa", "siswa", "siswa123"),
                                   ("Kepala sekolah", "kepsek", "kepsek123"), ("Orang tua", "ortu", "ortu123")],
                                  columns=["Peran", "Username", "Password"])))


PANDUAN = {
    "👩‍🏫 Guru": (
        "Pengajar mengelola presensi, jurnal, tugas, nilai, dan perangkat pembelajaran dalam satu tempat.",
        [("Beranda", "Ringkasan jadwal mengajar hari ini, tugas aktif, dan pengumpulan yang menunggu dinilai."),
         ("Presensi & jurnal", "Presensi masuk dan pulang guru, izin atau sakit, jurnal mengajar, dan presensi siswa per pertemuan."),
         ("Tugas & penilaian", "Terbitkan tugas untuk kelas, pantau siapa yang sudah mengumpulkan, beri nilai dan umpan balik."),
         ("Buku nilai", "Isi nilai harian dan ujian. Nilai akhir dihitung otomatis (40% harian, 60% ujian)."),
         ("Administrasi guru", "Unggah perangkat pembelajaran (Prota, Promes, modul ajar, dan lainnya), rekap hadir dan nilai, serta jadwal mengajar."),
         ("Asisten AI", "Susun modul ajar dan koreksi lembar jawaban tulisan tangan dengan bantuan AI.")],
        ["Masuk dengan akun guru dari sekolah.", "Lakukan presensi masuk di menu Presensi & jurnal.",
         "Isi jurnal dan presensi siswa setelah mengajar.", "Terbitkan tugas dan nilai pengumpulan siswa."]),
    "🎒 Siswa": (
        "Peserta didik melihat tugas, mengumpulkan pekerjaan, dan memantau kehadiran serta nilai sendiri.",
        [("Beranda", "Ringkasan kehadiran, tugas yang belum dikumpulkan, jadwal hari ini, dan pengumuman."),
         ("Tugas saya", "Daftar tugas dari guru. Unggah berkas (PDF, DOCX, atau foto) lalu tekan Kirim tugas."),
         ("Presensi saya", "Rekap kehadiran per mata pelajaran dan riwayat tiap pertemuan."),
         ("Nilai & rapor", "Nilai harian, ujian, nilai akhir, dan nilai tugas."),
         ("Tutor AI", "Tanya materi pelajaran dan dapatkan penjelasan sederhana beserta latihan.")],
        ["Masuk dengan akun siswa dari sekolah.", "Buka menu Tugas saya untuk melihat batas kumpul.",
         "Unggah berkas tugas sebelum batas waktu.", "Lihat nilai dan catatan guru setelah tugas dinilai."]),
    "🏫 Kepala sekolah": (
        "Kepala sekolah memantau kedisiplinan guru, kehadiran siswa, dan capaian akademik tanpa menunggu laporan.",
        [("Dashboard", "Guru hadir hari ini, tren kehadiran siswa, dan pengumuman. Kepala sekolah dapat menerbitkan pengumuman."),
         ("Presensi guru", "Rekap kehadiran, keterlambatan, sakit, dan izin guru dengan filter tanggal, dapat diunduh CSV."),
         ("Presensi siswa", "Rekap kehadiran per siswa dan per kelas, beserta siswa yang perlu perhatian."),
         ("Jurnal mengajar", "Semua pertemuan yang dicatat guru."),
         ("Nilai & tugas", "Rata-rata nilai per kelas dan mata pelajaran, serta pengumpulan tugas."),
         ("Administrasi guru", "Kelengkapan perangkat pembelajaran setiap guru.")],
        ["Masuk dengan akun kepala sekolah.", "Mulai dari Dashboard untuk gambaran hari ini.",
         "Buka menu rekap bila perlu rincian.", "Unduh CSV untuk laporan."]),
    "👨‍👩‍👧 Orang tua": (
        "Orang tua atau wali memantau perkembangan belajar anak secara langsung.",
        [("Beranda", "Profil anak, wali kelas, ringkasan kehadiran, tugas, nilai, dan peringatan bila ada tugas terlewat atau alpa."),
         ("Presensi anak", "Kehadiran anak pada setiap pertemuan dan rekap per mata pelajaran."),
         ("Tugas & pengumpulan", "Daftar tugas anak, status pengumpulan, nilai, dan catatan guru."),
         ("Nilai anak", "Nilai harian, ujian, nilai akhir, dan nilai tugas per mata pelajaran.")],
        ["Masuk dengan akun orang tua dari sekolah.", "Baca peringatan di Beranda.",
         "Periksa tugas yang belum dikumpulkan dan ingatkan anak.", "Hubungi wali kelas bila ada kendala."]),
}


def panduan_login():
    st.markdown("### Panduan penggunaan")
    st.caption("Pilih peran Anda untuk melihat menu dan cara memakainya. Akun diberikan oleh sekolah.")
    tabs = st.tabs(list(PANDUAN) + ["🕒 Cara absen"])
    for tab, (peran, (desc, menu, langkah)) in zip(tabs, PANDUAN.items()):
        with tab:
            H(panel(peran, f"<p style='margin:0'>{esc(desc)}</p>"))
            H(panel("Menu yang tersedia", tb(pd.DataFrame(menu, columns=["Menu", "Fungsi"])), True))
            H(panel("Langkah singkat", "<ol style='margin:0;padding-left:1.2rem'>"
                    + "".join(f"<li>{esc(x)}</li>" for x in langkah) + "</ol>"))
    with tabs[-1]:
        ol = "<ol style='margin:0;padding-left:1.2rem'>"
        H(panel("Absen guru (diisi guru sendiri)", ol
                + "<li>Buka menu <b>Presensi &amp; jurnal</b>, tab <b>Presensi saya</b>.</li>"
                "<li>Tekan ikon lokasi (izinkan akses lokasi) dan ambil foto selfie. Pertama kali, daftarkan wajah dahulu.</li>"
                f"<li>Tekan <b>Presensi masuk</b> saat tiba (lewat {BATAS_MASUK} dicatat Terlambat).</li>"
                "<li>Tekan <b>Presensi pulang</b> saat pulang.</li>"
                "<li>Bila tidak hadir, isi <b>Tidak dapat hadir?</b> (Sakit atau Izin) lalu <b>Kirim keterangan</b>.</li></ol>"))
        H(panel("Absen siswa (diisi guru setelah mengajar)", ol
                + "<li>Buka menu <b>Presensi &amp; jurnal</b>, tab <b>Jurnal &amp; presensi siswa</b>.</li>"
                "<li>Pilih kelas dan mata pelajaran, isi tanggal dan materi.</li>"
                "<li>Semua siswa awalnya Hadir. Ubah status siswa yang Sakit, Izin, atau Alpa.</li>"
                "<li>Tekan <b>Simpan jurnal dan presensi</b>.</li></ol>"))
        H(panel("Melihat hasil absen", ol
                + "<li><b>Siswa</b>: menu Presensi saya.</li><li><b>Orang tua</b>: menu Presensi anak.</li>"
                "<li><b>Kepala sekolah</b>: menu Presensi guru dan Presensi siswa.</li></ol>"))


def halaman_login():
    t1, t2 = st.tabs(["🔐 Masuk", "📘 Panduan penggunaan"])
    with t1:
        _form_login()
    with t2:
        panduan_login()


# ══════════════════════════ PORTAL GURU ══════════════════════════
def guru_home(u):
    hero("Beranda pengajar", f"{sapa()}, {u['nama']}. Ringkasan aktivitas mengajar Anda.")
    g, hi = u["id"], today().isoformat()
    n_sesi = one("SELECT COUNT(*) FROM sesi WHERE guru_id=? AND tanggal LIKE ?", (g, hi[:7] + "%"))[0]
    n_tugas = one("SELECT COUNT(*) FROM tugas WHERE guru_id=? AND deadline>=?", (g, hi))[0]
    n_nilai = one("""SELECT COUNT(*) FROM pengumpulan p JOIN tugas t ON t.id=p.tugas_id
                     WHERE t.guru_id=? AND p.nilai IS NULL""", (g,))[0]
    pr = one("SELECT status, jam_masuk FROM presensi_guru WHERE guru_id=? AND tanggal=?", (g, hi))
    kpis([("Pertemuan bulan ini", n_sesi, "jurnal mengajar tercatat", ""),
          ("Tugas aktif", n_tugas, "belum melewati batas kumpul", "t"),
          ("Menunggu penilaian", n_nilai, "pengumpulan dari siswa", "o"),
          ("Presensi hari ini", pr[0] if pr else "Belum", f"masuk {pr[1]} WIB" if pr and pr[1] else "buka menu Presensi & jurnal",
           "g" if pr else "r")])
    a, b = st.columns([1.25, 1])
    with a:
        hari = today().weekday()
        rows = sorted((s, k, m) for k in KELAS for s, m, gu in jadwal(k, hari) if gu == u["username"]) if hari < 5 else []
        H(panel("Jadwal mengajar hari ini", jadwal_html(rows, ["Jam", "Kelas", "Mata pelajaran"]), True))
        d = qdf("""SELECT si.nama AS Siswa, t.judul AS Tugas, t.kelas AS Kelas, p.waktu AS 'Waktu kumpul'
                   FROM pengumpulan p JOIN tugas t ON t.id=p.tugas_id JOIN siswa si ON si.id=p.siswa_id
                   WHERE t.guru_id=? AND p.nilai IS NULL ORDER BY p.waktu DESC LIMIT 8""", (g,))
        H(panel("Menunggu penilaian", tb(d, kosong="Semua pengumpulan sudah dinilai."), True))
    with b:
        H(panel("Pengumuman", pengumuman_html()))


def guru_presensi(u):
    hero("Presensi & jurnal mengajar", "Catat kehadiran Anda, jurnal pertemuan, dan presensi siswa.")
    t1, t2, t3 = st.tabs(["🕒 Presensi saya", "📝 Jurnal & presensi siswa", "📜 Riwayat"])
    g, hi = u["id"], today().isoformat()
    opsi_mp = [m for m, gu in MP if gu == u["username"]] or [m for m, _ in MP]
    with t1:
        note("info", "<b>Cara absen guru:</b><ol style='margin:6px 0 0;padding-left:1.2rem'>"
                     "<li>Tekan ikon lokasi dan izinkan akses lokasi di browser.</li>"
                     "<li>Ambil foto selfie dengan kamera, wajah menghadap lurus dan terlihat jelas.</li>"
                     "<li>Tekan <b>Presensi masuk</b> (lewat pukul " + BATAS_MASUK + " dicatat Terlambat). "
                     "Sistem memeriksa lokasi dan mencocokkan wajah.</li>"
                     "<li>Sebelum pulang, ulangi langkah 1 dan 2, lalu tekan <b>Presensi pulang</b>.</li>"
                     "<li>Bila tidak hadir, isi bagian <b>Tidak dapat hadir</b> di bawah.</li></ol>")
        pr = one("SELECT status, jam_masuk, jam_pulang FROM presensi_guru WHERE guru_id=? AND tanggal=?", (g, hi))
        if pr:
            ket_st = f'{badge(pr[0])}<p style="margin:10px 0 0">Masuk: <b>{esc(pr[1])}</b>, pulang: <b>{esc(pr[2])}</b></p>'
        else:
            ket_st = badge("Belum Presensi") + '<p style="margin:10px 0 0">Anda belum presensi hari ini.</p>'
        H(panel("Status hari ini", ket_st))
        titik = titik_sekolah()
        boleh_masuk = pr is None
        boleh_pulang = bool(pr) and pr[0] in ("Hadir", "Terlambat") and not pr[2]
        if not one("SELECT 1 FROM wajah_guru WHERE guru_id=?", (g,)):
            sec("Daftarkan wajah")
            daftar_wajah(g)
        else:
            lokasi, foto = None, None
            if boleh_masuk or boleh_pulang:
                sec("Lokasi dan foto selfie")
                lokasi, foto = ambil_bukti("masuk" if boleh_masuk else "pulang", titik)
            sec("Tombol presensi")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("✅ Presensi masuk", key="btn_masuk", type="primary", disabled=not boleh_masuk):
                    ok, pesan, wajah = proses_bukti(g, hi, "masuk", lokasi, foto, titik)
                    if ok:
                        jam = now().strftime("%H:%M")
                        run("INSERT OR REPLACE INTO presensi_guru(guru_id,tanggal,jam_masuk,status,keterangan) VALUES (?,?,?,?,?)",
                            (g, hi, jam, "Terlambat" if jam > BATAS_MASUK else "Hadir", ""))
                        flash(f"Presensi masuk tercatat pukul {jam} WIB. Wajah: {wajah}.")
                        st.rerun()
                    else:
                        st.error(pesan)
                st.caption("Aktif bila belum presensi hari ini." if boleh_masuk else "Sudah presensi masuk hari ini.")
            with c2:
                if st.button("🏠 Presensi pulang", key="btn_pulang", type="primary", disabled=not boleh_pulang):
                    ok, pesan, wajah = proses_bukti(g, hi, "pulang", lokasi, foto, titik)
                    if ok:
                        jam = now().strftime("%H:%M")
                        run("UPDATE presensi_guru SET jam_pulang=? WHERE guru_id=? AND tanggal=?", (jam, g, hi))
                        flash(f"Presensi pulang tercatat pukul {jam} WIB. Wajah: {wajah}.")
                        st.rerun()
                    else:
                        st.error(pesan)
                st.caption("Aktif setelah presensi masuk." if not pr else ("Sudah presensi pulang." if pr[2] else
                           ("Aktif, tekan saat pulang." if boleh_pulang else "Tidak tersedia untuk status ini.")))
        sec("Tidak dapat hadir (sakit atau izin)")
        with st.form("f_izin"):
            sts = st.radio("Alasan", ["Sakit", "Izin"], horizontal=True, disabled=not boleh_masuk)
            ket = st.text_input("Keterangan", disabled=not boleh_masuk)
            if st.form_submit_button("📨 Kirim keterangan", disabled=not boleh_masuk):
                run("INSERT OR REPLACE INTO presensi_guru(guru_id,tanggal,jam_masuk,status,keterangan) VALUES (?,?,?,?,?)",
                    (g, hi, None, sts, ket.strip()))
                flash("Keterangan ketidakhadiran tersimpan.")
                st.rerun()
    with t2:
        note("info", "<b>Cara mengabsen siswa:</b><ol style='margin:6px 0 0;padding-left:1.2rem'>"
                     "<li>Pilih <b>Kelas</b> dan <b>Mata pelajaran</b>.</li>"
                     "<li>Isi <b>Tanggal</b> dan <b>Materi pembahasan</b>.</li>"
                     "<li>Di tabel, semua siswa awalnya <b>Hadir</b>. Ketuk kolom <b>Status</b> siswa yang tidak hadir, "
                     "lalu pilih Sakit, Izin, atau Alpa.</li>"
                     "<li>Tekan <b>Simpan jurnal dan presensi</b> di bagian bawah.</li></ol>")
        c1, c2 = st.columns(2)
        kelas, mapel = c1.selectbox("Kelas", KELAS), c2.selectbox("Mata pelajaran", opsi_mp)
        sw = qdf("SELECT id, nis, nama FROM siswa WHERE kelas=? ORDER BY nama", (kelas,))
        with st.form(f"f_jurnal_{kelas}"):
            x, y = st.columns([1, 2])
            tgl = x.date_input("Tanggal", today())
            materi = y.text_input("Materi pembahasan", placeholder="contoh: Teorema Pythagoras")
            hasil = grid(pd.DataFrame({"NIS": sw.nis, "Nama siswa": sw.nama, "Status": "Hadir"}),
                         key=f"pres_{kelas}", disabled=["NIS", "Nama siswa"],
                         column_config={"Status": st.column_config.SelectboxColumn("Status", options=STATUS_HADIR, required=True)})
            catatan = st.text_area("Catatan atau kendala", height=80)
            if st.form_submit_button("Simpan jurnal dan presensi", type="primary"):
                if not materi.strip():
                    st.warning("Materi pembahasan wajib diisi.")
                else:
                    sid = run("INSERT INTO sesi(tanggal,kelas,mapel,materi,catatan,guru_id) VALUES (?,?,?,?,?,?)",
                              (tgl.isoformat(), kelas, mapel, materi.strip(), catatan.strip(), g))
                    run("INSERT INTO presensi_siswa(sesi_id,siswa_id,status) VALUES (?,?,?)",
                        [(sid, int(i), s) for i, s in zip(sw.id, hasil["Status"])], many=True)
                    flash(f"Jurnal {mapel} kelas {kelas} tersimpan, {len(sw)} siswa dicatat.")
                    st.rerun()
    with t3:
        d = qdf("""SELECT s.tanggal AS Tanggal, s.kelas AS Kelas, s.mapel AS 'Mata pelajaran', s.materi AS Materi,
                          COALESCE(SUM(p.status='Hadir'),0) AS Hadir, COALESCE(SUM(p.status='Sakit'),0) AS Sakit,
                          COALESCE(SUM(p.status='Izin'),0) AS Izin, COALESCE(SUM(p.status='Alpa'),0) AS Alpa
                   FROM sesi s LEFT JOIN presensi_siswa p ON p.sesi_id=s.id WHERE s.guru_id=?
                   GROUP BY s.id ORDER BY s.tanggal DESC, s.id DESC LIMIT 30""", (g,))
        d["Tanggal"] = d["Tanggal"].map(fmt)
        H(panel("Jurnal mengajar terakhir", tb(d), True))
        m = qdf("""SELECT tanggal AS Tanggal, status AS Status, jam_masuk AS Masuk, jam_pulang AS Pulang, keterangan AS Keterangan
                   FROM presensi_guru WHERE guru_id=? ORDER BY tanggal DESC LIMIT 15""", (g,))
        m["Tanggal"] = m["Tanggal"].map(fmt)
        H(panel("Presensi saya", tb(m, ("Status",)), True))


def guru_tugas(u):
    hero("Tugas & penilaian", "Terbitkan tugas, pantau pengumpulan, lalu beri nilai dan umpan balik.")
    t1, t2 = st.tabs(["📥 Pengumpulan & penilaian", "➕ Buat tugas baru"])
    opsi_mp = [m for m, gu in MP if gu == u["username"]] or [m for m, _ in MP]
    with t2:
        with st.form("f_tugas", clear_on_submit=True):
            judul = st.text_input("Judul tugas")
            a, b, c = st.columns(3)
            mapel = a.selectbox("Mata pelajaran", opsi_mp)
            kelas = b.selectbox("Kelas", KELAS)
            dl = c.date_input("Batas pengumpulan", today() + timedelta(days=7), min_value=today())
            desk = st.text_area("Instruksi tugas", height=110)
            if st.form_submit_button("Terbitkan tugas", type="primary"):
                if judul.strip():
                    run("INSERT INTO tugas(judul,deskripsi,mapel,kelas,deadline,guru_id) VALUES (?,?,?,?,?,?)",
                        (judul.strip(), desk.strip(), mapel, kelas, dl.isoformat(), u["id"]))
                    flash("Tugas diterbitkan dan langsung tampil di portal siswa dan orang tua.")
                    st.rerun()
                else:
                    st.warning("Judul tugas wajib diisi.")
    with t1:
        d = qdf("SELECT id, judul, kelas, deadline FROM tugas WHERE guru_id=? ORDER BY id DESC", (u["id"],))
        if d.empty:
            H('<div class="empty">Belum ada tugas. Buat tugas pertama di tab sebelah.</div>')
            return
        label = {int(i): f"{j} (kelas {k}, batas {fmt(dl, False)})" for i, j, k, dl in zip(d.id, d.judul, d.kelas, d.deadline)}
        tid = st.selectbox("Pilih tugas", list(label), format_func=label.get)
        kelas = d.loc[d.id == tid, "kelas"].iloc[0]
        df = qdf("""SELECT s.id AS sid, s.nama, p.id AS pid, p.waktu, p.nama_file, p.nilai, p.umpan_balik
                    FROM siswa s LEFT JOIN pengumpulan p ON p.siswa_id=s.id AND p.tugas_id=?
                    WHERE s.kelas=? ORDER BY s.nama""", (int(tid), kelas))
        dinilai = int(df.nilai.notna().sum())
        kpis([("Sudah mengumpulkan", f"{int(df.pid.notna().sum())} dari {len(df)}", f"kelas {kelas}", "g"),
              ("Belum mengumpulkan", int(df.pid.isna().sum()), "siswa", "o"), ("Sudah dinilai", dinilai, "pengumpulan", "t")])
        tampil = pd.DataFrame({
            "Nama siswa": df.nama,
            "Status": ["Belum Dikumpulkan" if pd.isna(p) else ("Sudah Dinilai" if pd.notna(n) else "Menunggu Penilaian")
                       for p, n in zip(df.pid, df.nilai)],
            "Waktu kumpul": df.waktu.fillna("–"), "Nilai": df.nilai, "Umpan balik": df.umpan_balik.fillna("")})
        ed = grid(tampil, key=f"nilai_tugas_{tid}", disabled=["Nama siswa", "Status", "Waktu kumpul"],
                  column_config={"Nilai": st.column_config.NumberColumn("Nilai", min_value=0, max_value=100, step=1)})
        if st.button("Simpan penilaian", type="primary"):
            for pid, nl, fb in zip(df.pid, ed["Nilai"], ed["Umpan balik"]):
                if pd.notna(pid):
                    run("UPDATE pengumpulan SET nilai=?, umpan_balik=? WHERE id=?",
                        (None if pd.isna(nl) else int(nl), (fb or "").strip() if isinstance(fb, str) else "", int(pid)))
            flash("Penilaian tersimpan.")
            st.rerun()
        sub = df[df.pid.notna()]
        if not sub.empty:
            nm = dict(zip(sub.pid.astype(int), sub.nama))
            pil = st.selectbox("Unduh berkas siswa", list(nm), format_func=nm.get)
            nf, bl = one("SELECT nama_file, berkas FROM pengumpulan WHERE id=?", (pil,))
            st.download_button("Unduh berkas", bl, file_name=nf)


def guru_nilai(u):
    hero("Buku nilai", "Nilai akhir dihitung dari 40% nilai harian dan 60% nilai ujian.")
    a, b = st.columns(2)
    kelas = a.selectbox("Kelas", KELAS)
    mapel = b.selectbox("Mata pelajaran", [m for m, gu in MP if gu == u["username"]] or [m for m, _ in MP])
    d = qdf("""SELECT s.id, s.nis, s.nama, n.harian, n.ujian FROM siswa s
               LEFT JOIN nilai n ON n.siswa_id=s.id AND n.mapel=? WHERE s.kelas=? ORDER BY s.nama""", (mapel, kelas))
    ed = grid(pd.DataFrame({"NIS": d.nis, "Nama siswa": d.nama, "Harian": d.harian, "Ujian": d.ujian}),
              key=f"nilai_{kelas}_{mapel}", disabled=["NIS", "Nama siswa"],
              column_config={"Harian": st.column_config.NumberColumn(min_value=0, max_value=100, step=1),
                             "Ujian": st.column_config.NumberColumn(min_value=0, max_value=100, step=1)})
    if st.button("Simpan nilai", type="primary"):
        run("INSERT OR REPLACE INTO nilai(siswa_id,mapel,harian,ujian) VALUES (?,?,?,?)",
            [(int(i), mapel, None if pd.isna(h) else int(h), None if pd.isna(x) else int(x))
             for i, h, x in zip(d.id, ed["Harian"], ed["Ujian"])], many=True)
        flash("Nilai berhasil disimpan.")
        st.rerun()
    v = ed.copy()
    v["Nilai akhir"] = (v["Harian"].fillna(0) * .4 + v["Ujian"].fillna(0) * .6).round(1)
    v["Huruf"] = v["Nilai akhir"].map(huruf)
    sec("Pratinjau nilai akhir")
    H(tb(v, ("Huruf",)))


def guru_ai(u):
    hero("Asisten AI", "Susun modul ajar dan koreksi lembar jawaban dengan bantuan AI.")
    t1, t2 = st.tabs(["✨ Modul ajar", "📸 Koreksi lembar jawaban"])
    with t1:
        with st.form("f_modul"):
            a, b, c = st.columns(3)
            mapel = a.text_input("Mata pelajaran", "Matematika")
            kls = b.text_input("Kelas / fase", "VIII / Fase D")
            n = c.number_input("Jumlah pertemuan", 1, 12, 3)
            topik = st.text_input("Topik", "Teorema Pythagoras")
            d, e = st.columns(2)
            alokasi = d.text_input("Alokasi waktu per pertemuan", "2 x 40 menit")
            pendekatan = e.selectbox("Pendekatan", ["Pembelajaran mendalam (mindful, meaningful, joyful)", "Kurikulum Merdeka (umum)"])
            kirim = st.form_submit_button("Susun modul ajar", type="primary")
        if kirim and topik.strip():
            with st.spinner("AI sedang menyusun modul ajar..."):
                hasil = ai(f"""Kamu pakar kurikulum Indonesia. Susun MODUL AJAR lengkap dan siap pakai dalam Bahasa Indonesia,
format Markdown (gunakan tabel untuk identitas modul dan rubrik).
Mata pelajaran: {mapel}. Kelas/Fase: {kls}. Topik: {topik}.
Jumlah pertemuan: {n}. Alokasi waktu per pertemuan: {alokasi}. Pendekatan: {pendekatan}.
Struktur wajib: 1) Identitas modul 2) Kompetensi awal dan dimensi Profil Lulusan yang dikembangkan 3) Sarana prasarana
4) Target peserta didik 5) Model dan metode pembelajaran 6) Tujuan pembelajaran 7) Pemahaman bermakna dan pertanyaan pemantik
8) Kegiatan pembelajaran per pertemuan (pendahuluan, inti: memahami, mengaplikasi, merefleksi, penutup, lengkap dengan alokasi menit)
9) Asesmen diagnostik, formatif, sumatif beserta rubrik 10) Pengayaan dan remedial 11) Refleksi guru dan peserta didik
12) Lampiran: LKPD, bahan bacaan, glosarium.""")
            if hasil:
                st.session_state["modul"] = (f"Modul Ajar {mapel} - {topik}", hasil)
        if st.session_state.get("modul"):
            judul, teks = st.session_state["modul"]
            st.download_button("Unduh modul ajar (.docx)", buat_docx(teks, judul), re.sub(r"[^A-Za-z0-9]+", "_", judul) + ".docx",
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
            with st.container(key="card_modul"):
                st.markdown(teks)
    with t2:
        f = st.file_uploader("Foto lembar jawaban siswa", type=["jpg", "jpeg", "png"])
        kunci = st.text_area("Kunci jawaban atau rubrik (opsional)", height=90)
        if f:
            img = Image.open(f).convert("RGB")
            img.thumbnail((1600, 1600))
            st.image(img, caption="Pratinjau lembar jawaban", width=380)
            if st.button("Koreksi lembar jawaban", type="primary"):
                with st.spinner("Membaca dan menilai tulisan tangan..."):
                    hasil = ai("Kamu guru yang teliti. 1) Transkripsikan tulisan tangan pada gambar. 2) Beri nilai 0-100 berdasarkan "
                               + (f"kunci atau rubrik berikut:\n{kunci}" if kunci.strip() else "ketepatan jawaban")
                               + ". 3) Jelaskan kesalahan per nomor. 4) Beri saran perbaikan singkat. Gunakan Bahasa Indonesia, format Markdown.", img)
                if hasil:
                    st.session_state["koreksi"] = hasil
        if st.session_state.get("koreksi"):
            with st.container(key="card_koreksi"):
                st.markdown(st.session_state["koreksi"])


def guru_admin(u):
    hero("Administrasi guru", "Kelola perangkat pembelajaran, rekap otomatis, dan jadwal mengajar Anda.")
    g = u["id"]
    opsi_mp = [m for m, gu in MP if gu == u["username"]] or [m for m, _ in MP]
    t1, t2, t3, t4 = st.tabs(["📁 Kelengkapan dokumen", "🤖 Susun dengan AI", "📊 Rekap otomatis", "🗓️ Jadwal & beban"])
    ada = qdf("SELECT jenis, nama_file, status, catatan, diperbarui FROM admin_guru WHERE guru_id=?", (g,))
    ada = ada[ada.jenis.isin(DOKUMEN_ADMIN)]
    with t1:
        lengkap, draf = int((ada.status == "Lengkap").sum()), int((ada.status == "Draf").sum())
        kpis([("Kelengkapan", f"{pct(lengkap, len(DOKUMEN_ADMIN))}%", f"{lengkap} dari {len(DOKUMEN_ADMIN)} dokumen", "g"),
              ("Berstatus draf", draf, "perlu dilengkapi", "o"),
              ("Belum ada", len(DOKUMEN_ADMIN) - lengkap - draf, "belum diunggah", "r")])
        info = ada.set_index("jenis")
        v = pd.DataFrame({"Dokumen": DOKUMEN_ADMIN,
                          "Status": [info.status.get(j, "Belum Ada") for j in DOKUMEN_ADMIN],
                          "Berkas": [info.nama_file.get(j) for j in DOKUMEN_ADMIN],
                          "Diperbarui": [info.diperbarui.get(j) for j in DOKUMEN_ADMIN],
                          "Catatan": [info.catatan.get(j) or None for j in DOKUMEN_ADMIN]})
        H(panel("Daftar perangkat pembelajaran", tb(v, ("Status",)), True))
        sec("Unggah atau perbarui dokumen")
        with st.form("f_admin", clear_on_submit=True):
            jenis = st.selectbox("Jenis dokumen", DOKUMEN_ADMIN)
            f = st.file_uploader("Berkas (PDF, DOCX, XLSX, PPTX, gambar, maksimal 5 MB)",
                                 type=["pdf", "docx", "xlsx", "pptx", "png", "jpg", "jpeg"])
            sts = st.radio("Status", ["Lengkap", "Draf"], horizontal=True)
            cat = st.text_input("Catatan (opsional)")
            if st.form_submit_button("Simpan dokumen", type="primary"):
                if f is None:
                    st.warning("Pilih berkas terlebih dahulu.")
                elif f.size > 5 * 1024 * 1024:
                    st.error("Ukuran berkas maksimal 5 MB.")
                else:
                    run("INSERT OR REPLACE INTO admin_guru(guru_id,jenis,nama_file,berkas,status,catatan,diperbarui) VALUES (?,?,?,?,?,?,?)",
                        (g, jenis, f.name, f.getvalue(), sts, cat.strip(), now().strftime("%Y-%m-%d %H:%M")))
                    flash(f'Dokumen "{jenis}" tersimpan.')
                    st.rerun()
        if not ada.empty:
            pil = st.selectbox("Pilih dokumen tersimpan", list(ada.jenis), key="adm_pilih")
            r = one("SELECT nama_file, berkas FROM admin_guru WHERE guru_id=? AND jenis=?", (g, pil))
            c1, c2 = st.columns(2)
            c1.download_button("Unduh dokumen", r[1], file_name=r[0], key="adm_unduh")
            if c2.button("Hapus dokumen", key="adm_hapus"):
                run("DELETE FROM admin_guru WHERE guru_id=? AND jenis=?", (g, pil))
                flash(f'Dokumen "{pil}" dihapus.')
                st.rerun()
    with t2:
        y = today().year
        ta0 = f"{y}/{y + 1}" if today().month >= 7 else f"{y - 1}/{y}"
        with st.form("f_adm_ai"):
            jenis = st.selectbox("Dokumen yang disusun", list(GENERATOR))
            a, b, c = st.columns(3)
            mapel = a.selectbox("Mata pelajaran", opsi_mp)
            kls = b.text_input("Kelas / fase", "VIII / Fase D")
            smt = c.selectbox("Semester", ["Ganjil", "Genap"])
            d, e = st.columns(2)
            ta = d.text_input("Tahun ajaran", ta0)
            jp = e.number_input("Jam pelajaran per minggu", 1, 12, 4)
            ket = st.text_area("Catatan tambahan (topik, KKM, dan sebagainya)", height=80)
            kirim = st.form_submit_button("Susun dokumen", type="primary")
        if kirim:
            with st.spinner("AI sedang menyusun dokumen..."):
                hasil = ai(f"""Kamu pakar administrasi guru di Indonesia (Kurikulum Merdeka). Susun {jenis} dalam Bahasa Indonesia,
format Markdown dengan tabel bila sesuai, rapi dan siap dicetak.
Sekolah: {NAMA_SEKOLAH}. Mata pelajaran: {mapel}. Kelas/Fase: {kls}. Semester: {smt}. Tahun ajaran: {ta}.
Jam pelajaran per minggu: {jp}. Nama guru: {u['nama']}.
Petunjuk isi dokumen: {GENERATOR[jenis]}
Catatan tambahan: {ket.strip() or '-'}""")
            if hasil:
                st.session_state["adm_hasil"] = (jenis, f"{jenis} {mapel} {kls} Semester {smt} {ta}", hasil)
        if st.session_state.get("adm_hasil"):
            jh, judul, teks = st.session_state["adm_hasil"]
            nama_f = re.sub(r"[^A-Za-z0-9]+", "_", judul) + ".docx"
            docx = buat_docx(teks, judul)
            c1, c2 = st.columns(2)
            c1.download_button("Unduh (.docx)", docx, nama_f,
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               type="primary", key="adm_dl_ai")
            if c2.button("Simpan ke kelengkapan dokumen", key="adm_simpan_ai"):
                run("INSERT OR REPLACE INTO admin_guru(guru_id,jenis,nama_file,berkas,status,catatan,diperbarui) VALUES (?,?,?,?,?,?,?)",
                    (g, jh, nama_f, docx, "Draf", "Disusun dengan AI, mohon dicek", now().strftime("%Y-%m-%d %H:%M")))
                flash(f'"{jh}" disimpan sebagai draf di kelengkapan dokumen.')
                st.rerun()
            with st.container(key="card_adm"):
                st.markdown(teks)
    with t3:
        c1, c2 = st.columns(2)
        kelas = c1.selectbox("Kelas", KELAS, key="adm_kelas")
        mapel = c2.selectbox("Mata pelajaran", opsi_mp, key="adm_mapel")
        h = qdf("""SELECT si.nis AS NIS, si.nama AS 'Nama siswa', COALESCE(SUM(x.status='Hadir'),0) AS Hadir,
                          COALESCE(SUM(x.status='Sakit'),0) AS Sakit, COALESCE(SUM(x.status='Izin'),0) AS Izin,
                          COALESCE(SUM(x.status='Alpa'),0) AS Alpa
                   FROM siswa si LEFT JOIN (SELECT p.siswa_id, p.status FROM presensi_siswa p JOIN sesi s ON s.id=p.sesi_id
                                            WHERE s.kelas=? AND s.mapel=?) x ON x.siswa_id=si.id
                   WHERE si.kelas=? GROUP BY si.id ORDER BY si.nama""", (kelas, mapel, kelas))
        h["Total"] = h[["Hadir", "Sakit", "Izin", "Alpa"]].sum(axis=1)
        h["% Hadir"] = [f"{pct(a, b)}%" for a, b in zip(h.Hadir, h.Total)]
        sec("Daftar hadir siswa")
        H(panel(f"{mapel}, kelas {kelas}", tb(h), True))
        st.download_button("Unduh daftar hadir (CSV)", h.to_csv(index=False).encode("utf-8-sig"),
                           f"daftar_hadir_{kelas}_{mapel}.csv", "text/csv", key="adm_csv_hadir")
        n = qdf("""SELECT si.nis AS NIS, si.nama AS 'Nama siswa', n.harian AS Harian, n.ujian AS Ujian FROM siswa si
                   LEFT JOIN nilai n ON n.siswa_id=si.id AND n.mapel=? WHERE si.kelas=? ORDER BY si.nama""", (mapel, kelas))
        n["Nilai akhir"] = (pd.to_numeric(n["Harian"]) * .4 + pd.to_numeric(n["Ujian"]) * .6).round(1)
        n["Huruf"] = n["Nilai akhir"].map(lambda x: "–" if pd.isna(x) else huruf(x))
        n["Ketuntasan"] = n["Nilai akhir"].map(lambda x: "–" if pd.isna(x) else ("Tuntas" if x >= KKM else "Remedial"))
        sec(f"Daftar nilai dan analisis (KKM {KKM})")
        nv = n["Nilai akhir"].dropna()
        if nv.empty:
            H('<div class="empty">Belum ada nilai untuk kelas dan mata pelajaran ini.</div>')
        else:
            kpis([("Rata-rata kelas", f"{nv.mean():.1f}", f"{len(nv)} siswa bernilai", ""),
                  ("Tertinggi", f"{nv.max():.1f}", "nilai akhir", "g"), ("Terendah", f"{nv.min():.1f}", "nilai akhir", "o"),
                  ("Perlu remedial", int((nv < KKM).sum()), f"di bawah KKM {KKM}", "r" if (nv < KKM).any() else "g")])
            H(panel("Daftar nilai", tb(n, ("Huruf", "Ketuntasan")), True))
            st.download_button("Unduh daftar nilai (CSV)", n.to_csv(index=False).encode("utf-8-sig"),
                               f"daftar_nilai_{kelas}_{mapel}.csv", "text/csv", key="adm_csv_nilai")
    with t4:
        rows = sorted((h_, s, k, m) for h_ in range(5) for k in KELAS for s, m, gu in jadwal(k, h_) if gu == u["username"])
        n_sesi = one("SELECT COUNT(*) FROM sesi WHERE guru_id=? AND tanggal LIKE ?", (g, today().isoformat()[:7] + "%"))[0]
        kpis([("Jam mengajar per pekan", len(rows), "jam pelajaran terjadwal", ""),
              ("Kelas diampu", len({r[2] for r in rows}), ", ".join(sorted({r[2] for r in rows})) or "belum ada", "t"),
              ("Jurnal terisi bulan ini", n_sesi, "pertemuan tercatat", "g")])
        v = pd.DataFrame([(HARI[a], b, c, d) for a, b, c, d in rows], columns=["Hari", "Jam", "Kelas", "Mata pelajaran"])
        H(panel("Jadwal mengajar mingguan", tb(v, kosong="Belum ada jadwal mengajar."), True))
        H(panel("Wali kelas", "".join(f'<div class="ann"><small>Kelas {k}</small><b>{esc(w)}</b></div>' for k, w in WALI.items())))


# ══════════════════════════ PORTAL SISWA ══════════════════════════
def siswa_home(u):
    sid = u["siswa_id"]
    nis, nama, kelas = one("SELECT nis, nama, kelas FROM siswa WHERE id=?", (sid,))
    hero("Beranda peserta didik", f"{sapa()}, {nama.split()[0]}. Kelas {kelas}, NIS {nis}.")
    tg, pr = daftar_tugas_siswa(sid), rekap_presensi(sid)
    hadir = int((pr.status == "Hadir").sum())
    belum, lewat = int((tg.status == "Belum Dikumpulkan").sum()), int((tg.status == "Melewati Deadline").sum())
    rata = one("SELECT AVG(harian*0.4+ujian*0.6) FROM nilai WHERE siswa_id=?", (sid,))[0]
    kpis([("Kehadiran", f"{pct(hadir, len(pr))}%", f"{hadir} dari {len(pr)} pertemuan", "g"),
          ("Tugas belum dikumpulkan", belum, "masih sebelum batas kumpul", "o" if belum else "g"),
          ("Tugas terlewat", lewat, "melewati batas kumpul", "r" if lewat else "g"),
          ("Rata-rata nilai akhir", f"{rata:.1f}" if rata else "–", "seluruh mata pelajaran", "t")])
    pg = peta_guru()
    att = qdf("""SELECT s.mapel, COUNT(*) AS t, SUM(p.status='Hadir') AS h FROM presensi_siswa p
                 JOIN sesi s ON s.id=p.sesi_id WHERE p.siswa_id=? GROUP BY s.mapel""", (sid,)).set_index("mapel")
    nl = qdf("SELECT mapel, harian*0.4+ujian*0.6 AS a FROM nilai WHERE siswa_id=?", (sid,)).set_index("mapel")["a"]
    kartu = ""
    for i, (m, _) in enumerate(MP):
        t_, h_ = (att.loc[m, "t"], att.loc[m, "h"]) if m in att.index else (0, 0)
        p_ = pct(h_, t_)
        n_ = f"{nl[m]:.1f}" if m in nl.index and pd.notna(nl[m]) else "–"
        kartu += (f'<div class="course"><div class="cband c{i % 4}">{IKON.get(m, "📘")}</div><div class="cbody"><b>{esc(m)}</b>'
                  f'<small>{esc(pg.get(m, "-"))}</small><div class="bar"><i style="width:{p_}%"></i></div>'
                  f'<small>Kehadiran {p_}%, nilai akhir {n_}</small></div></div>')
    sec("Mata pelajaran saya")
    H(f'<div class="courses">{kartu}</div>')
    a, b = st.columns(2)
    with a:
        up = tg[tg.status == "Belum Dikumpulkan"].sort_values("deadline").head(4)
        isi = "".join(f'<div class="ann"><small>Batas kumpul {fmt(r.deadline)}</small><b>{esc(r.judul)}</b><p>{esc(r.mapel)}</p></div>'
                      for r in up.itertuples())
        H(panel("Tugas yang perlu dikumpulkan", isi or '<div class="empty">Tidak ada tugas yang menunggu.</div>'))
        if st.button("📚 Buka halaman kumpul tugas", type="primary", key="ke_tugas"):
            st.session_state["menu"] = "📚 Tugas saya"
            st.rerun()
    with b:
        hari = today().weekday()
        jd = [(s, m, pg.get(m, "-")) for s, m, _ in jadwal(kelas, hari)] if hari < 5 else []
        H(panel("Jadwal hari ini", jadwal_html(jd, ["Jam", "Mata pelajaran", "Pengajar"]), True))
        H(panel("Pengumuman", pengumuman_html(2)))


def siswa_tugas(u):
    sid = u["siswa_id"]
    hero("Tugas saya", "Kumpulkan tugas sebelum batas waktu. Berkas yang belum dinilai masih bisa diganti.")
    tg = daftar_tugas_siswa(sid)
    if tg.empty:
        H('<div class="empty">Belum ada tugas dari guru.</div>')
        return
    for r in tg.itertuples():
        with st.container(key=f"card_{r.id}"):
            H(f'<div class="tugas-h"><b>{esc(r.judul)}</b>{badge(r.status)}</div>'
              f'<div class="tugas-m">{esc(r.mapel)}, batas kumpul {fmt(r.deadline)}</div><p>{esc(r.deskripsi)}</p>')
            sudah = pd.notna(r.waktu)
            if sudah:
                isi = f"Dikirim {esc(r.waktu)} ({esc(r.nama_file)})."
                if pd.notna(r.nilai):
                    isi += f" Nilai: <b>{int(r.nilai)}</b>. Catatan guru: {esc(r.umpan_balik)}"
                note("ok" if pd.notna(r.nilai) else "info", isi)
            if pd.isna(r.nilai):
                f = st.file_uploader("Ganti berkas" if sudah else "Unggah berkas tugas",
                                     type=["pdf", "docx", "png", "jpg", "jpeg"], key=f"up_{r.id}")
                if st.button("Kirim ulang" if sudah else "Kirim tugas", type="primary", key=f"kirim_{r.id}"):
                    if f is None:
                        st.warning("Pilih berkas terlebih dahulu.")
                    elif f.size > 5 * 1024 * 1024:
                        st.error("Ukuran berkas maksimal 5 MB.")
                    else:
                        run("""INSERT OR REPLACE INTO pengumpulan(tugas_id,siswa_id,waktu,nama_file,berkas,nilai,umpan_balik)
                               VALUES (?,?,?,?,?,NULL,NULL)""",
                            (int(r.id), sid, now().strftime("%Y-%m-%d %H:%M"), f.name, f.getvalue()))
                        flash(f'Tugas "{f.name}" berhasil dikirim.')
                        st.rerun()


def siswa_presensi(u):
    hero("Presensi saya", "Rekap kehadiran per mata pelajaran dan riwayat tiap pertemuan.")
    note("info", "Siswa tidak absen sendiri. Kehadiran dicatat oleh guru setiap selesai pelajaran, lalu tampil di halaman ini.")
    blok_presensi(u["siswa_id"])


def siswa_nilai(u):
    hero("Nilai & rapor", "Nilai harian, ujian, nilai akhir, dan nilai tugas.")
    blok_nilai(u["siswa_id"])


def siswa_tutor(u):
    hero("Tutor AI", "Tanyakan materi apa saja, dijelaskan dengan bahasa yang mudah dipahami.")
    riwayat = st.session_state.setdefault("chat", [])
    for m in riwayat:
        with st.chat_message(m["r"]):
            st.markdown(m["t"])
    if q := st.chat_input("Tulis pertanyaanmu di sini"):
        konteks = "\n".join(f"{'Siswa' if m['r'] == 'user' else 'Tutor'}: {m['t']}" for m in riwayat[-6:])
        riwayat.append({"r": "user", "t": q})
        with st.chat_message("user"):
            st.markdown(q)
        with st.chat_message("assistant"):
            with st.spinner("Sedang berpikir..."):
                jawab = ai("Kamu tutor yang ramah untuk siswa sekolah di Indonesia. Jelaskan dengan bahasa sederhana, beri contoh, "
                           f"dan akhiri dengan satu pertanyaan latihan singkat.\nRiwayat percakapan:\n{konteks}\nPertanyaan siswa: {q}")
            st.markdown(jawab or "Maaf, AI belum bisa dihubungi. Coba lagi nanti.")
        riwayat.append({"r": "assistant", "t": jawab or "Maaf, AI belum bisa dihubungi."})
    if riwayat and st.button("Hapus percakapan"):
        st.session_state["chat"] = []
        st.rerun()


# ══════════════════════════ PORTAL KEPALA SEKOLAH ══════════════════════════
GURU_HARI_SQL = """SELECT u.nama AS Guru, COALESCE(p.status,'Belum Presensi') AS Status, p.jam_masuk AS Masuk, p.jam_pulang AS Pulang
                   FROM users u LEFT JOIN presensi_guru p ON p.guru_id=u.id AND p.tanggal=?
                   WHERE u.role='guru' ORDER BY u.nama"""


def rentang_tanggal(kunci, hari=30):
    r = st.date_input("Rentang tanggal", (today() - timedelta(days=hari - 1), today()), format="DD/MM/YYYY", key=kunci)
    if not isinstance(r, (tuple, list)) or len(r) < 2:
        st.info("Pilih tanggal awal dan tanggal akhir.")
        return None
    return r[0].isoformat(), r[1].isoformat()


def kepsek_home(u):
    hero("Dashboard eksekutif", "Pantau kehadiran guru dan siswa, jurnal mengajar, dan capaian akademik.")
    ref, awal = hari_kerja(1)[0].isoformat(), hari_kerja(7)[-1].isoformat()
    g = qdf(GURU_HARI_SQL, (ref,))
    s = qdf("""SELECT s.tanggal AS Tanggal, COUNT(*) AS total, SUM(p.status='Hadir') AS hadir FROM presensi_siswa p
               JOIN sesi s ON s.id=p.sesi_id WHERE s.tanggal>=? GROUP BY s.tanggal ORDER BY s.tanggal""", (awal,))
    n_sesi = one("SELECT COUNT(*) FROM sesi WHERE tanggal LIKE ?", (today().isoformat()[:7] + "%",))[0]
    rata = one("SELECT AVG(harian*0.4+ujian*0.6) FROM nilai")[0]
    kpis([("Guru hadir", f"{int(g['Status'].isin(['Hadir', 'Terlambat']).sum())} dari {len(g)}", fmt(ref), "g"),
          ("Kehadiran siswa", f"{pct(s['hadir'].sum(), s['total'].sum())}%", "7 hari kerja terakhir", ""),
          ("Pertemuan bulan ini", n_sesi, "jurnal mengajar", "o"),
          ("Rata-rata nilai akhir", f"{rata:.1f}" if rata else "–", "seluruh mata pelajaran", "t")])
    a, b = st.columns([1.15, 1])
    with a:
        H(panel(f"Presensi guru, {fmt(ref)}", tb(g, ("Status",)), True))
        H(panel("Pengumuman", pengumuman_html()))
        with st.expander("Buat pengumuman baru"):
            with st.form("f_umum", clear_on_submit=True):
                j, i = st.text_input("Judul"), st.text_area("Isi pengumuman", height=90)
                if st.form_submit_button("Terbitkan pengumuman", type="primary"):
                    if j.strip() and i.strip():
                        run("INSERT INTO pengumuman(tanggal,judul,isi) VALUES (?,?,?)", (today().isoformat(), j.strip(), i.strip()))
                        flash("Pengumuman diterbitkan.")
                        st.rerun()
                    else:
                        st.warning("Judul dan isi pengumuman wajib diisi.")
    with b:
        sec("Tren kehadiran siswa (%)")
        if s.empty:
            H('<div class="empty">Belum ada data presensi siswa.</div>')
        else:
            s["Kehadiran (%)"] = (100 * s["hadir"] / s["total"]).round(1)
            st.line_chart(s.set_index("Tanggal")["Kehadiran (%)"], height=220, color="#1B4F9C")
        k = qdf("""SELECT si.kelas AS Kelas, ROUND(100.0*SUM(p.status='Hadir')/COUNT(*),1) AS Hadir FROM presensi_siswa p
                   JOIN siswa si ON si.id=p.siswa_id JOIN sesi s ON s.id=p.sesi_id WHERE s.tanggal>=? GROUP BY si.kelas""", (awal,))
        sec("Kehadiran siswa per kelas (%)")
        if not k.empty:
            st.bar_chart(k.set_index("Kelas")["Hadir"], height=200, color="#0F766E")


def kepsek_pguru(u):
    hero("Presensi guru", "Rekap kehadiran, keterlambatan, dan ketidakhadiran pengajar.")
    c1, c2 = st.columns(2)
    with c1:
        r = rentang_tanggal("r_guru")
    nama = qdf("SELECT nama FROM users WHERE role='guru' ORDER BY nama")["nama"].tolist()
    pilih = c2.selectbox("Guru", ["Semua guru"] + nama)
    if not r:
        return
    d = qdf("""SELECT p.tanggal, u.nama AS guru, p.status, p.jam_masuk, p.jam_pulang, p.keterangan
               FROM presensi_guru p JOIN users u ON u.id=p.guru_id WHERE p.tanggal BETWEEN ? AND ?
               ORDER BY p.tanggal DESC, u.nama""", r)
    if pilih != "Semua guru":
        d = d[d.guru == pilih]
    if d.empty:
        H('<div class="empty">Tidak ada data presensi pada rentang ini.</div>')
        return
    rk = d.pivot_table(index="guru", columns="status", values="tanggal", aggfunc="count", fill_value=0)
    rk = rk.reindex(columns=["Hadir", "Terlambat", "Sakit", "Izin", "Alpa"], fill_value=0)
    rk["Total"] = rk.sum(axis=1)
    rk["% Hadir"] = ((rk["Hadir"] + rk["Terlambat"]) * 100 / rk["Total"]).round().astype(int).astype(str) + "%"
    rk.columns.name, rk.index.name = None, "Guru"
    kpis([("Hadir tepat waktu", int((d.status == "Hadir").sum()), f"dari {len(d)} catatan", "g"),
          ("Terlambat", int((d.status == "Terlambat").sum()), f"presensi setelah {BATAS_MASUK}", "o"),
          ("Sakit atau izin", int(d.status.isin(["Sakit", "Izin"]).sum()), "dengan keterangan", ""),
          ("Alpa", int((d.status == "Alpa").sum()), "tanpa keterangan", "r")])
    t1, t2 = st.tabs(["Rekap per guru", "Detail harian"])
    with t1:
        H(panel("Rekap kehadiran guru", tb(rk.reset_index()), True))
    with t2:
        v = pd.DataFrame({"Tanggal": d.tanggal.map(fmt), "Guru": d.guru, "Status": d.status, "Masuk": d.jam_masuk,
                          "Pulang": d.jam_pulang, "Keterangan": d.keterangan.where(d.keterangan != "", None)})
        H(panel("Detail presensi guru", tb(v, ("Status",)), True))
    st.download_button("Unduh CSV", d.to_csv(index=False).encode("utf-8"), "presensi_guru.csv", "text/csv")


def kepsek_psiswa(u):
    hero("Presensi siswa", "Rekap kehadiran siswa per kelas, per siswa, dan per tanggal.")
    c1, c2 = st.columns(2)
    kelas = c1.selectbox("Kelas", ["Semua kelas"] + KELAS)
    with c2:
        r = rentang_tanggal("r_siswa")
    if not r:
        return
    d = qdf("""SELECT s.tanggal, si.kelas, si.nis, si.nama, s.mapel, p.status FROM presensi_siswa p
               JOIN sesi s ON s.id=p.sesi_id JOIN siswa si ON si.id=p.siswa_id
               WHERE s.tanggal BETWEEN ? AND ?""", r)
    if kelas != "Semua kelas":
        d = d[d.kelas == kelas]
    if d.empty:
        H('<div class="empty">Tidak ada data presensi pada rentang ini.</div>')
        return
    rk = d.pivot_table(index=["kelas", "nis", "nama"], columns="status", values="tanggal", aggfunc="count", fill_value=0)
    rk = rk.reindex(columns=STATUS_HADIR, fill_value=0)
    rk["Total"] = rk.sum(axis=1)
    persen = (100 * rk["Hadir"] / rk["Total"]).round().astype(int)
    rk["% Hadir"] = persen.astype(str) + "%"
    rk["Keterangan"] = persen.map(lambda x: "Baik" if x >= 90 else "Cukup" if x >= 80 else "Perlu Perhatian")
    rk = rk.reset_index()
    rk.columns.name = None
    rk = rk.rename(columns={"kelas": "Kelas", "nis": "NIS", "nama": "Nama siswa"})
    tot = len(d)
    kpis([("Kehadiran", f"{pct((d.status == 'Hadir').sum(), tot)}%", f"{tot} catatan presensi", "g"),
          ("Sakit", int((d.status == "Sakit").sum()), "catatan", "o"), ("Izin", int((d.status == "Izin").sum()), "catatan", ""),
          ("Alpa", int((d.status == "Alpa").sum()), f"{int((rk['Keterangan'] == 'Perlu Perhatian').sum())} siswa perlu perhatian", "r")])
    t1, t2 = st.tabs(["Rekap per siswa", "Per kelas dan tanggal"])
    with t1:
        H(panel("Rekap kehadiran siswa", tb(rk, ("Keterangan",)), True))
    with t2:
        h = d.assign(h=(d.status == "Hadir") * 100).groupby(["tanggal", "kelas"])["h"].mean().round(1).unstack("kelas")
        st.line_chart(h, height=240)
        h = h.sort_index(ascending=False)
        h.index = h.index.map(fmt)
        h.index.name, h.columns.name = "Tanggal", None
        H(panel("Persentase hadir per tanggal", tb(h.reset_index()), True))
    st.download_button("Unduh CSV", d.to_csv(index=False).encode("utf-8"), "presensi_siswa.csv", "text/csv")


def kepsek_jurnal(u):
    hero("Jurnal mengajar", "Semua pertemuan yang dicatat guru beserta rekap kehadiran siswa.")
    c1, c2 = st.columns(2)
    nama = qdf("SELECT nama FROM users WHERE role='guru' ORDER BY nama")["nama"].tolist()
    pilih = c1.selectbox("Guru", ["Semua guru"] + nama)
    with c2:
        r = rentang_tanggal("r_jurnal")
    if not r:
        return
    d = qdf("""SELECT s.tanggal AS Tanggal, u.nama AS Guru, s.kelas AS Kelas, s.mapel AS 'Mata pelajaran', s.materi AS Materi,
                      COALESCE(SUM(p.status='Hadir'),0) AS Hadir, COALESCE(SUM(p.status='Sakit'),0) AS Sakit,
                      COALESCE(SUM(p.status='Izin'),0) AS Izin, COALESCE(SUM(p.status='Alpa'),0) AS Alpa
               FROM sesi s JOIN users u ON u.id=s.guru_id LEFT JOIN presensi_siswa p ON p.sesi_id=s.id
               WHERE s.tanggal BETWEEN ? AND ? GROUP BY s.id ORDER BY s.tanggal DESC, s.id DESC""", r)
    if pilih != "Semua guru":
        d = d[d.Guru == pilih]
    kpis([("Total pertemuan", len(d), "pada rentang terpilih", ""), ("Guru aktif mengajar", int(d.Guru.nunique()), "mencatat jurnal", "t")])
    d["Tanggal"] = d["Tanggal"].map(fmt)
    H(panel("Daftar jurnal mengajar", tb(d), True))


def kepsek_nilai(u):
    hero("Nilai & tugas", "Capaian akademik per kelas dan pemantauan pengumpulan tugas.")
    t1, t2 = st.tabs(["Rata-rata nilai", "Monitoring tugas"])
    with t1:
        d = qdf("""SELECT si.kelas AS Kelas, n.mapel AS 'Mata pelajaran', ROUND(AVG(n.harian),1) AS Harian,
                          ROUND(AVG(n.ujian),1) AS Ujian, ROUND(AVG(n.harian*0.4+n.ujian*0.6),1) AS 'Nilai akhir'
                   FROM nilai n JOIN siswa si ON si.id=n.siswa_id GROUP BY si.kelas, n.mapel ORDER BY si.kelas, n.mapel""")
        H(panel("Rata-rata nilai per kelas dan mata pelajaran", tb(d), True))
    with t2:
        d = qdf("""SELECT t.judul AS Tugas, t.mapel AS 'Mata pelajaran', t.kelas AS Kelas, t.deadline AS 'Batas kumpul',
                          (SELECT COUNT(*) FROM siswa WHERE kelas=t.kelas) AS total, COUNT(p.id) AS terkumpul,
                          SUM(p.nilai IS NOT NULL) AS Dinilai, ROUND(AVG(p.nilai),1) AS 'Rata-rata nilai'
                   FROM tugas t LEFT JOIN pengumpulan p ON p.tugas_id=t.id GROUP BY t.id ORDER BY t.deadline DESC""")
        d["Pengumpulan"] = [f"{a} dari {b} ({pct(a, b)}%)" for a, b in zip(d.terkumpul, d.total)]
        d["Batas kumpul"] = d["Batas kumpul"].map(fmt)
        H(panel("Pengumpulan tugas", tb(d.drop(columns=["total", "terkumpul"])), True))


def kepsek_admin(u):
    hero("Administrasi guru", "Pantau kelengkapan perangkat pembelajaran setiap guru.")
    gr = qdf("SELECT id, nama FROM users WHERE role='guru' ORDER BY nama")
    ad = qdf("SELECT guru_id, jenis, status, diperbarui FROM admin_guru")
    ad = ad[ad.jenis.isin(DOKUMEN_ADMIN)]
    N, rows = len(DOKUMEN_ADMIN), []
    for gid, nm in zip(gr.id, gr.nama):
        s = ad[ad.guru_id == gid]
        lk, dr = int((s.status == "Lengkap").sum()), int((s.status == "Draf").sum())
        rows.append((nm, lk, dr, N - lk - dr, f"{pct(lk, N)}%"))
    rk = pd.DataFrame(rows, columns=["Guru", "Lengkap", "Draf", "Belum ada", "Kelengkapan"])
    tot = int((ad.status == "Lengkap").sum())
    kpis([("Kelengkapan rata-rata", f"{pct(tot, N * max(len(gr), 1))}%", f"{len(gr)} guru, {N} dokumen per guru", "g"),
          ("Dokumen lengkap", tot, "seluruh guru", ""),
          ("Masih draf", int((ad.status == "Draf").sum()), "perlu dilengkapi", "o")])
    H(panel("Rekap per guru", tb(rk), True))
    mat = pd.DataFrame({"Dokumen": DOKUMEN_ADMIN})
    tanda = {"Lengkap": "✔ Lengkap", "Draf": "… Draf"}
    for gid, nm in zip(gr.id, gr.nama):
        peta = dict(ad[ad.guru_id == gid][["jenis", "status"]].values.tolist())
        mat[nm] = [tanda.get(peta.get(j), "– Belum") for j in DOKUMEN_ADMIN]
    H(panel("Rincian per dokumen", tb(mat), True))


def kepsek_bukti(u):
    hero("Bukti presensi guru", "Lokasi GPS, foto selfie, dan hasil pencocokan wajah pada presensi guru.")
    t1, t2 = st.tabs(["📍 Bukti presensi", "⚙️ Pengaturan lokasi & wajah"])
    with t1:
        tgl = st.date_input("Tanggal", today(), format="DD/MM/YYYY", key="bukti_tgl")
        d = qdf("""SELECT b.id, u.nama AS Guru, b.jenis AS Jenis, b.waktu AS Waktu, b.lat, b.lon, b.akurasi,
                          b.jarak AS Jarak, b.wajah AS Wajah, b.catatan AS Catatan
                   FROM bukti_presensi b JOIN users u ON u.id=b.guru_id WHERE b.tanggal=?
                   ORDER BY u.nama, b.jenis""", (tgl.isoformat(),))
        if d.empty:
            H('<div class="empty">Belum ada bukti presensi pada tanggal ini.</div>')
        else:
            kpis([("Catatan presensi", len(d), fmt(tgl), ""),
                  ("Wajah cocok", int((d.Wajah == "Cocok").sum()), "terverifikasi", "g"),
                  ("Perlu verifikasi", int((d.Wajah == "Perlu Verifikasi").sum()), "cek manual", "o")])
            v = pd.DataFrame({"Guru": d.Guru, "Jenis": d.Jenis.str.capitalize(), "Waktu": d.Waktu,
                              "Jarak dari sekolah": d.Jarak.map(lambda x: "–" if pd.isna(x) else f"{int(x)} m"),
                              "Akurasi GPS": d.akurasi.map(lambda x: "–" if pd.isna(x) else f"{x:.0f} m"), "Wajah": d.Wajah})
            H(panel("Bukti presensi", tb(v, ("Wajah",)), True))
            sec("Peta lokasi presensi")
            st.map(d.rename(columns={"lat": "latitude", "lon": "longitude"})[["latitude", "longitude"]])
            label = {int(i): f"{g}, {j.capitalize()}, {w}" for i, g, j, w in zip(d.id, d.Guru, d.Jenis, d.Waktu)}
            pil = st.selectbox("Lihat rincian", list(label), format_func=label.get)
            r = d[d.id == pil].iloc[0]
            foto = one("SELECT foto FROM bukti_presensi WHERE id=?", (int(pil),))[0]
            c1, c2 = st.columns([1, 1.5])
            with c1:
                st.image(foto, caption=f"Foto selfie, {label[int(pil)]}", width=240)
                note("info", f"Hasil pencocokan: {badge(r.Wajah)}<br>{esc(r.Catatan)}")
            with c2:
                components.html(f'<iframe src="https://maps.google.com/maps?q={r.lat},{r.lon}&z=18&output=embed" '
                                'width="100%" height="300" style="border:0" loading="lazy"></iframe>', height=310)
                st.link_button("Buka di Google Maps", link_maps(r.lat, r.lon))
    with t2:
        titik = titik_sekolah()
        if titik:
            note("ok", f"Titik sekolah: {titik[0]:.6f}, {titik[1]:.6f}, radius {titik[2]} m.")
        else:
            note("warn", "Titik sekolah belum diatur. Selama belum diatur, lokasi tetap dicatat tetapi radius tidak diperiksa.")
        sec("Atur titik sekolah")
        st.caption("Berdirilah di halaman sekolah lalu ambil lokasi, atau isi koordinat secara manual.")
        lat0, lon0 = (titik[0], titik[1]) if titik else (0.0, 0.0)
        if streamlit_geolocation is not None:
            loc = _geo("kepsek")
            if isinstance(loc, dict) and loc.get("latitude") is not None:
                lat0, lon0 = float(loc["latitude"]), float(loc["longitude"])
                note("info", f"Lokasi Anda saat ini: {lat0:.6f}, {lon0:.6f}. Tekan Simpan untuk menjadikannya titik sekolah.")
        with st.form("f_titik"):
            a, b, c = st.columns(3)
            la = a.number_input("Lintang (latitude)", -90.0, 90.0, lat0, format="%.6f", key=f"la_{lat0:.6f}")
            lo = b.number_input("Bujur (longitude)", -180.0, 180.0, lon0, format="%.6f", key=f"lo_{lon0:.6f}")
            rd = c.number_input("Radius (meter)", 20, 2000, titik[2] if titik else RADIUS_DEFAULT, step=10)
            if st.form_submit_button("Simpan titik sekolah", type="primary"):
                if la == 0.0 and lo == 0.0:
                    st.warning("Koordinat belum diisi.")
                else:
                    set_set("lat", la)
                    set_set("lon", lo)
                    set_set("radius", int(rd))
                    flash("Titik sekolah tersimpan.")
                    st.rerun()
        sec("Wajah terdaftar")
        w = qdf("""SELECT u.id, u.nama, w.diperbarui FROM users u LEFT JOIN wajah_guru w ON w.guru_id=u.id
                   WHERE u.role='guru' ORDER BY u.nama""")
        v = pd.DataFrame({"Guru": w.nama, "Status": ["Terdaftar" if pd.notna(x) and x else "Belum Ada" for x in w.diperbarui],
                          "Terakhir diperbarui": w.diperbarui})
        H(panel("Pendaftaran wajah guru", tb(v, ("Status",)), True))
        reg = w[w.diperbarui.notna()]
        if not reg.empty:
            nm = dict(zip(reg.id.astype(int), reg.nama))
            pil = st.selectbox("Reset pendaftaran wajah", list(nm), format_func=nm.get, key="reset_wajah")
            if st.button("Hapus pendaftaran wajah", key="btn_reset_wajah"):
                run("DELETE FROM wajah_guru WHERE guru_id=?", (int(pil),))
                flash(f"Pendaftaran wajah {nm[pil]} dihapus. Guru dapat mendaftar ulang.")
                st.rerun()


# ══════════════════════════ PORTAL ORANG TUA ══════════════════════════
def ortu_anak(u):
    if not u["siswa_id"]:
        H('<div class="empty">Akun ini belum dihubungkan dengan data siswa. Hubungi admin sekolah.</div>')
        return None
    return u["siswa_id"]


def ortu_home(u):
    sid = u["siswa_id"]
    if not sid:
        hero("Beranda orang tua")
        ortu_anak(u)
        return
    nis, nama, kelas = one("SELECT nis, nama, kelas FROM siswa WHERE id=?", (sid,))
    hero("Beranda orang tua", f"{sapa()}, {u['nama']}. Pantau perkembangan belajar {nama.split()[0]} di sini.")
    H(f'<div class="profile"><div class="av big">{esc(initials(nama))}</div><div><b>{esc(nama)}</b>'
      f'<span>NIS {esc(nis)}, kelas {esc(kelas)}</span><span>Wali kelas: {esc(WALI.get(kelas, "-"))}</span></div></div>')
    tg, pr = daftar_tugas_siswa(sid), rekap_presensi(sid)
    hadir, alpa = int((pr.status == "Hadir").sum()), int((pr.status == "Alpa").sum())
    belum = tg[tg.status.isin(["Belum Dikumpulkan", "Melewati Deadline"])].sort_values("deadline")
    rata = one("SELECT AVG(harian*0.4+ujian*0.6) FROM nilai WHERE siswa_id=?", (sid,))[0]
    kpis([("Kehadiran", f"{pct(hadir, len(pr))}%", f"{hadir} dari {len(pr)} pertemuan", "g"),
          ("Tugas dikumpulkan", f"{int(tg.waktu.notna().sum())} dari {len(tg)}", "tugas kelas ini", ""),
          ("Belum dikumpulkan", len(belum), "perlu ditindaklanjuti", "o" if len(belum) else "g"),
          ("Rata-rata nilai akhir", f"{rata:.1f}" if rata else "–", "seluruh mata pelajaran", "t")])
    if alpa:
        note("bad", f"Tercatat alpa {alpa} kali. Mohon konfirmasi ke wali kelas bila ada kendala.")
    for r in belum.head(5).itertuples():
        sisa = (date.fromisoformat(r.deadline) - today()).days
        if sisa >= 0:
            note("warn", f"Tugas <b>{esc(r.judul)}</b> ({esc(r.mapel)}) belum dikumpulkan. Batas kumpul {fmt(r.deadline)}, sisa {sisa} hari.")
        else:
            note("bad", f"Tugas <b>{esc(r.judul)}</b> ({esc(r.mapel)}) melewati batas kumpul sejak {fmt(r.deadline)}.")
    a, b = st.columns([1.2, 1])
    with a:
        v = pd.DataFrame({"Tanggal": pr.tanggal.map(fmt), "Mata pelajaran": pr.mapel, "Status": pr.status}).head(6)
        H(panel("Presensi terbaru", tb(v, ("Status",), "Belum ada data presensi."), True))
    with b:
        H(panel("Pengumuman sekolah", pengumuman_html(2)))


def ortu_presensi(u):
    hero("Presensi anak", "Kehadiran anak pada setiap pertemuan dan rekap per mata pelajaran.")
    note("info", "Kehadiran anak dicatat oleh guru setiap selesai pelajaran. Anda cukup melihat hasilnya di halaman ini.")
    if ortu_anak(u):
        blok_presensi(u["siswa_id"])


def ortu_tugas(u):
    hero("Tugas & pengumpulan", "Daftar tugas anak, status pengumpulan, nilai, dan catatan guru.")
    if ortu_anak(u):
        tabel_tugas_anak(u["siswa_id"])


def ortu_nilai(u):
    hero("Nilai anak", "Hasil belajar anak per mata pelajaran.")
    if ortu_anak(u):
        blok_nilai(u["siswa_id"])


# ══════════════════════════ ENTRY POINT ══════════════════════════
PAGES = {
    "guru": {"🏠 Beranda": guru_home, "🕒 Presensi & jurnal": guru_presensi, "📤 Tugas & penilaian": guru_tugas,
             "📊 Buku nilai": guru_nilai, "🗂️ Administrasi guru": guru_admin, "✨ Asisten AI": guru_ai},
    "siswa": {"🏠 Beranda": siswa_home, "📚 Tugas saya": siswa_tugas, "🗓️ Presensi saya": siswa_presensi,
              "🎓 Nilai & rapor": siswa_nilai, "🤖 Tutor AI": siswa_tutor},
    "kepsek": {"🏠 Dashboard": kepsek_home, "🕘 Presensi guru": kepsek_pguru, "👥 Presensi siswa": kepsek_psiswa,
               "📖 Jurnal mengajar": kepsek_jurnal, "📈 Nilai & tugas": kepsek_nilai, "🗂️ Administrasi guru": kepsek_admin, "📍 Bukti presensi": kepsek_bukti},
    "ortu": {"🏠 Beranda": ortu_home, "🗓️ Presensi anak": ortu_presensi, "📤 Tugas & pengumpulan": ortu_tugas,
             "🎓 Nilai anak": ortu_nilai},
}


def navbar_atas(u, aktif):
    """Menu di halaman utama, supaya tetap terlihat di HP saat sidebar tertutup."""
    menu = list(PAGES[u["role"]])
    with st.container(key="topnav"):
        pilih = st.radio("Menu", menu, index=menu.index(aktif), horizontal=True,
                         label_visibility="collapsed", key=f"topnav_{aktif}")
    if pilih != aktif:
        st.session_state["menu"] = pilih
        st.rerun()


def main():
    pasang_css()
    pastikan_skema()
    init_db()
    u = st.session_state.get("user")
    if not u:
        halaman_login()
        return
    aktif = sidebar(u)
    navbar_atas(u, aktif)
    tampil_flash()
    PAGES[u["role"]][aktif](u)


main()
