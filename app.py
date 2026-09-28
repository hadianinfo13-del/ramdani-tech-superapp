import streamlit as st
import google.generativeai as genai
from docx import Document
from gtts import gTTS
import pandas as pd
import numpy as np
from PIL import Image
import io
import time
import sqlite3
import os

# Konfigurasi Halaman
st.set_page_config(page_title="Asisten Guru - Super App EdTech", layout="wide", initial_sidebar_state="expanded")

# AMBIL API KEY DARI SECRETS
API_KEY = os.environ.get("GEMINI_API_KEY") 
if API_KEY:
    genai.configure(api_key=API_KEY)
else:
    st.error("⚠️ GEMINI_API_KEY belum diatur di Secrets Streamlit Cloud!")

# SISTEM DATABASE
@st.cache_resource
def init_db():
    conn = sqlite3.connect('sekolah.db')
    cursor = conn.cursor()
    cursor.execute('CREATE TABLE IF NOT EXISTS nilai_siswa (nama TEXT, harian INTEGER, ujian INTEGER)')
    cursor.execute('CREATE TABLE IF NOT EXISTS users (username TEXT, password TEXT, role TEXT, nama_asli TEXT)')
    
    cursor.execute("SELECT COUNT(*) FROM nilai_siswa")
    if cursor.fetchone()[0] == 0:
        cursor.executemany("INSERT INTO nilai_siswa VALUES (?, ?, ?)", [('Budi', 80, 75), ('Siti', 95, 90), ('Andi', 60, 65)])
    
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        users = [
            ('guru', 'guru123', 'guru', 'Bpk. Ramdani'),
            ('siswa', 'siswa123', 'siswa', 'Budi'),
            ('kepsek', 'kepsek123', 'kepsek', 'Ibu Kepsek'),
            ('ortu', 'ortu123', 'ortu', 'Ortu Budi')
        ]
        cursor.executemany("INSERT INTO users VALUES (?, ?, ?, ?)", users)
    
    conn.commit()
    conn.close()

init_db()

@st.cache_data(ttl=5)
def load_data_nilai():
    conn = sqlite3.connect('sekolah.db')
    df = pd.read_sql_query("SELECT nama as 'Nama Siswa', harian as 'Harian', ujian as 'Ujian' FROM nilai_siswa", conn)
    conn.close()
    return df

def simpan_data_nilai(df):
    conn = sqlite3.connect('sekolah.db')
    df_sql = df.rename(columns={'Nama Siswa': 'nama', 'Harian': 'harian', 'Ujian': 'ujian'})
    df_sql.to_sql('nilai_siswa', conn, if_exists='replace', index=False)
    conn.close()
    st.cache_data.clear()

# FUNGSIONALITAS MODEL GEMINI AUTOMATIC FALLBACK + DYNAMIC DISCOVERY
def panggil_gemini(prompt, gambar=None):
    # Opsi nama model prioritas (Generasi Terbaru hingga Legacy)
    kandidat_model = [
        "gemini-3.8-flash",
        "gemini-3.5-flash",
        "gemini-2.5-flash",
        "gemini-1.5-flash",
        "models/gemini-3.8-flash",
        "models/gemini-3.5-flash"
    ]
    
    # 1. Coba daftar kandidat utama lebih dulu
    for nama_model in kandidat_model:
        try:
            m = genai.GenerativeModel(nama_model)
            res = m.generate_content([prompt, gambar] if gambar else prompt)
            return res.text
        except Exception:
            continue

    # 2. Jika kandidat gagal, temukan model aktif secara otomatis dari API Google
    try:
        available_models = [
            m.name for m in genai.list_models() 
            if 'generateContent' in m.supported_generation_methods
        ]
        for nama_model in available_models:
            try:
                m = genai.GenerativeModel(nama_model)
                res = m.generate_content([prompt, gambar] if gambar else prompt)
                return res.text
            except Exception:
                continue
    except Exception as e:
        st.error(f"Gagal mengambil daftar model otomatis: {str(e)}")

    st.error("Semua percobaan panggillan model Gemini gagal. Mohon periksa status API Key Anda.")
    return None

# SISTEM ANTI-SPAM
if 'ai_usage' not in st.session_state:
    st.session_state.ai_usage = 0
if 'ai_last' not in st.session_state:
    st.session_state.ai_last = 0

def cek_izin_ai():
    if st.session_state.ai_usage >= 15:
        st.error("🛑 Kuota AI habis untuk sesi ini. Refresh browser Anda.")
        return False
    if (time.time() - st.session_state.ai_last) < 5:
        st.warning("⏳ Mesin AI sedang pendinginan. Tunggu beberapa detik lagi.")
        return False
    return True

def catat_penggunaan_ai():
    st.session_state.ai_usage += 1
    st.session_state.ai_last = time.time()

# EXPORT DOCX
def buat_file_word(teks, judul):
    doc = Document()
    doc.add_heading(judul, 0)
    doc.add_paragraph(teks)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# AUTENTIKASI
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.nama_user = None

def login(username, password):
    conn = sqlite3.connect('sekolah.db')
    cursor = conn.cursor()
    cursor.execute("SELECT role, nama_asli FROM users WHERE username=? AND password=?", (username, password))
    user = cursor.fetchone()
    conn.close()
    if user:
        st.session_state.logged_in = True
        st.session_state.role = user[0]
        st.session_state.nama_user = user[1]
        st.rerun()
    else:
        st.error("❌ Username atau Password salah!")

def logout():
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.nama_user = None
    st.rerun()

# ------------------------------------------
# HALAMAN LOGIN
# ------------------------------------------
if not st.session_state.logged_in:
    st.title("🎓 Portal Login Asisten Guru")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("### Silakan Masuk")
        with st.form("form_login"):
            user_input = st.text_input("Username")
            pass_input = st.text_input("Password", type="password")
            submit = st.form_submit_button("Masuk", use_container_width=True)
            if submit:
                login(user_input, pass_input)
        
        st.info("**Akun Demo:**\n- Guru: `guru` / `guru123`\n- Siswa: `siswa` / `siswa123`\n- Kepsek: `kepsek` / `kepsek123`\n- Ortu: `ortu` / `ortu123`")

# ------------------------------------------
# PORTAL GURU
# ------------------------------------------
elif st.session_state.role == "guru":
    st.sidebar.title(f"👨‍🏫 Halo, {st.session_state.nama_user}")
    st.sidebar.button("🚪 Keluar", on_click=logout)
    
    st.title("👨‍🏫 Asisten Guru Co-Pilot")
    tab1, tab2, tab3, tab4 = st.tabs(["📄 Modul Ajar", "📚 Pembuat Soal", "📸 Scan Koreksi AI", "📊 Buku Nilai"])

    with tab1:
        st.subheader("Generator Administrasi Mengajar")
        mapel = st.text_input("Materi / Topik:", "Sistem Pencernaan Manusia")
        if st.button("Generate RPP") and mapel and cek_izin_ai():
            with st.spinner("Merakit RPP dengan Asisten Guru..."):
                prompt = f"Anda adalah Asisten Guru. Buatkan RPP Kurikulum Merdeka ringkas untuk topik {mapel}."
                hasil = panggil_gemini(prompt)
                if hasil:
                    st.success("RPP Berhasil Dibuat!")
                    st.write(hasil)
                    st.download_button("📥 Unduh RPP (.docx)", buat_file_word(hasil, "RPP"), f"RPP_{mapel}.docx")
                    catat_penggunaan_ai()

    with tab2:
        st.subheader("Pembuat Soal Otomatis")
        topik_soal = st.text_input("Topik Soal:", "Hukum Newton")
        tingkat = st.selectbox("Tingkat Kesulitan:", ["Mudah", "Sedang", "HOTS (Sulit)"])
        if st.button("Buat 5 Soal") and cek_izin_ai():
            with st.spinner("Membuat soal..."):
                prompt = f"Anda adalah Asisten Guru. Buatkan 5 soal pilihan ganda tentang {topik_soal} dengan tingkat {tingkat}, lengkap dengan kunci jawaban."
                hasil = panggil_gemini(prompt)
                if hasil:
                    st.write(hasil)
                    catat_penggunaan_ai()

    with tab3:
        st.subheader("📸 Koreksi Tugas dengan AI Vision")
        st.write("Unggah foto jawaban siswa. AI akan membaca dan mengevaluasinya.")
        file_foto = st.file_uploader("Upload Foto", type=["jpg", "png", "jpeg"])
        if file_foto is not None:
            img = Image.open(file_foto)
            st.image(img, caption="Foto Tugas Siswa", width=300)
            if st.button("Koreksi Gambar Ini") and cek_izin_ai():
                with st.spinner("Menganalisis tulisan..."):
                    prompt = "Anda adalah Asisten Guru. Tolong transkrip tulisan di gambar ini, lalu beri nilai 1-100 dan umpan balik singkat."
                    hasil = panggil_gemini(prompt, gambar=img)
                    if hasil:
                        st.write(hasil)
                        catat_penggunaan_ai()

    with tab4:
        st.subheader("Buku Nilai Permanen")
        df_edit = st.data_editor(load_data_nilai(), num_rows="dynamic", use_container_width=True)
        if st.button("💾 Simpan Perubahan", type="primary"):
            simpan_data_nilai(df_edit)
            st.success("Tersimpan ke Database!")

# ------------------------------------------
# PORTAL SISWA
# ------------------------------------------
elif st.session_state.role == "siswa":
    st.sidebar.title(f"👨‍🎓 Halo, {st.session_state.nama_user}")
    st.sidebar.button("🚪 Keluar", on_click=logout)
    
    st.title("👋 Ruang Belajar - Asisten Guru")
    st.subheader("🤖 Tanya AI Tutor")
    tanya = st.text_input("Ada materi yang belum kamu pahami?")
    if st.button("Tanya") and tanya and cek_izin_ai():
        with st.spinner("Tutor berpikir..."):
            prompt = f"Anda adalah Asisten Guru. Jelaskan materi berikut kepada siswa SMP dengan bahasa ramah dan mudah dipahami: {tanya}"
            hasil = panggil_gemini(prompt)
            if hasil:
                st.info(hasil)
                catat_penggunaan_ai()

# ------------------------------------------
# PORTAL KEPALA SEKOLAH
# ------------------------------------------
elif st.session_state.role == "kepsek":
    st.sidebar.title(f"👔 {st.session_state.nama_user}")
    st.sidebar.button("🚪 Keluar", on_click=logout)
    
    st.title("📊 Dasbor Kepala Sekolah (Real-Time)")
    conn = sqlite3.connect('sekolah.db')
    df_kepsek = pd.read_sql_query("SELECT * FROM nilai_siswa", conn)
    conn.close()

    total_siswa = len(df_kepsek)
    rata_harian = df_kepsek['harian'].mean() if total_siswa > 0 else 0
    rata_ujian = df_kepsek['ujian'].mean() if total_siswa > 0 else 0

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Siswa", f"{total_siswa} Anak")
    col2.metric("Rata-rata Nilai Harian", f"{rata_harian:.1f}")
    col3.metric("Rata-rata Ujian", f"{rata_ujian:.1f}")
    
    st.divider()
    st.write("📈 Grafik Distribusi Nilai Harian")
    if total_siswa > 0:
        st.bar_chart(df_kepsek.set_index('nama')['harian'])

# ------------------------------------------
# PORTAL ORANG TUA
# ------------------------------------------
elif st.session_state.role == "ortu":
    st.sidebar.title(f"👨‍👩‍👧 {st.session_state.nama_user}")
    st.sidebar.button("🚪 Keluar", on_click=logout)
    
    st.title("👨‍👩‍👧 Pantau Nilai Anak")
    nama_anak = st.text_input("🔍 Masukkan Nama Anak (contoh: Budi):")
    
    if st.button("Cari Data"):
        conn = sqlite3.connect('sekolah.db')
        df_ortu = pd.read_sql_query(f"SELECT * FROM nilai_siswa WHERE nama='{nama_anak}'", conn)
        conn.close()

        if not df_ortu.empty:
            nilai_h = df_ortu.iloc[0]['harian']
            nilai_u = df_ortu.iloc[0]['ujian']
            st.success(f"Data ditemukan untuk: **{nama_anak}**")
            col1, col2 = st.columns(2)
            col1.metric("Nilai Harian", nilai_h)
            col2.metric("Nilai Ujian", nilai_u)
            
            if nilai_h >= 75:
                st.info("Pesan Wali Kelas: Nilai anak Anda sudah baik, pertahankan!")
            else:
                st.warning("Pesan Wali Kelas: Anak Anda butuh bimbingan belajar tambahan di rumah.")
        else:
            st.error("Data tidak ditemukan. Pastikan ejaan nama sesuai buku absen.")
