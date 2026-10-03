import streamlit as st
from datetime import datetime
import math

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="LMS SMP Negeri 1 Cijambe",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- FUNGSI UTILITAS & SIMULASI DATABASE ---
def now():
    return datetime.now()

def jarak_m(lat1, lon1, lat2, lon2):
    # Simulasi perhitungan jarak Haversine sederhana dalam meter
    R = 6371000  # Radius bumi dalam meter
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda/2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

def kecilkan(foto):
    # Fungsi placeholder untuk kompresi/pengecilan ukuran foto
    return foto

def bandingkan_wajah(ref_foto, input_foto):
    # Simulasi pencocokan AI wajah
    return "Cocok", "Akurasi 95%"

def one(query, params=()):
    # Simulasi pengambilan 1 baris data dari database
    return ("dummy_foto_base64",)

def run(query, params=()):
    # Simulasi eksekusi query database (INSERT/UPDATE)
    pass

# --- INISIALISASI SESSION STATE ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "role" not in st.session_state:
    st.session_state.role = None
if "user_name" not in st.session_state:
    st.session_state.user_name = ""
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "Beranda"

# --- HALAMAN UTAMA / LOGIN ---
def login_page():
    st.title("🏫 Login LMS SMP Negeri 1 Cijambe")
    with st.form("login_form"):
        username = st.text_input("Username / NIP / NIS")
        password = st.text_input("Password", type="password")
        role_choice = st.selectbox("Masuk Sebagai", ["Siswa", "Guru", "Kepala Sekolah"])
        submit = st.form_submit_button("Masuk")
        
        if submit:
            if username and password:
                st.session_state.logged_in = True
                st.session_state.role = role_choice
                st.session_state.user_name = username
                st.success(f"Berhasil masuk sebagai {role_choice}!")
                st.rerun()
            else:
                st.error("Mohon isi username dan password dengan benar.")

# --- MODUL PRESENSI GURU (DENGAN VALIDASI GPS & WAJAH) ---
def modul_presensi_guru():
    st.header("📍 Modul Presensi & Validasi Wajah Guru")
    st.info("Pastikan perangkat Anda memberikan izin akses lokasi dan pencahayaan ruangan cukup.")
    
    MAX_AKURASI = 50  # Batas maksimal akurasi GPS dalam meter
    
    # Simulasi titik koordinat sekolah (Lat, Lon, Radius Max dalam meter)
    titik_sekolah = (-6.551234, 107.751234, 100) 

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Ambil Lokasi GPS")
        # Simulasi input/pengambilan koordinat perangkat
        simulasi_lat = st.number_input("Latitude", value=-6.551200, format="%.6f")
        simulasi_lon = st.number_input("Longitude", value=107.751200, format="%.6f")
        simulasi_akurasi = st.slider("Akurasi GPS (meter)", 5, 100, 15)
        
        lokasi_terkini = (simulasi_lat, simulasi_lon, simulasi_akurasi)

    with col2:
        st.subheader("2. Ambil Foto Wajah")
        foto_kamera = st.camera_input("Ambil foto selfie presensi")

    if st.button("Kirim Presensi Masuk", type="primary"):
        g_id = st.session_state.user_name
        tgl_hari_ini = now().strftime("%Y-%m-%d")
        
        # Validasi menggunakan alur yang diperbarui
        berhasil, pesan, info_wajah = proses_bukti(
            g_id, tgl_hari_ini, "Masuk", lokasi_terkini, foto_kamera, titik_sekolah
        )
        
        if not berhasil:
            st.error(f"Presensi Gagal: {pesan}")
        else:
            st.success(f"Presensi Berhasil dicatat! Status Wajah: {info_wajah}")

# --- MODUL TUGAS SISWA (DENGAN VALIDASI FORMAT & REAL-TIME STATUS) ---
def modul_tugas_siswa():
    st.header("📚 Pengumpulan Tugas Siswa")
    st.write("Unggah tugas Anda sesuai dengan format yang didukung (.pdf, .docx, .jpg, .png).")
    
    with st.form("form_upload_tugas"):
        mata_pelajaran = st.selectbox("Pelajaran", ["Matematika", "Bahasa Indonesia", "IPA", "IPS"])
        file_tugas = st.file_uploader("Pilih Berkas Tugas", type=["pdf", "docx", "jpg", "png"])
        submit_tugas = st.form_submit_button("Unggah Tugas")
        
        if submit_tugas:
            if file_tugas is not None:
                # Validasi tipe file tambahan di sisi backend/logic jika diperlukan
                st.success(f"Berhasil mengunggah berkas: {file_tugas.name}. Status tugas: Menunggu Penilaian.")
            else:
                st.warning("Silakan pilih berkas yang akan diunggah terlebih dahulu.")

# --- NAVIGATION & MAIN ROUTER ---
def main():
    if not st.session_state.logged_in:
        login_page()
    else:
        # Sidebar Navigasi
        st.sidebar.title(f"Halo, {st.session_state.user_name}")
        st.sidebar.write(Peran: **{st.session_state.role}**")
        
        menu_options = ["Beranda"]
        if st.session_state.role == "Guru":
            menu_options.append("Presensi & Wajah")
        elif st.session_state.role == "Siswa":
            menu_options.append("Pengumpulan Tugas")
        
        menu_options.append("Keluar")
        
        selected_menu = st.sidebar.radio("Navigasi Menu", menu_options, key="active_tab")
        
        if selected_menu == "Beranda":
            st.title("Selamat Datang di LMS SMP Negeri 1 Cijambe")
            st.write("Gunakan menu di sebelah kiri untuk mengakses fitur presensi, pembelajaran, atau pengumpulan tugas.")
        elif selected_menu == "Presensi & Wajah":
            modul_presensi_guru()
        elif selected_menu == "Pengumpulan Tugas":
            modul_tugas_siswa()
        elif selected_menu == "Keluar":
            st.session_state.logged_in = False
            st.session_state.role = None
            st.session_state.user_name = ""
            st.rerun()

if __name__ == "__main__":
    main()
