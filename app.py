import streamlit as st

# ==========================================
# --- Konfigurasi Halaman ---
# ==========================================
st.set_page_config(
    page_title="Ramdani Tech - Super App EdTech",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# --- Sidebar / Navigasi ---
# ==========================================
st.sidebar.title("🚀 Ramdani Tech")
st.sidebar.subheader("Super App EdTech")

menu = st.sidebar.radio(
    "Pilih Fitur:",
    ["Dashboard", "Materi Pembelajaran", "Latihan Soal / Kuis", "Pengaturan"]
)

# ==========================================
# --- Konten Utama ---
# ==========================================
if menu == "Dashboard":
    st.title("📌 Dashboard Utama")
    st.write("Selamat datang di **Ramdani Tech Super App EdTech**!")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="Total Siswa", value="1,250")
    with col2:
        st.metric(label="Kursus Aktif", value="18")
    with col3:
        st.metric(label="Tingkat Kelulusan", value="94%")

elif menu == "Materi Pembelajaran":
    st.title("📚 Materi Pembelajaran")
    st.info("Pilih modul di bawah ini untuk mulai belajar.")
    
    tab1, tab2 = st.tabs(["Pemrograman", "Desain Grafis"])
    with tab1:
        st.subheader("Python untuk Pemula")
        st.write("Pelajari dasar-dasar sintaks Python, tipe data, dan logika pemrograman.")
    with tab2:
        st.subheader("Dasar UI/UX Design")
        st.write("Pelajari prinsip desain antarmuka dan pengalaman pengguna.")

elif menu == "Latihan Soal / Kuis":
    st.title("📝 Kuis & Evaluasi")
    st.write("Uji pemahaman Anda melalui kuis interaktif.")
    
    jawaban = st.radio(
        "1. Bahasa pemrogramam apa yang digunakan untuk membuat aplikasi ini?",
        ["Java", "Python", "C++", "PHP"]
    )
    if st.button("Kirim Jawaban"):
        if jawaban == "Python":
            st.success("Jawaban Anda Benar! 🎉")
        else:
            st.error("Jawaban kurang tepat, coba lagi!")

elif menu == "Pengaturan":
    st.title("⚙️ Pengaturan Profil")
    st.text_input("Nama Lengkap", value="Pengguna Ramdani Tech")
    st.text_input("Email", value="user@ramdanitech.com")
    st.button("Simpan Perubahan")
