import streamlit as st

# ==========================================
# --- Konfigurasi Halaman ---
# ==========================================
st.set_page_config(
    page_title="Ramdani Tech - Asisten Guru EdTech",
    page_icon="👨‍🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# --- Sidebar Navigasi ---
# ==========================================
st.sidebar.title("👨‍🏫 Ramdani Tech")
st.sidebar.caption("Super App EdTech - Asisten Guru")

menu = st.sidebar.radio(
    "Fitur Asisten Guru:",
    [
        "📝 Pembuat RPP / Modul Ajar", 
        "❓ Generator Soal & Ujian", 
        "📊 Analisis Nilai Siswa", 
        "💡 Ide Metode Pembelajaran"
    ]
)

# ==========================================
# --- Fitur 1: Pembuat RPP / Modul Ajar ---
# ==========================================
if menu == "📝 Pembuat RPP / Modul Ajar":
    st.title("📝 Generator RPP & Modul Ajar")
    st.write("Buat rancangan pembelajaran dengan cepat dan terstruktur.")

    col1, col2 = st.columns(2)
    with col1:
        mata_pelajaran = st.text_input("Mata Pelajaran", placeholder="Contoh: Matematika / IPA")
        kelas = st.selectbox("Tingkat Kelas", ["Kelas 7", "Kelas 8", "Kelas 9", "Kelas 10", "Kelas 11", "Kelas 12"])
    with col2:
        topik = st.text_input("Materi / Topik Utama", placeholder="Contoh: Photosynthesis / Persamaan Kuadrat")
        alokasi_waktu = st.text_input("Alokasi Waktu", value="2 x 45 Menit")

    tujuan = st.text_area("Tujuan Pembelajaran", placeholder="Tuliskan tujuan yang ingin dicapai siswa...")

    if st.button("✨ Generasi Draf RPP", type="primary"):
        if mata_pelajaran and topik:
            st.success("RPP Berhasil Dibuat!")
            st.markdown("---")
            st.subheader(f"Draf RPP: {mata_pelajaran} - {topik} ({kelas})")
            st.markdown(f"**Alokasi Waktu:** {alokasi_waktu}")
            st.markdown(f"**Tujuan:** {tujuan}")
            
            st.markdown("### 📌 Kegiatan Pembelajaran")
            st.markdown("""
            1. **Pendahuluan (10 Menit):**
               - Guru membuka kelas, berdoa, dan mengecek kehadiran.
               - Apersepsi dan menyampaikan tujuan pembelajaran.
            2. **Kegiatan Inti (70 Menit):**
               - Eksplorasi materi bersama siswa.
               - Diskusi kelompok dan pengerjaan lembar kerja (LKPD).
               - Presentasi hasil diskusi kelompok.
            3. **Penutup (10 Menit):**
               - Kesimpulan bersama dan refleksi materi.
               - Penutupan dan doa.
            """)
        else:
            st.warning("Mohon isi Mata Pelajaran dan Topik terlebih dahulu!")

# ==========================================
# --- Fitur 2: Generator Soal ---
# ==========================================
elif menu == "❓ Generator Soal & Ujian":
    st.title("❓ Generator Soal Otomatis")
    st.write("Buat bank soal latihan atau ujian untuk siswa Anda.")

    col1, col2, col3 = st.columns(3)
    with col1:
        mapel_soal = st.text_input("Mata Pelajaran", value="Bahasa Indonesia")
    with col2:
        bentuk_soal = st.selectbox("Bentuk Soal", ["Pilihan Ganda", "Essay / Uraian", "Isian Singkat"])
    with col3:
        jumlah_soal = st.number_input("Jumlah Soal", min_value=1, max_value=20, value=5)

    tingkat_kesulitan = st.select_slider("Tingkat Kesulitan", options=["Mudah", "Sedang", "Sulit"])

    if st.button("🎯 Buat Soal Sekarang", type="primary"):
        st.subheader(f"Hasil Soal ({bentuk_soal} - {tingkat_kesulitan})")
        
        if bentuk_soal == "Pilihan Ganda":
            for i in range(1, jumlah_soal + 1):
                st.markdown(f"**{i}. Contoh soal pilihan ganda untuk {mapel_soal}?**")
                st.write("A. Pilihan Jawaban A")
                st.write("B. Pilihan Jawaban B")
                st.write("C. Pilihan Jawaban C")
                st.write("D. Pilihan Jawaban D")
                st.caption("Kunci Jawaban: A")
                st.write("")
        else:
            for i in range(1, jumlah_soal + 1):
                st.markdown(f"**{i}. Jelaskan dan uraikan mengenai konsep utama pada {mapel_soal}!**")
                st.caption("Rubrik Penilaian: Skor 10 jika menyebutkan 3 kata kunci utama.")
                st.write("")

# ==========================================
# --- Fitur 3: Analisis Nilai ---
# ==========================================
elif menu == "📊 Analisis Nilai Siswa":
    st.title("📊 Analisis & Rekap Nilai Siswa")
    st.write("Pantau perkembangan dan ketuntasan belajar siswa.")

    uploaded_file = st.file_uploader("Unggah File Nilai (CSV / Excel)", type=["csv"])
    
    st.info("Atau gunakan data contoh di bawah ini:")
    
    # Sample Data
    import pandas as pd
    data_nilai = pd.DataFrame({
        "Nama Siswa": ["Andi", "Budi", "Citra", "Dewi", "Eko"],
        "Tugas 1": [85, 70, 90, 60, 78],
        "Tugas 2": [88, 75, 95, 65, 80],
        "UTS": [80, 68, 92, 55, 75],
        "Status": ["Lulus", "Lulus", "Lulus", "Remedial", "Lulus"]
    })
    
    st.dataframe(data_nilai, use_container_width=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Rata-Rata Kelas", "76.6")
    with col2:
        st.metric("Siswa Perlu Remedial", "1 Orang", delta="-20%", delta_color="inverse")

# ==========================================
# --- Fitur 4: Ide Metode Pembelajaran ---
# ==========================================
elif menu == "💡 Ide Metode Pembelajaran":
    st.title("💡 Ide & Strategi Pembelajaran Interaktif")
    st.write("Cari inspirasi metode mengajar agar kelas tidak membosankan.")

    gaya_belajar = st.selectbox("Pilih Target / Suasana Kelas:", [
        "Kelas Kurang Aktif / Pasif",
        "Materi Sulit / Membutuhkan Praktik",
        "Pembelajaran Kelompok / Kolaboratif",
        "Kuis & Game Interaktif"
    ])

    if gaya_belajar == "Kelas Kurang Aktif / Pasif":
        st.lightbulb = st.success("💡 **Rekomendasi Metode: Think-Pair-Share (TPS)**")
        st.write("Berikan pertanyaan, minta siswa berpikir sendiri (1-2 menit), berpasangan dengan teman di sebelahnya, lalu bagikan hasil diskusinya ke depan kelas.")
    elif gaya_belajar == "Kuis & Game Interaktif":
        st.success("💡 **Rekomendasi Metode: Gamifikasi (Kahoot / Quizizz Style)**")
        st.write("Gunakan kuis berbatas waktu dengan poin tinggi untuk meningkatkan kompetisi positif di kelas.")
