import streamlit as st
import pandas as pd

# ==========================================
# 1. KONFIGURASI HALAMAN (BEBAS U+200B)
# ==========================================
st.set_page_config(
    page_title="Ramdani Tech - Asisten Guru EdTech",
    page_icon="👨‍🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS untuk tampilan lebih modern
st.markdown("""
    <style>
    .main { padding: 1rem 2rem; }
    .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. SIDEBAR NAVIGATION
# ==========================================
st.sidebar.title("👨‍🏫 Ramdani Tech")
st.sidebar.caption("Super App EdTech - Asisten Guru Multi-Fitur")
st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "📌 Menu Utama:",
    [
        "🏠 Dashboard Utama",
        "📝 Generator RPP / Modul Ajar",
        "❓ Generator Soal & Rubrik",
        "📊 Analisis Nilai & Ketuntasan",
        "🤖 Chatbot Konsultasi Guru",
        "💡 Ide Media & Alat Peraga",
        "⚙️ Pengaturan Profil Guru"
    ]
)

st.sidebar.markdown("---")
st.sidebar.info("💡 **Tips:** Pastikan data terisi lengkap untuk hasil generasi RPP dan Soal yang maksimal.")

# ==========================================
# 3. FITUR: DASHBOARD UTAMA
# ==========================================
if menu == "🏠 Dashboard Utama":
    st.title("🏠 Dashboard Asisten Guru")
    st.write("Selamat datang kembali! Berikut ringkasan aktivitas administrasi & mengajar Anda hari ini.")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Kelas Diampu", "6 Kelas", "+1 dari semester lalu")
    with col2:
        st.metric("Modul Ajar Dibuat", "24 Modul", "Kurikulum Merdeka")
    with col3:
        st.metric("Bank Soal Tersimpan", "150 Soal", "HOTS & LOTS")
    with col4:
        st.metric("Rata-rata Ketuntasan", "87.5%", "Sesuai KKTP")

    st.markdown("---")
    
    col_left, col_right = st.columns([2, 1])
    with col_left:
        st.subheader("📅 Jadwal Mengajar Hari Ini")
        jadwal = pd.DataFrame({
            "Jam": ["07.00 - 08.30", "08.45 - 10.15", "10.30 - 12.00"],
            "Kelas": ["X-IPA 1", "XI-IPS 2", "XII-IPA 3"],
            "Mata Pelajaran": ["Informatika", "Kewirausahaan", "Informatika Lanjut"],
            "Status RPP": ["✅ Siap", "✅ Siap", "⚠️ Belum Buat"]
        })
        st.dataframe(jadwal, use_container_width=True)

    with col_right:
        st.subheader("🔔 Pengingat Cepat")
        st.success("✔ Input Nilai Tugas 2 Kelas X-IPA 1")
        st.warning("⚠️ Buat Soal PTS Informatika Kelas XII")
        st.info("ℹ️ Rapat Evaluasi Kurikulum pukul 13.00 WIB")

# ==========================================
# 4. FITUR: GENERATOR RPP / MODUL AJAR
# ==========================================
elif menu == "📝 Generator RPP / Modul Ajar":
    st.title("📝 Generator RPP & Modul Ajar (Kurikulum Merdeka)")
    st.write("Isi formulir di bawah ini untuk menyusun Modul Ajar otomatis secara terstruktur.")

    with st.form("form_rpp"):
        c1, c2, c3 = st.columns(3)
        with c1:
            nama_guru = st.text_input("Nama Guru", "Ramdani, S.Pd.")
            sekolah = st.text_input("Nama Sekolah", "SMA Negeri 1 Tech")
        with c2:
            mapel = st.text_input("Mata Pelajaran", "Informatika")
            fase = st.selectbox("Fase / Kelas", ["Fase D (SMP)", "Fase E (SMA Kelas 10)", "Fase F (SMA Kelas 11-12)"])
        with c3:
            materi_pokok = st.text_input("Materi Pokok / Topik", "Algoritma Pemrograman Python")
            alokasi = st.text_input("Alokasi Waktu", "2 x 45 Menit (1 Pertemuan)")

        st.subheader("🎯 Capaian & Profil Pelajar")
        profil_pancasila = st.multiselect(
            "Profil Pelajar Pancasila yang Dikembangkan:",
            ["Bernalar Kritis", "Kreatif", "Gotong Royong", "Mandiri", "Beriman & Bertakwa", "Berkebinekaan Global"],
            default=["Bernalar Kritis", "Gotong Royong"]
        )
        
        tujuan_pembelajaran = st.text_area("Tujuan Pembelajaran (TP)", "Siswa mampu memahami struktur runtutan dan percabangan dalam Python serta menerapkannya dalam studi kasus sederhana.")
        
        btn_generate_rpp = st.form_submit_button("✨ Generasi Modul Ajar Lengkap")

    if btn_generate_rpp:
        st.success("Modul Ajar Berhasil Dibuat!")
        st.markdown("---")
        
        st.header(f"MODUL AJAR: {materi_pokok.upper()}")
        st.markdown(f"**Penyusun:** {nama_guru} | **Instansi:** {sekolah} | **Fase/Kelas:** {fase}")
        st.markdown(f"**Alokasi Waktu:** {alokasi}")
        st.markdown(f"**Profil Pelajar Pancasila:** {', '.join(profil_pancasila)}")
        
        st.markdown("---")
        st.subheader("1. TUJUAN PEMBELAJARAN")
        st.write(tujuan_pembelajaran)

        st.subheader("2. KEGIATAN PEMBELAJARAN")
        st.markdown("""
        * **A. Kegiatan Pendahuluan (15 Menit)**
          1. Guru mengucapkan salam, berdoa, dan melakukan presensi.
          2. Guru memberikan pertanyaan pemantik terkait topik hari ini.
          3. Guru menyampaikan tujuan pembelajaran dan penilaian.
        
        * **B. Kegiatan Inti (60 Menit)**
          1. **Orientasi Siswa:** Guru menyajikan masalah nyata yang membutuhkan solusi algoritma.
          2. **Pengorganisasian:** Siswa dibagi menjadi kelompok kecil (3-4 orang).
          3. **Penyelidikan:** Siswa mendiskusikan penyelesaian logika bersama kelompoknya.
          4. **Pengembangan & Presentasi:** Masing-masing kelompok mempresentasikan logika programnya.
        
        * **C. Kegiatan Penutup (15 Menit)**
          1. Guru dan siswa merangkum poin-poin utama materi.
          2. Guru memberikan refleksi singkat dan umpan balik.
          3. Doa penutup dan salam.
        """)

        st.subheader("3. ASESMEN / PENILAIAN")
        st.write("- **Asesmen Formatif:** Observasi keaktifan diskusi kelompok & Lembar Kerja Siswa (LKPD)")
        st.write("- **Asesmen Sumatif:** Kuis pilihan ganda dan tes praktik koding")

# ==========================================
# 5. FITUR: GENERATOR SOAL & RUBRIK
# ==========================================
elif menu == "❓ Generator Soal & Rubrik":
    st.title("❓ Generator Soal Ujian & Rubrik Penilaian")
    st.write("Buat paket soal beserta kunci jawaban dan bobot penilaian secara otomatis.")

    col1, col2, col3 = st.columns(3)
    with col1:
        mapel_soal = st.text_input("Mata Pelajaran", "Informatika")
        topik_soal = st.text_input("Topik / Bab", "Struktur Data & Algoritma")
    with col2:
        tipe_soal = st.selectbox("Jenis Soal", ["Pilihan Ganda", "Essay / Uraian HOTS", "Kombinasi"])
        tingkat_kesulitan = st.select_slider("Tingkat Kesulitan", options=["Mudah", "Sedang / LOTS", "Sangat Sulit / HOTS"])
    with col3:
        jumlah_soal = st.number_input("Jumlah Soal", min_value=1, max_value=20, value=3)

    if st.button("🚀 Buat Bank Soal", type="primary"):
        st.markdown("---")
        st.subheader(f"📋 Paket Soal {mapel_soal} - {topik_soal}")
        
        for i in range(1, jumlah_soal + 1):
            st.markdown(f"#### Soal Nomor {i} ({tingkat_kesulitan})")
            if tipe_soal == "Pilihan Ganda" or (tipe_soal == "Kombinasi" and i % 2 != 0):
                st.write(f"Berikut ini yang merupakan implementasi algoritma yang paling efisien untuk pencarian data terurut adalah...")
                st.write("A. Linear Search")
                st.write("B. Binary Search")
                st.write("C. Bubble Sort")
                st.write("D. Quick Sort")
                with st.expander(f"🔑 Lihat Kunci Jawaban & Pembahasan Soal {i}"):
                    st.write("**Kunci Jawaban:** B. Binary Search")
                    st.write("**Pembahasan:** Binary search membagi dua ruang pencarian pada setiap langkah, sehingga kecepatannya $O(\log n)$ yang jauh lebih efisien untuk data terurut dibanding Linear Search.")
            else:
                st.write(f"Jelaskan perbedaan mendasar antara struktur data Stack (Tumpukan) dan Queue (Antrean) beserta contoh penerapannya dalam kehidupan sehari-hari!")
                with st.expander(f"🔑 Lihat Rubrik Penilaian & Jawaban Soal {i}"):
                    st.write("**Rubrik Skor (Maksimal 10):**")
                    st.write("- **Skor 10:** Menjelaskan LIFO & FIFO dengan benar + contoh konkret kedua konsep.")
                    st.write("- **Skor 5:** Hanya menjelaskan konsep LIFO & FIFO tanpa contoh.")
                    st.write("- **Skor 2:** Jawaban kurang tepat tetapi mencoba menjawab.")

# ==========================================
# 6. FITUR: ANALISIS NILAI & KETUNTASAN
# ==========================================
elif menu == "📊 Analisis Nilai & Ketuntasan":
    st.title("📊 Analisis Nilai & Ketuntasan Belajar (KKTP)")
    st.write("Unggah data nilai siswa atau kelola sampel rekap nilai di bawah ini.")

    # Data Contoh
    data_awal = {
        "NIS": ["1001", "1002", "1003", "1004", "1005", "1006"],
        "Nama Siswa": ["Ahmad Rizky", "Siti Nurhaliza", "Budi Santoso", "Dian Sastro", "Eko Prasetyo", "Fani Amalia"],
        "Nilai Tugas": [85, 92, 60, 78, 90, 65],
        "Nilai UH": [80, 88, 55, 75, 85, 60],
        "Nilai PTS": [88, 95, 65, 80, 92, 70]
    }
    
    df = pd.DataFrame(data_awal)
    df["Rata-Rata"] = ((df["Nilai Tugas"] + df["Nilai UH"] + df["Nilai PTS"]) / 3).round(1)
    
    kktp = st.number_input("Batas Ketuntasan Minimum (KKTP / KKM):", value=75)
    df["Status Ketuntasan"] = df["Rata-Rata"].apply(lambda x: "✅ Tuntas" if x >= kktp else "❌ Remedial")

    st.subheader("📋 Tabel Rekapitulasi Nilai")
    st.dataframe(df, use_container_width=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Siswa", len(df))
    with col2:
        tuntas = len(df[df["Status Ketuntasan"] == "✅ Tuntas"])
        st.metric("Siswa Tuntas", f"{tuntas} Siswa")
    with col3:
        remed = len(df[df["Status Ketuntasan"] == "❌ Remedial"])
        st.metric("Siswa Perlu Remedial", f"{remed} Siswa", delta_color="inverse")

    st.markdown("---")
    st.subheader("🎯 Rekomendasi Tindak Lanjut")
    if remed > 0:
        st.warning(f"Terdapat **{remed} siswa** yang berada di bawah nilai KKTP ({kktp}). Disarankan untuk mengadakan kelas pengayaan atau pemberian tugas tambahan.")
    else:
        st.success("Luar biasa! Seluruh siswa telah mencapai standar ketuntasan minimum.")

# ==========================================
# 7. FITUR: CHATBOT KONSULTASI GURU
# ==========================================
elif menu == "🤖 Chatbot Konsultasi Guru":
    st.title("🤖 Asisten Pedagojik Guru (AI Chatbot)")
    st.write("Diskusikan kendala mengajar, manajemen kelas, atau ide pembelajaran interaktif.")

    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "assistant", "content": "Halo Bapak/Ibu Guru! Ada yang bisa saya bantu terkait kelas, metode mengajar, atau penanganan siswa hari ini?"}
        ]

    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    if prompt := st.chat_input("Tanyakan sesuatu (Contoh: Bagaimana cara menghadapi siswa yang pasif?)..."):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        # Respon Asisten
        response = f"Terima kasih atas pertanyaannya mengenai: **'{prompt}'**.\n\n" \
                   f"Secara pedagogis, langkah efektif yang bisa diterapkan adalah:\n" \
                   f"1. **Gunakan metode interaktif:** Cobalah teknik *Think-Pair-Share* atau kuis singkat berbasis game.\n" \
                   f"2. **Apresiasi Proses:** Berikan penguatan positif pada setiap usaha siswa tanpa hanya berfokus pada hasil akhir.\n" \
                   f"3. **Pendekatan Personal:** Lakukan komunikasi personal setelah jam pelajaran selesai untuk memahami kendala mendasarnya."

        with st.chat_message("assistant"):
            st.markdown(response)
        st.session_state.messages.append({"role": "assistant", "content": response})

# ==========================================
# 8. FITUR: IDE MEDIA & ALAT PERAGA
# ==========================================
elif menu == "💡 Ide Media & Alat Peraga":
    st.title("💡 Ide Media & Alat Peraga Pembelajaran")
    st.write("Cari ide kreativitas pemanfaatan barang sekitar dan teknologi digital untuk mengajar.")

    kategori_media = st.selectbox("Pilih Kategori Media:", ["Digital & Aplikasi Apps", "Bahan Bekas / Sederhana", "Game & Simulasi"])

    if kategori_media == "Digital & Aplikasi Apps":
        st.info("💻 **Media Pembelajaran Digital Recommended:**")
        st.markdown("""
        - **Canva for Education:** Membuat presentasi interaktif dan lembar kerja bergambar.
        - **Wordwall / Quizizz:** Membuat game edukasi berbasis kuis untuk penilaian harian.
        - **PhET Interactive Simulations:** Untuk peragaan simulasi Sains dan Matematika secara visual.
        """)
    elif kategori_media == "Bahan Bekas / Sederhana":
        st.info("📦 **Alat Peraga Berbahan Sederhana:**")
        st.markdown("""
        - **Kotak Karton & Strik:** Membangun maket algoritma atau alur sistem.
        - **Kartu Domino Soal:** Membuat pasang-pasangan rumus/kata kunci menggunakan kartu kertas bekas.
        """)
    else:
        st.info("🎲 **Simulasi & Game Edukasi:**")
        st.markdown("""
        - **Roleplaying (Bermain Peran):** Masing-masing siswa memerankan elemen tertentu dalam materi.
        - **Treasure Hunt (Mencari Jejak):** Menyembunyikan petunjuk soal di sekeliling ruang kelas.
        """)

# ==========================================
# 9. FITUR: PENGATURAN PROFIL
# ==========================================
elif menu == "⚙️ Pengaturan Profil Guru":
    st.title("⚙️ Pengaturan Akun & Sekolah")
    
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Nama Lengkap & Gelar", value="Ramdani, S.Pd.")
        st.text_input("NIP / NUPTK", value="19920815 202012 1 002")
        st.text_input("Email Edukasi", value="ramdani@sekolah.sch.id")
    with col2:
        st.text_input("Nama Sekolah / Instansi", value="SMA Negeri 1 Tech")
        st.text_input("Mata Pelajaran Utama", value="Informatika & EdTech")
        st.selectbox("Tahun Ajaran Aktif", ["2025/2026", "2026/2027"])
        
    if st.button("💾 Simpan Perubahan Profil", type="primary"):
        st.success("Profil guru berhasil diperbarui!")
