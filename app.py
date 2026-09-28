import os
import google.generativeai as genai
from typing import Optional

# Configuration
DEFAULT_MODEL = "gemini-1.5-flash"
FALLBACK_MODELS = ["gemini-1.5-pro", "gemini-2.0-flash"]

class AsistenGuruRPPGenerator:
    """
    Modul Asisten Guru untuk membuat RPP/Modul Ajar menggunakan Gemini API.
    """
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.getenv("GEMINI_API_KEY")
        if not self.api_key:
            raise ValueError("API Key Google Gemini belum dikonfigurasi. Harap tentukan GEMINI_API_KEY.")
        
        genai.configure(api_key=self.api_key)

    def generate_rpp(self, mata_pelajaran: str, kelas: str, topik: str, alokasi_waktu: str = "2x45 menit") -> str:
        """
        Menghasilkan RPP/Modul Ajar Kurikulum Merdeka.
        """
        prompt = f"""
        Anda adalah Asisten Guru, asisten AI interaktif dan profesional yang membantu pendidik membuat Perencanaan Pembelajaran (RPP/Modul Ajar).
        
        Buatkan RPP / Modul Ajar Kurikulum Merdeka yang lengkap dan terstruktur berdasarkan informasi berikut:
        - Mata Pelajaran: {mata_pelajaran}
        - Kelas/Fase: {kelas}
        - Topik/Materi Utama: {topik}
        - Alokasi Waktu: {alokasi_waktu}

        Format RPP harus mencakup:
        1. Identitas Modul
        2. Capaian Pembelajaran & Tujuan Pembelajaran
        3. Profil Pelajar Pancasila
        4. Model & Metode Pembelajaran
        5. Kegiatan Pembelajaran (Pendahuluan, Inti, Penutup)
        6. Asesmen/Penilaian (Formatif & Sumatif)
        7. Remedial & Pengayaan
        
        Gunakan bahasa Indonesia yang profesional, jelas, dan mudah dipahami oleh guru.
        """

        # Mencoba model utama terlebih dahulu
        models_to_try = [DEFAULT_MODEL] + FALLBACK_MODELS
        last_error = None

        for model_name in models_to_try:
            try:
                model = genai.GenerativeModel(model_name)
                response = model.generate_content(prompt)
                if response and response.text:
                    return response.text
            except Exception as e:
                last_error = e
                continue

        raise RuntimeError(f"Gagal generate RPP dari Asisten Guru. Error terakhir: {str(last_error)}")


# Contoh Penggunaan:
if __name__ == "__main__":
    # Ganti 'YOUR_API_KEY' dengan API key Anda atau atur environment variable GEMINI_API_KEY
    API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_API_KEY_HERE")
    
    try:
        asisten = AsistenGuruRPPGenerator(api_key=API_KEY)
        print("Sedang membuat RPP via Asisten Guru...")
        
        rpp_result = asisten.generate_rpp(
            mata_pelajaran="Matematika",
            kelas="X (Fase E)",
            topik="Persamaan dan Pertidaksamaan Linear",
            alokasi_waktu="2 x 45 menit"
        )
        print("\n--- HASIL GENERATE RPP (ASISTEN GURU) ---\n")
        print(rpp_result)
        
    except Exception as err:
        print(f"Error: {err}")
