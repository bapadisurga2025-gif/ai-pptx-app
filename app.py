import streamlit as st
from google import genai
from pptx import Presentation
import os

# Konfigurasi halaman agar lebih responsif di mobile
st.set_page_config(page_title="AI PPTX Generator", layout="centered")

st.title("AI PPTX Generator Online")
st.write("Upload data pendukung dan masukkan prompt untuk menghasilkan file presentasi secara otomatis.")

# Inisialisasi session state untuk menjaga kestabilan data saat rerun
if "ai_output" not in st.session_state:
    st.session_state.ai_output = None
if "output_path" not in st.session_state:
    st.session_state.output_path = None

# Input Prompt
prompt = st.text_area("Prompt / Instruksi AI (EM / Admin)", key="prompt_input")

# Upload File Pendukung
uploaded_file = st.file_uploader("Upload File Data / Dokumen", type=["csv", "xlsx", "pdf"], key="file_upload")
uploaded_photo = st.file_uploader("Upload Foto Pendukung", type=["png", "jpg", "jpeg"], key="photo_upload")

if st.button("Generate & Download PPTX"):
    if not prompt:
        st.warning("Mohon masukkan prompt terlebih dahulu!")
    else:
        with st.spinner("Sedang memproses dengan Gemini AI..."):
            try:
                # Menggunakan API Key dari st.secrets
                client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                
                # Menggunakan model gemini-3.5-flash yang stabil
                response = client.models.generate_content(
                    model='gemini-3.5-flash',
                    contents=f"Buatkan kerangka materi presentasi (Judul Slide dan Poin-poin isi) berdasarkan instruksi: {prompt}"
                )
                st.session_state.ai_output = response.text

                # Buat PPTX menggunakan python-pptx
                prs = Presentation()
                slide = prs.slides.add_slide(prs.slide_layouts[0])
                slide.shapes.title.text = "Hasil Presentasi AI"
                if slide.placeholders and len(slide.placeholders) > 1:
                    slide.placeholders[1].text = st.session_state.ai_output[:500]

                st.session_state.output_path = "output_presentation.pptx"
                prs.save(st.session_state.output_path)
                st.success("Berhasil! Silakan klik tombol download di bawah.")
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

# Tampilkan tombol download jika file sudah siap di session state
if st.session_state.output_path and os.path.exists(st.session_state.output_path):
    with open(st.session_state.output_path, "rb") as f:
        st.download_button(
            label="Download File PPTX",
            data=f,
            file_name="Presentasi_AI.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
