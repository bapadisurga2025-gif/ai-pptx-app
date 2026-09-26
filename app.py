import streamlit as st
from google import genai
from pptx import Presentation
import os

st.title("AI PPTX Generator Online")
st.write("Upload data pendukung dan masukkan prompt untuk menghasilkan file presentasi secara otomatis.")

# Input Prompt
prompt = st.text_area("Prompt / Instruksi AI (EM / Admin)")

# Upload File Pendukung
uploaded_file = st.file_uploader("Upload File Data / Dokumen", type=["csv", "xlsx", "pdf"])
uploaded_photo = st.file_uploader("Upload Foto Pendukung", type=["png", "jpg", "jpeg"])

if st.button("Generate & Download PPTX"):
    if not prompt:
        st.warning("Mohon masukkan prompt terlebih dahulu!")
    else:
        with st.spinner("Sedang memproses dengan Gemini AI..."):
            try:
                # Menggunakan API Key dari st.secrets (disimpan aman di cloud Streamlit)
                client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                
                # Menggunakan model gemini-3.5-flash yang lebih stabil
                response = client.models.generate_content(
                    model='gemini-3.5-flash',
                    contents=f"Buatkan kerangka materi presentasi (Judul Slide dan Poin-poin isi) berdasarkan instruksi: {prompt}"
                )
                ai_output = response.text

                # Buat PPTX menggunakan python-pptx
                prs = Presentation()
                slide = prs.slides.add_slide(prs.slide_layouts[0])
                slide.shapes.title.text = "Hasil Presentasi AI"
                if slide.placeholders and len(slide.placeholders) > 1:
                    slide.placeholders[1].text = ai_output[:500]

                output_path = "output_presentation.pptx"
                prs.save(output_path)

                with open(output_path, "rb") as f:
                    st.download_button(
                        label="Download File PPTX",
                        data=f,
                        file_name="Presentasi_AI.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                st.success("Berhasil! Silakan klik tombol download di atas.")
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")
