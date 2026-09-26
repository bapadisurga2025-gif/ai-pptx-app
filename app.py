import streamlit as st
from google import genai
from pptx import Presentation
import os
import pypdf
import pandas as pd

# Konfigurasi halaman agar responsif
st.set_page_config(page_title="AI PPTX Generator", layout="centered")

st.title("AI PPTX Generator Online")
st.write("Upload dokumen pendukung, template, dan masukkan instruksi untuk menghasilkan presentasi otomatis.")

# Inisialisasi session state agar data tidak hilang saat halaman refresh/rerun
if "prompt" not in st.session_state:
    st.session_state.prompt = ""
if "ai_output" not in st.session_state:
    st.session_state.ai_output = None
if "output_path" not in st.session_state:
    st.session_state.output_path = None

# Input Prompt (Tersimpan otomatis)
prompt = st.text_area("Prompt / Instruksi AI (EM / Admin)", value=st.session_state.prompt, key="prompt_input")
st.session_state.prompt = prompt

# Fungsi pembantu untuk membaca isi file yang di-upload
def extract_file_content(uploaded_file):
    if uploaded_file is None:
        return ""
    content = ""
    try:
        file_extension = uploaded_file.name.split('.')[-1].lower()
        if file_extension == 'pdf':
            reader = pypdf.PdfReader(uploaded_file)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    content += text + "\n"
        elif file_extension in ['xlsx', 'xls']:
            df = pd.read_excel(uploaded_file)
            content = df.to_string()
        elif file_extension == 'csv':
            df = pd.read_csv(uploaded_file)
            content = df.to_string()
    except Exception as e:
        content = f"[Gagal membaca file {uploaded_file.name}: {e}]"
    return content

st.markdown("---")
st.subheader("📁 Upload Dokumen & Foto Pendukung")

# 1. Upload File EB & Foto Pendukungnya
col1, col2 = st.columns(2)
with col1:
    file_eb = st.file_uploader("Upload File EB", type=["csv", "xlsx", "pdf"], key="eb_file")
with col2:
    photo_eb = st.file_uploader("Upload Foto EB", type=["png", "jpg", "jpeg"], key="eb_photo")

# 2. Upload File Jaskug & Foto Pendukungnya
col3, col4 = st.columns(2)
with col3:
    file_jaskug = st.file_uploader("Upload File Jaskug", type=["csv", "xlsx", "pdf"], key="jaskug_file")
with col4:
    photo_jaskug = st.file_uploader("Upload Foto Jaskug", type=["png", "jpg", "jpeg"], key="jaskug_photo")

# 3. Upload File Ritel & Foto Pendukungnya
col5, col6 = st.columns(2)
with col5:
    file_ritel = st.file_uploader("Upload File Ritel", type=["csv", "xlsx", "pdf"], key="ritel_file")
with col6:
    photo_ritel = st.file_uploader("Upload Foto Ritel", type=["png", "jpg", "jpeg"], key="ritel_photo")

st.markdown("---")
# Upload Template PowerPoint
uploaded_template = st.file_uploader("Upload Template PowerPoint (.pptx)", type=["pptx"], key="template_file")

if st.button("Generate & Download PPTX"):
    if not prompt:
        st.warning("Mohon masukkan prompt terlebih dahulu!")
    else:
        with st.spinner("Membaca dokumen dan memproses dengan Gemini AI..."):
            try:
                # Ekstrak isi dari ketiga file dokumen
                content_eb = extract_file_content(file_eb)
                content_jaskug = extract_file_content(file_jaskug)
                content_ritel = extract_file_content(file_ritel)

                combined_data = f"""
                DATA EB:
                {content_eb[:3000]}
                
                DATA JASKUG:
                {content_jaskug[:3000]}
                
                DATA RITEL:
                {content_ritel[:3000]}
                """

                # Panggil Gemini API
                client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                response = client.models.generate_content(
                    model='gemini-3.5-flash',
                    contents=f"Berdasarkan instruksi: '{prompt}' dan data berikut:\n{combined_data}\nBuatkan kerangka materi presentasi terstruktur yang mencakup Judul Slide dan Poin-poin penjelasannya secara mendetail."
                )
                st.session_state.ai_output = response.text

                # Buat atau gunakan template PPTX
                if uploaded_template is not None:
                    prs = Presentation(uploaded_template)
                else:
                    prs = Presentation()

                # Tambahkan slide hasil ringkasan AI
                slide = prs.slides.add_slide(prs.slide_layouts[0])
                if slide.shapes.title:
                    slide.shapes.title.text = "Hasil Analisis Presentasi AI"
                if slide.placeholders and len(slide.placeholders) > 1:
                    slide.placeholders[1].text = st.session_state.ai_output[:1000]

                st.session_state.output_path = "output_presentation.pptx"
                prs.save(st.session_state.output_path)
                st.success("Berhasil! Silakan klik tombol download di bawah.")
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

# Tampilkan tombol download jika file sudah tersedia di session state
if st.session_state.output_path and os.path.exists(st.session_state.output_path):
    with open(st.session_state.output_path, "rb") as f:
        st.download_button(
            label="Download File PPTX",
            data=f,
            file_name="Presentasi_Hasil_AI.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
