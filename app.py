import streamlit as st
from google import genai
from pptx import Presentation
import os
import pypdf
import pandas as pd

# Konfigurasi halaman
st.set_page_config(page_title="AI PPTX Generator Multi-Level", layout="centered")

st.title("AI PPTX Generator Online (Admin / SPV / EM)")
st.write("Upload data dan foto pendukung secara terpisah, lalu EM dapat langsung men-generate presentasi.")

# Buat folder penyimpanan permanen jika belum ada
UPLOAD_DIR = "uploaded_data"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Fungsi helper untuk menyimpan file secara permanen di server
def save_uploaded_file(uploaded_file, prefix):
    if uploaded_file is not None:
        file_path = os.path.join(UPLOAD_DIR, f"{prefix}_{uploaded_file.name}")
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return file_path
    return None

# Fungsi untuk membaca isi file yang sudah tersimpan
def read_saved_file(file_path):
    if not file_path or not os.path.exists(file_path):
        return ""
    content = ""
    try:
        ext = file_path.split('.')[-1].lower()
        if ext == 'pdf':
            reader = pypdf.PdfReader(file_path)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    content += text + "\n"
        elif ext in ['xlsx', 'xls']:
            df = pd.read_excel(file_path)
            content = df.to_string()
        elif ext == 'csv':
            df = pd.read_csv(file_path)
            content = df.to_string()
    except Exception as e:
        content = f"[Gagal membaca {file_path}: {e}]"
    return content

# Inisialisasi session state untuk prompt dan template
if "prompt" not in st.session_state:
    st.session_state.prompt = ""
if "output_path" not in st.session_state:
    st.session_state.output_path = None

# Input Prompt (Tersimpan otomatis)
prompt = st.text_area("Prompt / Instruksi AI (Khusus EM)", value=st.session_state.prompt, key="prompt_input")
st.session_state.prompt = prompt

st.markdown("---")
st.subheader("📁 Upload File & Foto Pendukung (Admin / SPV)")

# 1. File & Foto EB
col1, col2 = st.columns(2)
with col1:
    f_eb = st.file_uploader("Upload File EB", type=["csv", "xlsx", "pdf"], key="eb_f")
    if f_eb:
        st.session_state.path_eb = save_uploaded_file(f_eb, "eb")
with col2:
    p_eb = st.file_uploader("Upload Foto EB", type=["png", "jpg", "jpeg"], key="eb_p")
    if p_eb:
        st.session_state.photo_eb = save_uploaded_file(p_eb, "photo_eb")

# 2. File & Foto Jaskug
col3, col4 = st.columns(2)
with col3:
    f_jaskug = st.file_uploader("Upload File Jaskug", type=["csv", "xlsx", "pdf"], key="jaskug_f")
    if f_jaskug:
        st.session_state.path_jaskug = save_uploaded_file(f_jaskug, "jaskug")
with col4:
    p_jaskug = st.file_uploader("Upload Foto Jaskug", type=["png", "jpg", "jpeg"], key="jaskug_p")
    if p_jaskug:
        st.session_state.photo_jaskug = save_uploaded_file(p_jaskug, "photo_jaskug")

# 3. File & Foto Ritel
col5, col6 = st.columns(2)
with col5:
    f_ritel = st.file_uploader("Upload File Ritel", type=["csv", "xlsx", "pdf"], key="ritel_f")
    if f_ritel:
        st.session_state.path_ritel = save_uploaded_file(f_ritel, "ritel")
with col6:
    p_ritel = st.file_uploader("Upload Foto Ritel", type=["png", "jpg", "jpeg"], key="ritel_p")
    if p_ritel:
        st.session_state.photo_ritel = save_uploaded_file(p_ritel, "photo_ritel")

st.markdown("---")
# Upload Template PowerPoint
uploaded_template = st.file_uploader("Upload Template PowerPoint (.pptx)", type=["pptx"], key="template_file")
if uploaded_template:
    st.session_state.path_template = save_uploaded_file(uploaded_template, "template")

st.markdown("---")
# Tombol Eksekusi oleh EM
if st.button("Generate & Download PPTX (Akses EM)"):
    if not prompt:
        st.warning("Mohon masukkan instruksi/prompt terlebih dahulu!")
    else:
        with st.spinner("Menggabungkan data dan memproses dengan Gemini AI..."):
            try:
                # Baca file yang tersimpan secara permanen di server
                data_eb = read_saved_file(st.session_state.get("path_eb"))
                data_jaskug = read_saved_file(st.session_state.get("path_jaskug"))
                data_ritel = read_saved_file(st.session_state.get("path_ritel"))

                combined_data = f"""
                DATA EB:
                {data_eb[:2000]}
                
                DATA JASKUG:
                {data_jaskug[:2000]}
                
                DATA RITEL:
                {data_ritel[:2000]}
                """

                # Panggil Gemini AI dengan format terstruktur per slide
                client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                response = client.models.generate_content(
                    model='gemini-3.5-flash',
                    contents=f"""Berdasarkan instruksi: '{prompt}' dan data di bawah ini, buatkan materi presentasi lengkap yang dibagi menjadi beberapa slide. 
                    Gunakan format persis seperti ini untuk setiap slide:
                    ---SLIDE---
                    Judul: [Judul Slide]
                    Isi: [Poin-poin penjelasan lengkap]

                    Data pendukung:
                    {combined_data}
                    """
                )
                ai_output = response.text

                # Inisialisasi PowerPoint (menggunakan template jika ada)
                if "path_template" in st.session_state and os.path.exists(st.session_state.path_template):
                    prs = Presentation(st.session_state.path_template)
                else:
                    prs = Presentation()

                # Pecah hasil AI menjadi beberapa slide secara otomatis
                slides_data = ai_output.split("---SLIDE---")
                for s_data in slides_data:
                    if "Judul:" in s_data and "Isi:" in s_data:
                        try:
                            parts = s_data.split("Isi:")
                            title_part = parts[0].replace("Judul:", "").strip()
                            content_part = parts[1].strip()

                            slide = prs.slides.add_slide(prs.slide_layouts[1]) # Layout Title & Content
                            if slide.shapes.title:
                                slide.shapes.title.text = title_part
                            if slide.placeholders and len(slide.placeholders) > 1:
                                slide.placeholders[1].text = content_part
                        except:
                            continue

                # Jika gagal parsing, buat slide cadangan dari teks mentah
                if len(prs.slides) == 0:
                    slide = prs.slides.add_slide(prs.slide_layouts[0])
                    slide.shapes.title.text = "Hasil Presentasi AI"
                    if slide.placeholders and len(slide.placeholders) > 1:
                        slide.placeholders[1].text = ai_output[:1000]

                st.session_state.output_path = "output_presentation.pptx"
                prs.save(st.session_state.output_path)
                st.success("Berhasil! Presentasi multi-slide siap di-download.")
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

# Tombol download file hasil
if st.session_state.output_path and os.path.exists(st.session_state.output_path):
    with open(st.session_state.output_path, "rb") as f:
        st.download_button(
            label="Download File PPTX",
            data=f,
            file_name="Presentasi_Executive_Summary.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
