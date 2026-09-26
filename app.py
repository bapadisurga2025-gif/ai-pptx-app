import streamlit as st
from google import genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import os
import pypdf
import pandas as pd

# Konfigurasi halaman
st.set_page_config(page_title="AI PPTX Executive Generator", layout="centered")

st.title("AI PPTX Generator - Executive Dashboard Style")
st.write("Sistem otomatis pengolahan data EB, Jaskug, dan Ritel menjadi presentasi profesional.")

# Buat folder penyimpanan permanen jika belum ada
UPLOAD_DIR = "uploaded_data"
os.makedirs(UPLOAD_DIR, exist_ok=True)

def save_uploaded_file(uploaded_file, prefix):
    if uploaded_file is not None:
        file_path = os.path.join(UPLOAD_DIR, f"{prefix}_{uploaded_file.name}")
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return file_path
    return None

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

# Session state initialization
if "prompt" not in st.session_state:
    st.session_state.prompt = ""
if "output_path" not in st.session_state:
    st.session_state.output_path = None

prompt = st.text_area("Prompt / Instruksi Khusus EM", value=st.session_state.prompt, key="prompt_input")
st.session_state.prompt = prompt

st.markdown("---")
st.subheader("📁 Upload Data & Foto Pendukung (Admin / SPV)")

# Kolom Upload Sesuai Kanal
col1, col2 = st.columns(2)
with col1:
    f_eb = st.file_uploader("Upload File EB", type=["csv", "xlsx", "pdf"], key="eb_f")
    if f_eb: st.session_state.path_eb = save_uploaded_file(f_eb, "eb")
with col2:
    p_eb = st.file_uploader("Upload Foto EB", type=["png", "jpg", "jpeg"], key="eb_p")
    if p_eb: st.session_state.photo_eb = save_uploaded_file(p_eb, "photo_eb")

col3, col4 = st.columns(2)
with col3:
    f_jaskug = st.file_uploader("Upload File Jaskug", type=["csv", "xlsx", "pdf"], key="jaskug_f")
    if f_jaskug: st.session_state.path_jaskug = save_uploaded_file(f_jaskug, "jaskug")
with col4:
    p_jaskug = st.file_uploader("Upload Foto Jaskug", type=["png", "jpg", "jpeg"], key="jaskug_p")
    if p_jaskug: st.session_state.photo_jaskug = save_uploaded_file(p_jaskug, "photo_jaskug")

col5, col6 = st.columns(2)
with col5:
    f_ritel = st.file_uploader("Upload File Ritel", type=["csv", "xlsx", "pdf"], key="ritel_f")
    if f_ritel: st.session_state.path_ritel = save_uploaded_file(f_ritel, "ritel")
with col6:
    p_ritel = st.file_uploader("Upload Foto Ritel", type=["png", "jpg", "jpeg"], key="ritel_p")
    if p_ritel: st.session_state.photo_ritel = save_uploaded_file(p_ritel, "photo_ritel")

st.markdown("---")
uploaded_template = st.file_uploader("Upload Template PowerPoint (.pptx)", type=["pptx"], key="template_file")
if uploaded_template:
    st.session_state.path_template = save_uploaded_file(uploaded_template, "template")

st.markdown("---")
if st.button("Generate Executive PPTX"):
    if not prompt:
        st.warning("Mohon masukkan instruksi terlebih dahulu!")
    else:
        with st.spinner("Menyusun format presentasi eksekutif dengan Gemini AI..."):
            try:
                data_eb = read_saved_file(st.session_state.get("path_eb"))
                data_jaskug = read_saved_file(st.session_state.get("path_jaskug"))
                data_ritel = read_saved_file(st.session_state.get("path_ritel"))

                combined_data = f"""
                DATA EB: {data_eb[:1500]}
                DATA JASKUG: {data_jaskug[:1500]}
                DATA RITEL: {data_ritel[:1500]}
                """

                client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                response = client.models.generate_content(
                    model='gemini-3.5-flash',
                    contents=f"""Anda adalah tenaga ahli pembuat dashboard eksekutif korporat. Berdasarkan instruksi '{prompt}' dan data berikut:
                    {combined_data}
                    
                    Buatkan materi presentasi terstruktur yang bersih tanpa simbol markdown liar (** atau * berlebih). 
                    Pecah menjadi beberapa slide dengan format persis berikut untuk setiap slide:
                    ---SLIDE---
                    KATEGORI: [Judul Panel Utama, misal: Kinerja KANAL EB / Rencana Aksi]
                    METRIK: [Ringkasan Angka penting, misal: Realisasi: 731 Juta | Capaian: 83.7%]
                    POIN_UTAMA:
                    1. [Poin penjelas pertama yang rapi]
                    2. [Poin penjelas kedua yang rapi]
                    3. [Poin penjelas ketiga yang rapi]
                    """
                )
                ai_output = response.text

                # Inisialisasi Presentation
                if "path_template" in st.session_state and os.path.exists(st.session_state.path_template):
                    prs = Presentation(st.session_state.path_template)
                else:
                    prs = Presentation()

                slides_data = ai_output.split("---SLIDE---")
                for s_data in slides_data:
                    if "KATEGORI:" in s_data:
                        try:
                            lines = s_data.strip().split("\n")
                            cat, met, points = "", "", []
                            for line in lines:
                                if "KATEGORI:" in line: cat = line.replace("KATEGORI:", "").strip()
                                elif "METRIK:" in line: met = line.replace("METRIK:", "").strip()
                                elif line.strip().startswith(("1.", "2.", "3.", "4.", "5.", "-")): points.append(line.strip())

                            slide = prs.slides.add_slide(prs.slide_layouts[1])
                            if slide.shapes.title:
                                slide.shapes.title.text = cat
                            
                            if slide.placeholders and len(slide.placeholders) > 1:
                                formatted_text = f"RINGKASAN METRIK:\n{met}\n\nRINCIAN PROGRAM & KINERJA:\n" + "\n".join(points)
                                slide.placeholders[1].text = formatted_text
                        except:
                            continue

                if len(prs.slides) == 0:
                    slide = prs.slides.add_slide(prs.slide_layouts[0])
                    slide.shapes.title.text = "Executive Summary"
                    if slide.placeholders and len(slide.placeholders) > 1:
                        slide.placeholders[1].text = ai_output[:1000]

                st.session_state.output_path = "output_executive.pptx"
                prs.save(st.session_state.output_path)
                st.success("Berhasil! File presentasi gaya eksekutif siap di-download.")
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

if st.session_state.output_path and os.path.exists(st.session_state.output_path):
    with open(st.session_state.output_path, "rb") as f:
        st.download_button(
            label="Download File PPTX Eksekutif",
            data=f,
            file_name="Presentasi_Executive_Dashboard.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
