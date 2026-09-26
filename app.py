import streamlit as st
import os
import tempfile
import pandas as pd
import pypdf
import copy
from google import genai
from pptx import Presentation
from pptx.util import Inches, Pt

# -----------------------------------------------------------------------------
# KONFIGURASI HALAMAN STREAMLIT
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="AI PPTX Executive Generator",
    page_icon="📊",
    layout="wide"
)

# Custom Styling
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        width: 100%;
        background-color: #1f1f1f;
        color: white;
        border-radius: 6px;
        font-weight: bold;
        padding: 0.6rem;
    }
    .stButton>button:hover {
        background-color: #333333;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 AI PPTX Generator - Precision Template Adaptor")
st.write("Sistem otomatis pengolahan data yang mengadaptasi 100% tata letak, warna, dan gaya visual dari template PowerPoint Anda.")

# -----------------------------------------------------------------------------
# FUNGSI PEMBACAAN FILE DATASOURCE
# -----------------------------------------------------------------------------
def extract_text_from_file(file_obj):
    if file_obj is None:
        return ""
    content = ""
    try:
        ext = file_obj.name.split('.')[-1].lower()
        if ext == 'pdf':
            reader = pypdf.PdfReader(file_obj)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    content += text + "\n"
        elif ext in ['xlsx', 'xls']:
            df = pd.read_excel(file_obj)
            content = df.to_string()
        elif ext == 'csv':
            df = pd.read_csv(file_obj)
            content = df.to_string()
    except Exception as e:
        content = f"[Gagal membaca file {file_obj.name}: {e}]"
    return content

# -----------------------------------------------------------------------------
# INPUT PROMPT / INSTRUKSI EKSEKUTIF
# -----------------------------------------------------------------------------
st.subheader("📝 Prompt / Instruksi Khusus EM")
prompt_text = st.text_area(
    "Masukkan instruksi penyusunan presentasi:",
    value="Buat laporan eksekutif performa kinerja keuangan berdasarkan data yang diunggah, lengkap dengan ringkasan pencapaian utama, segmen top growth, area underperform/kritis, serta rencana aksi pemulihan strategis.",
    height=120,
    key="prompt_input"
)

st.markdown("---")
st.subheader("📁 Upload Data & Foto Pendukung (Admin / SPV)")

# Kolom Upload Sesuai Kanal (EB, Jaskug, Ritel)
col1, col2 = st.columns(2)
with col1:
    f_eb = st.file_uploader("Upload File EB (.csv, .xlsx, .pdf)", type=["csv", "xlsx", "pdf"], key="eb_f")
with col2:
    p_eb = st.file_uploader("Upload Foto EB (.png, .jpg, .jpeg)", type=["png", "jpg", "jpeg"], key="eb_p")

col3, col4 = st.columns(2)
with col3:
    f_jaskug = st.file_uploader("Upload File Jaskug (.csv, .xlsx, .pdf)", type=["csv", "xlsx", "pdf"], key="jaskug_f")
with col4:
    p_jaskug = st.file_uploader("Upload Foto Jaskug (.png, .jpg, .jpeg)", type=["png", "jpg", "jpeg"], key="jaskug_p")

col5, col6 = st.columns(2)
with col5:
    f_ritel = st.file_uploader("Upload File Ritel (.csv, .xlsx, .pdf)", type=["csv", "xlsx", "pdf"], key="ritel_f")
with col6:
    p_ritel = st.file_uploader("Upload Foto Ritel (.png, .jpg, .jpeg)", type=["png", "jpg", "jpeg"], key="ritel_p")

st.markdown("---")
st.subheader("🎨 Upload Template PowerPoint Acuan (.pptx)")
uploaded_template = st.file_uploader("Upload File Template PowerPoint (.pptx)", type=["pptx"], key="template_file")

st.markdown("---")

# -----------------------------------------------------------------------------
# PROSES GENERATE PRESENTASI
# -----------------------------------------------------------------------------
if st.button("🚀 Generate Executive PPTX Presisi Template"):
    if not (f_eb or f_jaskug or f_ritel):
        st.warning("⚠️ Mohon unggah setidaknya satu file data sumber (EB, Jaskug, atau Ritel) terlebih dahulu!")
    elif not uploaded_template:
        st.warning("⚠️ Mohon unggah file Template PowerPoint (.pptx) acuan agar hasilnya sesuai dengan desain yang Anda inginkan!")
    elif not prompt_text.strip():
        st.warning("⚠️ Mohon masukkan instruksi prompt terlebih dahulu!")
    else:
        with st.spinner("🤖 Mengolah data dan menerapkan isi ke dalam struktur template PowerPoint..."):
            try:
                # 1. Ekstraksi teks dari file data
                data_eb_text = extract_text_from_file(f_eb)
                data_jaskug_text = extract_text_from_file(f_jaskug)
                data_ritel_text = extract_text_from_file(f_ritel)

                combined_data = f"""
                DATA EB:
                {data_eb_text[:3000]}

                DATA JASKUG:
                {data_jaskug_text[:3000]}

                DATA RITEL:
                {data_ritel_text[:3000]}
                """

                # 2. Pemanggilan Gemini API via SDK google-genai
                api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
                if not api_key:
                    st.error("🔑 GEMINI_API_KEY tidak ditemukan pada st.secrets atau environment variables.")
                    st.stop()

                client = genai.Client(api_key=api_key)
                
                system_instruction = f"""Anda adalah ahli penyusun presentasi eksekutif. 
Berdasarkan instruksi: '{prompt_text}' dan data berikut:
{combined_data}

Buatkan isi materi untuk slide presentasi. 
Pecah menjadi beberapa bagian slide terstruktur persis dengan format:

---SLIDE---
JUDUL: [Judul Utama Slide]
SUBJUDUL: [Subjudul / Kategori Slide]
POIN_1: [Isi Ringkas Poin Utama Pertama]
POIN_2: [Isi Ringkas Poin Utama Kedua]
POIN_3: [Isi Ringkas Poin Utama Ketiga]
POIN_4: [Isi Ringkas Poin Utama Keempat]
"""

                response = client.models.generate_content(
                    model='gemini-3.5-flash',
                    contents=system_instruction
                )
                ai_output = response.text

                # 3. Baca Template PowerPoint Asli Pengguna
                temp_template = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
                temp_template.write(uploaded_template.getbuffer())
                temp_template.close()

                prs = Presentation(temp_template.name)
                slides_data = [s for s in ai_output.split("---SLIDE---") if "JUDUL:" in s]

                # Memadankan isi dari Gemini langsung ke bentuk kotak teks slide asli
                for index, s_data in enumerate(slides_data):
                    if index < len(prs.slides):
                        slide = prs.slides[index]
                    else:
                        break # Menggunakan jumlah slide dari template yang tersedia

                    lines = s_data.strip().split("\n")
                    judul, subjudul, points = "", "", []
                    
                    for line in lines:
                        if line.startswith("JUDUL:"):
                            judul = line.replace("JUDUL:", "").strip()
                        elif line.startswith("SUBJUDUL:"):
                            subjudul = line.replace("SUBJUDUL:", "").strip()
                        elif line.strip().startswith(("POIN_", "1.", "2.", "3.", "4.", "5.", "-")):
                            content_point = line.split(":", 1)[-1].strip() if ":" in line else line.strip()
                            points.append(content_point)

                    # Memasukkan teks secara aman ke dalam elemen/shape yang sudah ada di template
                    text_boxes = [shape for shape in slide.shapes if shape.has_text_frame]

                    if len(text_boxes) > 0 and judul:
                        # Tempatkan judul pada text box utama/teratas
                        text_boxes[0].text_frame.text = judul

                    if len(text_boxes) > 1 and subjudul:
                        # Tempatkan subjudul pada text box kedua
                        text_boxes[1].text_frame.text = subjudul

                    # Masukkan poin-poin data ke dalam sisa text box yang ada di slide template tersebut
                    point_idx = 0
                    for tb in text_boxes[2:]:
                        if point_idx < len(points):
                            tb.text_frame.text = points[point_idx]
                            point_idx += 1

                # 4. Simpan ke File Sementara
                output_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
                prs.save(output_temp.name)

                st.success("✨ File presentasi berhasil dibuat dengan mempertahankan 100% gaya & desain template Anda!")

                # 5. Tombol Unduh
                with open(output_temp.name, "rb") as f:
                    st.download_button(
                        label="📥 Download File PPTX (Presisi Template)",
                        data=f,
                        file_name="Presentasi_Eksekutif_Sesuai_Template.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses data: {e}")
