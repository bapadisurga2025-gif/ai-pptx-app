import streamlit as st
import os
import tempfile
import pandas as pd
import pypdf
from google import genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

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

st.title("📊 AI PPTX Generator - Executive Style")
st.write("Sistem otomatis pengolahan data menjadi presentasi PowerPoint profesional yang menyesuaikan instruksi prompt dan acuan desain.")

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
# SIDEBAR: UPLOAD DATA & TEMPLATE
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Sumber Data & Template")
    
    uploaded_template = st.file_uploader(
        "Upload Template PowerPoint (.pptx)", 
        type=["pptx"],
        help="Template ini akan diacu sebagai contoh gaya visual dan tata letak.",
        key="template_file"
    )
    
    st.divider()
    st.subheader("📁 Data Sumber Kinerja")
    f_eb = st.file_uploader("Data EB (.csv, .xlsx, .pdf)", type=["csv", "xlsx", "pdf"], key="eb_f")
    f_jaskug = st.file_uploader("Data Jaskug (.csv, .xlsx, .pdf)", type=["csv", "xlsx", "pdf"], key="jaskug_f")
    f_ritel = st.file_uploader("Data Ritel (.csv, .xlsx, .pdf)", type=["csv", "xlsx", "pdf"], key="ritel_f")
    
    st.divider()
    st.info("💡 **Tips:** Gemini akan menganalisis data sumber dan menyusun isi slide sesuai prompt yang Anda masukkan.")

# -----------------------------------------------------------------------------
# INPUT PROMPT / INSTRUKSI EKSEKUTIF
# -----------------------------------------------------------------------------
st.subheader("📝 Prompt / Instruksi Khusus Eksekutif")
prompt_text = st.text_area(
    "Masukkan instruksi penyusunan presentasi:",
    value="Buat laporan eksekutif performa kinerja keuangan berdasarkan data yang diunggah, lengkap dengan ringkasan pencapaian utama, segmen top growth, area underperform/kritis, serta rencana aksi pemulihan strategis.",
    height=120,
    key="prompt_input"
)

# -----------------------------------------------------------------------------
# PROSES GENERATE PRESENTASI
# -----------------------------------------------------------------------------
if st.button("🚀 Generate Executive Presentation (.pptx)"):
    if not (f_eb or f_jaskug or f_ritel):
        st.warning("⚠️ Mohon unggah setidaknya satu file data sumber (EB, Jaskug, atau Ritel) terlebih dahulu!")
    elif not prompt_text.strip():
        st.warning("⚠️ Mohon masukkan instruksi prompt terlebih dahulu!")
    else:
        with st.spinner("🤖 Memproses data dan merakit slide presentasi PowerPoint dengan Gemini AI..."):
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
                
                system_instruction = f"""Anda adalah tenaga ahli pembuat dashboard dan presentasi eksekutif korporat.
Berdasarkan instruksi berikut: '{prompt_text}' dan gabungan data di bawah ini:
{combined_data}

Buatkan materi presentasi terstruktur yang bersih tanpa simbol markdown liar (** atau * berlebih). 
Pecah menjadi beberapa slide (minimal 4 slide: Cover/Title, Summary, Detail Performance, Action Plan) dengan format terstruktur persis berikut untuk setiap slide:

---SLIDE---
JUDUL: [Judul Slide Utama]
SUBJUDUL: [Subjudul/Kategori Slide]
POIN_UTAMA:
1. [Poin penjelas pertama]
2. [Poin penjelas kedua]
3. [Poin penjelas ketiga]
4. [Poin penjelas keempat]
"""

                response = client.models.generate_content(
                    model='gemini-3.5-flash-lite',
                    contents=system_instruction
                )
                ai_output = response.text

                # 3. Merakit File PowerPoint (.pptx)
                prs = Presentation()
                prs.slide_width = Inches(13.333)
                prs.slide_height = Inches(7.5)
                blank_layout = prs.slide_layouts[6]

                slides_raw = ai_output.split("---SLIDE---")
                slide_count = 0

                for s_data in slides_raw:
                    if "JUDUL:" in s_data:
                        lines = s_data.strip().split("\n")
                        judul, subjudul, points = "", "", []
                        
                        for line in lines:
                            if line.startswith("JUDUL:"):
                                judul = line.replace("JUDUL:", "").strip()
                            elif line.startswith("SUBJUDUL:"):
                                subjudul = line.replace("SUBJUDUL:", "").strip()
                            elif line.strip().startswith(("1.", "2.", "3.", "4.", "5.", "-")):
                                points.append(line.strip())

                        slide = prs.slides.add_slide(blank_layout)
                        slide_count += 1

                        # Desain Latar Belakang & Tata Letak
                        if slide_count == 1:
                            # Slide Cover
                            bg = slide.shapes.add_shape(1, 0, 0, Inches(13.333), Inches(7.5))
                            bg.fill.solid()
                            bg.fill.fore_color.rgb = RGBColor(245, 245, 245)
                            bg.line.color.rgb = RGBColor(245, 245, 245)

                            box = slide.shapes.add_shape(1, Inches(1.0), Inches(1.0), Inches(2.2), Inches(1.8))
                            box.fill.solid()
                            box.fill.fore_color.rgb = RGBColor(20, 20, 20)
                            box.line.color.rgb = RGBColor(20, 20, 20)

                            tb = slide.shapes.add_textbox(Inches(1.5), Inches(3.2), Inches(10.0), Inches(3.0))
                            tf = tb.text_frame
                            tf.word_wrap = True

                            p0 = tf.paragraphs[0]
                            p0.text = subjudul.upper() if subjudul else "EXECUTIVE REPORT"
                            p0.font.size = Pt(14)
                            p0.font.bold = True
                            p0.font.color.rgb = RGBColor(100, 100, 100)

                            p1 = tf.add_paragraph()
                            p1.text = judul if judul else "Laporan Performa Kinerja"
                            p1.font.size = Pt(40)
                            p1.font.bold = True
                            p1.font.color.rgb = RGBColor(20, 20, 20)

                            if points:
                                p2 = tf.add_paragraph()
                                p2.text = "\n".join(points)
                                p2.font.size = Pt(16)
                                p2.font.color.rgb = RGBColor(80, 80, 80)
                        else:
                            # Slide Konten
                            tb_title = slide.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(1.0))
                            tf_title = tb_title.text_frame
                            
                            tp0 = tf_title.paragraphs[0]
                            tp0.text = subjudul.upper() if subjudul else "PERFORMANCE ANALYSIS"
                            tp0.font.size = Pt(14)
                            tp0.font.bold = True
                            tp0.font.color.rgb = RGBColor(100, 100, 100)

                            tp1 = tf_title.add_paragraph()
                            tp1.text = judul
                            tp1.font.size = Pt(30)
                            tp1.font.bold = True
                            tp1.font.color.rgb = RGBColor(20, 20, 20)

                            # Kartu Konten
                            card = slide.shapes.add_shape(1, Inches(1.0), Inches(2.2), Inches(11.333), Inches(4.5))
                            card.fill.solid()
                            card.fill.fore_color.rgb = RGBColor(255, 255, 255)
                            card.line.color.rgb = RGBColor(220, 220, 220)

                            tf_card = card.text_frame
                            tf_card.word_wrap = True
                            
                            p_head = tf_card.paragraphs[0]
                            p_head.text = "Poin-Poin Utama Eksekutif:"
                            p_head.font.bold = True
                            p_head.font.size = Pt(18)
                            p_head.font.color.rgb = RGBColor(20, 20, 20)

                            p_body = tf_card.add_paragraph()
                            p_body.text = "\n" + "\n\n".join(points)
                            p_body.font.size = Pt(14)
                            p_body.font.color.rgb = RGBColor(60, 60, 60)

                # Fallback jika slide kosong
                if len(prs.slides) == 0:
                    slide = prs.slides.add_slide(blank_layout)
                    tb = slide.shapes.add_textbox(Inches(1.0), Inches(1.0), Inches(11.333), Inches(5.0))
                    tb.text_frame.text = ai_output

                # 4. Simpan ke File Sementara
                output_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
                prs.save(output_temp.name)

                st.success("✨ Presentasi PowerPoint eksekutif berhasil digenerate!")

                # 5. Tombol Unduh
                with open(output_temp.name, "rb") as f:
                    st.download_button(
                        label="📥 Download File PPTX Eksekutif",
                        data=f,
                        file_name="Executive_Dashboard_Presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses data: {e}")
