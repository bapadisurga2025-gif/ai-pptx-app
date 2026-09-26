import streamlit as st
import os
import tempfile
from google import genai

# -----------------------------------------------------------------------------
# KONFIGURASI HALAMAN STREAMLIT
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="AI PPTX Generator (Full Gemini Engine)",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AI PPTX Generator - Full Gemini Engine")
st.write("Aplikasi mengunggah file data, foto, dan template langsung ke Gemini AI. Gemini menganalisis seluruh konteks dan merancang file presentasi.")

# -----------------------------------------------------------------------------
# INPUT PROMPT / INSTRUKSI EKSEKUTIF
# -----------------------------------------------------------------------------
st.subheader("📝 Prompt / Instruksi Khusus EM")
prompt_text = st.text_area(
    "Masukkan instruksi penyusunan presentasi:",
    value="Buat laporan eksekutif performa kinerja keuangan berdasarkan data yang diunggah. Adaptasi gaya visual, tata letak, dan skema warna dari template yang diunggah.",
    height=120,
    key="prompt_input"
)

st.markdown("---")
st.subheader("📁 Upload Data, Foto, & Template Acuan")

col1, col2 = st.columns(2)
with col1:
    f_eb = st.file_uploader("Upload File EB", type=["csv", "xlsx", "pdf"], key="eb_f")
with col2:
    p_eb = st.file_uploader("Upload Foto EB", type=["png", "jpg", "jpeg"], key="eb_p")

col3, col4 = st.columns(2)
with col3:
    f_jaskug = st.file_uploader("Upload File Jaskug", type=["csv", "xlsx", "pdf"], key="jaskug_f")
with col4:
    p_jaskug = st.file_uploader("Upload Foto Jaskug", type=["png", "jpg", "jpeg"], key="jaskug_p")

col5, col6 = st.columns(2)
with col5:
    f_ritel = st.file_uploader("Upload File Ritel", type=["csv", "xlsx", "pdf"], key="ritel_f")
with col6:
    p_ritel = st.file_uploader("Upload Foto Ritel", type=["png", "jpg", "jpeg"], key="ritel_p")

st.markdown("---")
uploaded_template = st.file_uploader("Upload Template Contoh / Referensi Visual (.pdf atau .pptx)", type=["pdf", "pptx"], key="template_file")

st.markdown("---")

# -----------------------------------------------------------------------------
# PROSES GENERATE VIA GEMINI MULTIMODAL (WITH AUTOMATIC FALLBACK)
# -----------------------------------------------------------------------------
if st.button("🚀 Proses Semua di Gemini AI & Generate PPTX"):
    if not (f_eb or f_jaskug or f_ritel):
        st.warning("⚠️ Mohon unggah setidaknya satu file data sumber!")
    elif not prompt_text.strip():
        st.warning("⚠️ Mohon masukkan instruksi prompt terlebih dahulu!")
    else:
        with st.spinner("🤖 Mengirim file ke Gemini AI untuk dianalisis dan dirakit menjadi presentasi..."):
            try:
                api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
                if not api_key:
                    st.error("🔑 GEMINI_API_KEY tidak ditemukan pada st.secrets.")
                    st.stop()

                client = genai.Client(api_key=api_key)

                # 1. Upload File ke Gemini API (File API)
                uploaded_files_gemini = []
                all_inputs = [f_eb, p_eb, f_jaskug, p_jaskug, f_ritel, p_ritel, uploaded_template]
                
                for file_item in all_inputs:
                    if file_item is not None:
                        ext = os.path.splitext(file_item.name)[1]
                        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                            tmp.write(file_item.getbuffer())
                            tmp_path = tmp.name
                        
                        g_file = client.files.upload(file=tmp_path)
                        uploaded_files_gemini.append(g_file)

                # 2. Instruksi ke Gemini untuk menghasilkan skrip python-pptx yang presisi
                prompt_instructions = f"""
Anda adalah pakar desain presentasi korporat dan pemrograman Python.
Pengguna memberikan berkas data, foto pendukung, serta berkas contoh template/referensi visual.

Instruksi khusus pengguna:
"{prompt_text}"

Tugas Anda:
1. Pelajari data keuangan dari berkas yang diunggah.
2. Analisis gaya desain, warna, dan tata letak dari berkas template referensi yang diunggah.
3. Hasilkan KODE PYTHON LENGKAP menggunakan library `python-pptx` yang membuat file presentasi bernama `output_presentation.pptx`.
4. Berikan HANYA kode Python di dalam pembungkus ```python ...
