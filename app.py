import streamlit as st
import os
import tempfile
import json
from google import genai
from pptx import Presentation

# -----------------------------------------------------------------------------
# KONFIGURASI HALAMAN STREAMLIT
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="AI PPTX Dashboard Engine",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AI PPTX Generator - Precision Dashboard Engine")
st.write("Aplikasi membaca data dan memetakan angka-angkanya secara presisi ke dalam elemen visual template Anda.")

# -----------------------------------------------------------------------------
# INPUT PROMPT & FILE UPLOAD
# -----------------------------------------------------------------------------
prompt_text = st.text_area(
    "📝 Instruksi Penyusunan Data:",
    value="Olah data keuangan berikut menjadi laporan kinerja bisnis berformat Executive Dashboard.",
    height=100
)

st.markdown("---")
col1, col2, col3 = st.columns(3)
with col1:
    f_eb = st.file_uploader("Upload File EB", type=["csv", "xlsx", "pdf"])
with col2:
    f_jaskug = st.file_uploader("Upload File Jaskug", type=["csv", "xlsx", "pdf"])
with col3:
    f_ritel = st.file_uploader("Upload File Ritel", type=["csv", "xlsx", "pdf"])

st.markdown("---")
uploaded_template = st.file_uploader("🎨 Upload Template PPTX Asli (Wajib .pptx)", type=["pptx"])

# -----------------------------------------------------------------------------
# EKSEKUSI PENEMPATAN DATA KE TEMPLATE
# -----------------------------------------------------------------------------
if st.button("🚀 Process & Map Data into Template"):
    if not (f_eb or f_jaskug or f_ritel):
        st.warning("⚠️ Mohon unggah minimal satu file data!")
    elif not uploaded_template:
        st.warning("⚠️ Mohon unggah file Template PowerPoint (.pptx) yang ber-layout dashboard!")
    else:
        with st.spinner("🤖 Gemini sedang menganalisis data dan menyuntikkannya ke template..."):
            try:
                api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
                client = genai.Client(api_key=api_key)

                # 1. Upload File Data ke Gemini API
                all_files = [f for f in [f_eb, f_jaskug, f_ritel] if f is not None]
                uploaded_gemini_files = []
                for file_item in all_files:
                    ext = os.path.splitext(file_item.name)[1]
                    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                        tmp.write(file_item.getbuffer())
                        tmp_path = tmp.name
                    g_file = client.files.upload(file=tmp_path)
                    uploaded_gemini_files.append(g_file)

                # 2. Minta Gemini Mengembalikan Struktur Teks Terstruktur
                prompt_instructions = f"""
Anda adalah Business Analyst. Analisis data keuangan dari file yang diunggah berdasarkan instruksi: "{prompt_text}".

Buatkan isi laporan untuk slide presentasi secara ringkas, padat angka, dan presisi.
Format output HARUS persis seperti ini (gunakan pemisah ---SLIDE---):

---SLIDE---
JUDUL: [Judul Utama Slide]
SUBJUDUL: [Subjudul / Periode Data]
KPI_1: [Label KPI 1] | [Nilai Realisasi] | [Target/Growth]
KPI_2: [Label KPI 2] | [Nilai Realisasi] | [Target/Growth]
KPI_3: [Label KPI 3] | [Nilai Realisasi] | [Target/Growth]
KONTEN:
[Poin-poin analisis data, ringkasan per channel, atau tabel ringkas]
"""

                models_to_try = ['gemini-3.5-flash-lite', 'gemini-3.5-flash', 'gemini-2.5-flash']
                response = None
                for m in models_to_try:
                    try:
                        response = client.models.generate_content(
                            model=m,
                            contents=uploaded_gemini_files + [prompt_instructions]
                        )
                        if response and response.text:
                            break
                    except:
                        continue

                if not response or not response.text:
                    st.error("Server Gemini sedang sibuk. Silakan coba lagi nanti.")
                    st.stop()

                # 3. Buka File Template PPTX Asli
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_tmpl:
                    tmp_tmpl.write(uploaded_template.getbuffer())
                    tmpl_path = tmp_tmpl.name

                prs = Presentation(tmpl_path)
                slides_content = [s for s in response.text.split("---SLIDE---") if "JUDUL:" in s]

                # 4. In-Place Replacement: Memasukkan teks Gemini langsung ke shape template asli
                for idx, slide in enumerate(prs.slides):
                    if idx >= len(slides_content):
                        break

                    raw_text = slides_content[idx].strip()
                    lines = raw_text.split("\n")

                    judul = ""
                    body_text = []

                    for line in lines:
                        if line.startswith("JUDUL:"):
                            judul = line.replace("JUDUL:", "").strip()
                        elif line.startswith("SUBJUDUL:"):
                            judul += " - " + line.replace("SUBJUDUL:", "").strip()
                        else:
                            body_text.append(line)

                    # Ambil semua kotak teks yang ada di slide template
                    text_shapes = [shape for shape in slide.shapes if shape.has_text_frame]

                    if text_shapes:
                        # Ganti judul di shape pertama tanpa merusak desain
                        if judul:
                            text_shapes[0].text_frame.text = judul

                        # Ganti isi konten di shape-shape berikutnya
                        if len(text_shapes) > 1 and body_text:
                            text_shapes[1].text_frame.text = "\n".join(body_text)

                # 5. Simpan File Hasil
                output_filename = "Hasil_Dashboard_Presisi.pptx"
                prs.save(output_filename)

                st.success("✨ Berhasil memetakan data Gemini ke dalam template dashboard Anda!")
                with open(output_filename, "rb") as f:
                    st.download_button(
                        label="📥 Download File PPTX (Presisi Template)",
                        data=f,
                        file_name="Laporan_Dashboard_Presisi.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")
