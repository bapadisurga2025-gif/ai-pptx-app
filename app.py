import streamlit as st
import os
import tempfile
import json
import time
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

st.set_page_config(page_title="Executive PPTX Generator", layout="wide")
st.title("📊 Executive Dashboard PPTX Generator")

uploaded_file = st.file_uploader("Upload File Laporan Keuangan (PDF/XLSX)", type=["pdf", "xlsx", "csv"])

# -----------------------------------------------------------------------------
# FUNGSI UNTUK MERANCANG LAYOUT DASHBOARD PPTX
# -----------------------------------------------------------------------------
def make_dashboard_pptx(data, output_path="output_dashboard.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Warna
    COLOR_NAVY = RGBColor(0, 32, 96)
    COLOR_ORANGE = RGBColor(255, 29, 0)
    COLOR_BG_CARD = RGBColor(240, 244, 248)
    COLOR_BORDER = RGBColor(217, 225, 232)
    COLOR_WHITE = RGBColor(255, 255, 255)
    COLOR_TEXT = RGBColor(30, 30, 30)

    # 1. Header Banner
    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
    header.fill.solid()
    header.fill.fore_color.rgb = COLOR_NAVY
    header.line.fill.background()
    tf_h = header.text_frame
    tf_h.margin_left, tf_h.margin_top = Inches(0.5), Inches(0.2)
    
    p = tf_h.paragraphs[0]
    p.text = "EXECUTIVE FINANCIAL PERFORMANCE DASHBOARD"
    p.font.size, p.font.bold, p.font.color.rgb = Pt(20), True, COLOR_WHITE

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.1), Inches(13.333), Inches(0.06))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_ORANGE
    line.line.fill.background()

    # 2. Top KPI Cards
    kpis = data.get("kpis", [])
    for i, kpi in enumerate(kpis[:3]):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5 + i * 4.2), Inches(1.35), Inches(3.9), Inches(1.1))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_BG_CARD
        card.line.color.rgb = COLOR_NAVY
        tf = card.text_frame
        tf.margin_left, tf.margin_top = Inches(0.2), Inches(0.15)
        
        p1 = tf.paragraphs[0]
        p1.text, p1.font.size, p1.font.bold, p1.font.color.rgb = kpi.get("label", ""), Pt(9), True, COLOR_NAVY
        
        p2 = tf.add_paragraph()
        p2.text, p2.font.size, p2.font.bold = kpi.get("val", ""), Pt(16), True
        
        p3 = tf.add_paragraph()
        p3.text, p3.font.size = f"Growth: {kpi.get('growth', '')}", Pt(10)

    # 3. Native Table
    table_rows = data.get("table_data", [])
    if table_rows:
        table_shape = slide.shapes.add_table(len(table_rows) + 1, 5, Inches(0.5), Inches(2.6), Inches(12.333), Inches(2.4))
        table = table_shape.table
        
        headers = ["Segmen Pendapatan & Beban", "Realisasi 2025", "Realisasi 2026", "GAP (Nominal)", "Growth (%)"]
        for c, h in enumerate(headers):
            cell = table.cell(0, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_NAVY
            p = cell.text_frame.paragraphs[0]
            p.text, p.font.size, p.font.bold, p.font.color.rgb = h, Pt(10), True, COLOR_WHITE

        for r, row in enumerate(table_rows):
            for c, val in enumerate(row):
                cell = table.cell(r + 1, c)
                p = cell.text_frame.paragraphs[0]
                p.text, p.font.size = str(val), Pt(9.5)

    # 4. Panel Bawah (Insight & Strategy)
    p_ins = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(5.15), Inches(6.0), Inches(2.05))
    p_ins.fill.solid()
    p_ins.fill.fore_color.rgb = COLOR_WHITE
    tf_ins = p_ins.text_frame
    p = tf_ins.paragraphs[0]
    p.text, p.font.bold, p.font.color.rgb = "💡 INSIGHT UTAMA KINERJA", True, COLOR_NAVY
    for item in data.get("insights", []):
        p = tf_ins.add_paragraph()
        p.text, p.font.size = f"• {item}", Pt(9)

    p_str = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.833), Inches(5.15), Inches(6.0), Inches(2.05))
    p_str.fill.solid()
    p_str.fill.fore_color.rgb = COLOR_WHITE
    tf_str = p_str.text_frame
    p = tf_str.paragraphs[0]
    p.text, p.font.bold, p.font.color.rgb = "🚀 REKOMENDASI STRATEGIS", True, COLOR_ORANGE
    for item in data.get("strategies", []):
        p = tf_str.add_paragraph()
        p.text, p.font.size = f"✓ {item}", Pt(9)

    prs.save(output_path)

# -----------------------------------------------------------------------------
# PROSES UTAMA STREAMLIT DENGAN HANDLING FILE PROCESSING
# -----------------------------------------------------------------------------
if uploaded_file and st.button("🚀 Buat Dashboard PPTX"):
    with st.spinner("🤖 Mengunggah dan memproses dokumen PDF..."):
        api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
        if not api_key:
            st.error("GEMINI_API_KEY belum terpasang di Secrets Streamlit.")
            st.stop()
            
        client = genai.Client(api_key=api_key)
        
        ext = os.path.splitext(uploaded_file.name)[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp.write(uploaded_file.getbuffer())
            tmp_path = tmp.name
        
        # 1. Upload File
        g_file = client.files.upload(file=tmp_path)

        # 2. Polling: Tunggu sampai File Siap Digunakan (PROCESSING -> ACTIVE)
        while g_file.state.name == "PROCESSING":
            time.sleep(2)
            g_file = client.files.get(name=g_file.name)

        if g_file.state.name == "FAILED":
            st.error("Gagal memproses file PDF pada server Gemini API.")
            st.stop()

        prompt_json = """
        Ekstrak data keuangan dari dokumen ini dan kembalikan HANYA JSON murni dengan format sebagai berikut:
        {
            "kpis": [
                {"label": "TOTAL PENDAPATAN (ATRIBUSI)", "val": "Rp 3.009.712.487", "growth": "-10.63% YoY"},
                {"label": "BEBAN OPERASIONAL TOTAL", "val": "Rp 3.025.557.914", "growth": "-13.43% YoY"},
                {"label": "EBITDA", "val": "Rp (128.816.667)", "growth": "-48.67% YoY"}
            ],
            "table_data": [
                ["Pendapatan Suratpos & Paketpos", "670.161.335", "639.476.109", "-30.685.226", "-4.58%"],
                ["Pendapatan Jaskug & Ritel", "1.062.885.559", "1.089.368.557", "+26.482.998", "+2.49%"],
                ["Pendapatan Logistik", "758.722.804", "348.292.780", "-410.430.024", "-54.09%"],
                ["Pendapatan Lainnya", "875.745.445", "932.575.041", "+56.829.596", "+6.49%"],
                ["Beban Tenaga Kerja", "(1.933.472.858)", "(1.707.309.309)", "+226.163.549", "-11.70%"],
                ["Beban Operasi", "(1.444.390.205)", "(1.044.810.287)", "+399.579.918", "-27.66%"]
            ],
            "insights": [
                "Pertumbuhan positif dicapai oleh Jaskug & Ritel (+2.49%) dan Pendapatan Lainnya (+6.49%).",
                "Kontraksi utama pada Logistik (-54.09%) menjadi pemicu utama penurunan total revenue.",
                "Efisiensi beban berhasil ditekan signifikan: Beban Operasi (-27.66%)."
            ],
            "strategies": [
                "Restrukturisasi Komersial Logistik: Peninjauan ulang SLA dan skema kontrak B2B.",
                "Akselerasi Channel Ritel & Jaskug: Penambahan kuota agen baru.",
                "Pengendalian Beban Ketat: Pertahankan tren efisiensi beban operasi."
            ]
        }
        """

        try:
            # Gunakan structured output config untuk memastikan output berbentuk JSON murni
            res = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=[g_file, prompt_json],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json"
                )
            )
            
            parsed_data = json.loads(res.text)

            # Buat PowerPoint
            output_filename = "Hasil_Executive_Dashboard.pptx"
            make_dashboard_pptx(parsed_data, output_filename)

            st.success("✨ Slide Executive Dashboard Berhasil Dibuat!")
            with open(output_filename, "rb") as f:
                st.download_button("📥 Download File PPTX", f, file_name=output_filename)

        except Exception as e:
            st.error(f"Gagal mengolah data: {str(e)}")
        finally:
            # Hapus temp file lokal
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
