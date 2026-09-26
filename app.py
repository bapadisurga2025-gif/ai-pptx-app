import os
from flask import Flask, render_template, request, send_file
from google import genai
from pptx import Presentation

app = Flask(__name__)

# Inisialisasi Gemini Client
# Di server online nanti, GEMINI_API_KEY akan diambil dari Environment Variables
client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    # Menerima data dari form web
    prompt = request.form.get('prompt', '')
    data_file = request.files.get('dataFile')
    pdf_file = request.files.get('pdfFile')
    photo_file = request.files.get('photoFile')
    
    # Ekstraksi teks sederhana dari file yang di-upload jika ada
    file_context = ""
    if data_file:
        file_context += f"\n[Data File: {data_file.filename}]"
    if pdf_file:
        file_context += f"\n[PDF File: {pdf_file.filename}]"

    # Gabungkan prompt dengan konteks file
    full_prompt = f"{prompt}\n{file_context}"

    try:
        # Panggil Gemini AI untuk membuat struktur presentasi
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=f"Buatkan kerangka materi presentasi (Judul Slide dan Poin-poin isi) berdasarkan instruksi berikut: {full_prompt}"
        )
        ai_output = response.text
    except Exception as e:
        ai_output = f"Gagal memproses dengan Gemini AI: {str(e)}"

    # Generate file .pptx otomatis berdasarkan hasil AI
    prs = Presentation()
    
    # Tambahkan Slide Judul
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Hasil Presentasi AI"
    slide.placeholders[1].text = "Dibuat otomatis via Gemini AI"

    # Tambahkan Slide Konten dari AI
    content_layout = prs.slide_layouts[1]
    content_slide = prs.slides.add_slide(content_layout)
    content_slide.shapes.title.text = "Ringkasan Materi"
    content_slide.placeholders[1].text = ai_output

    output_path = "output_presentation.pptx"
    prs.save(output_path)

    # Kirim file ke browser untuk didownload otomatis
    return send_file(output_path, as_attachment=True, download_name="Presentasi_AI.pptx")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)