from flask import Flask, render_template, request, send_file, session
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
import io
import os
import subprocess
from datetime import datetime

app = Flask(__name__)
app.secret_key = "secret_key"  
TEMPLATE_FOLDER = 'templates'

# Fungsi untuk mengonversi Word ke PDF menggunakan LibreOffice di Linux
def convert_to_pdf_linux(word_mem, output_pdf_path):
    temp_word_path = "temp.docx"
    with open(temp_word_path, "wb") as temp_file:
        temp_file.write(word_mem.getvalue())
    
    # Gunakan LibreOffice untuk mengonversi ke PDF
    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", temp_word_path, "--outdir", os.getcwd()])
    
    # Baca hasil PDF
    with open(output_pdf_path, "rb") as pdf_file:
        pdf_mem = io.BytesIO(pdf_file.read())
    
    # Hapus file sementara
    os.remove(temp_word_path)
    os.remove(output_pdf_path)
    
    return pdf_mem

@app.route('/generate', methods=['POST'])
def generate():
    nama_petugas = request.form['nama_petugas'].strip().replace(" ", "_")
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    template_path = os.path.join(TEMPLATE_FOLDER, 'template.docx')
    
    if not os.path.exists(template_path):
        return "Template tidak ditemukan."
    
    doc = Document(template_path)
    
    word_mem = io.BytesIO()
    doc.save(word_mem)
    word_mem.seek(0)
    
    if 'generate_pdf' in request.form:
        output_pdf_path = f"Laporan_Pendataan_{nama_petugas}_{timestamp}.pdf"
        pdf_mem = convert_to_pdf_linux(word_mem, output_pdf_path)
        pdf_mem.seek(0)
        return send_file(pdf_mem, as_attachment=True, download_name=output_pdf_path, mimetype="application/pdf")
    
    return send_file(word_mem, as_attachment=True, download_name=f"Laporan_Pendataan_{nama_petugas}_{timestamp}.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
