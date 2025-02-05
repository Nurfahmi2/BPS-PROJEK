from flask import Flask, render_template, request, send_file, session
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
import io
import os
import subprocess
from datetime import datetime

def convert_docx_to_pdf(input_path, output_path):
    """Konversi DOCX ke PDF menggunakan LibreOffice di Linux."""
    command = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", os.path.dirname(output_path),
        input_path
    ]
    subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

def change_font(doc):
    for para in doc.paragraphs:
        for run in para.runs:
            run.font.name = "Arial"
            run.font.size = Pt(11)
            r = run._element
            r.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

def format_tanggal(tanggal_str):
    bulan_inggris_ke_indonesia = {
        "January": "Januari", "February": "Februari", "March": "Maret",
        "April": "April", "May": "Mei", "June": "Juni",
        "July": "Juli", "August": "Agustus", "September": "September",
        "October": "Oktober", "November": "November", "December": "Desember"
    }
    try:
        dt = datetime.strptime(tanggal_str, '%Y-%m-%d')
        tanggal_format = dt.strftime('%d %B %Y')
        for eng, ind in bulan_inggris_ke_indonesia.items():
            tanggal_format = tanggal_format.replace(eng, ind)
        return tanggal_format
    except ValueError:
        return tanggal_str

app = Flask(__name__)
app.secret_key = "secret_key"

@app.route('/')
def home():
    return render_template("index.html")

@app.route('/generate', methods=['POST'])
def generate():
    nama_petugas = request.form['nama_petugas'].strip().replace(" ", "_")
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    template_path = 'templates/template.docx'
    
    if not os.path.exists(template_path):
        return "Template tidak ditemukan.", 400
    
    doc = Document(template_path)
    change_font(doc)
    
    # Simpan dokumen sementara
    temp_word_path = f"temp_{nama_petugas}_{timestamp}.docx"
    temp_pdf_path = temp_word_path.replace(".docx", ".pdf")
    doc.save(temp_word_path)
    
    # Konversi ke PDF menggunakan LibreOffice
    convert_docx_to_pdf(temp_word_path, temp_pdf_path)
    
    # Kirim file PDF
    return send_file(temp_pdf_path, as_attachment=True, download_name=f"Laporan_{nama_petugas}_{timestamp}.pdf", mimetype="application/pdf")

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
