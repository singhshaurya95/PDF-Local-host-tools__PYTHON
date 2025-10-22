import os
import zipfile
import uuid
import pythoncom
from flask import Flask, render_template, request, send_file
from PyPDF2 import PdfReader, PdfWriter
from pdf2docx import Converter

# Try importing docx2pdf (Windows/Mac with MS Word)
try:
    from docx2pdf import convert as docx_to_pdf
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
TEMPLATES_FOLDER = os.path.join(BASE_DIR, 'templates')
STATIC_FOLDER = os.path.join(BASE_DIR, 'static')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__, template_folder=TEMPLATES_FOLDER, static_folder=STATIC_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


@app.route('/')
def index():
    return render_template('index.html', docx2pdf_available=DOCX2PDF_AVAILABLE)


# ---------- MERGE PDFs ----------
@app.route('/merge', methods=['POST'])
def merge():
    files = request.files.getlist('pdfs')
    if not files:
        return "❌ No files uploaded", 400

    merger = PdfWriter()
    file_count = 0
    for file in files:
        if file and file.filename.lower().endswith('.pdf'):
            reader = PdfReader(file)
            for page in reader.pages:
                merger.add_page(page)
            file_count += 1

    if file_count == 0:
        return "❌ No valid PDF files provided", 400

    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'merged_{uuid.uuid4().hex}.pdf')
    with open(output_path, 'wb') as f:
        merger.write(f)

    return send_file(output_path, as_attachment=True)


# ---------- Parse Page Ranges ----------
def parse_page_numbers(pages_str, total_pages):
    pages = set()
    for part in pages_str.split(','):
        part = part.strip()
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                for p in range(start, end + 1):
                    if 1 <= p <= total_pages:
                        pages.add(p)
            except ValueError:
                continue
        elif part.isdigit():
            p = int(part)
            if 1 <= p <= total_pages:
                pages.add(p)
    return sorted(pages)


# ---------- SPLIT PDFs ----------
@app.route('/split', methods=['POST'])
def split():
    file = request.files.get('pdf')
    pages_str = request.form.get('pages', '').strip()

    if not file or not file.filename.lower().endswith('.pdf'):
        return "❌ Invalid file", 400

    reader = PdfReader(file)
    total_pages = len(reader.pages)

    if not pages_str:
        return "❌ Please enter pages to split", 400

    page_numbers = parse_page_numbers(pages_str, total_pages)
    if not page_numbers:
        return f"❌ Invalid page selection. Total pages: {total_pages}", 400

    output_files = []
    for p in page_numbers:
        writer = PdfWriter()
        writer.add_page(reader.pages[p - 1])
        split_path = os.path.join(app.config['UPLOAD_FOLDER'], f'page_{p}_{uuid.uuid4().hex}.pdf')
        with open(split_path, 'wb') as f:
            writer.write(f)
        output_files.append(split_path)

    zip_path = os.path.join(app.config['UPLOAD_FOLDER'], f'split_pages_{uuid.uuid4().hex}.zip')
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for pdf in output_files:
            zipf.write(pdf, os.path.basename(pdf))

    return send_file(zip_path, as_attachment=True)


# ---------- PDF → WORD ----------
@app.route('/pdf_to_word', methods=['POST'])
def pdf_to_word():
    file = request.files.get('pdf')
    if not file or not file.filename.lower().endswith('.pdf'):
        return "❌ Invalid file", 400

    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{uuid.uuid4().hex}.pdf')
    file.save(pdf_path)
    docx_path = os.path.splitext(pdf_path)[0] + '.docx'

    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
    except Exception as e:
        return f"❌ Error converting PDF to Word: {str(e)}", 500

    return send_file(docx_path, as_attachment=True)


# ---------- WORD → PDF ----------
@app.route('/word_to_pdf', methods=['POST'])
def word_to_pdf():
    if not DOCX2PDF_AVAILABLE:
        return "❌ Word-to-PDF not available on this system (requires Windows/Mac with MS Word installed)", 500

    file = request.files.get('word')
    if not file or not file.filename.lower().endswith('.docx'):
        return "❌ Invalid Word file", 400

    word_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{uuid.uuid4().hex}.docx')
    file.save(word_path)
    pdf_path = os.path.splitext(word_path)[0] + '.pdf'

    try:
        pythoncom.CoInitialize()  # FIX for COM init
        docx_to_pdf(word_path, pdf_path)
    except Exception as e:
        if "corrupted" in str(e).lower():
            return "❌ The .docx file appears corrupted or invalid. Open it in MS Word, save again, and retry.", 400
        return f"❌ Error converting Word to PDF: {str(e)}", 500
    finally:
        pythoncom.CoUninitialize()

    return send_file(pdf_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
