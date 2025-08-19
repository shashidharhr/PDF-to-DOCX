from flask import Flask, render_template, request, send_file
import fitz  # PyMuPDF
from docx import Document
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'pdf_file' not in request.files:
            return "No file uploaded"

        pdf_file = request.files['pdf_file']
        if pdf_file.filename == '':
            return "No selected file"

        pdf_path = os.path.join(UPLOAD_FOLDER, pdf_file.filename)
        pdf_file.save(pdf_path)

        docx_filename = pdf_file.filename.rsplit('.', 1)[0] + '.docx'
        docx_path = os.path.join(UPLOAD_FOLDER, docx_filename)

        # Extract text and write to DOCX
        doc = Document()
        with fitz.open(pdf_path) as pdf:
            for page in pdf:
                text = page.get_text()
                if text.strip():  # avoid blank paragraphs
                    doc.add_paragraph(text)

        doc.save(docx_path)

        return send_file(docx_path, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True, port=8000)

