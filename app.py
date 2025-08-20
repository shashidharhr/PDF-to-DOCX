from flask import Flask, render_template, request, send_file
import fitz  # PyMuPDF
from docx import Document
import os
import camelot
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
from openpyxl import Workbook

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

        # Get conversion type from form
        conversion_type = request.form.get("conversion_type")

        if conversion_type == "word":
            docx_filename = pdf_file.filename.rsplit('.', 1)[0] + '.docx'
            docx_path = os.path.join(UPLOAD_FOLDER, docx_filename)

            # Extract text
            doc = Document()
            with fitz.open(pdf_path) as pdf:
                for page in pdf:
                    text = page.get_text()
                    if text.strip():
                        doc.add_paragraph(text)

            doc.save(docx_path)
            return send_file(docx_path, as_attachment=True)

        elif conversion_type == "excel":
            excel_filename = pdf_file.filename.rsplit('.', 1)[0] + '.xlsx'
            excel_path = os.path.join(UPLOAD_FOLDER, excel_filename)

            try:
                tables = camelot.read_pdf(pdf_path, pages='all')
                writer = pd.ExcelWriter(excel_path, engine='openpyxl')
                for i, table in enumerate(tables):
                    table.df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
                writer.close()
            except Exception as e:
                return f"Error extracting tables: {e}"

            return send_file(excel_path, as_attachment=True)

        elif conversion_type == "ocr":
            docx_filename = pdf_file.filename.rsplit('.', 1)[0] + '_ocr.docx'
            docx_path = os.path.join(UPLOAD_FOLDER, docx_filename)

            images = convert_from_path(pdf_path)
            doc = Document()

            for img in images:
                text = pytesseract.image_to_string(img)
                if text.strip():
                    doc.add_paragraph(text)

            doc.save(docx_path)
            return send_file(docx_path, as_attachment=True)

        else:
            return "Invalid conversion type selected"

    return render_template('index.html')


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True, port=8000)

