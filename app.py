import os
import cv2
import numpy as np
import ezdxf
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import tempfile
from PyPDF2 import PdfMerger
from pdf2docx import Converter
import pdfplumber
from openpyxl import Workbook, load_workbook
from docx import Document
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
import fitz

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

def image_to_dxf(image_path, output_path):
    # Read image
    img = cv2.imread(image_path)
    if img is None:
        raise ValueError("Could not read image")
    
    # Convert to grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Edge detection using Canny
    # Adjust thresholds as needed
    edges = cv2.Canny(gray, 50, 150)
    
    # Find contours
    contours, hierarchy = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # Create DXF document
    doc = ezdxf.new('R2010')
    msp = doc.modelspace()
    
    height, width = edges.shape
    
    for contour in contours:
        # contour is a numpy array of shape (n, 1, 2)
        # ezdxf expects a list of (x, y) tuples
        # We need to flip Y because image coordinates (0,0) is top-left, 
        # while CAD (0,0) is usually bottom-left or center.
        # Here we just mirror Y to keep orientation consistent visually
        points = []
        for point in contour:
            x, y = point[0]
            points.append((float(x), float(height - y)))
            
        if len(points) > 1:
            msp.add_lwpolyline(points)
            
    doc.saveas(output_path)

def pdf_to_docx(pdf_path, output_path):
    cv = Converter(pdf_path)
    cv.convert(output_path)
    cv.close()

def pdf_to_xlsx(pdf_path, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "PDF"
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        ws.append([str(cell) if cell else "" for cell in row])
                    ws.append([])
            else:
                text = page.extract_text() or ""
                for line in text.splitlines():
                    ws.append([line])
                ws.append([])
    wb.save(output_path)

def docx_to_pdf(docx_path, output_path):
    d = Document(docx_path)
    c = canvas.Canvas(output_path, pagesize=A4)
    w, h = A4
    y = h - 50
    for p in d.paragraphs:
        t = p.text
        for line in t.split("\n"):
            c.drawString(50, y, line)
            y -= 14
            if y < 50:
                c.showPage()
                y = h - 50
    c.save()

def xlsx_to_pdf(xlsx_path, output_path):
    wb = load_workbook(xlsx_path, data_only=True)
    doc = SimpleDocTemplate(output_path, pagesize=landscape(A4))
    elements = []
    for sheet in wb.worksheets:
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append(["" if v is None else str(v) for v in row])
        if data:
            tbl = Table(data)
            tbl.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('FONTSIZE', (0,0), (-1,-1), 8),
            ]))
            elements.append(tbl)
    doc.build(elements)

def merge_pdfs(paths, output_path):
    merger = PdfMerger()
    for p in paths:
        merger.append(p)
    merger.write(output_path)
    merger.close()

def compress_pdf(input_path, output_path):
    doc = fitz.open(input_path)
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()

@app.route('/')
def index():
    return "Backend is running!"

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
        
    if file:
        try:
            # Create temporary files for processing
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_img:
                file.save(tmp_img.name)
                input_path = tmp_img.name
                
            output_path = input_path + '.dxf'
            
            # Process image to DXF
            image_to_dxf(input_path, output_path)
            
            # Send file back
            return send_file(
                output_path,
                as_attachment=True,
                download_name=f"{os.path.splitext(file.filename)[0]}.dxf",
                mimetype='application/dxf'
            )
            
        except Exception as e:
            return jsonify({'error': str(e)}), 500
            
        finally:
            # Cleanup temp files
            if 'input_path' in locals() and os.path.exists(input_path):
                os.unlink(input_path)
            # We don't delete output_path immediately if we stream it? 
            # send_file usually keeps file open. 
            # For simplicity in this script, we might leave a temp file or use a cleanup approach.
            # But Flask's send_file doesn't auto-delete.
            # For a production app, consider using a background task or byte stream.
            # Here we will just let OS temp cleanup handle it eventually or leave it for now.
            pass

@app.route('/pdf/merge', methods=['POST'])
def pdf_merge_route():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'No files'}), 400
    tmp_paths = []
    try:
        for f in files:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as t:
                f.save(t.name)
                tmp_paths.append(t.name)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as out:
            merge_pdfs(tmp_paths, out.name)
            return send_file(out.name, as_attachment=True, download_name='merged.pdf', mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        for p in tmp_paths:
            if os.path.exists(p):
                os.unlink(p)

@app.route('/pdf/convert', methods=['POST'])
def pdf_convert_route():
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    target = request.form.get('target')
    f = request.files['file']
    if not target:
        return jsonify({'error': 'Missing target'}), 400
    ext = os.path.splitext(f.filename.lower())[1]
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as src:
            f.save(src.name)
            input_path = src.name
        if ext == '.pdf' and target == 'docx':
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as out:
                pdf_to_docx(input_path, out.name)
                return send_file(out.name, as_attachment=True, download_name='converted.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        if ext == '.pdf' and target == 'xlsx':
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as out:
                pdf_to_xlsx(input_path, out.name)
                return send_file(out.name, as_attachment=True, download_name='converted.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        if ext == '.docx' and target == 'pdf':
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as out:
                docx_to_pdf(input_path, out.name)
                return send_file(out.name, as_attachment=True, download_name='converted.pdf', mimetype='application/pdf')
        if ext == '.xlsx' and target == 'pdf':
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as out:
                xlsx_to_pdf(input_path, out.name)
                return send_file(out.name, as_attachment=True, download_name='converted.pdf', mimetype='application/pdf')
        return jsonify({'error': 'Unsupported conversion'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if 'input_path' in locals() and os.path.exists(input_path):
            os.unlink(input_path)

@app.route('/pdf/compress', methods=['POST'])
def pdf_compress_route():
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    f = request.files['file']
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as src:
            f.save(src.name)
            input_path = src.name
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as out:
            compress_pdf(input_path, out.name)
            return send_file(out.name, as_attachment=True, download_name='compressed.pdf', mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if 'input_path' in locals() and os.path.exists(input_path):
            os.unlink(input_path)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
