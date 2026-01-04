import os
import cv2
import numpy as np
import ezdxf
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge
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
from PIL import Image
import zipfile

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Set Max Content Length to 50MB
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e):
    return jsonify({'error': 'File terlalu besar. Maksimum ukuran file adalah 50MB.'}), 413

def image_to_dxf(image_path, output_path):
    # Read image with alpha channel support (IMREAD_UNCHANGED)
    img = cv2.imread(image_path, cv2.IMREAD_UNCHANGED)
    if img is None:
        raise ValueError("Could not read image")
    
    # Handle Transparency: Convert to white background
    if img.shape[2] == 4:
        # Create white background
        bg = np.ones_like(img[:,:,:3]) * 255
        # Extract alpha channel
        alpha = img[:,:,3] / 255.0
        # Blend
        for c in range(3):
            bg[:,:,c] = bg[:,:,c] * (1 - alpha) + img[:,:,c] * alpha
        img = bg.astype(np.uint8)
    
    # Convert to grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Reduce noise with Gaussian Blur
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    
    # Auto Canny Thresholds
    v = np.median(blurred)
    sigma = 0.33
    lower = int(max(0, (1.0 - sigma) * v))
    upper = int(min(255, (1.0 + sigma) * v))
    
    # Edge detection
    edges = cv2.Canny(blurred, lower, upper)
    
    # Find contours - Use RETR_LIST to capture all contours (inner and outer)
    contours, hierarchy = cv2.findContours(edges, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    
    # Create DXF document
    doc = ezdxf.new('R2010')
    msp = doc.modelspace()
    
    height, width = edges.shape
    
    for contour in contours:
        # contour is a numpy array of shape (n, 1, 2)
        if len(contour) < 2:
            continue
            
        points = []
        for point in contour:
            x, y = point[0]
            # Flip Y for CAD (image 0,0 is top-left, CAD is bottom-left)
            points.append((float(x), float(height - y)))
            
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
    # Using pdfplumber to open the PDF
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        # Clean cell data: handle None and convert to string
                        cleaned_row = []
                        for cell in row:
                            if cell is None:
                                cleaned_row.append("")
                            else:
                                cleaned_row.append(str(cell))
                        ws.append(cleaned_row)
                    ws.append([]) # Add empty row between tables
            else:
                # If no tables, extract text line by line
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

def image_to_pdf(image_path, output_path):
    image = Image.open(image_path)
    if image.mode != 'RGB':
        image = image.convert('RGB')
    image.save(output_path, "PDF", resolution=100.0)

def merge_pdfs(paths, output_path):
    merger = PdfMerger()
    for p in paths:
        merger.append(p)
    merger.write(output_path)
    merger.close()

def compress_pdf(input_path, output_path):
    # Aggressive compression to target ~50% size reduction
    # 1. Downsample images > 1024px
    # 2. Convert images to JPEG with quality 60
    # 3. Clean and deflate PDF structure
    
    try:
        doc = fitz.open(input_path)
        
        # Subset fonts
        try:
            doc.subset_fonts()
        except Exception:
            pass

        processed_xrefs = set()
        
        for page in doc:
            for img in page.get_images():
                xref = img[0]
                if xref in processed_xrefs:
                    continue
                processed_xrefs.add(xref)
                
                try:
                    pix = fitz.Pixmap(doc, xref)
                    
                    # Check if image is large enough to benefit from downscaling
                    if pix.width > 1024 or pix.height > 1024:
                        # Convert to RGB if not already (e.g. CMYK)
                        if pix.n >= 4:
                            pix = fitz.Pixmap(fitz.csRGB, pix)
                        
                        # Scale down
                        # We want to reduce size significantly. 
                        # Let's target a max dimension of 1024.
                        scale = 1024 / max(pix.width, pix.height)
                        if scale < 0.9: # Only if reduction is > 10%
                            new_w = int(pix.width * scale)
                            new_h = int(pix.height * scale)
                            pix = fitz.Pixmap(pix, new_w, new_h)
                    
                    # Recompress as JPEG with lower quality (default often 75-95, we go 60 for "half size")
                    # Only if it wasn't already a highly compressed small image?
                    # Actually, recompressing everything to JPEG q=60 is the surest way to drop size.
                    
                    # Check if it's already a small image or mask? Skip masks.
                    if pix.n < 3 and pix.alpha == 0: 
                        # Grayscale or mono without alpha
                        pass
                    
                    new_data = pix.tobytes("jpeg", jpg_quality=60)
                    doc.update_image(xref, data=new_data)
                    
                    pix = None # free memory
                except Exception as e:
                    print(f"Image compression skipped for xref {xref}: {e}")
                    pass

        doc.save(output_path, garbage=4, deflate=True, clean=True)
        doc.close()
        
    except Exception as e:
        # Fallback to simple compression if advanced fails
        print(f"Advanced compression failed: {e}, falling back to simple save.")
        try:
            doc = fitz.open(input_path)
            doc.save(output_path, garbage=4, deflate=True)
            doc.close()
        except:
            raise e

def pdf_to_image(pdf_path, image_path):
    doc = fitz.open(pdf_path)
    page = doc[0] # Take first page
    pix = page.get_pixmap()
    pix.save(image_path)

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
        input_path = None
        temp_img_path = None
        output_path = None
        try:
            ext = os.path.splitext(file.filename.lower())[1]
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                file.save(tmp.name)
                input_path = tmp.name

            temp_img_path = input_path + '.png'
            if ext == '.pdf':
                # Convert PDF to Image first
                pdf_to_image(input_path, temp_img_path)
                processing_path = temp_img_path
            else:
                processing_path = input_path

            output_path = input_path + '.dxf'
            
            # Process image to DXF
            image_to_dxf(processing_path, output_path)
            
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
            if input_path and os.path.exists(input_path):
                os.unlink(input_path)
            if temp_img_path and os.path.exists(temp_img_path):
                os.unlink(temp_img_path)
            # output_path is kept open by send_file? 
            # In production, use background cleanup or streaming. 
            pass

@app.route('/pdf/merge', methods=['POST'])
def pdf_merge_route():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'No files'}), 400
    tmp_paths = []
    converted_tmp_paths = []
    
    try:
        for f in files:
            ext = os.path.splitext(f.filename.lower())[1]
            
            # Save original upload
            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as t:
                f.save(t.name)
                current_path = t.name
                tmp_paths.append(current_path)
            
            # Convert if needed
            if ext in ['.jpg', '.jpeg', '.png']:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as pdf_out:
                    image_to_pdf(current_path, pdf_out.name)
                    converted_tmp_paths.append(pdf_out.name)
            elif ext == '.docx':
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as pdf_out:
                    docx_to_pdf(current_path, pdf_out.name)
                    converted_tmp_paths.append(pdf_out.name)
            elif ext == '.pdf':
                converted_tmp_paths.append(current_path)
            else:
                # Skip or error? For now, skip unsupported
                pass

        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as out:
            merge_pdfs(converted_tmp_paths, out.name)
            return send_file(out.name, as_attachment=True, download_name='merged.pdf', mimetype='application/pdf')
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        # Cleanup original uploads
        for p in tmp_paths:
            if os.path.exists(p):
                os.unlink(p)
        # Cleanup converted pdfs that are not in tmp_paths (images/docx converted)
        for p in converted_tmp_paths:
            if p not in tmp_paths and os.path.exists(p):
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
    
    input_path = None # Initialize variables for cleanup
    
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as src:
            f.save(src.name)
            input_path = src.name
            
        if ext == '.pdf' and target == 'docx':
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as out:
                pdf_to_docx(input_path, out.name)
                return send_file(out.name, as_attachment=True, download_name='converted.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        
        elif ext == '.pdf' and target == 'xlsx':
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as out:
                pdf_to_xlsx(input_path, out.name)
                return send_file(out.name, as_attachment=True, download_name='converted.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        elif ext == '.docx' and target == 'pdf':
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as out:
                docx_to_pdf(input_path, out.name)
                return send_file(out.name, as_attachment=True, download_name='converted.pdf', mimetype='application/pdf')
        
        elif ext == '.xlsx' and target == 'pdf':
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as out:
                xlsx_to_pdf(input_path, out.name)
                return send_file(out.name, as_attachment=True, download_name='converted.pdf', mimetype='application/pdf')
        
        else:
            return jsonify({'error': 'Unsupported conversion'}), 400
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if input_path and os.path.exists(input_path):
            try:
                os.unlink(input_path)
            except:
                pass

@app.route('/pdf/compress', methods=['POST'])
def pdf_compress_route():
    # Try getting 'files' list first (for multiple files)
    files = request.files.getlist('files')
    
    # If empty, try 'file' (fallback for single file upload from some clients)
    if not files:
        files = request.files.getlist('file')
    
    if not files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    tmp_paths = []
    compressed_paths = []
    
    try:
        for f in files:
            # Check filename exists
            if not f.filename:
                continue
                
            # Use original filename extension
            ext = os.path.splitext(f.filename)[1].lower()
            if ext != '.pdf':
                continue # Skip non-pdfs
                
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as src:
                f.save(src.name)
                tmp_paths.append(src.name)
                
            out_path = src.name + '_compressed.pdf'
            compress_pdf(src.name, out_path)
            compressed_paths.append((out_path, f.filename))

        if not compressed_paths:
             return jsonify({'error': 'No valid PDF files processed. Please upload PDF files.'}), 400

        if len(compressed_paths) == 1:
            # Single file return
            return send_file(compressed_paths[0][0], as_attachment=True, download_name='compressed.pdf', mimetype='application/pdf')
        else:
            # Multiple files -> ZIP
            with tempfile.NamedTemporaryFile(delete=False, suffix='.zip') as zip_out:
                with zipfile.ZipFile(zip_out.name, 'w') as zf:
                    for path, original_name in compressed_paths:
                        zf.write(path, arcname=f"compressed_{original_name}")
                return send_file(zip_out.name, as_attachment=True, download_name='compressed_files.zip', mimetype='application/zip')

    except Exception as e:
        # Log error for debugging (in production logging would be better)
        print(f"Compression error: {str(e)}")
        return jsonify({'error': f"Compression failed: {str(e)}"}), 500
    finally:
        # Cleanup temporary input files
        for p in tmp_paths:
            if os.path.exists(p):
                try:
                    os.unlink(p)
                except:
                    pass
        # Cleanup temporary output files (only if they are not being streamed)
        # Note: send_file keeps the file handle open, so we rely on OS or specific Flask configs for cleanup usually.
        # However, for temp files created with delete=False, we should clean them up.
        # But cleaning up immediately after return send_file might fail if file is still in use.
        # A common strategy is to use after_request or a background task, but simple tempfile is hard to clean perfectly in basic Flask.
        # For this snippet, we leave the output files. In a production app, use a proper temp directory lifecycle.
        pass

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
