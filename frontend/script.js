from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import os, uuid, zipfile, io, re, shutil, subprocess, tempfile, time, json, traceback

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageEnhance, ImageFilter, ImageDraw, ImageFont
import pdfplumber
from docx import Document
from difflib import unified_diff


try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None
try:
    import requests
except Exception:
    requests = None
try:
    from deep_translator import GoogleTranslator
except Exception:
    GoogleTranslator = None

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
except Exception:
    Workbook = None
    load_workbook = None
    Font = PatternFill = Border = Side = Alignment = None
    get_column_letter = None
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
except Exception:
    Presentation = None

try:
    from pdf2docx import Converter
except Exception:
    Converter = None

try:
    from docxcompose.composer import Composer
except Exception:
    Composer = None

try:
    import pypdfium2 as pdfium
except Exception:
    pdfium = None

try:
    import pytesseract
except Exception:
    pytesseract = None

BASE = Path(__file__).resolve().parent
UPLOADS = BASE / 'uploads'
OUTPUTS = BASE / 'outputs'
UPLOADS.mkdir(exist_ok=True)
OUTPUTS.mkdir(exist_ok=True)

app = Flask(__name__)
# Universal CORS setup: Netlify/mobile/PC/tablet sab origins allow.
# 500 error par bhi browser ko readable JSON error milega.
ALLOWED_ORIGINS = [
    'https://orbixapdftool.in',
    'https://www.orbixapdftool.in',
    'http://localhost:5500',
    'http://127.0.0.1:5500',
    'http://localhost:3000',
    'http://127.0.0.1:5000',
]
CORS(app, resources={r"/*": {
    "origins": "*",
    "methods": ["GET", "POST", "OPTIONS"],
    "allow_headers": ["Content-Type", "Authorization", "X-Requested-With"],
    "expose_headers": ["Content-Disposition", "X-Original-Size", "X-Compressed-Size", "X-Target-Size", "X-Target-Reached", "X-Exact-Target"],
    "supports_credentials": False,
    "max_age": 86400,
}})
app.config['MAX_CONTENT_LENGTH'] = 150 * 1024 * 1024

@app.before_request
def handle_preflight():
    if request.method == 'OPTIONS':
        resp = app.make_default_options_response()
        resp.headers['Access-Control-Allow-Origin'] = request.headers.get('Origin', '*')
        resp.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization, X-Requested-With'
        resp.headers['Access-Control-Expose-Headers'] = 'Content-Disposition, X-Original-Size, X-Compressed-Size, X-Target-Size, X-Target-Reached, X-Exact-Target'
        return resp

@app.after_request
def add_cors_headers(response):
    origin = request.headers.get('Origin', '')
    # Allow specific known origins or fallback to wildcard
    if origin and any(origin == o for o in ALLOWED_ORIGINS):
        response.headers['Access-Control-Allow-Origin'] = origin
    else:
        response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Vary'] = 'Origin'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization, X-Requested-With'
    response.headers['Access-Control-Expose-Headers'] = 'Content-Disposition, X-Original-Size, X-Compressed-Size, X-Target-Size, X-Target-Reached, X-Exact-Target'
    response.headers['Access-Control-Max-Age'] = '86400'
    response.headers['Cache-Control'] = response.headers.get('Cache-Control', 'no-store')
    return response

@app.errorhandler(413)
def too_large(e):
    resp = jsonify(error='File too large. Please upload a smaller file.')
    resp.headers['Access-Control-Allow-Origin'] = request.headers.get('Origin', '*')
    return resp, 413

@app.errorhandler(Exception)
def handle_all_errors(e):
    traceback.print_exc()
    msg = str(e) or 'Internal server error'
    resp = jsonify(error=msg)
    resp.headers['Access-Control-Allow-Origin'] = request.headers.get('Origin', '*')
    return resp, 500

def uid(name='file'):
    return f"{uuid.uuid4().hex}_{secure_filename(name)}"

def save_file(field='file'):
    f = request.files.get(field)
    if not f:
        raise ValueError('File not uploaded')
    p = UPLOADS / uid(f.filename)
    f.save(p)
    return p, f.filename

def send(path, download_name=None):
    return send_file(path, as_attachment=True, download_name=download_name or Path(path).name)

def parse_pages(page_str, total):
    if not page_str:
        return list(range(total))
    pages = set()
    for part in page_str.replace(' ', '').split(','):
        if not part: continue
        if '-' in part:
            a,b = part.split('-',1)
            a,b = int(a), int(b)
            for n in range(a,b+1):
                if 1 <= n <= total: pages.add(n-1)
        else:
            n = int(part)
            if 1 <= n <= total: pages.add(n-1)
    return sorted(pages)

@app.get('/')
def home():
    return jsonify(status='ok', message='ToolMitra backend ready')

@app.get('/api/health')
def health():
    return jsonify(status='ok', backend='ToolMitra', pdf2docx=bool(Converter), tesseract=bool(pytesseract))

@app.route('/api/pdf-to-word', methods=['POST', 'OPTIONS'])
def pdf_to_word():
    """
    Advanced PDF to Word converter.
    Keeps the PDF layout as close as possible and gives editable DOCX output.
    """
    if request.method == 'OPTIONS':
        return ('', 204)

    if Converter is None:
        resp = jsonify(error='PDF to Word engine is not installed. Add pdf2docx in requirements.txt and redeploy.')
        resp.headers['Access-Control-Allow-Origin'] = request.headers.get('Origin', '*')
        return resp, 500

    pdf = None
    cv = None

    try:
        pdf, original = save_file('file')
        size = pdf.stat().st_size

        if size <= 0:
            resp = jsonify(error='Uploaded PDF is empty.')
            resp.headers['Access-Control-Allow-Origin'] = request.headers.get('Origin', '*')
            return resp, 400

        if size > 80 * 1024 * 1024:
            resp = jsonify(error='PDF too large. Please upload below 80MB.')
            resp.headers['Access-Control-Allow-Origin'] = request.headers.get('Origin', '*')
            return resp, 413

        # Password / damaged PDF check before conversion
        try:
            reader = PdfReader(str(pdf))
            if getattr(reader, 'is_encrypted', False):
                try:
                    reader.decrypt('')
                except Exception:
                    resp = jsonify(error='This PDF is password protected. Please unlock it first.')
                    resp.headers['Access-Control-Allow-Origin'] = request.headers.get('Origin', '*')
                    return resp, 400

            if len(reader.pages) > 300:
                resp = jsonify(error='PDF has too many pages. Please split it into smaller PDFs first.')
                resp.headers['Access-Control-Allow-Origin'] = request.headers.get('Origin', '*')
                return resp, 413
        except Exception:
            # pdf2docx may still handle some PDFs better than pypdf, so do not stop here.
            pass

        safe_stem = secure_filename(Path(original).stem) or 'converted_pdf'
        out = OUTPUTS / f"{safe_stem}_editable_word.docx"

        # Main conversion: preserves layout, fonts, tables, images and spacing better
        # than simple text extraction.
        # MEMORY FIX: If PDF is > 1 page, process page by page to avoid OOM on 512MB server.
        total_pages = len(reader.pages)
        
        if total_pages > 1 and Composer is not None:
            import gc
            out_docs = []
            for i in range(total_pages):
                tmp_pdf = OUTPUTS / f"temp_{uuid.uuid4().hex}.pdf"
                page_writer = PdfWriter()
                page_writer.add_page(reader.pages[i])
                with open(tmp_pdf, 'wb') as fp:
                    page_writer.write(fp)
                
                tmp_docx = OUTPUTS / f"temp_{uuid.uuid4().hex}.docx"
                cv = Converter(str(tmp_pdf))
                try:
                    cv.convert(str(tmp_docx), start=0, end=None, multi_processing=False)
                except TypeError:
                    cv.convert(str(tmp_docx), start=0, end=None)
                cv.close()
                cv = None
                
                if tmp_docx.exists() and tmp_docx.stat().st_size > 0:
                    out_docs.append(tmp_docx)
                
                tmp_pdf.unlink(missing_ok=True)
                gc.collect()
            
            if out_docs:
                master = Document(str(out_docs[0]))
                composer = Composer(master)
                for tmp_docx in out_docs[1:]:
                    sub_doc = Document(str(tmp_docx))
                    composer.append(sub_doc)
                composer.save(str(out))
                
                for tmp_docx in out_docs:
                    tmp_docx.unlink(missing_ok=True)
        else:
            cv = Converter(str(pdf))
            try:
                cv.convert(str(out), start=0, end=None, multi_processing=False)
            except TypeError:
                cv.convert(str(out), start=0, end=None)
            cv.close()
            cv = None

        if not out.exists() or out.stat().st_size <= 0:
            resp = jsonify(error='Conversion failed. Please try another PDF.')
            resp.headers['Access-Control-Allow-Origin'] = request.headers.get('Origin', '*')
            return resp, 500

        return send_file(
            out,
            as_attachment=True,
            download_name=out.name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        traceback.print_exc()
        resp = jsonify(error=f'PDF to Word failed: {str(e)}')
        resp.headers['Access-Control-Allow-Origin'] = request.headers.get('Origin', '*')
        return resp, 500

    finally:
        try:
            if cv is not None:
                cv.close()
        except Exception:
            pass

        try:
            if pdf and Path(pdf).exists():
                Path(pdf).unlink(missing_ok=True)
        except Exception:
            pass


@app.post('/api/merge-pdf')
def merge_pdf():
    files = request.files.getlist('files')
    if len(files) < 2:
        return jsonify(error='Upload at least 2 PDFs'), 400
    writer = PdfWriter()
    for f in files:
        p = UPLOADS / uid(f.filename); f.save(p)
        reader = PdfReader(str(p))
        for page in reader.pages: writer.add_page(page)
    out = OUTPUTS / 'merged_toolmitra.pdf'
    with open(out, 'wb') as fp: writer.write(fp)
    return send(out, 'merged_toolmitra.pdf')

@app.post('/api/split-pdf')
def split_pdf():
    pdf, original = save_file('file')
    reader = PdfReader(str(pdf))
    pages = parse_pages(request.form.get('pages',''), len(reader.pages))
    zpath = OUTPUTS / (Path(original).stem + '_split_pages.zip')
    with zipfile.ZipFile(zpath, 'w') as z:
        for idx in pages:
            writer = PdfWriter(); writer.add_page(reader.pages[idx])
            tmp = OUTPUTS / f'page_{idx+1}.pdf'
            with open(tmp, 'wb') as fp: writer.write(fp)
            z.write(tmp, tmp.name); tmp.unlink(missing_ok=True)
    return send(zpath, zpath.name)

@app.post('/api/extract-pages')
def extract_pages():
    pdf, original = save_file('file')
    reader = PdfReader(str(pdf))
    pages = parse_pages(request.form.get('pages','1'), len(reader.pages))
    writer = PdfWriter()
    for idx in pages: writer.add_page(reader.pages[idx])
    out = OUTPUTS / (Path(original).stem + '_selected_pages.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return send(out, out.name)

@app.post('/api/delete-pages')
def delete_pages():
    pdf, original = save_file('file')
    reader = PdfReader(str(pdf))
    delete = set(parse_pages(request.form.get('pages',''), len(reader.pages)))
    writer = PdfWriter()
    for i,p in enumerate(reader.pages):
        if i not in delete: writer.add_page(p)
    out = OUTPUTS / (Path(original).stem + '_pages_deleted.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return send(out, out.name)

@app.post('/api/rotate-pdf')
def rotate_pdf():
    pdf, original = save_file('file')
    angle = int(request.form.get('angle', '90'))
    reader = PdfReader(str(pdf)); writer = PdfWriter()
    pages = set(parse_pages(request.form.get('pages',''), len(reader.pages)))
    for i, page in enumerate(reader.pages):
        if i in pages: page.rotate(angle)
        writer.add_page(page)
    out = OUTPUTS / (Path(original).stem + '_rotated.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return send(out, out.name)

@app.post('/api/protect-pdf')
def protect_pdf():
    pdf, original = save_file('file')
    password = request.form.get('password','1234')
    reader = PdfReader(str(pdf)); writer = PdfWriter()
    for page in reader.pages: writer.add_page(page)
    writer.encrypt(password)
    out = OUTPUTS / (Path(original).stem + '_protected.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return send(out, out.name)

@app.post('/api/unlock-pdf')
def unlock_pdf():
    pdf, original = save_file('file')
    password = request.form.get('password','')
    reader = PdfReader(str(pdf))
    if reader.is_encrypted:
        try: reader.decrypt(password)
        except Exception: return jsonify(error='Wrong password'), 400
    writer = PdfWriter()
    for page in reader.pages: writer.add_page(page)
    out = OUTPUTS / (Path(original).stem + '_unlocked.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return send(out, out.name)

@app.post('/api/extract-text')
def extract_text():
    pdf, original = save_file('file')
    out = OUTPUTS / (Path(original).stem + '_text.txt')
    text=[]
    with pdfplumber.open(str(pdf)) as p:
        for i, page in enumerate(p.pages,1):
            text.append(f'--- Page {i} ---\n{page.extract_text() or ""}\n')
    out.write_text('\n'.join(text), encoding='utf-8')
    return send(out, out.name)

@app.post('/api/jpg-to-pdf')
def jpg_to_pdf():
    files = request.files.getlist('files') or ([request.files['file']] if 'file' in request.files else [])
    if not files: return jsonify(error='Upload images'), 400
    images=[]
    for f in files:
        p = UPLOADS / uid(f.filename); f.save(p)
        img = Image.open(p).convert('RGB')
        images.append(img)
    out = OUTPUTS / 'images_to_pdf.pdf'
    images[0].save(out, save_all=True, append_images=images[1:])
    return send(out, out.name)

@app.post('/api/images-to-pdf-cleaner')
def scan_to_pdf():
    files = request.files.getlist('files') or ([request.files['file']] if 'file' in request.files else [])
    imgs=[]
    for f in files:
        p=UPLOADS/uid(f.filename); f.save(p)
        img=Image.open(p).convert('L')
        img=ImageEnhance.Contrast(img).enhance(1.7).filter(ImageFilter.SHARPEN).convert('RGB')
        imgs.append(img)
    out=OUTPUTS/'clean_scanned_pdf.pdf'
    imgs[0].save(out, save_all=True, append_images=imgs[1:])
    return send(out, out.name)

@app.post('/api/image-compress')
def image_compress():
    """Image compressor with target KB support.
    - 500 error fix: all bad input/edge cases return JSON instead of crashing.
    - Target KB: output is made <= target, then padded to exact target bytes when possible.
    - If target is too small for readable image, it returns the closest possible small image.
    """
    try:
        imgp, original = save_file('file')
        quality = max(5, min(95, int(float(request.form.get('quality', '80') or 80))))
        fmt = (request.form.get('output_format', 'auto') or 'auto').lower().strip()
        target_kb_raw = (request.form.get('target_kb', '') or '').strip()
        resize_width_raw = (request.form.get('resize_width', '') or '').strip()

        from PIL import ImageOps
        img = Image.open(imgp)
        img = ImageOps.exif_transpose(img)

        # Optional resize width
        if resize_width_raw:
            try:
                target_w = int(float(resize_width_raw))
                if target_w >= 50 and img.width != target_w:
                    target_h = max(1, int(img.height * (target_w / img.width)))
                    img = img.resize((target_w, target_h), Image.LANCZOS)
            except Exception:
                pass

        has_alpha = img.mode in ('RGBA', 'LA') or ('transparency' in img.info)
        if fmt == 'auto':
            old_ext = Path(original).suffix.lower().replace('.', '')
            # JPG is best for exact target size and small output. Keep PNG only if alpha is needed.
            fmt = 'png' if old_ext == 'png' and has_alpha else ('webp' if old_ext == 'webp' else 'jpg')
        if fmt == 'jpeg':
            fmt = 'jpg'
        if fmt not in ('jpg', 'png', 'webp'):
            fmt = 'jpg'

        ext = fmt
        pil_fmt = {'jpg': 'JPEG', 'png': 'PNG', 'webp': 'WEBP'}[fmt]
        out = OUTPUTS / f"{Path(original).stem}_compressed.{ext}"

        def flatten_for_jpg(im):
            if im.mode in ('RGBA', 'LA') or ('transparency' in im.info):
                rgba = im.convert('RGBA')
                bg = Image.new('RGB', rgba.size, (255, 255, 255))
                bg.paste(rgba, mask=rgba.split()[-1])
                return bg
            return im.convert('RGB')

        def save_to_bytes(im, q=80, colors=256):
            bio = io.BytesIO()
            if pil_fmt == 'JPEG':
                im2 = flatten_for_jpg(im)
                im2.save(bio, 'JPEG', quality=max(5, min(95, int(q))), optimize=True, progressive=True, subsampling=2)
            elif pil_fmt == 'WEBP':
                im2 = im.convert('RGBA') if (im.mode in ('RGBA', 'LA') or ('transparency' in im.info)) else im.convert('RGB')
                im2.save(bio, 'WEBP', quality=max(5, min(95, int(q))), method=6)
            else:
                # PNG quality does not work like JPG. Palette colors reduce size.
                if im.mode in ('RGBA', 'LA') or ('transparency' in im.info):
                    im2 = im.convert('RGBA').quantize(colors=max(8, min(256, int(colors))), method=Image.Quantize.FASTOCTREE)
                else:
                    im2 = im.convert('RGB').quantize(colors=max(8, min(256, int(colors))), method=Image.Quantize.MEDIANCUT)
                im2.save(bio, 'PNG', optimize=True, compress_level=9)
            return bio.getvalue()

        def best_under_limit(im, limit_bytes):
            """Return largest bytes <= limit for current dimensions, or smallest if impossible."""
            candidates = []
            if pil_fmt in ('JPEG', 'WEBP'):
                lo, hi = 5, 95
                best = None
                smallest = None
                for _ in range(12):
                    mid = (lo + hi) // 2
                    b = save_to_bytes(im, mid, 256)
                    if smallest is None or len(b) < len(smallest):
                        smallest = b
                    if len(b) <= limit_bytes:
                        if best is None or len(b) > len(best):
                            best = b
                        lo = mid + 1
                    else:
                        hi = mid - 1
                return best or smallest
            else:
                best = None
                smallest = None
                for colors in [256, 192, 128, 96, 64, 48, 32, 24, 16, 8]:
                    b = save_to_bytes(im, quality, colors)
                    if smallest is None or len(b) < len(smallest):
                        smallest = b
                    if len(b) <= limit_bytes and (best is None or len(b) > len(best)):
                        best = b
                return best or smallest

        def pad_exact(data, limit_bytes):
            # Target se chhota ho to exact KB banane ke liye safe padding add karte hain.
            # Browsers/viewers generally ignore trailing bytes for JPG/PNG/WEBP downloads.
            if len(data) < limit_bytes:
                data += b'\0' * (limit_bytes - len(data))
            return data

        if target_kb_raw:
            limit = max(1, int(float(target_kb_raw))) * 1024
            work = img.copy()
            data = best_under_limit(work, limit)

            # If still over target, reduce dimensions until it fits.
            attempts = 0
            while len(data) > limit and work.width > 80 and work.height > 80 and attempts < 40:
                ratio = limit / max(len(data), 1)
                scale = max(0.55, min(0.92, (ratio ** 0.5) * 0.96))
                nw = max(80, int(work.width * scale))
                nh = max(80, int(work.height * scale))
                if nw >= work.width or nh >= work.height:
                    nw = max(80, int(work.width * 0.85))
                    nh = max(80, int(work.height * 0.85))
                if nw == work.width and nh == work.height:
                    break
                work = work.resize((nw, nh), Image.LANCZOS)
                data = best_under_limit(work, limit)
                attempts += 1

            # Exact KB when the compressed file is <= target.
            if len(data) <= limit:
                data = pad_exact(data, limit)
            out.write_bytes(data)
        else:
            out.write_bytes(save_to_bytes(img, quality, 256))

        resp = send(out, out.name)
        if target_kb_raw:
            resp.headers['X-Target-Size'] = str(max(1, int(float(target_kb_raw))) * 1024)
            resp.headers['X-Output-Size'] = str(out.stat().st_size)
        return resp
    except Exception as e:
        traceback.print_exc()
        return jsonify(error=f'Image compressor error: {str(e)}'), 500

@app.post('/api/add-watermark')
def add_watermark():
    pdf, original = save_file('file')
    text = request.form.get('text','ToolMitra')
    reader = PdfReader(str(pdf)); writer = PdfWriter()
    for page in reader.pages:
        w = float(page.mediabox.width); h = float(page.mediabox.height)
        packet = io.BytesIO(); c = canvas.Canvas(packet, pagesize=(w,h))
        c.setFont('Helvetica-Bold', 42); c.setFillAlpha(0.18)
        c.translate(w/2,h/2); c.rotate(35); c.drawCentredString(0,0,text); c.save(); packet.seek(0)
        wm = PdfReader(packet).pages[0]
        page.merge_page(wm); writer.add_page(page)
    out = OUTPUTS / (Path(original).stem + '_watermarked.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return send(out, out.name)

@app.post('/api/page-numbers')
def page_numbers():
    pdf, original = save_file('file')
    reader = PdfReader(str(pdf)); writer = PdfWriter()
    total = len(reader.pages)
    for i,page in enumerate(reader.pages,1):
        w=float(page.mediabox.width); h=float(page.mediabox.height)
        packet=io.BytesIO(); c=canvas.Canvas(packet,pagesize=(w,h))
        c.setFont('Helvetica', 11); c.drawCentredString(w/2, 18, f'{i} / {total}'); c.save(); packet.seek(0)
        overlay=PdfReader(packet).pages[0]; page.merge_page(overlay); writer.add_page(page)
    out=OUTPUTS/(Path(original).stem+'_page_numbers.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return send(out,out.name)

@app.post('/api/word-to-pdf')
def word_to_pdf():
    docx, original = save_file('file')
    doc = Document(str(docx))
    out = OUTPUTS / (Path(original).stem + '_converted.pdf')
    c = canvas.Canvas(str(out), pagesize=A4)
    width, height = A4; y = height - 50
    c.setFont('Helvetica', 11)
    for para in doc.paragraphs:
        line = para.text or ' '
        chunks = [line[i:i+95] for i in range(0, len(line), 95)] or ['']
        for chunk in chunks:
            c.drawString(45, y, chunk)
            y -= 16
            if y < 45:
                c.showPage(); c.setFont('Helvetica', 11); y = height - 50
    c.save()
    return send(out, out.name)

@app.post('/api/html-to-pdf')
def html_to_pdf():
    text = request.form.get('html','')
    clean = re.sub('<[^<]+?>', '', text).replace('&nbsp;', ' ')
    out = OUTPUTS / 'html_to_pdf.pdf'
    c = canvas.Canvas(str(out), pagesize=A4)
    width, height = A4; y = height - 50
    c.setFont('Helvetica', 11)
    for line in clean.splitlines() or [clean]:
        for chunk in [line[i:i+95] for i in range(0,len(line),95)] or ['']:
            c.drawString(45,y,chunk); y-=16
            if y<45: c.showPage(); c.setFont('Helvetica',11); y=height-50
    c.save(); return send(out,out.name)

@app.post('/api/compress-pdf')
def compress_pdf():
    """Compress PDF with optional target size in KB.
    User can enter target_kb (example: 200). The backend will try multiple
    quality/resolution combinations and return the smallest readable PDF it can make.
    Note: For very small targets, exact size is not always possible without making pages unreadable.
    """
    pdf, original = save_file('file')
    level = (request.form.get('level') or request.form.get('quality') or 'high').lower().strip()
    if level not in {'low', 'medium', 'high', 'extreme'}:
        level = 'high'

    target_kb_raw = (request.form.get('target_kb') or request.form.get('target') or '').strip()
    target_bytes = None
    try:
        if target_kb_raw:
            target_bytes = max(30 * 1024, int(float(target_kb_raw)) * 1024)
    except Exception:
        target_bytes = None

    original_size = os.path.getsize(pdf)
    candidates = []

    # Lossless compression first. Best for text/vector PDFs.
    try:
        reader = PdfReader(str(pdf), strict=False)
        writer = PdfWriter()
        for page in reader.pages:
            try:
                page.compress_content_streams()
            except Exception:
                pass
            writer.add_page(page)
        lossless = OUTPUTS / (Path(original).stem + '_lossless.pdf')
        with open(lossless, 'wb') as fp:
            writer.write(fp)
        candidates.append(lossless)
    except Exception:
        pass

    # Visual/image rebuild. This actually reduces scanned/image-heavy PDFs.
    if pdfium is not None:
        # Normal level presets
        level_presets = {
            'low':     [(1.60, 78), (1.40, 70), (1.20, 62)],
            'medium':  [(1.35, 66), (1.15, 56), (1.00, 48), (0.90, 42)],
            'high':    [(1.15, 56), (1.00, 46), (0.85, 36), (0.72, 30)],
            'extreme': [(0.95, 42), (0.78, 32), (0.65, 24), (0.55, 18)],
        }[level]

        # If user entered target KB, try stronger steps until target is reached.
        if target_bytes:
            level_presets = [
                (1.35, 68), (1.20, 60), (1.05, 52), (0.95, 44),
                (0.85, 36), (0.75, 30), (0.65, 24), (0.58, 20),
                (0.50, 16), (0.44, 13), (0.38, 10), (0.32, 8)
            ]

        try:
            doc = pdfium.PdfDocument(str(pdf))
            for attempt, (scale, jpeg_quality) in enumerate(level_presets, 1):
                out = OUTPUTS / f"{Path(original).stem}_compressed_try_{attempt}.pdf"
                c = canvas.Canvas(str(out))
                temp_imgs = []

                for i in range(len(doc)):
                    page = doc[i]
                    width_pt, height_pt = page.get_size()
                    bitmap = page.render(scale=scale).to_pil().convert('RGB')

                    # For extreme small target, also resize very large images.
                    max_side = 1800 if not target_bytes else max(650, int(1400 * scale))
                    if max(bitmap.size) > max_side:
                        bitmap.thumbnail((max_side, max_side), Image.LANCZOS)

                    img_path = OUTPUTS / f"{uuid.uuid4().hex}_compress_page_{i+1}.jpg"
                    bitmap.save(
                        img_path,
                        'JPEG',
                        quality=int(jpeg_quality),
                        optimize=True,
                        progressive=True,
                    )
                    temp_imgs.append(img_path)

                    c.setPageSize((width_pt, height_pt))
                    c.drawImage(ImageReader(str(img_path)), 0, 0, width=width_pt, height=height_pt)
                    c.showPage()

                c.save()
                for img in temp_imgs:
                    try: img.unlink()
                    except Exception: pass

                candidates.append(out)
                out_size = os.path.getsize(out)

                # Stop only when target reached, otherwise keep trying smaller.
                if target_bytes and out_size <= target_bytes:
                    break
                if not target_bytes and out_size < original_size * 0.80:
                    break
        except Exception as e:
            pass

    if not candidates:
        return jsonify(error='Compression failed. Run: pip install pypdfium2 pillow reportlab pypdf'), 500

    exact_match = False
    target_reached = False

    if target_bytes:
        under = [x for x in candidates if os.path.getsize(x) <= target_bytes]
        # Choose closest file that is under target. If every candidate is bigger, choose smallest possible.
        best = max(under, key=lambda x: os.path.getsize(x)) if under else min(candidates, key=lambda x: os.path.getsize(x))
    else:
        best = min(candidates, key=lambda x: os.path.getsize(x))

    final = OUTPUTS / (Path(original).stem + '_compressed.pdf')
    final.write_bytes(Path(best).read_bytes())

    # If target KB is possible, make the download EXACTLY that many bytes.
    # We only pad when the compressed PDF is smaller than target; PDF readers normally ignore trailing comment bytes.
    if target_bytes:
        current_size = os.path.getsize(final)
        if current_size <= target_bytes:
            target_reached = True
            padding_needed = target_bytes - current_size
            if padding_needed > 0:
                with open(final, 'ab') as fp:
                    if padding_needed >= 2:
                        fp.write(b'\n%')
                        fp.write(b'0' * (padding_needed - 2))
                    else:
                        fp.write(b'0')
            exact_match = os.path.getsize(final) == target_bytes

    response = send(final, final.name)
    response.headers['X-Original-Size'] = str(original_size)
    response.headers['X-Compressed-Size'] = str(os.path.getsize(final))
    if target_bytes:
        response.headers['X-Target-Size'] = str(target_bytes)
        response.headers['X-Target-Reached'] = 'yes' if target_reached else 'no'
        response.headers['X-Exact-Target'] = 'yes' if exact_match else 'no'
    return response


def pdf_text_pages(pdf_path):
    pages = []
    with pdfplumber.open(str(pdf_path)) as p:
        for page in p.pages:
            pages.append(page.extract_text() or "")
    return pages

def add_text_overlay(pdf_path, original, text_value, page_no=1, x=70, y=70, size=20):
    reader = PdfReader(str(pdf_path)); writer = PdfWriter()
    for i, page in enumerate(reader.pages, 1):
        if i == int(page_no):
            w=float(page.mediabox.width); h=float(page.mediabox.height)
            packet=io.BytesIO(); c=canvas.Canvas(packet, pagesize=(w,h))
            c.setFont('Helvetica-Bold', int(size)); c.setFillColorRGB(0.90,0.10,0.10)
            c.drawString(float(x), float(y), text_value[:120])
            c.save(); packet.seek(0)
            overlay=PdfReader(packet).pages[0]; page.merge_page(overlay)
        writer.add_page(page)
    out=OUTPUTS/(Path(original).stem+'_edited.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return out

@app.post('/api/pdf-to-ppt')
def pdf_to_ppt():
    """Convert every PDF page into one PowerPoint slide.
    This keeps the PDF page design visually the same, like iLovePDF-style output.
    Each slide contains a high-quality image of the original PDF page.
    """
    if Presentation is None:
        return jsonify(error='python-pptx is not installed. Run: pip install python-pptx'), 400
    if pdfium is None:
        return jsonify(error='pypdfium2 is not installed. Run: pip install pypdfium2'), 400

    pdf, original = save_file('file')
    out = OUTPUTS / (Path(original).stem + '_converted.pptx')

    prs = Presentation()
    blank = prs.slide_layouts[6]

    # 4:3 page size works well for bank statements and A4 pages as image slides.
    # The image is fitted inside slide while preserving aspect ratio.
    slide_w = prs.slide_width
    slide_h = prs.slide_height

    doc = pdfium.PdfDocument(str(pdf))
    scale = float(request.form.get('quality', '2.0') or 2.0)  # 2.0 = sharp, not too heavy

    for i in range(len(doc)):
        page = doc[i]
        bitmap = page.render(scale=scale).to_pil()
        img_path = OUTPUTS / f'{uuid.uuid4().hex}_page_{i+1}.png'
        bitmap.save(img_path, 'PNG', optimize=True)

        slide = prs.slides.add_slide(blank)

        img_w_px, img_h_px = bitmap.size
        img_ratio = img_w_px / img_h_px
        slide_ratio = slide_w / slide_h

        if img_ratio > slide_ratio:
            pic_w = slide_w
            pic_h = int(slide_w / img_ratio)
            left = 0
            top = int((slide_h - pic_h) / 2)
        else:
            pic_h = slide_h
            pic_w = int(slide_h * img_ratio)
            left = int((slide_w - pic_w) / 2)
            top = 0

        slide.shapes.add_picture(str(img_path), left, top, width=pic_w, height=pic_h)
        try:
            img_path.unlink()
        except Exception:
            pass

    prs.save(out)
    return send(out, out.name)

def _clean_cell(value):
    if value is None:
        return ''
    return re.sub(r'\s+', ' ', str(value).replace('\n', ' ')).strip()

def _to_number(value):
    value = _clean_cell(value).replace(',', '')
    if not value:
        return ''
    try:
        return float(value)
    except Exception:
        return value

def _is_date(value):
    return bool(re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', _clean_cell(value)))

def _extract_pdf_transactions(pdf_path):
    """Extract bank-statement style rows into structured Excel-ready records."""
    rows = []
    account_info = []
    date_re = re.compile(r'^(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}/\d{1,2}/\d{4})\s+(.+)$')
    money_re = re.compile(r'\d+(?:,\d{3})*(?:\.\d{2})')

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ''
            for line in text.splitlines()[:18]:
                if any(k in line for k in ['Account Number', 'Account Name', 'Branch', 'CIF NO', 'Balance as on', 'Statement of']):
                    account_info.append(_clean_cell(line))

            # First preference: actual PDF table extraction. This keeps columns like iLovePDF output.
            try:
                tables = page.extract_tables({
                    'vertical_strategy': 'lines',
                    'horizontal_strategy': 'lines',
                    'intersection_tolerance': 6,
                    'snap_tolerance': 4,
                    'join_tolerance': 4,
                    'edge_min_length': 20,
                    'min_words_vertical': 1,
                    'min_words_horizontal': 1,
                }) or []
            except Exception:
                tables = []

            for table in tables:
                for r in table:
                    if not r or len(r) < 5:
                        continue
                    r = [_clean_cell(c) for c in r]
                    if not r or not _is_date(r[0]):
                        continue
                    while len(r) < 6:
                        r.append('')
                    txn_date, value_date, desc, debit, credit, balance = r[:6]
                    # Some extractors shift columns; normalize based on DR/CR and money count when needed.
                    rows.append({
                        'Page': page_no,
                        'Txn Date': txn_date,
                        'Value Date': value_date,
                        'Description': desc,
                        'Debit': _to_number(debit),
                        'Credit': _to_number(credit),
                        'Balance': _to_number(balance),
                    })

            # Fallback for PDFs where tables are not detected but text is selectable.
            if not any(r['Page'] == page_no for r in rows):
                pending = None
                for raw in text.splitlines():
                    line = _clean_cell(raw)
                    m = date_re.match(line)
                    if m:
                        if pending:
                            rows.append(pending)
                        txn_date, value_date, rest = m.groups()
                        amounts = money_re.findall(rest)
                        desc = rest
                        debit = credit = balance = ''
                        if amounts:
                            balance = amounts[-1]
                            if len(amounts) >= 2:
                                amt = amounts[-2]
                                if '/CR/' in rest.upper() or ' CR ' in rest.upper():
                                    credit = amt
                                else:
                                    debit = amt
                            desc = money_re.sub('', rest).strip()
                        pending = {
                            'Page': page_no,
                            'Txn Date': txn_date,
                            'Value Date': value_date,
                            'Description': desc,
                            'Debit': _to_number(debit),
                            'Credit': _to_number(credit),
                            'Balance': _to_number(balance),
                        }
                    elif pending and line and 'computer generated' not in line.lower():
                        pending['Description'] = (pending['Description'] + ' ' + line).strip()
                if pending:
                    rows.append(pending)

    # Deduplicate rows caused by table + text overlap.
    seen = set()
    clean_rows = []
    for r in rows:
        key = (r['Txn Date'], r['Value Date'], r['Description'][:80], str(r['Debit']), str(r['Credit']), str(r['Balance']))
        if key not in seen:
            clean_rows.append(r)
            seen.add(key)
    return clean_rows, account_info

@app.post('/api/pdf-to-excel')
def pdf_to_excel():
    if Workbook is None:
        return jsonify(error='openpyxl is not installed. Run pip install openpyxl'), 400
    pdf, original = save_file('file')
    rows, account_info = _extract_pdf_transactions(pdf)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Statement Data'

    # Professional title area like bank-statement Excel output.
    ws.merge_cells('A1:G1')
    ws['A1'] = 'PDF to Excel - Extracted Statement Data'
    ws['A1'].font = Font(bold=True, size=16, color='FFFFFF')
    ws['A1'].fill = PatternFill('solid', fgColor='1F4E79')
    ws['A1'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A2:G2')
    ws['A2'] = Path(original).name
    ws['A2'].font = Font(italic=True, color='666666')
    ws['A2'].alignment = Alignment(horizontal='center')

    start_row = 4
    if account_info:
        ws['A4'] = 'PDF Details'
        ws['A4'].font = Font(bold=True)
        for idx, info in enumerate(account_info[:8], start=5):
            ws.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=7)
            ws.cell(idx, 1).value = info
            ws.cell(idx, 1).alignment = Alignment(wrap_text=True)
        start_row = 14

    headers = ['Page', 'Txn Date', 'Value Date', 'Description', 'Debit', 'Credit', 'Balance']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(start_row, col, h)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='4F81BD')
        cell.alignment = Alignment(horizontal='center', vertical='center')

    if rows:
        for r_idx, item in enumerate(rows, start=start_row + 1):
            ws.cell(r_idx, 1, item['Page'])
            ws.cell(r_idx, 2, item['Txn Date'])
            ws.cell(r_idx, 3, item['Value Date'])
            ws.cell(r_idx, 4, item['Description'])
            ws.cell(r_idx, 5, item['Debit'])
            ws.cell(r_idx, 6, item['Credit'])
            ws.cell(r_idx, 7, item['Balance'])
    else:
        # Last fallback: put raw text in a separate readable sheet.
        ws.cell(start_row + 1, 1, 'No clear table detected. Raw text saved in Raw Text sheet.')
        raw = wb.create_sheet('Raw Text')
        raw.append(['Page', 'Text'])
        with pdfplumber.open(str(pdf)) as p:
            for i, page in enumerate(p.pages, 1):
                raw.append([i, page.extract_text() or ''])
        raw.column_dimensions['A'].width = 12
        raw.column_dimensions['B'].width = 100

    # Add totals and style.
    last = max(ws.max_row, start_row + 1)
    total_row = last + 2
    ws.cell(total_row, 4, 'Total')
    ws.cell(total_row, 5, f'=SUM(E{start_row+1}:E{last})')
    ws.cell(total_row, 6, f'=SUM(F{start_row+1}:F{last})')
    for c in range(4, 7):
        ws.cell(total_row, c).font = Font(bold=True)
        ws.cell(total_row, c).fill = PatternFill('solid', fgColor='E2F0D9')

    thin = Side(style='thin', color='D9E2F3')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=start_row, max_row=last, min_col=1, max_col=7):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical='top', wrap_text=(cell.column == 4))
    for col in [5, 6, 7]:
        for cell in ws.iter_cols(min_col=col, max_col=col, min_row=start_row+1, max_row=total_row):
            for c in cell:
                c.number_format = '#,##0.00'

    widths = {'A':10, 'B':14, 'C':14, 'D':48, 'E':14, 'F':14, 'G':14}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = ws.cell(start_row + 1, 1)
    try:
        ws.auto_filter.ref = f'A{start_row}:G{last}'
    except Exception:
        pass

    out = OUTPUTS / (Path(original).stem + '_excel_statement.xlsx')
    wb.save(out)
    return send(out, out.name)

@app.post('/api/ppt-to-pdf')
def ppt_to_pdf():
    if Presentation is None:
        return jsonify(error='python-pptx is not installed. Run pip install python-pptx'), 400
    ppt, original = save_file('file')
    prs = Presentation(str(ppt))
    out = OUTPUTS/(Path(original).stem+'_converted.pdf')
    c = canvas.Canvas(str(out), pagesize=A4)
    width, height = A4
    for idx, slide in enumerate(prs.slides, 1):
        y = height - 50
        c.setFont('Helvetica-Bold', 18); c.drawString(45, y, f'Slide {idx}'); y -= 30
        c.setFont('Helvetica', 11)
        for shape in slide.shapes:
            if hasattr(shape, 'text') and shape.text:
                for line in shape.text.splitlines():
                    for chunk in [line[i:i+90] for i in range(0, len(line), 90)] or ['']:
                        c.drawString(45, y, chunk); y -= 15
                        if y < 45:
                            c.showPage(); c.setFont('Helvetica', 11); y = height - 50
        c.showPage()
    c.save()
    return send(out, out.name)

@app.post('/api/excel-to-pdf')
def excel_to_pdf():
    if load_workbook is None:
        return jsonify(error='openpyxl is not installed. Run pip install openpyxl'), 400
    xlsx, original = save_file('file')
    wb = load_workbook(str(xlsx), data_only=True)
    out = OUTPUTS/(Path(original).stem+'_converted.pdf')
    c = canvas.Canvas(str(out), pagesize=A4)
    width, height = A4
    for ws in wb.worksheets:
        y = height - 45
        c.setFont('Helvetica-Bold', 16); c.drawString(40, y, ws.title); y -= 25
        c.setFont('Helvetica', 9)
        for row in ws.iter_rows(values_only=True):
            line = ' | '.join('' if v is None else str(v) for v in row)
            for chunk in [line[i:i+115] for i in range(0, len(line), 115)] or ['']:
                c.drawString(35, y, chunk); y -= 12
                if y < 40:
                    c.showPage(); c.setFont('Helvetica', 9); y = height - 45
        c.showPage()
    c.save()
    return send(out, out.name)

@app.post('/api/edit-pdf')
def edit_pdf():
    pdf, original = save_file('file')
    out = add_text_overlay(pdf, original, request.form.get('text','Edited with ToolMitra'), request.form.get('page','1'), 70, 70, 18)
    return send(out, out.name)

def add_signature_all_pages(pdf_path, original, text_value, size=22):
    reader = PdfReader(str(pdf_path)); writer = PdfWriter()
    for i, page in enumerate(reader.pages, 1):
        w=float(page.mediabox.width); h=float(page.mediabox.height)
        packet=io.BytesIO(); c=canvas.Canvas(packet, pagesize=(w,h))
        c.setFont('Helvetica-Bold', int(size)); c.setFillColorRGB(0.10,0.10,0.90)
        
        est_width = len(text_value[:120]) * (size * 0.5)
        x = w - est_width - 40
        y = 50
        
        c.drawString(max(20, x), float(y), text_value[:120])
        c.save(); packet.seek(0)
        overlay=PdfReader(packet).pages[0]; page.merge_page(overlay)
        writer.add_page(page)
    out=OUTPUTS/(Path(original).stem+'_edited.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return out

@app.post('/api/sign-pdf')
def sign_pdf():
    pdf, original = save_file('file')
    out = add_signature_all_pages(pdf, original, request.form.get('signature','Signed'), 22)
    out2 = OUTPUTS/(Path(original).stem+'_signed.pdf')
    out.replace(out2)
    return send(out2, out2.name)

@app.post('/api/pdf-to-jpg')
def pdf_to_jpg():
    """Render every selected PDF page as real JPG/PNG images and return a ZIP.
    This is visual conversion, so bank statements/scanned PDFs also convert exactly as images.
    Options from frontend/form:
      - pages: blank/all or e.g. 1,3-5
      - dpi: 96-300, default 180
      - quality: JPG quality 40-100, default 92
      - format: jpg or png, default jpg
    """
    if pdfium is None:
        return jsonify(error='pypdfium2 is not installed. Run: pip install pypdfium2'), 400

    pdf, original = save_file('file')
    stem = Path(original).stem or 'pdf'
    try:
        dpi = int(float(request.form.get('dpi', '180') or 180))
    except Exception:
        dpi = 180
    dpi = max(72, min(300, dpi))

    try:
        quality = int(float(request.form.get('quality', '92') or 92))
    except Exception:
        quality = 92
    quality = max(40, min(100, quality))

    fmt = (request.form.get('format', 'jpg') or 'jpg').lower().strip()
    if fmt not in {'jpg', 'jpeg', 'png'}:
        fmt = 'jpg'
    ext = 'png' if fmt == 'png' else 'jpg'

    try:
        doc = pdfium.PdfDocument(str(pdf))
        selected_pages = parse_pages(request.form.get('pages', ''), len(doc))
        if not selected_pages:
            selected_pages = list(range(len(doc)))

        scale = dpi / 72.0
        temp_files = []

        def render_page_to_image(page_index):
            page = doc[page_index]
            bitmap = page.render(scale=scale).to_pil()
            # Always output JPG for the simple one-click PDF to JPG tool.
            if bitmap.mode in ('RGBA', 'LA'):
                bg = Image.new('RGB', bitmap.size, 'white')
                bg.paste(bitmap, mask=bitmap.split()[-1])
                bitmap = bg
            else:
                bitmap = bitmap.convert('RGB')
            img_path = OUTPUTS / f'{stem}_page_{page_index + 1}.jpg'
            bitmap.save(img_path, 'JPEG', quality=quality, optimize=True, progressive=True)
            return img_path

        if len(selected_pages) == 1:
            img_path = render_page_to_image(selected_pages[0])
            return send(img_path, img_path.name)

        zpath = OUTPUTS / f'{stem}_jpg_images.zip'
        with zipfile.ZipFile(zpath, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for page_index in selected_pages:
                img_path = render_page_to_image(page_index)
                temp_files.append(img_path)
                z.write(img_path, img_path.name)

        for f in temp_files:
            try:
                f.unlink()
            except Exception:
                pass
        return send(zpath, zpath.name)
    except Exception as e:
        return jsonify(error=f'PDF to JPG failed: {str(e)}'), 500

@app.post('/api/organize-pdf')
def organize_pdf():
    pdf, original = save_file('file')
    reader = PdfReader(str(pdf)); writer = PdfWriter()
    order = request.form.get('order','').replace(' ','')
    if order:
        indexes = [int(x)-1 for x in order.split(',') if x.isdigit()]
    else:
        indexes = list(range(len(reader.pages)))
    for idx in indexes:
        if 0 <= idx < len(reader.pages):
            writer.add_page(reader.pages[idx])
    out=OUTPUTS/(Path(original).stem+'_organized.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return send(out,out.name)

@app.post('/api/pdf-to-pdfa')
def pdf_to_pdfa():
    pdf, original = save_file('file')
    reader=PdfReader(str(pdf)); writer=PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    out=OUTPUTS/(Path(original).stem+'_pdfa_ready.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return send(out,out.name)

@app.post('/api/repair-pdf')
def repair_pdf():
    pdf, original = save_file('file')
    reader=PdfReader(str(pdf), strict=False); writer=PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    out=OUTPUTS/(Path(original).stem+'_repaired.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return send(out,out.name)

def _auto_config_tesseract():
    """Find Tesseract on Windows/Linux and configure pytesseract."""
    if pytesseract is None:
        return False, "Python package pytesseract missing. Run: pip install pytesseract"
    exe = shutil.which('tesseract')
    if not exe and os.name == 'nt':
        candidates = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        ]
        exe = next((p for p in candidates if Path(p).exists()), None)
    if not exe:
        return False, "tesseract_missing"
    pytesseract.pytesseract.tesseract_cmd = exe
    try:
        langs = subprocess.check_output([exe, '--list-langs'], text=True, stderr=subprocess.STDOUT, timeout=8)
        if 'eng' not in langs.split():
            return False, "Tesseract me English language data missing hai. English language pack install karo."
    except Exception:
        pass
    return True, "ok"

def _enhance_for_ocr(img):
    """Clean scanned page for better English OCR while keeping page readable."""
    if img.mode != 'RGB':
        img = img.convert('RGB')
    # white background if transparency exists
    bg = Image.new('RGB', img.size, 'white')
    bg.paste(img)
    img = bg
    gray = img.convert('L')
    gray = ImageEnhance.Contrast(gray).enhance(1.65)
    gray = ImageEnhance.Sharpness(gray).enhance(1.45)
    # light denoise/sharpen
    gray = gray.filter(ImageFilter.MedianFilter(size=3))
    gray = gray.filter(ImageFilter.SHARPEN)
    return gray.convert('RGB')

def _render_pdf_pages(pdf_path, dpi=260):
    if pdfium is None:
        raise RuntimeError('pypdfium2 missing hai. Run: pip install pypdfium2')
    doc = pdfium.PdfDocument(str(pdf_path))
    scale = max(1.0, min(float(dpi), 350.0)) / 72.0
    try:
        for i in range(len(doc)):
            page = doc[i]
            bitmap = page.render(scale=scale)
            pil = bitmap.to_pil()
            yield i + 1, pil
            try:
                page.close()
            except Exception:
                pass
    finally:
        try:
            doc.close()
        except Exception:
            pass


def _ocr_page_to_pdf_bytes(page_tuple, dpi=260):
    page_no, pil = page_tuple
    enhanced = _enhance_for_ocr(pil)
    pdf_bytes = pytesseract.image_to_pdf_or_hocr(
        enhanced,
        extension='pdf',
        lang='eng',
        config=f'--oem 3 --psm 6 -c preserve_interword_spaces=1 --dpi {dpi}'
    )
    return page_no, pdf_bytes

def _ocr_single_pdf_advanced(pdf_path, original_name, dpi=260, mode='auto', enhance='yes'):
    """Create an English searchable OCR PDF. Uses OCRmyPDF when available for best original layout."""
    out = OUTPUTS / (Path(original_name).stem + '_english_searchable_ocr.pdf')
    dpi = max(180, min(int(dpi or 260), 350))

    # Best quality path: OCRmyPDF preserves original PDF layout and adds hidden text layer.
    # Works on Render/VPS if ocrmypdf + qpdf + ghostscript are installed.
    ocrmypdf_exe = shutil.which('ocrmypdf')
    if ocrmypdf_exe and mode in ('auto', 'original'):
        cmd = [
            ocrmypdf_exe,
            '--language', 'eng',
            '--deskew',
            '--clean',
            '--optimize', '1',
            '--output-type', 'pdf',
            '--jobs', str(max(1, min((os.cpu_count() or 2), 4))),
            '--tesseract-timeout', '180',
            str(pdf_path),
            str(out)
        ]
        try:
            subprocess.check_call(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
            if out.exists() and out.stat().st_size > 0:
                return out
        except Exception:
            # Fallback below keeps the tool working even when OCRmyPDF is not installed properly.
            pass

    # Fallback path: render each page to high-DPI image and add Tesseract searchable layer.
    pages = list(_render_pdf_pages(pdf_path, dpi=dpi))
    if not pages:
        raise RuntimeError('PDF pages read nahi hui. Dusri PDF try karo.')

    results = {}
    workers = max(1, min((os.cpu_count() or 2), 4, len(pages)))
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futures = [ex.submit(_ocr_page_to_pdf_bytes, p, dpi) for p in pages]
        for fut in as_completed(futures):
            page_no, pdf_bytes = fut.result()
            results[page_no] = pdf_bytes

    writer = PdfWriter()
    for page_no in sorted(results):
        tmp_reader = PdfReader(io.BytesIO(results[page_no]))
        writer.add_page(tmp_reader.pages[0])

    with open(out, 'wb') as fp:
        writer.write(fp)
    return out

@app.post('/api/ocr-pdf')
def ocr_pdf():
    """Ultra Advanced English OCR: single/batch PDF -> searchable PDF or ZIP."""
    ok, msg = _auto_config_tesseract()
    if not ok and msg != "tesseract_missing":
        return jsonify(error=msg), 500

    files = request.files.getlist('files')
    if not files and 'file' in request.files:
        files = [request.files['file']]
    if not files:
        return jsonify(error='PDF file upload karo.'), 400

    try:
        dpi = int(request.form.get('dpi') or 260)
    except Exception:
        dpi = 260
    dpi = max(180, min(dpi, 350))
    mode = (request.form.get('mode') or 'auto').lower()
    enhance = (request.form.get('enhance') or 'yes').lower()

    outputs = []
    try:
        for f in files:
            if not f.filename.lower().endswith('.pdf'):
                return jsonify(error='Sirf PDF file upload karo.'), 400
            pdf_path = UPLOADS / uid(f.filename)
            f.save(pdf_path)
            if not ok and msg == "tesseract_missing":
                import shutil
                out = OUTPUTS / (Path(f.filename).stem + '_english_searchable_ocr.pdf')
                shutil.copy(pdf_path, out)
                outputs.append(out)
            else:
                outputs.append(_ocr_single_pdf_advanced(pdf_path, f.filename, dpi=dpi, mode=mode, enhance=enhance))

        if len(outputs) == 1:
            return send(outputs[0], outputs[0].name)

        zpath = OUTPUTS / f'toolmitra_batch_english_ocr_{uuid.uuid4().hex[:8]}.zip'
        with zipfile.ZipFile(zpath, 'w', zipfile.ZIP_DEFLATED) as z:
            for p in outputs:
                z.write(p, p.name)
        return send(zpath, zpath.name)
    except Exception as e:
        return jsonify(error=f'English OCR PDF failed: {str(e)}'), 500


@app.post('/api/compare-pdf')
def compare_pdf():
    files = request.files.getlist('files')
    if len(files) < 2:
        return jsonify(error='Upload 2 PDF files'), 400
    paths=[]
    for f in files[:2]:
        p=UPLOADS/uid(f.filename); f.save(p); paths.append(p)
    a='\n'.join(pdf_text_pages(paths[0])).splitlines()
    b='\n'.join(pdf_text_pages(paths[1])).splitlines()
    diff='\n'.join(unified_diff(a,b,fromfile=files[0].filename,tofile=files[1].filename,lineterm=''))
    out=OUTPUTS/'pdf_compare_report.txt'
    out.write_text(diff or 'No text difference found.', encoding='utf-8')
    return send(out,out.name)

@app.post('/api/redact-pdf')
def redact_pdf():
    pdf, original = save_file('file')
    needle = (request.form.get('text','') or '').strip().lower()
    reader=PdfReader(str(pdf)); writer=PdfWriter()
    with pdfplumber.open(str(pdf)) as p:
        for i,page in enumerate(reader.pages):
            w=float(page.mediabox.width); h=float(page.mediabox.height)
            packet=io.BytesIO(); c=canvas.Canvas(packet,pagesize=(w,h))
            if needle:
                words = p.pages[i].extract_words() or []
                for word in words:
                    if needle in word.get('text','').lower():
                        x0=float(word['x0']); x1=float(word['x1'])
                        top=float(word['top']); bottom=float(word['bottom'])
                        c.setFillColorRGB(0,0,0)
                        c.rect(x0, h-bottom, max(8,x1-x0), max(8,bottom-top), fill=1, stroke=0)
            # Important: create a real blank overlay page even if no word matched.
            # Without showPage(), ReportLab may produce a PDF with 0 pages, causing
            # IndexError: Sequence index out of range in PdfReader(packet).pages[0].
            c.showPage()
            c.save(); packet.seek(0)
            overlay_reader = PdfReader(packet)
            if len(overlay_reader.pages) > 0:
                overlay = overlay_reader.pages[0]
                page.merge_page(overlay)
            writer.add_page(page)
    out=OUTPUTS/(Path(original).stem+'_redacted.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return send(out,out.name)

@app.post('/api/crop-pdf')
def crop_pdf():
    pdf, original = save_file('file')
    margin=float(request.form.get('margin','36') or 36)
    reader=PdfReader(str(pdf)); writer=PdfWriter()
    for page in reader.pages:
        page.mediabox.lower_left = (float(page.mediabox.left)+margin, float(page.mediabox.bottom)+margin)
        page.mediabox.upper_right = (float(page.mediabox.right)-margin, float(page.mediabox.top)-margin)
        writer.add_page(page)
    out=OUTPUTS/(Path(original).stem+'_cropped.pdf')
    with open(out,'wb') as fp: writer.write(fp)
    return send(out,out.name)

@app.post('/api/pdf-forms')
def pdf_forms():
    pdf, original = save_file('file')
    reader=PdfReader(str(pdf))
    fields = reader.get_fields() or {}
    out=OUTPUTS/(Path(original).stem+'_form_fields.txt')
    if fields:
        data='\n'.join(f'{k}: {v.get("/V","")}' for k,v in fields.items())
    else:
        data='No interactive form fields found in this PDF.'
    out.write_text(data, encoding='utf-8')
    return send(out,out.name)

@app.post('/api/ai-summarizer')
def ai_summarizer():
    pdf, original = save_file('file')
    text = '\n'.join(pdf_text_pages(pdf))
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    summary = '\n'.join(sentences[:8]) if sentences and sentences[0] else 'No selectable text found for summary.'
    out=OUTPUTS/(Path(original).stem+'_summary.txt')
    out.write_text('ToolMitra Summary\n\n'+summary, encoding='utf-8')
    return send(out,out.name)

LANG_MAP = {
    'hindi':'hi','hi':'hi','english':'en','en':'en','spanish':'es','french':'fr','arabic':'ar','gujarati':'gu','marathi':'mr','bengali':'bn','tamil':'ta','telugu':'te','urdu':'ur','punjabi':'pa','nepali':'ne','german':'de','italian':'it','portuguese':'pt','russian':'ru','chinese':'zh','japanese':'ja','korean':'ko'
}

def _target_code(value):
    v=(value or 'Hindi').strip().lower()
    return LANG_MAP.get(v, v[:2] if len(v)>=2 else 'hi')

def _find_font_for_lang(code):
    paths=[]
    if code == 'hi' or code == 'mr' or code == 'ne':
        paths += [
            '/usr/share/fonts/truetype/noto/NotoSansDevanagari-Regular.ttf',
            '/usr/share/fonts/truetype/noto/NotoSansDevanagari-CondensedRegular.ttf',
            '/usr/share/fonts/truetype/noto/NotoSansDevanagari-ExtraCondensedMedium.ttf',
            '/usr/share/fonts/truetype/lohit-devanagari/Lohit-Devanagari.ttf',
            'C:/Windows/Fonts/Nirmala.ttf',
            'C:/Windows/Fonts/mangal.ttf',
        ]
    elif code == 'gu':
        paths += ['/usr/share/fonts/truetype/noto/NotoSansGujarati-Regular.ttf','C:/Windows/Fonts/Nirmala.ttf']
    elif code == 'bn':
        paths += ['/usr/share/fonts/truetype/noto/NotoSansBengali-Regular.ttf','C:/Windows/Fonts/Nirmala.ttf']
    elif code == 'ta':
        paths += ['/usr/share/fonts/truetype/noto/NotoSansTamil-Regular.ttf','C:/Windows/Fonts/Nirmala.ttf']
    elif code == 'te':
        paths += ['/usr/share/fonts/truetype/noto/NotoSansTelugu-Regular.ttf','C:/Windows/Fonts/Nirmala.ttf']
    elif code == 'ar' or code == 'ur':
        paths += ['/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf','/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf','C:/Windows/Fonts/Nirmala.ttf']
    paths += [
        '/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf',
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        '/usr/share/fonts/truetype/freefont/FreeSans.ttf',
        'C:/Windows/Fonts/arial.ttf'
    ]
    for p in paths:
        if Path(p).exists(): return p
    return None

def _translate_batch(texts, target='hi'):
    """Real translator priority: Google Cloud -> Azure -> LibreTranslate -> deep-translator fallback."""
    texts=[(t or '').strip() for t in texts]
    results=['' for _ in texts]
    todo=[(i,t) for i,t in enumerate(texts) if t]
    if not todo: return results

    # Keep technical tokens useful for statements/resumes; translate full sentences but don't destroy IDs.
    google_key=os.environ.get('GOOGLE_TRANSLATE_API_KEY','').strip()
    if google_key and requests:
        try:
            url='https://translation.googleapis.com/language/translate/v2'
            payload={'q':[t for _,t in todo], 'target':target, 'format':'text', 'key':google_key}
            r=requests.post(url, data=payload, timeout=40)
            r.raise_for_status()
            arr=r.json().get('data',{}).get('translations',[])
            for (i,_), item in zip(todo, arr): results[i]=item.get('translatedText','')
            return results
        except Exception:
            pass

    azure_key=os.environ.get('AZURE_TRANSLATOR_KEY','').strip()
    azure_region=os.environ.get('AZURE_TRANSLATOR_REGION','').strip()
    if azure_key and requests:
        try:
            endpoint=os.environ.get('AZURE_TRANSLATOR_ENDPOINT','https://api.cognitive.microsofttranslator.com').rstrip('/')
            url=f'{endpoint}/translate?api-version=3.0&to={target}'
            headers={'Ocp-Apim-Subscription-Key':azure_key,'Content-Type':'application/json'}
            if azure_region: headers['Ocp-Apim-Subscription-Region']=azure_region
            body=[{'text':t} for _,t in todo]
            r=requests.post(url, headers=headers, json=body, timeout=40)
            r.raise_for_status()
            for (i,_), item in zip(todo, r.json()):
                results[i]=item.get('translations',[{}])[0].get('text','')
            return results
        except Exception:
            pass

    libre=os.environ.get('LIBRETRANSLATE_URL','').strip().rstrip('/')
    libre_key=os.environ.get('LIBRETRANSLATE_API_KEY','').strip()
    if libre and requests:
        try:
            for i,t in todo:
                data={'q':t,'source':'auto','target':target,'format':'text'}
                if libre_key: data['api_key']=libre_key
                rr=requests.post(libre+'/translate', json=data, timeout=30)
                rr.raise_for_status(); results[i]=rr.json().get('translatedText','')
            return results
        except Exception:
            pass

    if GoogleTranslator:
        def translate_single(item):
            idx, text = item
            try:
                tr = GoogleTranslator(source='auto', target=target)
                return idx, (tr.translate(text[:4500]) or text)
            except Exception:
                return idx, text
        try:
            with ThreadPoolExecutor(max_workers=5) as ex:
                futures = [ex.submit(translate_single, item) for item in todo]
                for fut in as_completed(futures):
                    idx, text = fut.result()
                    results[idx] = text
            return results
        except Exception:
            pass
    # safe fallback: return original, not corrupted question marks
    for i,t in todo: results[i]=t
    return results

def _wrap_text(draw, text, font, max_w):
    words=(text or '').replace('\n',' ').split()
    if not words: return ['']
    lines=[]; line=''
    for w in words:
        test=(line+' '+w).strip()
        try: width=draw.textbbox((0,0), test, font=font)[2]
        except Exception: width=len(test)*8
        if width <= max_w or not line:
            line=test
        else:
            lines.append(line); line=w
    if line: lines.append(line)
    return lines[:6]

def _render_translated_pdf_image_mode(pdf_path, original, target_lang='Hindi', mode='layout'):
    if fitz is None:
        raise RuntimeError('PyMuPDF missing. Run: pip install pymupdf')
    target=_target_code(target_lang)
    font_path=_find_font_for_lang(target)
    if not font_path:
        raise RuntimeError('Unicode font missing. Install Noto fonts or Windows Nirmala UI font.')

    doc=fitz.open(str(pdf_path))
    out_images=[]
    zoom=float(request.form.get('quality','3') or 3)
    matrix=fitz.Matrix(zoom, zoom)

    for page_index, page in enumerate(doc):
        pix=page.get_pixmap(matrix=matrix, alpha=False)
        img=Image.open(io.BytesIO(pix.tobytes('png'))).convert('RGB')
        draw=ImageDraw.Draw(img)
        blocks=page.get_text('blocks')
        usable=[]
        for b in blocks:
            if len(b)<5: continue
            x0,y0,x1,y1,text=b[:5]
            text=re.sub(r'\s+',' ',str(text)).strip()
            if len(text)<2: continue
            # skip mostly numeric/ids only blocks to preserve statement amounts and codes
            letters=len(re.findall(r'[A-Za-z\u0900-\u097F]', text))
            if letters < 3: continue
            usable.append((x0,y0,x1,y1,text))

        translated=_translate_batch([u[4] for u in usable], target)
        for (x0,y0,x1,y1,text), trans in zip(usable, translated):
            trans=(trans or text).strip()
            X0,Y0,X1,Y1=[int(v*zoom) for v in (x0,y0,x1,y1)]
            pad=max(2,int(2*zoom))
            # white-cover only text area, keep lines/images around it
            draw.rectangle([X0-pad,Y0-pad,X1+pad,Y1+pad], fill='white')
            box_w=max(20, X1-X0)
            box_h=max(12, Y1-Y0)
            base_size=max(8, min(22, int((box_h/zoom)*0.85)))
            # shrink to fit
            for fs in range(base_size, 6, -1):
                try: font=ImageFont.truetype(font_path, int(fs*zoom))
                except Exception: font=ImageFont.load_default()
                lines=_wrap_text(draw, trans, font, box_w)
                line_h=max(10, int((fs+3)*zoom))
                if len(lines)*line_h <= max(box_h, line_h): break
            y=Y0
            fill=(20,20,20)
            for line in lines:
                draw.text((X0, y), line, font=font, fill=fill, direction='rtl' if target in ('ar','ur') else None)
                y += line_h
                if y > Y1 + line_h: break
        # footer
        try:
            foot_font=ImageFont.truetype(font_path, int(7*zoom))
            draw.text((int(10*zoom), img.height-int(16*zoom)), f'Translated to {target_lang} by ToolMitra', font=foot_font, fill=(90,90,90))
        except Exception:
            pass
        tmp=OUTPUTS/f'{uuid.uuid4().hex}_translated_page_{page_index+1}.jpg'
        img.save(tmp, 'JPEG', quality=92, optimize=True)
        out_images.append(tmp)

    out=OUTPUTS/(Path(original).stem+f'_translated_{target}.pdf')
    if not out_images:
        raise RuntimeError('No pages found')
    first=Image.open(out_images[0]).convert('RGB')
    rest=[Image.open(p).convert('RGB') for p in out_images[1:]]
    first.save(out, save_all=True, append_images=rest, resolution=144)
    for p in out_images:
        try: Path(p).unlink()
        except Exception: pass
    return out

def _extract_paragraphs_from_pdf(pdf_path):
    """Extract meaningful paragraphs from each PDF page, merging short fragments into coherent sentences."""
    import pdfplumber
    pages_paras = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            raw = page.extract_text(x_tolerance=2, y_tolerance=4) or ''
            # Split on double newlines first (clear paragraph breaks)
            blocks = re.split(r'\n\s*\n', raw.strip())
            paras = []
            for block in blocks:
                # Merge internal single newlines if they appear to be word wraps
                # (line does not end with sentence-ending punctuation)
                lines = block.splitlines()
                merged = ''
                for line in lines:
                    line = line.strip()
                    if not line:
                        if merged:
                            paras.append(merged.strip())
                            merged = ''
                        continue
                    if merged and not re.search(r'[.!?:;]$', merged):
                        merged = merged + ' ' + line
                    else:
                        if merged:
                            paras.append(merged.strip())
                        merged = line
                if merged:
                    paras.append(merged.strip())
            # Filter out very short fragments (numbers, page refs, etc)
            paras = [p for p in paras if len(p.split()) >= 3]
            pages_paras.append(paras)
    return pages_paras

def _render_translated_text_pdf(pdf_path, original, target_lang='Hindi'):
    """Extract paragraphs, translate each paragraph as a whole unit (parallel), render clean PDF.
    This produces Google-Translate quality results because each paragraph has full context."""
    target = _target_code(target_lang)
    font_path = _find_font_for_lang(target)

    # Register font once
    font_name = 'Helvetica'
    if font_path:
        try:
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
            pdfmetrics.registerFont(TTFont('TMUnicode', font_path))
            font_name = 'TMUnicode'
        except Exception:
            font_name = 'Helvetica'

    # Extract paragraphs per page
    pages_paras = _extract_paragraphs_from_pdf(pdf_path)
    if not pages_paras or all(not p for p in pages_paras):
        # Fallback: use simple page text
        raw_texts = pdf_text_pages(pdf_path)
        pages_paras = [[t] for t in raw_texts if t.strip()]

    # Flatten all paragraphs for batch translation (preserves context)
    # Map: (page_index, para_index) -> flat_index
    flat_paras = []
    index_map = []  # (page_i, para_i)
    for pi, paras in enumerate(pages_paras):
        for qi, para in enumerate(paras):
            flat_paras.append(para)
            index_map.append((pi, qi))

    # Translate in parallel batches of ~10 paragraphs for speed
    def translate_chunk(args):
        chunk_indices, chunk_texts = args
        results = _translate_batch(chunk_texts, target)
        return list(zip(chunk_indices, results))

    CHUNK = 10
    chunks = []
    for i in range(0, len(flat_paras), CHUNK):
        chunk_texts = flat_paras[i:i+CHUNK]
        chunk_indices = list(range(i, i+len(chunk_texts)))
        chunks.append((chunk_indices, chunk_texts))

    translated_flat = [''] * len(flat_paras)
    workers = min(5, len(chunks)) if chunks else 1
    if chunks:
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futures = [ex.submit(translate_chunk, c) for c in chunks]
            for fut in as_completed(futures):
                for idx, text in fut.result():
                    translated_flat[idx] = text

    # Re-map translated paragraphs back to pages
    translated_pages = [[] for _ in pages_paras]
    for flat_i, (pi, qi) in enumerate(index_map):
        translated_pages[pi].append(translated_flat[flat_i])

    # Render clean PDF
    out = OUTPUTS / (Path(original).stem + f'_translated_{target}.pdf')
    W, H = A4
    c = canvas.Canvas(str(out), pagesize=A4)
    MARGIN_L = 50
    MARGIN_R = W - 50
    LINE_W = MARGIN_R - MARGIN_L
    FONT_SIZE = 12
    LINE_H = FONT_SIZE + 5
    HEAD_SIZE = 13

    def draw_header(page_num):
        """Draw page header."""
        c.setFont('Helvetica-Bold', HEAD_SIZE)
        c.setFillColorRGB(0.39, 0.40, 0.95)  # indigo
        c.drawString(MARGIN_L, H - 38, f'Translated to {target_lang}  |  Page {page_num}  |  OrbixaPDFTool')
        c.setFillColorRGB(0, 0, 0)
        c.line(MARGIN_L, H - 46, MARGIN_R, H - 46)
        return H - 60

    def wrap_text(text, max_chars=78):
        """Wrap text at word boundaries for clean lines."""
        words = (text or '').split()
        lines = []; current = ''
        for w in words:
            test = (current + ' ' + w).strip()
            if len(test) <= max_chars:
                current = test
            else:
                if current:
                    lines.append(current)
                current = w
        if current:
            lines.append(current)
        return lines or ['']

    page_num = 1
    y = draw_header(page_num)

    c.setFont(font_name, FONT_SIZE)

    for pi, paras in enumerate(translated_pages):
        if not paras:
            continue
        # Page divider (not on first page)
        if pi > 0:
            if y < H * 0.4:
                c.showPage()
                page_num += 1
                y = draw_header(page_num)
                c.setFont(font_name, FONT_SIZE)
            else:
                c.setFillColorRGB(0.5, 0.5, 0.5)
                c.setFont('Helvetica', 9)
                c.drawString(MARGIN_L, y, f'— Original Page {pi + 1} —')
                c.setFillColorRGB(0, 0, 0)
                c.setFont(font_name, FONT_SIZE)
                y -= LINE_H + 4

        for para in paras:
            lines = wrap_text(para)
            for line in lines:
                if y < 55:
                    c.showPage()
                    page_num += 1
                    y = draw_header(page_num)
                    c.setFont(font_name, FONT_SIZE)
                c.drawString(MARGIN_L, y, line)
                y -= LINE_H
            y -= 6  # paragraph spacing

    c.save()
    return out

@app.post('/api/translate-pdf')
def translate_pdf():
    pdf, original = save_file('file')
    lang = request.form.get('language','Hindi')
    mode = (request.form.get('mode') or 'same-layout').lower()
    try:
        if mode in ('image','same-layout','layout'):
            out=_render_translated_pdf_image_mode(pdf, original, lang, mode)
        else:
            # Default: clean paragraph-level translation (best quality + fast)
            out=_render_translated_text_pdf(pdf, original, lang)
        return send(out, out.name)
    except Exception as e:
        return jsonify(error=f'Translate PDF failed: {e}. Install requirements and set GOOGLE_TRANSLATE_API_KEY or AZURE_TRANSLATOR_KEY for best quality.'), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)


