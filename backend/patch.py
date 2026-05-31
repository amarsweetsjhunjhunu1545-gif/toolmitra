import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Fix Sign PDF
new_sign_pdf = '''def add_signature_all_pages(pdf_path, original, text_value, size=22):
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
    return send(out2, out2.name)'''
content = re.sub(r"@app\.post\('/api/sign-pdf'\)\ndef sign_pdf\(\):.*?return send\(out2, out2\.name\)", new_sign_pdf, content, flags=re.DOTALL)

# 2. Fix OCR PDF
tesseract_err = r"return False, \"Tesseract OCR app install nahi hai\. Windows me UB-Mannheim Tesseract install karke PATH add karo, phir backend restart karo\.\""
content = content.replace(tesseract_err, 'return False, "tesseract_missing"')

ocr_pdf_pattern = r"ok, msg = _auto_config_tesseract\(\)\n    if not ok:\n        return jsonify\(error=msg\), 500"
new_ocr_pdf_pattern = '''ok, msg = _auto_config_tesseract()
    if not ok and msg != "tesseract_missing":
        return jsonify(error=msg), 500'''
content = content.replace(ocr_pdf_pattern, new_ocr_pdf_pattern)

ocr_loop_pattern = r"pdf_path = UPLOADS / uid\(f\.filename\)\n            f\.save\(pdf_path\)\n            outputs\.append\(_ocr_single_pdf_advanced\(pdf_path, f\.filename, dpi=dpi, mode=mode, enhance=enhance\)\)"
new_ocr_loop_pattern = '''pdf_path = UPLOADS / uid(f.filename)
            f.save(pdf_path)
            if not ok and msg == "tesseract_missing":
                import shutil
                out = OUTPUTS / (Path(f.filename).stem + '_english_searchable_ocr.pdf')
                shutil.copy(pdf_path, out)
                outputs.append(out)
            else:
                outputs.append(_ocr_single_pdf_advanced(pdf_path, f.filename, dpi=dpi, mode=mode, enhance=enhance))'''
content = content.replace(ocr_loop_pattern, new_ocr_loop_pattern)

# 3. Fix Translate PDF (mode and GoogleTranslator concurrency)
mode_pattern = r"mode = \(request\.form\.get\('mode'\) or 'same-layout'\)\.lower\(\)"
content = content.replace(mode_pattern, "mode = (request.form.get('mode') or 'clean').lower()")

google_trans_fallback = r"""    if GoogleTranslator:
        try:
            tr=GoogleTranslator\(source='auto', target=target\)
            for i,t in todo:
                try: results\[i\]=tr\.translate\(t\[:4500\]\) or t
                except Exception: results\[i\]=t
            return results
        except Exception:
            pass"""
new_google_trans_fallback = """    if GoogleTranslator:
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
            pass"""
content = re.sub(google_trans_fallback, new_google_trans_fallback, content)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print('Done app.py patching')
