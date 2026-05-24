# ToolMitra Professional Updated

Updated for PC, mobile and tablet compatibility.

## Fixes included
- PDF to Word backend now has safe fallback: if exact layout conversion fails, editable DOCX still downloads.
- Mobile-friendly file download flow for Android/iPhone browsers.
- Long timeout handling for Render backend tools.
- Better error messages instead of silent failure.
- Responsive CSS improvements for modal, processing popup, tablet and phone screens.
- Clean ZIP without old uploaded/output files.

## Deploy
### Frontend Netlify
Upload only the `frontend` folder files to Netlify or keep publish directory as `frontend`.

### Backend Render
Root directory: `backend`  
Build command: `pip install -r requirements.txt`  
Start command: `gunicorn app:app`  
Apt packages: use `apt.txt` for Tesseract/OCR tools.

## Local run
```bash
cd backend
pip install -r requirements.txt
python app.py
```
Then open `frontend/index.html`.
