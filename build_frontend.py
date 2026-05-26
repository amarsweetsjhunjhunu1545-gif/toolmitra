import json
import re
import os

tools = [
 {'t':'Merge PDF','e':'/api/merge-pdf','d':'Combine PDFs in the order you want.','i':'⇄','type':'multi','accept':'.pdf'},
 {'t':'Split PDF','e':'/api/split-pdf','d':'Separate pages into independent PDF files.','i':'✂','type':'single','accept':'.pdf','extra':[['pages','Pages e.g. 1-3,5', '']]},
 {'t':'Compress PDF','e':'/api/compress-pdf','d':'Reduce PDF size. Enter target KB like 200, 500, 1024 or choose level.','i':'↘','type':'single','accept':'.pdf','extra':[['target_kb','Target size in KB (optional) e.g. 200',''],['level','Compression level: low / medium / high / extreme','high']]},
 {'t':'PDF to Word','e':'/api/pdf-to-word','d':'Convert PDF into editable DOCX file.','i':'W','type':'single','accept':'.pdf'},
 {'t':'PDF to PowerPoint','e':'/api/pdf-to-ppt','d':'Convert each PDF page into a PPT slide with same visual layout.','i':'P','type':'single','accept':'.pdf','extra':[['quality','Quality 1.5-3.0','2.0']]},
 {'t':'PDF to Excel','e':'/api/pdf-to-excel','d':'Extract tables/text from PDF into XLSX.','i':'X','type':'single','accept':'.pdf'},
 {'t':'Word to PDF','e':'/api/word-to-pdf','d':'Convert DOCX text into PDF.','i':'W','type':'single','accept':'.docx'},
 {'t':'PowerPoint to PDF','e':'/api/ppt-to-pdf','d':'Convert PPTX text content into PDF.','i':'P','type':'single','accept':'.pptx'},
 {'t':'Excel to PDF','e':'/api/excel-to-pdf','d':'Convert spreadsheet rows into PDF.','i':'X','type':'single','accept':'.xlsx,.xls'},
 {'t':'Edit PDF','e':'editor','d':'Advanced PDF editor: add text, image, drawing, white cover box, resize/drag objects and save edited PDF.','i':'✎','type':'editor','accept':'.pdf'},
 {'t':'PDF to JPG','e':'/api/pdf-to-jpg','d':'Upload PDF and click Convert JPG Image. Single page PDF downloads JPG directly; multi-page PDF downloads ZIP with JPG images.','i':'JPG','type':'single','accept':'.pdf','btn':'Convert JPG Image'},
 {'t':'JPG to PDF','e':'/api/jpg-to-pdf','d':'Convert JPG/PNG images into one PDF.','i':'IMG','type':'multi','accept':'image/*'}, 
 {'t':'Watermark','e':'/api/add-watermark','d':'Stamp text watermark over PDF pages.','i':'◆','type':'single','accept':'.pdf','extra':[['text','Watermark text','OrbixaPDFTool']]},
 {'t':'Rotate PDF','e':'/api/rotate-pdf','d':'Rotate all or selected pages.','i':'⟳','type':'single','accept':'.pdf','extra':[['angle','Angle 90/180/270','90'],['pages','Pages blank = all', '']]},
 {'t':'HTML to PDF','e':'/api/html-to-pdf','d':'Paste HTML/text and download PDF.','i':'HTML','type':'html'},
 {'t':'Unlock PDF','e':'/api/unlock-pdf','d':'Remove password when you know it.','i':'🔓','type':'single','accept':'.pdf','extra':[['password','Current password', '']]},
 {'t':'Protect PDF','e':'/api/protect-pdf','d':'Add password protection to PDF.','i':'🔒','type':'single','accept':'.pdf','extra':[['password','Password','1234']]},
 {'t':'Organize PDF','e':'/api/organize-pdf','d':'Reorder PDF pages using custom order.','i':'A↕B','type':'single','accept':'.pdf','extra':[['order','New order e.g. 3,1,2 or blank for same', '']]},
 {'t':'PDF to PDF/A','e':'/api/pdf-to-pdfa','d':'Create a clean archival-style PDF copy.','i':'/A','type':'single','accept':'.pdf'},
 {'t':'Repair PDF','e':'/api/repair-pdf','d':'Rewrite readable pages into a repaired PDF.','i':'🛠','type':'single','accept':'.pdf'},
 {'t':'Page numbers','e':'/api/page-numbers','d':'Add page numbers into your PDF.','i':'123','type':'single','accept':'.pdf'},
 {'t':'Scan to PDF','e':'/api/images-to-pdf-cleaner','d':'Clean scanned images and make PDF.','i':'▣','type':'multi','accept':'image/*'},
 {'t':'Compare PDF','e':'/api/compare-pdf','d':'Compare two PDFs and download text report.','i':'≡','type':'multi','accept':'.pdf'},
 {'t':'Redact PDF','e':'/api/redact-pdf','d':'Cover selected text with black boxes.','i':'▰','type':'single','accept':'.pdf','extra':[['text','Text to redact','secret']]},
 {'t':'Crop PDF','e':'/api/crop-pdf','d':'Crop margins from PDF pages.','i':'⌗','type':'single','accept':'.pdf','extra':[['margin','Crop margin points','36']]},
 {'t':'PDF Forms','e':'/api/pdf-forms','d':'Extract form fields / create form info report.','i':'Ab','type':'single','accept':'.pdf'},
 {'t':'Smart Summarizer','e':'/api/ai-summarizer','d':'Generate a quick text summary from PDF.','i':'TXT','type':'single','accept':'.pdf'},
 {'t':'Delete Pages','e':'/api/delete-pages','d':'Remove unwanted pages.','i':'🗑','type':'single','accept':'.pdf','extra':[['pages','Pages to delete e.g. 2,4-5', '']]},
 {'t':'Extract Pages','e':'/api/extract-pages','d':'Save selected pages as a new PDF.','i':'📌','type':'single','accept':'.pdf','extra':[['pages','Pages e.g. 1-3,5','1']]},
 {'t':'Extract PDF Text','e':'/api/extract-text','d':'Download text from PDF.','i':'TXT','type':'single','accept':'.pdf'},
 {'t':'Image Compressor','e':'/api/image-compress','d':'Image upload karo, target size KB enter karo aur compressed image download karo.','i':'⚡','type':'image','accept':'image/*','btn':'Compress to KB & Download','extra':[['target_kb','Target size in KB e.g. 100','']]}
]

def get_slug(name):
    return name.lower().replace(' ', '-').replace('/', '-')

def tool_category(x):
    s = (x['t'] + ' ' + x['d']).lower()
    if 'image' in s or 'jpg' in s or 'png' in s or 'webp' in s: return 'image'
    if 'protect' in s or 'unlock' in s or 'redact' in s: return 'secure'
    if 'to ' in s or 'convert' in s: return 'convert'
    return 'pdf'

frontend_dir = os.path.join(os.path.dirname(__file__), 'frontend')
index_path = os.path.join(frontend_dir, 'index.html')

with open(index_path, 'r', encoding='utf-8') as f:
    html = f.read()

# 1. Update Index HTML: Replace empty <div class="grid" id="toolGrid"></div> with static grid
grid_html = '<div class="grid" id="toolGrid">\n'
for tool in tools:
    slug = get_slug(tool['t']) + '.html'
    cat = tool_category(tool).upper()
    if tool['t'] == 'Edit PDF':
        slug = 'index.html#editor' # Handled specifically later
    
    grid_html += f"""    <a href="{slug if tool['t'] != 'Edit PDF' else 'javascript:openPdfEditor()'}" class="tool" style="text-decoration: none; color: inherit; display: block;">
      <span class="tag">{cat}</span>
      <div class="icon">{tool['i']}</div>
      <h3>{tool['t']}</h3>
      <p>{tool['d']}</p>
      <small class="toolAction">Open tool →</small>
    </a>\n"""
grid_html += '</div>'

# Inject static grid into index.html
new_index_html = re.sub(r'<div class="grid" id="toolGrid">.*?</div>', grid_html, html, flags=re.DOTALL)

with open(os.path.join(frontend_dir, 'index.html'), 'w', encoding='utf-8') as f:
    f.write(new_index_html)

# 2. Generate Tool Pages
# We will create a template from index.html by replacing the main content with the tool form
# Extract everything before <main> and after </main>
head_match = re.search(r'(.*?)(<main>.*?</main>)(.*)', new_index_html, re.DOTALL)
if head_match:
    pre_main_full = head_match.group(1)
    post_main = head_match.group(3)

    # Remove hero, stats-bar, and categoryChips for tool pages
    pre_main_tool = re.sub(r'<!-- ✅ HERO SECTION.*?<!-- Category Chips -->.*?</div>', '', pre_main_full, flags=re.DOTALL)
    
    # Ensure any remaining stray sections are removed just in case
    pre_main_tool = re.sub(r'<section class="hero exact-hero">.*?</section>', '', pre_main_tool, flags=re.DOTALL)
    pre_main_tool = re.sub(r'<div class="stats-bar">.*?</div>', '', pre_main_tool, flags=re.DOTALL)
    pre_main_tool = re.sub(r'<div class="categoryChips".*?</div>', '', pre_main_tool, flags=re.DOTALL)

    for tool in tools:
        if tool['t'] == 'Edit PDF':
            continue # Edit PDF is special client-side editor
        
        slug = get_slug(tool['t'])
        
        # Build tool specific inputs
        inputs_html = ''
        if tool.get('type') == 'html':
            inputs_html += '<div class="field"><label>HTML/Text</label><textarea name="html" rows="7" placeholder="Paste HTML or text here"></textarea></div>'
        elif tool.get('type') == 'image':
            inputs_html += f"""
            <div class="imageToolPanel">
                <label class="dropUpload imageDrop" for="imageInput">
                <input id="imageInput" name="file" type="file" accept="image/jpeg,image/png,image/webp,image/*" required>
                <span class="uploadEmoji">🖼️</span>
                <b id="imageDropTitle">Choose image or drag & drop</b>
                <small>JPG / PNG / WEBP supported</small>
                <em id="selectedImageName">No image selected</em>
                </label>
                <div id="imagePreviewBox" class="imagePreviewBox hidden">
                <img id="imagePreview" alt="Image preview">
                <div class="imageInfo"><b id="imageMeta">Preview ready</b><small id="imageSubMeta">Original image loaded successfully.</small></div>
                </div>
            </div>
            <div class="settingGrid imageSettings">
                <div class="field"><label>Target image size in KB</label><input name="target_kb" type="number" min="5" placeholder="e.g. 100" required><small>Example: 50, 100, 200, 500 KB</small></div>
            </div>
            """
        else:
            accept_attr = f'accept="{tool.get("accept", "")}"' if tool.get('accept') else ''
            multiple = 'multiple' if tool.get('type') == 'multi' else ''
            label_text = 'Upload files' if tool.get('type') == 'multi' else 'Upload PDF file'
            file_name = 'files' if tool.get('type') == 'multi' else 'file'
            
            inputs_html += f'<div class="field"><label>{label_text}</label><input name="{file_name}" type="file" {multiple} {accept_attr} required></div>'
            
            for extra in tool.get('extra', []):
                val = extra[2] if len(extra) > 2 else ''
                inputs_html += f'<div class="field"><label>{extra[1]}</label><input name="{extra[0]}" value="{val}" placeholder="{extra[1]}"></div>'
                
        btn_text = tool.get('btn', 'Process & Download')
        
        main_content = f"""
<main>
  <section class="tool-page-hero" style="padding: 100px 20px 40px; text-align: center; max-width: 800px; margin: 0 auto;">
    <div class="icon" style="font-size: 48px; margin-bottom: 20px;">{tool['i']}</div>
    <h1>{tool['t']}</h1>
    <p style="font-size: 1.1rem; color: var(--text-muted);">{tool['d']}</p>
    
    <div class="tool-form-container" style="background: var(--bg-card); padding: 30px; border-radius: 16px; box-shadow: 0 10px 30px var(--shadow-color); margin-top: 40px; text-align: left; border: 1px solid var(--border-color);">
        <form id="toolForm">
            {inputs_html}
            <button class="primary" type="submit" style="width: 100%; margin-top: 20px;">{btn_text}</button>
        </form>
        <p id="status" style="margin-top: 15px; text-align: center; font-weight: 500;"></p>
    </div>
  </section>
  <script>
      window.currentTool = {json.dumps(tool)};
  </script>
</main>
"""
        
        # Update Title and Meta for SEO
        page_pre_main = pre_main_tool.replace(
            '<title>Free PDF to Word Converter Online | Compress, Merge, Split PDF – OrbixaPDFTool</title>',
            f'<title>{tool["t"]} Free Online | OrbixaPDFTool</title>'
        )
        page_pre_main = re.sub(
            r'<meta name="description" content=".*?">',
            f'<meta name="description" content="{tool["d"]} Free online tool without watermark and registration.">',
            page_pre_main
        )
        
        with open(os.path.join(frontend_dir, f'{slug}.html'), 'w', encoding='utf-8') as f:
            f.write(page_pre_main + main_content + post_main)

# 3. Generate sitemap.xml
sitemap = '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
sitemap += '  <url>\n    <loc>https://orbixapdftool.in/</loc>\n    <changefreq>daily</changefreq>\n    <priority>1.0</priority>\n  </url>\n'
for tool in tools:
    if tool['t'] != 'Edit PDF':
        sitemap += f'  <url>\n    <loc>https://orbixapdftool.in/{get_slug(tool["t"])}.html</loc>\n    <changefreq>weekly</changefreq>\n    <priority>0.8</priority>\n  </url>\n'
sitemap += '</urlset>'

with open(os.path.join(frontend_dir, 'sitemap.xml'), 'w', encoding='utf-8') as f:
    f.write(sitemap)

print("Static pages and sitemap generated successfully.")
