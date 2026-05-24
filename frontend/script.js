// Backend API URL
// Netlify par deployed frontend ke liye Render backend URL use hoga.
// Local testing me http://127.0.0.1:5000 auto use hoga.
const API = (location.hostname === 'localhost' || location.hostname === '127.0.0.1')
  ? 'http://127.0.0.1:5000'
   : 'https://toolmitra-api.onrender.com';

function safeFileName(name){
  return (name || 'toolmitra_download').replace(/[\\/:*?"<>|]+/g,'_').replace(/\s+/g,'_');
}

async function fetchWithTimeout(url, options={}, timeoutMs=180000){
  const controller = new AbortController();
  const timer = setTimeout(()=>controller.abort(), timeoutMs);
  try{
    return await fetch(url, {...options, signal: controller.signal, mode:'cors'});
  }finally{
    clearTimeout(timer);
  }
}

async function downloadBlobMobileFriendly(blob, filename){
  filename = safeFileName(filename);
  const nav = window.navigator;
  if(nav && typeof nav.msSaveOrOpenBlob === 'function'){
    nav.msSaveOrOpenBlob(blob, filename);
    return;
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.rel = 'noopener';
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  // Mobile Safari/older Android fallback: open file if download attribute is ignored.
  setTimeout(()=>{
    if(/iPhone|iPad|iPod|Android/i.test(navigator.userAgent)){
      try { window.open(url, '_blank'); } catch(e) {}
    }
  }, 350);
  setTimeout(()=>URL.revokeObjectURL(url), 60000);
  a.remove();
}
const tools = [
 {t:'Merge PDF',e:'/api/merge-pdf',d:'Combine PDFs in the order you want.',i:'⇄',type:'multi',accept:'.pdf'},
 {t:'Split PDF',e:'/api/split-pdf',d:'Separate pages into independent PDF files.',i:'✂',type:'single',accept:'.pdf',extra:[['pages','Pages e.g. 1-3,5']]},
 {t:'Compress PDF',e:'/api/compress-pdf',d:'Reduce PDF size. Enter target KB like 200, 500, 1024 or choose level.',i:'↘',type:'single',accept:'.pdf',extra:[['target_kb','Target size in KB (optional) e.g. 200',''],['level','Compression level: low / medium / high / extreme','high']]},
 {t:'PDF to Word',e:'/api/pdf-to-word',d:'Convert PDF into editable DOCX file.',i:'W',type:'single',accept:'.pdf'},
 {t:'PDF to PowerPoint',e:'/api/pdf-to-ppt',d:'Convert each PDF page into a PPT slide with same visual layout.',i:'P',type:'single',accept:'.pdf',extra:[['quality','Quality 1.5-3.0','2.0']]},
 {t:'PDF to Excel',e:'/api/pdf-to-excel',d:'Extract tables/text from PDF into XLSX.',i:'X',type:'single',accept:'.pdf'},
 {t:'Word to PDF',e:'/api/word-to-pdf',d:'Convert DOCX text into PDF.',i:'W',type:'single',accept:'.docx'},
 {t:'PowerPoint to PDF',e:'/api/ppt-to-pdf',d:'Convert PPTX text content into PDF.',i:'P',type:'single',accept:'.pptx'},
 {t:'Excel to PDF',e:'/api/excel-to-pdf',d:'Convert spreadsheet rows into PDF.',i:'X',type:'single',accept:'.xlsx,.xls'},
 {t:'Edit PDF',e:'editor',d:'Advanced PDF editor: add text, image, drawing, white cover box, resize/drag objects and save edited PDF.',i:'✎',type:'editor',accept:'.pdf'},
 {t:'PDF to JPG',e:'/api/pdf-to-jpg',d:'Upload PDF and click Convert JPG Image. Single page PDF downloads JPG directly; multi-page PDF downloads ZIP with JPG images.',i:'JPG',type:'single',accept:'.pdf',btn:'Convert JPG Image'},
 {t:'JPG to PDF',e:'/api/jpg-to-pdf',d:'Convert JPG/PNG images into one PDF.',i:'IMG',type:'multi',accept:'image/*'}, {t:'Watermark',e:'/api/add-watermark',d:'Stamp text watermark over PDF pages.',i:'◆',type:'single',accept:'.pdf',extra:[['text','Watermark text','ToolMitra']]},
 {t:'Rotate PDF',e:'/api/rotate-pdf',d:'Rotate all or selected pages.',i:'⟳',type:'single',accept:'.pdf',extra:[['angle','Angle 90/180/270','90'],['pages','Pages blank = all']]},
 {t:'HTML to PDF',e:'/api/html-to-pdf',d:'Paste HTML/text and download PDF.',i:'HTML',type:'html'},
 {t:'Unlock PDF',e:'/api/unlock-pdf',d:'Remove password when you know it.',i:'🔓',type:'single',accept:'.pdf',extra:[['password','Current password']]},
 {t:'Protect PDF',e:'/api/protect-pdf',d:'Add password protection to PDF.',i:'🔒',type:'single',accept:'.pdf',extra:[['password','Password','1234']]},
 {t:'Organize PDF',e:'/api/organize-pdf',d:'Reorder PDF pages using custom order.',i:'A↕B',type:'single',accept:'.pdf',extra:[['order','New order e.g. 3,1,2 or blank for same']]},
 {t:'PDF to PDF/A',e:'/api/pdf-to-pdfa',d:'Create a clean archival-style PDF copy.',i:'/A',type:'single',accept:'.pdf'},
 {t:'Repair PDF',e:'/api/repair-pdf',d:'Rewrite readable pages into a repaired PDF.',i:'🛠',type:'single',accept:'.pdf'},
 {t:'Page numbers',e:'/api/page-numbers',d:'Add page numbers into your PDF.',i:'123',type:'single',accept:'.pdf'},
 {t:'Scan to PDF',e:'/api/images-to-pdf-cleaner',d:'Clean scanned images and make PDF.',i:'▣',type:'multi',accept:'image/*'},
 {t:'Compare PDF',e:'/api/compare-pdf',d:'Compare two PDFs and download text report.',i:'≡',type:'multi',accept:'.pdf'},
 {t:'Redact PDF',e:'/api/redact-pdf',d:'Cover selected text with black boxes.',i:'▰',type:'single',accept:'.pdf',extra:[['text','Text to redact','secret']]},
 {t:'Crop PDF',e:'/api/crop-pdf',d:'Crop margins from PDF pages.',i:'⌗',type:'single',accept:'.pdf',extra:[['margin','Crop margin points','36']]},
 {t:'PDF Forms',e:'/api/pdf-forms',d:'Extract form fields / create form info report.',i:'Ab',type:'single',accept:'.pdf'},
 {t:'AI Summarizer',e:'/api/ai-summarizer',d:'Generate a quick text summary from PDF.',i:'AI',type:'single',accept:'.pdf'},
 {t:'Delete Pages',e:'/api/delete-pages',d:'Remove unwanted pages.',i:'🗑',type:'single',accept:'.pdf',extra:[['pages','Pages to delete e.g. 2,4-5']]},
 {t:'Extract Pages',e:'/api/extract-pages',d:'Save selected pages as a new PDF.',i:'📌',type:'single',accept:'.pdf',extra:[['pages','Pages e.g. 1-3,5','1']]},
 {t:'Extract PDF Text',e:'/api/extract-text',d:'Download text from PDF.',i:'TXT',type:'single',accept:'.pdf'},
 {t:'Image Compressor',e:'/api/image-compress',d:'Image upload karo, target size KB enter karo aur compressed image download karo.',i:'⚡',type:'image',accept:'image/*',btn:'Compress to KB & Download',extra:[['target_kb','Target size in KB e.g. 100','']]}
];
let current=null;
let activeCategory='all';
function toolCategory(x){const s=(x.t+' '+x.d).toLowerCase(); if(s.includes('image')||s.includes('jpg')||s.includes('png')||s.includes('webp')) return 'image'; if(s.includes('protect')||s.includes('unlock')||s.includes('redact')) return 'secure'; if(s.includes('to ')||s.includes('convert')) return 'convert'; return 'pdf';}
function setCategory(cat,btn){activeCategory=cat; document.querySelectorAll('.categoryChips button').forEach(b=>b.classList.remove('active')); if(btn)btn.classList.add('active'); filterTools();}
function render(list=tools){document.getElementById('toolGrid').innerHTML=list.map((x)=>{const cat=toolCategory(x).toUpperCase();return `<div class="tool" onclick="openTool(${tools.indexOf(x)})"><span class="tag">${cat}</span><div class="icon">${x.i}</div><h3>${x.t}</h3><p>${x.d}</p><small class="toolAction">Open tool →</small></div>`}).join('')}
function filterTools(){const q=document.getElementById('search').value.toLowerCase();render(tools.filter(x=>(activeCategory==='all'||toolCategory(x)===activeCategory) && (x.t+x.d).toLowerCase().includes(q)))}
function quickOpen(name){const i=tools.findIndex(x=>x.t===name);if(i>-1)openTool(i)}
function openTool(i){
  current=tools[i];
  if(current.type==='editor'){openPdfEditor();return;}
  document.getElementById('modalTitle').textContent=current.t;
  document.getElementById('modalDesc').textContent=current.d;
  document.getElementById('status').textContent='';
  let html='';
  if(current.type==='html'){
    html+=`<div class="field"><label>HTML/Text</label><textarea name="html" rows="7" placeholder="Paste HTML or text here"></textarea></div>`;
  }else if(current.type==='image'){
    html+=`<div class="imageToolHero"><div class="imageToolIcon">⚡</div><div><h3>Image Size Compressor</h3><p>Image select karo, target KB likho, aur compressed image direct download karo.</p></div></div>
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
      </div>`;
  }else{
    html+=`<div class="field"><label>${current.type==='multi'?'Upload files':'Upload PDF file'}</label><input name="${current.type==='multi'?'files':'file'}" type="file" ${current.type==='multi'?'multiple':''} accept="${current.accept||''}" required></div>`;
    (current.extra||[]).forEach(e=>{html+=`<div class="field"><label>${e[1]}</label><input name="${e[0]}" value="${e[2]||''}" placeholder="${e[1]}"></div>`});
  }
  document.getElementById('inputs').innerHTML=html;
  if(current.type==='image') setupImageToolPreview();
  const btn=document.querySelector('#toolForm .primary'); if(btn) btn.textContent=current.btn||'Process & Download';
  document.getElementById('modal').classList.remove('hidden');
}
function setupImageToolPreview(){
  const input=document.getElementById('imageInput');
  const box=document.getElementById('imagePreviewBox');
  const img=document.getElementById('imagePreview');
  const name=document.getElementById('selectedImageName');
  const meta=document.getElementById('imageMeta');
  const sub=document.getElementById('imageSubMeta');
  const drop=document.querySelector('.imageDrop');
  const title=document.getElementById('imageDropTitle');
  if(!input || !box || !img || !drop) return;

  const humanSize=(bytes)=>{
    if(bytes<1024) return bytes+' B';
    if(bytes<1024*1024) return (bytes/1024).toFixed(1)+' KB';
    return (bytes/1024/1024).toFixed(2)+' MB';
  };

  const loadFile=(file)=>{
    if(!file) return;
    if(!file.type.startsWith('image/')){
      alert('Please JPG, PNG ya WEBP image select karo.');
      return;
    }
    name.textContent=file.name;
    title.textContent='Image selected successfully';
    drop.classList.add('selected');
    const url=URL.createObjectURL(file);
    img.onload=()=>{
      meta.textContent=file.name;
      sub.textContent=`Original: ${img.naturalWidth} × ${img.naturalHeight}px • ${humanSize(file.size)}`;
      URL.revokeObjectURL(url);
    };
    img.src=url;
    box.classList.remove('hidden');
    const btn=document.querySelector('#toolForm .primary');
    if(btn) btn.textContent='Compress / Convert & Download';
  };

  input.addEventListener('change',()=>loadFile(input.files && input.files[0]));

  ['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{
    e.preventDefault();
    drop.classList.add('drag');
  }));
  ['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{
    e.preventDefault();
    drop.classList.remove('drag');
  }));
  drop.addEventListener('drop',e=>{
    const file=e.dataTransfer.files && e.dataTransfer.files[0];
    if(!file) return;
    const dt=new DataTransfer();
    dt.items.add(file);
    input.files=dt.files;
    loadFile(file);
  });
}
function closeModal(){document.getElementById('modal').classList.add('hidden')}


function ensureProcessingBox(){
  let el=document.getElementById('processingBox');
  if(el) return el;
  el=document.createElement('div');
  el.id='processingBox';
  el.className='processingBox hidden';
  el.innerHTML=`<div class="processingCard">
    <div class="processingTop"><span id="processingIcon" class="processingIcon">⚙</span><button type="button" onclick="hideProcessingBox()">×</button></div>
    <h3 id="processingTitle">Processing PDF...</h3>
    <p id="processingText">Please wait, your file is being processed.</p>
    <div class="progressTrack"><div id="progressBar" class="progressBar"></div></div>
    <ul id="processingSteps" class="processingSteps"></ul>
  </div>`;
  document.body.appendChild(el);
  return el;
}
let progressTimer=null;
function getProcessingSteps(tool){
  const name=(tool?.t||'PDF Tool').toLowerCase();
  if(name.includes('translate')) return ['Uploading PDF securely','Reading text and page layout','Translating selected language','Rendering same-layout PDF','Preparing download'];
  if(name.includes('image')) return ['Uploading image securely','Reading image details','Compressing and converting','Optimizing final file','Preparing download'];
  if(name.includes('compress')) return ['Uploading PDF securely','Analysing PDF size','Compressing pages/images','Optimizing final file','Preparing download'];
  if(name.includes('merge')) return ['Uploading PDF files','Checking file order','Combining pages','Creating merged PDF','Preparing download'];
  if(name.includes('split')) return ['Uploading PDF securely','Reading page range','Splitting selected pages','Creating ZIP/PDF output','Preparing download'];
  if(name.includes('word')||name.includes('excel')||name.includes('powerpoint')||name.includes('jpg')||name.includes('html')) return ['Uploading file securely','Reading file content','Converting format','Creating output file','Preparing download'];
  if(name.includes('protect')||name.includes('unlock')) return ['Uploading PDF securely','Checking PDF security','Applying password settings','Saving secure PDF','Preparing download'];
  if(name.includes('watermark')||name.includes('page numbers')||name.includes('rotate')||name.includes('crop')||name.includes('redact')) return ['Uploading PDF securely','Reading PDF pages','Applying selected changes','Rendering updated PDF','Preparing download'];
  return ['Uploading file securely','Reading PDF pages','Processing selected tool','Creating output file','Preparing download'];
}
function showProcessingBox(title, steps, icon='⚙'){
  const el=ensureProcessingBox();
  document.getElementById('processingTitle').textContent=title||'Processing PDF...';
  document.getElementById('processingIcon').textContent=icon||'⚙';
  document.getElementById('processingText').textContent='File upload ho rahi hai. Window close mat karo.';
  const ul=document.getElementById('processingSteps');
  ul.innerHTML=(steps||['Uploading PDF','Processing pages','Creating output','Preparing download']).map((s,i)=>`<li class="${i===0?'active':''}">${s}</li>`).join('');
  const bar=document.getElementById('progressBar');
  bar.style.width='8%';
  el.classList.remove('hidden');
  let pct=8, idx=0;
  clearInterval(progressTimer);
  progressTimer=setInterval(()=>{
    pct=Math.min(92,pct+Math.random()*8+3);
    bar.style.width=pct+'%';
    const lis=[...ul.querySelectorAll('li')];
    const next=Math.min(lis.length-1, Math.floor((pct/100)*lis.length));
    if(next!==idx){ lis[idx]?.classList.remove('active'); lis[idx]?.classList.add('done'); idx=next; lis[idx]?.classList.add('active'); }
    if(pct>88) document.getElementById('processingText').textContent='Almost done... file download ready ho rahi hai.';
  },700);
}
function finishProcessingBox(ok=true, msg='Done! File downloaded.'){
  clearInterval(progressTimer);
  const bar=document.getElementById('progressBar'); if(bar) bar.style.width='100%';
  const txt=document.getElementById('processingText'); if(txt) txt.textContent=msg;
  const lis=document.querySelectorAll('#processingSteps li'); lis.forEach(x=>{x.classList.remove('active');x.classList.add(ok?'done':'error')});
  setTimeout(()=>hideProcessingBox(), ok?1200:2500);
}
function hideProcessingBox(){const el=document.getElementById('processingBox'); if(el) el.classList.add('hidden'); clearInterval(progressTimer);}

document.getElementById('toolForm').addEventListener('submit',async ev=>{
  ev.preventDefault();
  if(!current)return;
  const st=document.getElementById('status');
  const btn=document.querySelector('#toolForm .primary');
  const isImage=current.t==='Image Compressor';
  // PDF Translator tool has been removed, but this flag keeps the shared submit handler safe for all tools.
  const isTranslate=current.e==='/api/translate-pdf' || /translate/i.test(current.t||'');
  st.textContent=isTranslate?'Translating PDF... please wait.':(isImage?'Compressing image... please wait.':'Processing...');
  if(btn){btn.disabled=true; btn.dataset.oldText=btn.textContent; btn.textContent=isTranslate?'Translating...':(isImage?'Compressing...':'Processing...');}
  showProcessingBox((isTranslate?'Translating PDF...':(isImage?'Compressing Image...':current.t+' Processing...')), getProcessingSteps(current), current.i||'⚙');
  try{
    const fd=new FormData(ev.target);
    const fileInput = ev.target.querySelector('input[type="file"]');
    if(fileInput && fileInput.files && fileInput.files.length){
      const maxMB = current.t === 'PDF to Word' ? 25 : 120;
      for(const f of fileInput.files){
        if(f.size > maxMB * 1024 * 1024){
          throw new Error(`File too large for mobile/browser upload. Please use below ${maxMB}MB file.`);
        }
      }
    }
    const res=await fetchWithTimeout(API+current.e,{method:'POST',body:fd}, 240000);
    if(!res.ok){let j=await res.json().catch(()=>({error:'Tool error'}));throw new Error(j.error||'Something went wrong')}
    const blob=await res.blob();
    if(!blob || blob.size===0) throw new Error('Empty file received from server');
    let name='toolmitra_download';
    const disp=res.headers.get('content-disposition')||'';
    const m=disp.match(/filename\*?=(?:UTF-8''|")?([^";]+)/i);
    if(m)name=decodeURIComponent(m[1].replace(/"/g,''));
    await downloadBlobMobileFriendly(blob,name);
    st.textContent='Done! File downloaded.';
    finishProcessingBox(true,'Done! File downloaded.');
  }catch(err){
    st.textContent='Error: '+err.message;
    finishProcessingBox(false,'Error: '+err.message);
  }finally{
    if(btn){btn.disabled=false; btn.textContent=btn.dataset.oldText||'Process & Download';}
  }
});
render();

// ---------------- Modern iLovePDF-style client-side PDF editor ----------------
const editor = {
  pdf:null, fileName:'edited.pdf', scale:1.55, zoom:1, canvases:[], fabricPages:[], pageWraps:[], activePage:0,
  tool:'select', color:'#111827', size:24, stroke:3, history:[], currentFile:null
};
function libReady(){ return window.pdfjsLib && window.fabric && window.jspdf; }
function openPdfEditor(){
  closeModal();
  let old=document.getElementById('pdfEditorApp'); if(old) old.remove();
  const div=document.createElement('div'); div.id='pdfEditorApp'; div.innerHTML=`
    <div class="pe-header">
      <div class="pe-brand"><span class="pe-logo">✦</span><div><b>Advanced Edit PDF</b><small>Smart dropdown tools: select, insert, annotate, shapes, style, zoom & save</small></div></div>
      <div class="pe-actions"><button onclick="peHelpToggle()">Help</button><button onclick="peClose()" class="pe-close">×</button></div>
    </div>
    <div class="pe-ribbon pe-dropdown-ribbon">
      <div class="pe-quick">
        <button onclick="peTool('select')" id="peSelect" class="on">👆 Select / Move</button>
        <button onclick="peUndo()">↶ Undo</button>
        <button onclick="peDelete()">🗑 Delete</button>
      </div>
      <details class="pe-dd" open><summary>SELECT</summary><div class="pe-menu">
        <button onclick="peTool('select')">👆 Select object</button>
        <button onclick="peSelectAll()">☑ Select all on page</button>
        <button onclick="peDeselect()">✖ Deselect</button>
        <button onclick="peDuplicate()">⧉ Duplicate</button>
        <button onclick="peBringFront()">⬆ Bring front</button>
        <button onclick="peSendBack()">⬇ Send back</button>
        <button onclick="peClearPage()">🧹 Clear current page</button>
      </div></details>
      <details class="pe-dd"><summary>INSERT</summary><div class="pe-menu">
        <button onclick="peAddText()">A Add Text</button>
        <button onclick="peAddBigText()">T Heading Text</button>
        <button onclick="peAddCoverText()">▣ Cover + New Text</button>
        <button onclick="document.getElementById('peImage').click()">🖼 Insert Image</button>
        <button onclick="peAddSignature()">✒ Signature text</button>
        <input id="peImage" type="file" accept="image/*" hidden onchange="peAddImage(event)">
      </div></details>
      <details class="pe-dd"><summary>ANNOTATE</summary><div class="pe-menu">
        <button onclick="peAddHighlight()">🟨 Highlight</button>
        <button onclick="peAddRedact()">⬛ Redact / Black cover</button>
        <button onclick="peTool('draw')" id="peDraw">✎ Free Draw</button>
        <button onclick="peAddNote()">💬 Sticky Note</button>
        <button onclick="peAddUnderline()">▔ Underline</button>
        <button onclick="peAddStrike()">S̶ Strike line</button>
      </div></details>
      <details class="pe-dd"><summary>SHAPES</summary><div class="pe-menu">
        <button onclick="peAddBox('white')">□ White Cover</button>
        <button onclick="peAddBox('border')">▢ Border Box</button>
        <button onclick="peAddBox('fill')">▣ Color Fill Box</button>
        <button onclick="peAddCircle()">○ Circle / Oval</button>
        <button onclick="peAddLine()">／ Line</button>
        <button onclick="peAddArrow()">➜ Arrow</button>
      </div></details>
      <details class="pe-dd pe-style"><summary>STYLE</summary><div class="pe-menu">
        <label>Color <input id="peColor" type="color" value="#111827" onchange="peSetColor(this.value)"></label>
        <label>Size <input id="peSize" type="number" min="8" max="120" value="24" onchange="peSetSize(this.value)"></label>
        <label>Stroke <input id="peStroke" type="number" min="1" max="30" value="3" onchange="peSetStroke(this.value)"></label>
        <label>Opacity <input id="peOpacity" type="range" min="10" max="100" value="100" onchange="peSetOpacity(this.value)"></label>
        <label>Font <select id="peFont" onchange="peSetFont(this.value)"><option>Arial</option><option>Times New Roman</option><option>Courier New</option><option>Georgia</option><option>Verdana</option></select></label>
        <button onclick="peMakeBold()">B Bold Text</button>
        <button onclick="peMakeItalic()">I Italic Text</button>
      </div></details>
      <details class="pe-dd"><summary>ZOOM</summary><div class="pe-menu pe-zoommenu">
        <button onclick="peZoomOut()">− Zoom Out</button><b id="peZoomText">100%</b><button onclick="peZoomIn()">+ Zoom In</button><button onclick="peZoomFit()">Fit Width</button><button onclick="peZoomTo(.75)">75%</button><button onclick="peZoomReset()">100%</button><button onclick="peZoomTo(1.5)">150%</button>
      </div></details>
      <button class="pe-save" onclick="peSave()">💾 Save edited PDF</button>
    </div>
    <div class="pe-body">
      <aside id="peThumbs"><div class="pe-mini-title">Pages</div></aside>
      <main id="pePages"><div class="pe-empty"><div class="pe-empty-icon">📄</div><h3>Upload PDF to start editing</h3><p>Text add karo, old text ko white cover se hide karo, highlight/annotate/shapes/image/draw use karo, phir Save edited PDF.</p><label class="pe-upload pe-upload-center"><input id="peFile" type="file" accept="application/pdf,.pdf"><span>📄 Upload PDF</span></label></div></main>
      <section class="pe-help" id="peHelp"><h2>How to edit?</h2><div class="pe-step"><b>1. Select tool</b><p>Dropdown open karo: Select, Insert, Annotate, Shapes, Style, Zoom. Mouse bahar le jaoge ya next dropdown khologe to purana auto close hoga.</p></div><div class="pe-step"><b>2. Drag/Resize</b><p>Object select karke move, resize, rotate kar sakte ho. Text pe double click se edit hoga.</p></div><div class="pe-step"><b>3. Old text change</b><p>PDF ka original text direct edit nahi hota. White Cover se old text hide karke naya text add karo.</p></div><div class="pe-step"><b>4. Save</b><p>Save edited PDF se edited file download hogi.</p></div></section>
    </div>`;
  document.body.appendChild(div);
  if(window.pdfjsLib){ pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'; }
  document.getElementById('peFile').addEventListener('change', peLoadPdf);
  document.addEventListener('keydown', peKeyHandler);
  peInitDropdowns();
}
function peClose(){ document.removeEventListener('keydown', peKeyHandler); const e=document.getElementById('pdfEditorApp'); if(e)e.remove(); }
function peHelpToggle(){ const h=document.getElementById('peHelp'); if(h) h.classList.toggle('hide'); }
function peKeyHandler(e){ if(!document.getElementById('pdfEditorApp'))return; if(e.key==='Delete')peDelete(); if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='z'){e.preventDefault();peUndo();} }

function peCloseDropdowns(except=null){ document.querySelectorAll('#pdfEditorApp .pe-dd[open]').forEach(d=>{ if(d!==except) d.open=false; }); }
function peInitDropdowns(){
  const app=document.getElementById('pdfEditorApp'); if(!app)return;
  app.querySelectorAll('.pe-dd').forEach(dd=>{
    let timer=null;
    const summary=dd.querySelector('summary');
    summary.addEventListener('click',(e)=>{ e.preventDefault(); const willOpen=!dd.open; peCloseDropdowns(dd); dd.open=willOpen; });
    dd.addEventListener('mouseenter',()=>{ if(timer) clearTimeout(timer); });
    dd.addEventListener('mouseleave',()=>{ timer=setTimeout(()=>{ dd.open=false; },260); });
  });
  app.addEventListener('click',(e)=>{
    const menuBtn=e.target.closest('.pe-menu button');
    if(menuBtn) setTimeout(()=>peCloseDropdowns(),120);
    if(!e.target.closest('.pe-dd')) peCloseDropdowns();
  });
  document.addEventListener('click', peOutsideDropdownClose, true);
}
function peOutsideDropdownClose(e){ const app=document.getElementById('pdfEditorApp'); if(!app){document.removeEventListener('click',peOutsideDropdownClose,true);return;} if(!app.contains(e.target)) peCloseDropdowns(); }

function peSetColor(v){editor.color=v; editor.fabricPages.forEach(c=>{c.freeDrawingBrush.color=v;}); const c=activeCanvas(); const o=c?.getActiveObject(); if(o){ if(o.type==='i-text') o.set('fill',v); else if(['line','rect','ellipse','triangle','group'].includes(o.type)) { if(o.type==='group') o.getObjects().forEach(x=>{ if(x.stroke) x.set('stroke',v); if(x.fill && x.type!=='rect') x.set('fill',v); }); else o.set('stroke',v); } c.renderAll(); }}
function peSetSize(v){editor.size=parseInt(v)||24; editor.fabricPages.forEach(c=>{c.freeDrawingBrush.width=Math.max(2,editor.size/5);}); const c=activeCanvas(); const o=c?.getActiveObject(); if(o){ if(o.type==='i-text') o.set('fontSize',editor.size); if(o.type==='group') o.getObjects().forEach(x=>{if(x.type==='i-text')x.set('fontSize',editor.size)}); c.renderAll(); }}
function peSetStroke(v){editor.stroke=parseInt(v)||3; editor.fabricPages.forEach(c=>{c.freeDrawingBrush.width=editor.stroke;}); const c=activeCanvas(); const o=c?.getActiveObject(); if(o&&o.set){ if(o.type==='group') o.getObjects().forEach(x=>{if(x.strokeWidth)x.set('strokeWidth',editor.stroke)}); else if(o.strokeWidth!==undefined) o.set('strokeWidth',editor.stroke); c.renderAll(); }}
function peSetOpacity(v){ const c=activeCanvas(); const o=c?.getActiveObject(); if(o){o.set('opacity',(parseInt(v)||100)/100); c.renderAll();}}
function peSetFont(v){ const c=activeCanvas(); editor.font=v||'Arial'; if(c){ const o=c.getActiveObject(); if(o&&o.set){ if(o.type==='i-text') o.set('fontFamily',editor.font); if(o.type==='group') o.getObjects().forEach(x=>{if(x.type==='i-text')x.set('fontFamily',editor.font)}); c.renderAll(); } } }
function peMakeBold(){ const c=activeCanvas(); const o=c?.getActiveObject(); if(!o)return; if(o.type==='i-text') o.set('fontWeight', o.fontWeight==='bold'?'normal':'bold'); if(o.type==='group') o.getObjects().forEach(x=>{if(x.type==='i-text')x.set('fontWeight', x.fontWeight==='bold'?'normal':'bold')}); c.renderAll(); }
function peMakeItalic(){ const c=activeCanvas(); const o=c?.getActiveObject(); if(!o)return; if(o.type==='i-text') o.set('fontStyle', o.fontStyle==='italic'?'normal':'italic'); if(o.type==='group') o.getObjects().forEach(x=>{if(x.type==='i-text')x.set('fontStyle', x.fontStyle==='italic'?'normal':'italic')}); c.renderAll(); }
function peTool(t){ editor.tool=t; editor.fabricPages.forEach(c=>{c.isDrawingMode=(t==='draw'); c.selection=(t==='select'); c.defaultCursor=t==='draw'?'crosshair':'default';}); document.querySelectorAll('.pe-ribbon button').forEach(b=>b.classList.remove('on')); const id=t==='draw'?'peDraw':'peSelect'; const el=document.getElementById(id); if(el)el.classList.add('on'); }
async function peLoadPdf(ev){
  if(!libReady()){ alert('PDF editor libraries load nahi hui. Internet ON karke page refresh karo.'); return; }
  const file=ev.target.files[0]; if(!file)return; editor.currentFile=file; editor.fileName=file.name.replace(/\.pdf$/i,'')+'_edited.pdf';
  const buf=await file.arrayBuffer(); editor.pdf=await pdfjsLib.getDocument({data:buf}).promise;
  editor.canvases=[]; editor.fabricPages=[]; editor.pageWraps=[]; editor.activePage=0; editor.zoom=1;
  document.getElementById('pePages').innerHTML=''; document.getElementById('peThumbs').innerHTML='<div class="pe-mini-title">Pages</div>';
  for(let p=1;p<=editor.pdf.numPages;p++) await peRenderPage(p);
  peTool('select'); peApplyZoom(); peStatus('PDF loaded. Ab toolbar se edit karo.');
}
async function peRenderPage(num){
  const page=await editor.pdf.getPage(num); const viewport=page.getViewport({scale:editor.scale});
  const wrap=document.createElement('div'); wrap.className='pe-pageWrap'; wrap.dataset.page=num; wrap.onclick=()=>{editor.activePage=num-1;peMarkThumb();};
  const pageNo=document.createElement('div'); pageNo.className='pe-pageNo'; pageNo.textContent='Page '+num;
  const base=document.createElement('canvas'); base.width=viewport.width; base.height=viewport.height; base.className='pe-base';
  await page.render({canvasContext:base.getContext('2d'), viewport}).promise;
  const overlay=document.createElement('canvas'); overlay.width=viewport.width; overlay.height=viewport.height; overlay.className='pe-overlay';
  wrap.style.width=viewport.width+'px'; wrap.style.height=viewport.height+'px'; wrap.appendChild(pageNo); wrap.appendChild(base); wrap.appendChild(overlay); document.getElementById('pePages').appendChild(wrap);
  const fc=new fabric.Canvas(overlay,{preserveObjectStacking:true,selection:true,stopContextMenu:true}); fc.setWidth(viewport.width); fc.setHeight(viewport.height); fc.freeDrawingBrush.color=editor.color; fc.freeDrawingBrush.width=Math.max(2,editor.size/5);
  fc.on('mouse:down',()=>{editor.activePage=num-1;peMarkThumb();}); fc.on('mouse:dblclick',()=>{ const o=fc.getActiveObject(); if(o&&o.type==='i-text') o.enterEditing(); }); fc.on('object:added',()=>peMarkThumb());
  editor.canvases.push(base); editor.fabricPages.push(fc); editor.pageWraps.push(wrap);
  const th=document.createElement('canvas'); th.width=132; th.height=Math.round(base.height*(132/base.width)); th.getContext('2d').drawImage(base,0,0,th.width,th.height); const td=document.createElement('div'); td.className='pe-thumb'; td.dataset.page=num-1; td.innerHTML=`<span>${num}</span>`; td.prepend(th); td.onclick=()=>{editor.activePage=num-1;peMarkThumb();wrap.scrollIntoView({behavior:'smooth',block:'center'});}; document.getElementById('peThumbs').appendChild(td); peMarkThumb();
}
function peMarkThumb(){ document.querySelectorAll('.pe-thumb').forEach(t=>t.classList.toggle('active',Number(t.dataset.page)===editor.activePage)); }
function activeCanvas(){ return editor.fabricPages[editor.activePage] || editor.fabricPages[0]; }
function needPdf(){ if(!activeCanvas()){alert('Pehle PDF upload karo'); return false;} return true; }

function peSelectAll(){ const c=activeCanvas(); if(!c)return; const objs=c.getObjects(); if(!objs.length)return; const sel=new fabric.ActiveSelection(objs,{canvas:c}); c.setActiveObject(sel); c.renderAll(); }
function peDeselect(){ const c=activeCanvas(); if(c){c.discardActiveObject(); c.renderAll();} }

function peAddText(){ if(!needPdf())return; const c=activeCanvas(); const t=new fabric.IText('Type here',{left:90,top:90,fill:editor.color,fontSize:editor.size,fontFamily:'Arial',backgroundColor:'rgba(255,255,255,0.35)',padding:6}); c.add(t).setActiveObject(t); t.enterEditing(); c.renderAll(); peTool('select'); }
function peAddBigText(){ if(!needPdf())return; const c=activeCanvas(); const t=new fabric.IText('Heading text',{left:90,top:90,fill:editor.color,fontSize:Math.max(34,editor.size+10),fontFamily:editor.font||'Arial',fontWeight:'bold',backgroundColor:'rgba(255,255,255,0.2)',padding:6}); c.add(t).setActiveObject(t); t.enterEditing(); c.renderAll(); peTool('select'); }
function peAddCoverText(){ if(!needPdf())return; const c=activeCanvas(); const r=new fabric.Rect({left:85,top:85,width:260,height:54,fill:'#ffffff',stroke:'#ffffff',strokeWidth:1,rx:2,ry:2}); const t=new fabric.IText('New text',{left:98,top:96,fill:editor.color,fontSize:editor.size,fontFamily:'Arial'}); const g=new fabric.Group([r,t],{left:85,top:85}); c.add(g).setActiveObject(g); c.renderAll(); peTool('select'); peStatus('Old text ko cover karke naya text add ho gaya.'); }
function peAddBox(mode){ if(!needPdf())return; const c=activeCanvas(); const r=new fabric.Rect({left:90,top:120,width:250,height:80,fill:mode==='white'?'#ffffff':(mode==='fill'?editor.color:'rgba(255,255,255,0)'),stroke:mode==='border'?editor.color:(mode==='fill'?editor.color:'#ffffff'),strokeWidth:mode==='border'?editor.stroke:1,rx:2,ry:2}); c.add(r).setActiveObject(r); c.renderAll(); peTool('select'); }
function peAddHighlight(){ if(!needPdf())return; const c=activeCanvas(); const r=new fabric.Rect({left:100,top:150,width:280,height:34,fill:'rgba(255,230,0,0.45)',stroke:'rgba(255,196,0,0.9)',strokeWidth:1,rx:3,ry:3}); c.add(r).setActiveObject(r); c.renderAll(); peTool('select'); }
function peAddRedact(){ if(!needPdf())return; const c=activeCanvas(); const r=new fabric.Rect({left:100,top:150,width:280,height:34,fill:'#000000',stroke:'#000000',strokeWidth:1,rx:2,ry:2}); c.add(r).setActiveObject(r); c.renderAll(); peTool('select'); }
function peAddNote(){ if(!needPdf())return; const c=activeCanvas(); const r=new fabric.Rect({left:110,top:110,width:230,height:95,fill:'#fff7ad',stroke:'#eab308',strokeWidth:2,rx:10,ry:10}); const t=new fabric.IText('Note: write here',{left:128,top:132,fill:'#111827',fontSize:18,fontFamily:'Arial'}); const g=new fabric.Group([r,t],{left:110,top:110}); c.add(g).setActiveObject(g); c.renderAll(); peTool('select'); }
function peAddCircle(){ if(!needPdf())return; const c=activeCanvas(); const o=new fabric.Ellipse({left:120,top:120,rx:90,ry:55,fill:'rgba(255,255,255,0)',stroke:editor.color,strokeWidth:3}); c.add(o).setActiveObject(o); c.renderAll(); peTool('select'); }
function peAddLine(){ if(!needPdf())return; const c=activeCanvas(); const l=new fabric.Line([120,160,360,160],{stroke:editor.color,strokeWidth:Math.max(2,editor.stroke||3),strokeLineCap:'round'}); c.add(l).setActiveObject(l); c.renderAll(); peTool('select'); }
function peAddImage(ev){ const file=ev.target.files[0]; if(!file||!needPdf())return; const c=activeCanvas(); const reader=new FileReader(); reader.onload=()=>fabric.Image.fromURL(reader.result,img=>{ img.set({left:100,top:100,cornerStyle:'circle'}); const max=260; if(img.width>max) img.scale(max/img.width); c.add(img).setActiveObject(img); c.renderAll(); peTool('select'); }); reader.readAsDataURL(file); ev.target.value=''; }
function peDelete(){ const c=activeCanvas(); if(!c)return; const objs=c.getActiveObjects(); if(!objs.length)return; objs.forEach(o=>c.remove(o)); c.discardActiveObject(); c.renderAll(); }
function peUndo(){ const c=activeCanvas(); if(!c)return; const arr=c.getObjects(); if(arr.length){c.remove(arr[arr.length-1]); c.renderAll();} }
function peDuplicate(){ const c=activeCanvas(); const o=c?.getActiveObject(); if(!o)return; o.clone(cl=>{cl.set({left:(o.left||0)+25,top:(o.top||0)+25}); c.add(cl); c.setActiveObject(cl); c.renderAll();}); }
function peBringFront(){ const c=activeCanvas(); const o=c?.getActiveObject(); if(o){c.bringToFront(o); c.renderAll();} }
function peSendBack(){ const c=activeCanvas(); const o=c?.getActiveObject(); if(o){c.sendToBack(o); c.renderAll();} }
function peClearPage(){ const c=activeCanvas(); if(c && confirm('Current page ke sab added edits remove karne hain?')){c.clear(); c.renderAll();} }
function peAddSignature(){ if(!needPdf())return; const c=activeCanvas(); const t=new fabric.IText('Signature',{left:120,top:160,fill:editor.color,fontSize:Math.max(28,editor.size),fontFamily:'Georgia',fontStyle:'italic',backgroundColor:'rgba(255,255,255,0)',padding:4}); c.add(t).setActiveObject(t); t.enterEditing(); c.renderAll(); peTool('select'); }
function peAddUnderline(){ if(!needPdf())return; const c=activeCanvas(); const l=new fabric.Line([100,190,380,190],{stroke:editor.color,strokeWidth:Math.max(3,editor.stroke||3),strokeLineCap:'round'}); c.add(l).setActiveObject(l); c.renderAll(); peTool('select'); }
function peAddStrike(){ if(!needPdf())return; const c=activeCanvas(); const l=new fabric.Line([100,190,380,190],{stroke:editor.color,strokeWidth:Math.max(2,editor.stroke||3),strokeLineCap:'round'}); c.add(l).setActiveObject(l); c.renderAll(); peTool('select'); }
function peAddArrow(){ if(!needPdf())return; const c=activeCanvas(); const line=new fabric.Line([110,160,350,160],{stroke:editor.color,strokeWidth:Math.max(3,editor.stroke||3),strokeLineCap:'round'}); const tri=new fabric.Triangle({left:350,top:160,width:18,height:18,fill:editor.color,angle:90,originX:'center',originY:'center'}); const g=new fabric.Group([line,tri],{left:110,top:150}); c.add(g).setActiveObject(g); c.renderAll(); peTool('select'); }
function peZoomTo(v){ editor.zoom=Math.max(0.45,Math.min(2.5,Number(v)||1)); peApplyZoom(); }
function peZoomReset(){ editor.zoom=1; peApplyZoom(); }
function peZoomIn(){ editor.zoom=Math.min(2.5,+(editor.zoom+0.15).toFixed(2)); peApplyZoom(); }
function peZoomOut(){ editor.zoom=Math.max(0.45,+(editor.zoom-0.15).toFixed(2)); peApplyZoom(); }
function peZoomFit(){
  const main=document.getElementById('pePages');
  const base=editor.canvases[0];
  const logicalWidth=base ? base.width : 900;
  if(main) editor.zoom=Math.max(.45,Math.min(1.8,(main.clientWidth-90)/logicalWidth));
  else editor.zoom=1;
  peApplyZoom();
}
function peApplyZoom(){
  const z=editor.zoom;
  const zt=document.getElementById('peZoomText');
  if(zt) zt.textContent=Math.round(z*100)+'%';
  editor.pageWraps.forEach((wrap,i)=>{
    const base=editor.canvases[i], fc=editor.fabricPages[i];
    if(!base || !fc) return;
    const sw=base.width*z, sh=base.height*z;
    // IMPORTANT: do not use CSS transform for zoom. Fabric mouse coordinates
    // become wrong after dragging/resizing and saved PDF position shifts.
    // We resize only the CSS display size; internal canvas coordinates stay same.
    wrap.style.transform='none';
    wrap.style.width=sw+'px';
    wrap.style.height=sh+'px';
    wrap.style.marginBottom='34px';
    base.style.width=sw+'px'; base.style.height=sh+'px';
    fc.setDimensions({width:sw,height:sh},{cssOnly:true});
    fc.calcOffset();
    fc.renderAll();
  });
}
function peStatus(msg){ let s=document.getElementById('peStatus'); if(!s){s=document.createElement('div');s.id='peStatus';document.querySelector('.pe-header')?.appendChild(s);} s.textContent=msg; setTimeout(()=>{if(s.textContent===msg)s.textContent='';},3500); }
async function peSave(){
  if(!editor.canvases.length){alert('Pehle PDF upload karo'); return;}
  const btn=document.querySelector('#pdfEditorApp .pe-save'); const old=btn.textContent; btn.textContent='Saving...'; btn.disabled=true;
  try{
    const { jsPDF } = window.jspdf; let pdf=null;
    for(let i=0;i<editor.canvases.length;i++){
      const base=editor.canvases[i], fc=editor.fabricPages[i]; fc.discardActiveObject(); fc.renderAll();
      const merged=document.createElement('canvas'); merged.width=base.width; merged.height=base.height; const ctx=merged.getContext('2d'); ctx.fillStyle='#fff'; ctx.fillRect(0,0,merged.width,merged.height); ctx.drawImage(base,0,0); ctx.drawImage(fc.lowerCanvasEl,0,0);
      const img=merged.toDataURL('image/jpeg',0.94); const w=base.width, h=base.height; const orient=w>h?'l':'p';
      if(!pdf) pdf=new jsPDF({orientation:orient,unit:'px',format:[w,h],compress:true}); else pdf.addPage([w,h],orient);
      pdf.addImage(img,'JPEG',0,0,w,h);
    }
    pdf.save(editor.fileName || 'edited.pdf'); peStatus('Edited PDF download ho gayi.');
  }catch(e){ alert('Save error: '+e.message); }
  btn.textContent=old; btn.disabled=false;
}


// Dark / Light mode
function toggleTheme(){
  document.body.classList.toggle('dark');
  localStorage.setItem('tm_theme', document.body.classList.contains('dark') ? 'dark' : 'light');
}
if(localStorage.getItem('tm_theme') === 'dark'){
  document.body.classList.add('dark');
}
