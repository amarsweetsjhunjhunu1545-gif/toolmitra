// Backend API URL
// Netlify par deployed frontend ke liye Render backend URL use hoga.
// Local testing me http://127.0.0.1:5000 auto use hoga.
const API = (location.hostname === 'localhost' || location.hostname === '127.0.0.1' || location.protocol === 'file:')
  ? 'http://127.0.0.1:5000'
  : 'https://toolmitra-api.onrender.com';

function safeFileName(name){
  return (name || 'orbixapdftool_download').replace(/[\\/:*?"<>|]+/g,'_').replace(/\s+/g,'_');
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
 {t:'Sign PDF',e:'/api/sign-pdf',d:'Draw signature, add company seal, place anywhere on PDF.',i:'✍',type:'single',accept:'.pdf'},
 {t:'PDF to JPG',e:'/api/pdf-to-jpg',d:'Upload PDF and click Convert JPG Image. Single page PDF downloads JPG directly; multi-page PDF downloads ZIP with JPG images.',i:'JPG',type:'single',accept:'.pdf',btn:'Convert JPG Image'},
 {t:'JPG to PDF',e:'/api/jpg-to-pdf',d:'Convert JPG/PNG images into one PDF.',i:'IMG',type:'multi',accept:'image/*'}, {t:'Watermark',e:'/api/add-watermark',d:'Stamp text watermark over PDF pages.',i:'◆',type:'single',accept:'.pdf',extra:[['text','Watermark text','OrbixaPDFTool']]},
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
 {t:'Smart Summarizer',e:'/api/ai-summarizer',d:'Generate a quick text summary from PDF.',i:'TXT',type:'single',accept:'.pdf'},
 {t:'Delete Pages',e:'/api/delete-pages',d:'Remove unwanted pages.',i:'🗑',type:'single',accept:'.pdf',extra:[['pages','Pages to delete e.g. 2,4-5']]},
 {t:'Extract Pages',e:'/api/extract-pages',d:'Save selected pages as a new PDF.',i:'📌',type:'single',accept:'.pdf',extra:[['pages','Pages e.g. 1-3,5','1']]},
 {t:'Extract PDF Text',e:'/api/extract-text',d:'Download text from PDF.',i:'TXT',type:'single',accept:'.pdf'},
 {t:'Image Compressor',e:'/api/image-compress',d:'Image upload karo, target size KB enter karo aur JPG/PNG format me download karo.',i:'⚡',type:'image',accept:'image/*',btn:'Compress to KB & Download',extra:[['target_kb','Target size in KB e.g. 100',''],['output_format','Output format: auto / jpg / png','auto']]},
 {t:'QR Generator',e:'#client-qr',d:'Advanced QR: link, UPI/PhonePe, WhatsApp, WiFi, contact, email, SMS, location. PNG/JPG/SVG/PDF download.',i:'QR',type:'client'},
];
let current=window.currentTool||null;
let activeCategory='all';
function toolCategory(x){const s=(x.t+' '+x.d).toLowerCase(); if(s.includes('qr')||s.includes('camera')) return 'image'; if(s.includes('image')||s.includes('jpg')||s.includes('png')||s.includes('webp')||s.includes('photo')) return 'image'; if(s.includes('protect')||s.includes('unlock')||s.includes('redact')) return 'secure'; if(s.includes('to ')||s.includes('convert')) return 'convert'; return 'pdf';}
function setCategory(cat,btn){activeCategory=cat; document.querySelectorAll('.categoryChips button').forEach(b=>b.classList.remove('active')); if(btn)btn.classList.add('active'); filterTools();}
function getSlug(name) { return name.toLowerCase().replace(/ /g, '-').replace(/\//g, '-'); }
function render(list=tools){const el=document.getElementById('toolGrid');if(!el)return;el.innerHTML=list.map((x)=>{const cat=toolCategory(x).toUpperCase();return `<a href=\"${getSlug(x.t)+'.html'}\" class="tool" style="text-decoration:none;color:inherit;display:block;"><span class="tag">${cat}</span><div class="icon">${x.i}</div><h3>${x.t}</h3><p>${x.d}</p><small class="toolAction">Open tool →</small></a>`}).join('')}
function filterTools(){const q=document.getElementById('search').value.toLowerCase();render(tools.filter(x=>(activeCategory==='all'||toolCategory(x)===activeCategory) && (x.t+x.d).toLowerCase().includes(q)))}
function quickOpen(name){const i=tools.findIndex(x=>x.t===name);if(i>-1)window.location.href=getSlug(tools[i].t)+'.html'}
function openTool(i){
  
  window.location.href=getSlug(tools[i].t)+'.html';
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
  st.textContent=isImage?'Compressing image... please wait.':'Processing...';
  if(btn){btn.disabled=true; btn.dataset.oldText=btn.textContent; btn.textContent=isImage?'Compressing...':'Processing...';}
  showProcessingBox((isImage?'Compressing Image...':current.t+' Processing...'), getProcessingSteps(current), current.i||'⚙');
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
    
    // Client-Side Processing for instant (< 2s) response
    let handledClientSide = false;
    if (window.PDFLib) {
      const { PDFDocument } = window.PDFLib;
      try {
        if (current.t === 'Merge PDF') {
          const mergedPdf = await PDFDocument.create();
          for (const file of fileInput.files) {
            const buf = await file.arrayBuffer();
            const pdf = await PDFDocument.load(buf);
            const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
            copiedPages.forEach((page) => mergedPdf.addPage(page));
          }
          const pdfBytes = await mergedPdf.save();
          await downloadBlobMobileFriendly(new Blob([pdfBytes], { type: 'application/pdf' }), 'merged.pdf');
          handledClientSide = true;
        } else if (current.t === 'Rotate PDF') {
          const angle = parseInt(fd.get('angle') || '90');
          const buf = await fileInput.files[0].arrayBuffer();
          const pdf = await PDFDocument.load(buf);
          const pages = pdf.getPages();
          pages.forEach(page => page.setRotation({type: 'degrees', angle: (page.getRotation().angle + angle) % 360}));
          const pdfBytes = await pdf.save();
          await downloadBlobMobileFriendly(new Blob([pdfBytes], { type: 'application/pdf' }), 'rotated.pdf');
          handledClientSide = true;
        } else if (current.t === 'Delete Pages') {
          const buf = await fileInput.files[0].arrayBuffer();
          const pdf = await PDFDocument.load(buf);
          let toDelete = [];
          const pagesStr = fd.get('pages') || '';
          if(pagesStr) {
            pagesStr.split(',').forEach(p => {
              if(p.includes('-')) {
                 let [s,e] = p.split('-');
                 for(let i=parseInt(s); i<=parseInt(e); i++) toDelete.push(i-1);
              } else {
                 toDelete.push(parseInt(p)-1);
              }
            });
          }
          toDelete.sort((a,b)=>b-a).forEach(idx => {
             if(idx >= 0 && idx < pdf.getPageCount()) pdf.removePage(idx);
          });
          const pdfBytes = await pdf.save();
          await downloadBlobMobileFriendly(new Blob([pdfBytes], { type: 'application/pdf' }), 'deleted_pages.pdf');
          handledClientSide = true;
        }
      } catch(e) {
        console.error("Client side processing failed:", e);
      }
    }
    
    if (handledClientSide) {
      st.textContent='Done! File downloaded instantly.';
      finishProcessingBox(true,'Done! File downloaded instantly.');
      if(btn){btn.disabled=false; btn.textContent=btn.dataset.oldText||'Process & Download';}
      return;
    }

    const res=await fetchWithTimeout(API+current.e,{method:'POST',body:fd}, 240000);
    if(!res.ok){let j=await res.json().catch(()=>({error:'Tool error'}));throw new Error(j.error||'Something went wrong')}
    const blob=await res.blob();
    if(!blob || blob.size===0) throw new Error('Empty file received from server');
    let name='orbixapdftool_download';
    const disp=res.headers.get('content-disposition')||'';
    const m=disp.match(/filename\*?=(?:UTF-8''|")?([^";]+)/i);
    if(m)name=decodeURIComponent(m[1].replace(/"/g,''));
    await downloadBlobMobileFriendly(blob,name);
    if(isImage && blob.size){
      const sizeKB = (blob.size/1024).toFixed(1);
      st.textContent=`✅ Done! Compressed image: ${sizeKB} KB downloaded.`;
      finishProcessingBox(true,`Done! Compressed image: ${sizeKB} KB`);
    } else if(current && current.t === 'Compress PDF'){
      const sizeKB = (blob.size/1024).toFixed(0);
      const target = res.headers.get('X-Target-Size');
      const exact = res.headers.get('X-Exact-Target') === 'yes';
      const reached = res.headers.get('X-Target-Reached') === 'yes';
      let msg = `✅ Done! Compressed PDF: ${sizeKB} KB downloaded.`;
      if(target && exact){
        msg = `✅ Done! Exact target PDF: ${(Number(target)/1024).toFixed(0)} KB downloaded.`;
      } else if(target && !reached){
        msg = `⚠️ PDF compressed to ${sizeKB} KB. Is PDF ko target se chhota karna possible nahi hua, smallest result download ho gaya.`;
      }
      st.textContent = msg;
      finishProcessingBox(true, msg);
    } else {
      st.textContent='Done! File downloaded.';
      finishProcessingBox(true,'Done! File downloaded.');
    }
  }catch(err){
    st.textContent='Error: '+err.message;
    finishProcessingBox(false,'Error: '+err.message);
  }finally{
    if(btn){btn.disabled=false; btn.textContent=btn.dataset.oldText||'Process & Download';}
  }
});
if(window.currentTool) {
  if(window.currentTool.type === 'image') setupImageToolPreview();
} else {
  render();
}

// ---------------- Modern iLovePDF-style client-side PDF editor ----------------
const editor = {
  pdf:null, fileName:'edited.pdf', scale:1.55, zoom:1, canvases:[], fabricPages:[], pageWraps:[], activePage:0,
  tool:'select', color:'#111827', size:24, stroke:3, history:[], currentFile:null
};
function libReady(){ return window.pdfjsLib && window.fabric && window.jspdf; }
