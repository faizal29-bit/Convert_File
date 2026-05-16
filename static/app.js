/* LocalConvert — Frontend Application */

// ── TOOL DEFINITIONS ──────────────────────────────────────────────────────

const TOOLS = [
  {
    id: 'image-to-pdf',
    icon: '🖼️',
    title: 'Gambar ke PDF',
    desc: 'JPG, PNG, WEBP, BMP → PDF',
    accept: 'image/*',
    acceptLabel: 'JPG, PNG, WEBP, GIF, BMP, TIFF',
    multi: true,
    color: '#f97316',
    colorBg: 'rgba(249,115,22,0.15)',
    colorBorder: 'rgba(249,115,22,0.3)',
    cardColorA: 'rgba(249,115,22,0.07)',
    resultIcon: '📄',
  },
  {
    id: 'image-to-word',
    icon: '🖼️',
    title: 'Gambar ke Word',
    desc: 'JPG, PNG, WEBP → DOCX',
    accept: 'image/*',
    acceptLabel: 'JPG, PNG, WEBP, GIF, BMP, TIFF',
    multi: true,
    color: '#3b82f6',
    colorBg: 'rgba(59,130,246,0.15)',
    colorBorder: 'rgba(59,130,246,0.3)',
    cardColorA: 'rgba(59,130,246,0.07)',
    resultIcon: '📝',
  },
  {
    id: 'word-to-pdf',
    icon: '📝',
    title: 'Word ke PDF',
    desc: 'DOC, DOCX → PDF',
    accept: '.doc,.docx',
    acceptLabel: 'DOC, DOCX',
    multi: false,
    color: '#6366f1',
    colorBg: 'rgba(99,102,241,0.15)',
    colorBorder: 'rgba(99,102,241,0.3)',
    cardColorA: 'rgba(99,102,241,0.07)',
    resultIcon: '📄',
  },
  {
    id: 'pdf-to-word',
    icon: '📄',
    title: 'PDF ke Word',
    desc: 'PDF → DOCX yang dapat diedit',
    accept: '.pdf',
    acceptLabel: 'PDF',
    multi: false,
    color: '#ec4899',
    colorBg: 'rgba(236,72,153,0.15)',
    colorBorder: 'rgba(236,72,153,0.3)',
    cardColorA: 'rgba(236,72,153,0.07)',
    resultIcon: '📝',
  },
  {
    id: 'ppt-to-pdf',
    icon: '📊',
    title: 'PPT ke PDF',
    desc: 'PPT, PPTX → PDF',
    accept: '.ppt,.pptx',
    acceptLabel: 'PPT, PPTX',
    multi: false,
    color: '#f59e0b',
    colorBg: 'rgba(245,158,11,0.15)',
    colorBorder: 'rgba(245,158,11,0.3)',
    cardColorA: 'rgba(245,158,11,0.07)',
    resultIcon: '📄',
  },
  {
    id: 'pdf-to-ppt',
    icon: '📄',
    title: 'PDF ke PPT',
    desc: 'PDF → PPTX (Hanya PDF berteks digital)', 
    accept: '.pdf',
    acceptLabel: 'PDF',
    multi: false,
    color: '#14b8a6',
    colorBg: 'rgba(20,184,166,0.15)',
    colorBorder: 'rgba(20,184,166,0.3)',
    cardColorA: 'rgba(20,184,166,0.07)',
    resultIcon: '📊',
  },
  {
    id: 'word-to-ppt',
    icon: '📝',
    title: 'Word ke PPT',
    desc: 'DOCX → Presentasi PPTX',
    accept: '.doc,.docx',
    acceptLabel: 'DOC, DOCX',
    multi: false,
    color: '#8b5cf6',
    colorBg: 'rgba(139,92,246,0.15)',
    colorBorder: 'rgba(139,92,246,0.3)',
    cardColorA: 'rgba(139,92,246,0.07)',
    resultIcon: '📊',
  },
  {
    id: 'excel-to-pdf',
    icon: '📊',
    title: 'Excel ke PDF',
    desc: 'XLS, XLSX → PDF',
    accept: '.xls,.xlsx',
    acceptLabel: 'XLS, XLSX',
    multi: false,
    color: '#10b981',
    colorBg: 'rgba(16,185,129,0.15)',
    colorBorder: 'rgba(16,185,129,0.3)',
    cardColorA: 'rgba(16,185,129,0.07)',
    resultIcon: '📄',
  },
  {
    id: 'image-to-ppt',
    icon: '🖼️',
    title: 'Gambar ke PPT',
    desc: 'JPG, PNG, WEBP → PPTX',
    accept: 'image/*',
    acceptLabel: 'JPG, PNG, WEBP, GIF, BMP, TIFF',
    multi: true,
    color: '#ef4444',
    colorBg: 'rgba(239,68,68,0.15)',
    colorBorder: 'rgba(239,68,68,0.3)',
    cardColorA: 'rgba(239,68,68,0.07)',
    resultIcon: '📊',
  },
  {
    id: 'compress-file',
    icon: '🗜️', // Ikon mesin pres / vakum
    title: 'Kompres Semua File',
    desc: 'Kecilkan ukuran Gambar, PDF, atau bungkus file lain jadi ZIP',
    accept: '*/*', // Bintang/Bintang artinya menerima SEMUA jenis file tanpa batas
    acceptLabel: 'Semua File Bebas',
    multi: false,
    color: '#ec4899', // Warna pink utama kamu
    colorBg: 'rgba(236,72,153,0.15)',
    colorBorder: 'rgba(236,72,153,0.3)',
    cardColorA: 'rgba(236,72,153,0.07)',
    resultIcon: '📦',
  },
];

// ── STATE ─────────────────────────────────────────────────────────────────

let currentTool  = null;
let selectedFiles = [];  // Array of File objects
let lastJobId    = null;
let lastJobFiles = [];

// ── DOM HELPERS ───────────────────────────────────────────────────────────

const $ = id => document.getElementById(id);
const showEl  = el => { if (el) el.style.display = ''; };
const hideEl  = el => { if (el) el.style.display = 'none'; };
const blockEl = el => { if (el) el.style.display = 'block'; };
const flexEl  = el => { if (el) el.style.display = 'flex'; };

function showView(id) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function formatBytes(bytes) {
  if (bytes < 1024)       return bytes + ' B';
  if (bytes < 1048576)    return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(1) + ' MB';
}

function fileExt(name) {
  return name.split('.').pop().toLowerCase();
}

function fileTypeIcon(name) {
  const ext = fileExt(name);
  const map = {
    pdf: '📄', doc: '📝', docx: '📝',
    ppt: '📊', pptx: '📊',
    xls: '📗', xlsx: '📗',
    jpg: '🖼️', jpeg: '🖼️', png: '🖼️',
    gif: '🖼️', bmp: '🖼️', webp: '🖼️', tiff: '🖼️', tif: '🖼️',
  };
  return map[ext] || '📁';
}

function showToast(msg, duration = 5000) {
  let t = document.querySelector('.toast');
  if (!t) {
    t = document.createElement('div');
    t.className = 'toast';
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove('show'), duration);
}

// ── BUILD HOME ────────────────────────────────────────────────────────────

function buildToolGrid() {
  const grid = $('tools-grid');
  grid.innerHTML = TOOLS.map(tool => `
    <div class="tool-card"
         data-id="${tool.id}"
         style="
           --card-color-a: ${tool.cardColorA};
           --card-border: ${tool.colorBorder};
         "
         onclick="selectTool('${tool.id}')">
      <div class="tool-card-icon"
           style="background:${tool.colorBg}; color:${tool.color};">
        ${tool.icon}
      </div>
      <div class="tool-card-body">
        <div class="tool-card-title">${tool.title}</div>
        <div class="tool-card-desc">${tool.desc}</div>
      </div>
      <div class="tool-card-arrow">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
          <path d="M5 12h14M12 5l7 7-7 7"/>
        </svg>
      </div>
    </div>
  `).join('');
}

// ── SELECT TOOL ───────────────────────────────────────────────────────────

function selectTool(toolId) {
  currentTool = TOOLS.find(t => t.id === toolId);
  if (!currentTool) return;

  // Reset state
  selectedFiles = [];
  lastJobId = null;
  lastJobFiles = [];

  // Update header
  $('th-icon').textContent = currentTool.icon;
  $('th-icon').style.cssText = `
    background: ${currentTool.colorBg};
    color: ${currentTool.color};
    width:64px;height:64px;border-radius:16px;
    display:flex;align-items:center;justify-content:center;font-size:30px;flex-shrink:0;
  `;
  $('th-title').textContent = currentTool.title;
  $('th-desc').textContent  = `Konversi file ${currentTool.acceptLabel}`;
  $('upload-hint').textContent = `Format didukung: ${currentTool.acceptLabel}${currentTool.multi ? ' · Bisa pilih banyak file sekaligus' : ''}`;

  // File input setup
  const fi = $('file-input');
  fi.accept   = currentTool.accept;
  fi.multiple = currentTool.multi;
  fi.value    = '';

  // Hide sections
  hideEl($('file-list-wrap'));
  hideEl($('action-bar'));
  hideEl($('progress-wrap'));
  hideEl($('results-wrap'));
  showEl($('upload-area'));
  renderFileList();

  showView('view-tool');
}

// ── FILE HANDLING ─────────────────────────────────────────────────────────

function addFiles(fileList) {
  const files = Array.from(fileList);
  if (!currentTool.multi) {
    // Non-multi: ganti semua
    selectedFiles = files.slice(0, 1);
  } else {
    // Multi: tambah (hindari duplikat nama)
    for (const f of files) {
      if (!selectedFiles.find(x => x.name === f.name && x.size === f.size)) {
        selectedFiles.push(f);
      }
    }
  }
  renderFileList();
  updateActionBar();
}

function removeFile(index) {
  selectedFiles.splice(index, 1);
  renderFileList();
  updateActionBar();
}

function renderFileList() {
  const list  = $('file-list');
  const wrap  = $('file-list-wrap');
  const count = $('file-count-label');

  if (selectedFiles.length === 0) {
    hideEl(wrap);
    list.innerHTML = '';
    return;
  }

  flexEl($('upload-area'));  // keep visible to allow adding more
  blockEl(wrap);
  count.textContent = `${selectedFiles.length} file dipilih`;

  list.innerHTML = selectedFiles.map((f, i) => `
    <div class="file-item">
      <div class="file-thumb">${fileTypeIcon(f.name)}</div>
      <div class="file-info">
        <div class="file-name">${f.name}</div>
        <div class="file-size">${formatBytes(f.size)}</div>
      </div>
      <button class="file-remove" onclick="removeFile(${i})" title="Hapus">×</button>
    </div>
  `).join('');
}

function updateActionBar() {
  if (selectedFiles.length > 0) {
    flexEl($('action-bar'));
  } else {
    hideEl($('action-bar'));
  }
}

// ── CONVERSION ────────────────────────────────────────────────────────────

async function doConvert() {
  if (!selectedFiles.length || !currentTool) return;

  // Hide actions, show progress
  hideEl($('action-bar'));
  hideEl($('results-wrap'));
  blockEl($('progress-wrap'));

  const fill = $('progress-fill');
  const msg  = $('progress-msg');
  fill.classList.add('running');
  msg.textContent = `Sedang mengkonversi ${selectedFiles.length} file…`;

  // Build FormData
  const fd = new FormData();
  fd.append('type', currentTool.id);
  for (const f of selectedFiles) {
    fd.append('files', f, f.name);
  }

  try {
    const res = await fetch('/api/convert', {
      method: 'POST',
      body: fd
    });
    const data = await res.json();

    fill.classList.remove('running');
    hideEl($('progress-wrap'));

    if (!res.ok || data.error) {
      showToast('❌ Error: ' + (data.error || 'Gagal mengkonversi'));
      flexEl($('action-bar'));
      return;
    }

    // Show results
    lastJobId    = data.job_id;
    lastJobFiles = data.files;
    showResults(data.job_id, data.files);

  } catch (err) {
    fill.classList.remove('running');
    hideEl($('progress-wrap'));
    showToast('❌ Koneksi gagal: ' + err.message);
    flexEl($('action-bar'));
  }
}

function showResults(jobId, files) {
  const list = $('results-list');
  const wrap = $('results-wrap');
  const sub  = $('results-sub');
  const zipBtn = $('btn-zip');

  sub.textContent = `${files.length} file berhasil dikonversi`;

  list.innerHTML = files.map(f => `
    <div class="result-item">
      <div class="result-icon">${currentTool.resultIcon}</div>
      <div class="result-name">${f.name}</div>
      <a class="btn-dl"
         href="/api/download/${jobId}/${encodeURIComponent(f.name)}"
         download="${f.name}">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
          <polyline points="7 10 12 15 17 10"/>
          <line x1="12" y1="15" x2="12" y2="3"/>
        </svg>
        Unduh
      </a>
    </div>
  `).join('');

  // Show ZIP button only if multiple files
  if (files.length > 1) {
    flexEl(zipBtn);
    zipBtn.onclick = () => {
      window.location.href = `/api/download-zip/${jobId}`;
    };
  } else {
    hideEl(zipBtn);
  }

  blockEl(wrap);
  wrap.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function resetTool() {
  selectedFiles = [];
  lastJobId = null;
  lastJobFiles = [];
  renderFileList();
  hideEl($('action-bar'));
  hideEl($('progress-wrap'));
  hideEl($('results-wrap'));
  $('file-input').value = '';
}

// ── STATUS CHECK ──────────────────────────────────────────────────────────

async function checkStatus() {
  try {
    const res  = await fetch('/api/status');
    const data = await res.json();
    const badge = $('lo-badge');
    const sbar  = $('status-bar');

    if (data.libreoffice) {
      badge.textContent = '✅ LibreOffice OK';
      badge.className   = 'badge-local ok';
    } else {
      badge.textContent = '⚠️ LibreOffice Tidak Ada';
      badge.className   = 'badge-local warn';
      if (sbar) {
        sbar.innerHTML = `
          <div style="
            display:inline-block;
            background:rgba(245,158,11,0.1);
            border:1px solid rgba(245,158,11,0.3);
            color:#fcd34d;
            font-size:12px;
            padding:8px 18px;
            border-radius:8px;
          ">
            ⚠️ <strong>LibreOffice</strong> tidak terdeteksi.
            Fitur Word/PPT/Excel→PDF membutuhkannya.
            <a href="https://www.libreoffice.org/download/libreoffice/"
               target="_blank"
               style="color:#60a5fa;text-decoration:underline;margin-left:8px;">
              Download di sini
            </a>
          </div>
        `;
      }
    }
  } catch {
    // Server belum siap, abaikan
  }
}

// ── EVENT LISTENERS ───────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {

  // Build home grid
  buildToolGrid();

  // Logo / back
  $('logo-link').addEventListener('click', () => showView('view-home'));
  $('back-btn').addEventListener('click', () => showView('view-home'));

  // File input change
  $('file-input').addEventListener('change', e => {
    if (e.target.files.length) addFiles(e.target.files);
    e.target.value = '';
  });

  // Upload area click
  $('upload-area').addEventListener('click', e => {
    if (e.target.closest('.btn-pick')) return;
    $('file-input').click();
  });

  // Pick button
  $('btn-pick').addEventListener('click', e => {
    e.stopPropagation();
    $('file-input').click();
  });

  // Add more
  $('btn-add-more').addEventListener('click', () => $('file-input').click());

  // Convert
  $('btn-convert').addEventListener('click', doConvert);

  // Clear all
  $('btn-clear-all').addEventListener('click', resetTool);

  // New conversion
  $('btn-new-conv').addEventListener('click', resetTool);

  // Drag & drop
  const ua = $('upload-area');
  ua.addEventListener('dragenter', e => { e.preventDefault(); ua.classList.add('drag-over'); });
  ua.addEventListener('dragover',  e => { e.preventDefault(); });
  ua.addEventListener('dragleave', e => {
    if (!ua.contains(e.relatedTarget)) ua.classList.remove('drag-over');
  });
  ua.addEventListener('drop', e => {
    e.preventDefault();
    ua.classList.remove('drag-over');
    if (e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
  });

  // Global drag & drop prevention
  document.addEventListener('dragover',  e => e.preventDefault());
  document.addEventListener('drop',      e => e.preventDefault());

  // Check status
  checkStatus();
});
