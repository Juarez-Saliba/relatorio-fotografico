const $ = (s) => document.querySelector(s);
const $$ = (s) => Array.from(document.querySelectorAll(s));

const state = {
  items: [],
  params: {
    ORIENTACAO: 'retrato',
    MARGENS_CM: { sup: 1.27, inf: 1.27, esq: 1.27, dir: 1.27 },
    ORGAO_NOME: '',
    NOME_ARQUIVO: 'itens_imagens.docx',
    ITEM_OFFSET: 0,
    PARAMS_LOCKED: false,
  },
};

let DRAGGING_ID = null;

// ─── Persistência de sessão (IndexedDB + localStorage, TTL 24h) ───────────────
const _DB_NAME = 'relatorio_fotografico_db';
const _DB_VER  = 1;
const _STORE   = 'imagens';
const _LS_KEY  = 'relatorio_state_v1';
const _TTL_MS  = 24 * 60 * 60 * 1000;
let   _restoring = false;

function _openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(_DB_NAME, _DB_VER);
    req.onupgradeneeded = (e) => e.target.result.createObjectStore(_STORE, { keyPath: 'id' });
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror   = (e) => reject(e.target.error);
  });
}

function _dbPut(db, record) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(_STORE, 'readwrite');
    tx.objectStore(_STORE).put(record);
    tx.oncomplete = resolve;
    tx.onerror = (e) => reject(e.target.error);
  });
}

function _dbGet(db, id) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(_STORE, 'readonly');
    const req = tx.objectStore(_STORE).get(id);
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror   = (e) => reject(e.target.error);
  });
}

function _dbPurge(db, validIds) {
  return new Promise((resolve) => {
    const tx = db.transaction(_STORE, 'readwrite');
    const store = tx.objectStore(_STORE);
    const req = store.getAllKeys();
    req.onsuccess = (e) => {
      e.target.result.filter(k => !validIds.has(k)).forEach(k => store.delete(k));
    };
    tx.oncomplete = resolve;
    tx.onerror    = resolve;
  });
}

let _saveTimer = null;
function scheduleSave() {
  if (_restoring) return;
  clearTimeout(_saveTimer);
  _saveTimer = setTimeout(saveState, 600);
}

async function saveState() {
  try {
    const db = await _openDB();
    const validIds = new Set();
    for (const item of state.items) {
      for (const img of item.imagens) {
        validIds.add(img.id);
        if (img.file) {
          const data = await img.file.arrayBuffer();
          await _dbPut(db, { id: img.id, data, name: img.file.name, type: img.file.type });
        }
      }
    }
    await _dbPurge(db, validIds);
    const payload = {
      ts: Date.now(),
      params: state.params,
      items: state.items.map(it => ({
        id: it.id, nome: it.nome, config: it.config, imported: it.imported || false,
        imagens: it.imagens.map(img => ({
          id: img.id,
          name: img.file?.name || 'image.png',
          type: img.file?.type || 'image/png',
          w: img.w, h: img.h,
        })),
      })),
    };
    localStorage.setItem(_LS_KEY, JSON.stringify(payload));
  } catch (e) { console.warn('saveState:', e); }
}

async function restoreState() {
  try {
    const raw = localStorage.getItem(_LS_KEY);
    if (!raw) return false;
    const saved = JSON.parse(raw);
    if (!saved?.ts || Date.now() - saved.ts > _TTL_MS) {
      localStorage.removeItem(_LS_KEY);
      return false;
    }
    const db = await _openDB();
    Object.assign(state.params, saved.params || {});
    const items = [];
    for (const si of (saved.items || [])) {
      const imagens = [];
      for (const imgMeta of (si.imagens || [])) {
        const rec = await _dbGet(db, imgMeta.id);
        if (!rec) continue;
        const blob = new Blob([rec.data], { type: rec.type || imgMeta.type || 'image/png' });
        const file = new File([blob], rec.name || imgMeta.name || 'image.png', { type: blob.type });
        const url  = URL.createObjectURL(blob);
        imagens.push({ id: imgMeta.id, file, url, w: imgMeta.w, h: imgMeta.h });
      }
      items.push({ id: si.id, nome: si.nome, config: si.config, imported: si.imported || false, imagens });
    }
    state.items = items;
    return true;
  } catch (e) { console.warn('restoreState:', e); return false; }
}

function syncParamsToUI() {
  const p = state.params;
  const orEl = $('#orientacaoSelect'); if (orEl) orEl.value = p.ORIENTACAO || 'retrato';
  const ogEl = $('#orgaoNome');        if (ogEl) ogEl.value = p.ORGAO_NOME  || '';
  const nfEl = $('#nomeArquivo');      if (nfEl) nfEl.value = p.NOME_ARQUIVO || 'itens_imagens.docx';
  const ms = p.MARGENS_CM || {};
  const mS = $('#mSup'); if (mS) mS.value = ms.sup ?? 1.27;
  const mI = $('#mInf'); if (mI) mI.value = ms.inf ?? 1.27;
  const mE = $('#mEsq'); if (mE) mE.value = ms.esq ?? 1.27;
  const mD = $('#mDir'); if (mD) mD.value = ms.dir ?? 1.27;
  updateOffsetUI();
  setFieldsLocked(!!p.PARAMS_LOCKED);
}
// ─────────────────────────────────────────────────────────────────────────────

function setGenerateEnabled(enabled) {
  const topBtn = $('#generateBtn');
  if (topBtn) topBtn.disabled = !enabled;
  const bottomBtn = $('#generateBtnBottom');
  if (bottomBtn) bottomBtn.disabled = !enabled;
}

async function ensureDocx() {
  if (window.docx) return true;
  const candidates = [
    // Prioriza arquivo local, se existir
    './lib/docx.umd.min.js',
    './lib/docx.umd.js',
    // UMD válidos nas CDNs (nome correto é index.umd.js)
    'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js',
    'https://unpkg.com/docx@8.5.0/build/index.umd.js',
  ];
  const load = (src) => new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    // Evita exigir CORS para script tradicional
    // s.crossOrigin = 'anonymous';
    const to = setTimeout(() => { s.remove(); reject(new Error('timeout')); }, 8000);
    s.onload = () => { clearTimeout(to); resolve(true); };
    s.onerror = () => { clearTimeout(to); reject(new Error('load error')); };
    document.head.appendChild(s);
  });
  for (const url of candidates) {
    try { await load(url); if (window.docx) return true; } catch {}
  }
  // Fallback: tenta ESM dinâmico
  const esmCandidates = [
    'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.mjs',
    'https://unpkg.com/docx@8.5.0/build/index.mjs',
  ];
  for (const url of esmCandidates) {
    try {
      const mod = await import(/* @vite-ignore */ url);
      if (mod) {
        window.docx = mod;
        return true;
      }
    } catch {}
  }
  return false;
}

function cmToTwips(cm) { return Math.round((cm / 2.54) * 1440); }
function mmToTwips(mm) { return cmToTwips(mm / 10); }
function ptToPx(pt) { return Math.round(pt * (96 / 72)); }
function mmToPx(mm) { return Math.round((mm / 25.4) * 96); }
function cmToPx(cm) { return Math.round((cm / 2.54) * 96); }
const RENDER_SCALE = 2; // renderiza com mais pixels para melhorar nitidez no Word
function ptToTwips(pt) { return Math.round(pt * 20); }
function ptToCm(pt) { return (pt * 25.4) / 72 / 10; } // pt -> mm -> cm
function pxToTwips(px) { return Math.round((px / 96) * 1440); }

function parseColor(input) {
  if (!input) return '#000000';
  const s = String(input).trim();
  if (/^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(s)) return s;
  const m = s.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
  if (m) {
    const r = Math.min(255, Math.max(0, parseInt(m[1], 10)));
    const g = Math.min(255, Math.max(0, parseInt(m[2], 10)));
    const b = Math.min(255, Math.max(0, parseInt(m[3], 10)));
    return `#${[r,g,b].map(x=>x.toString(16).padStart(2,'0')).join('')}`;
  }
  return '#000000';
}

function newItem(idx) {
  return {
    id: crypto.randomUUID(),
    nome: `Item ${String(idx).padStart(2, '0')}`,
    imagens: [],
    config: {
      autoSize: true,
      perOriEnabled: false,
      larguraValor: 8,
      larguraUn: 'cm',
      alturaValor: 6,
      alturaUn: 'cm',
      bordaCor: '#000000',
      modoAjuste: 'contain',
      hLarguraValor: 0,
      hLarguraUn: 'cm',
      hAlturaValor: 0,
      hAlturaUn: 'cm',
      vLarguraValor: 0,
      vLarguraUn: 'cm',
      vAlturaValor: 0,
      vAlturaUn: 'cm',
    },
  };
}

function updateProgress(p) {
  const val = Math.max(0, Math.min(100, Math.round(p)));
  const fill = $('#progressPopupFill');
  if (fill) fill.style.width = `${val}%`;
  const txt = $('#progressPopupText');
  if (txt) {
    const base = txt.dataset.base || txt.textContent || 'Gerando documento…';
    txt.textContent = `${base.replace(/\s+—\s+\d+%$/, '')} — ${val}%`.trim();
    txt.dataset.base = base.replace(/\s+—\s+\d+%$/, '') || base;
  }
}

function showStatus(msg) {
  const p = $('#progressPopup');
  if (p) {
    // garante que não há outro popout visível atrás
    const dl = $('#dlPopup');
    if (dl) dl.classList.add('hidden');
    const logo = $('#pwLogo');
    const pl = $('#progressPopupLogo');
    if (logo && pl && logo.src) pl.src = logo.src;
    const txt = $('#progressPopupText');
    if (txt) { txt.textContent = msg || 'Gerando documento…'; txt.dataset.base = msg || 'Gerando documento…'; }
    const fill = $('#progressPopupFill');
    if (fill) fill.style.width = '0%';
    p.classList.remove('hidden');
    p.style.display = 'block';
  }
}
function hideStatus() {
  const p = $('#progressPopup');
  if (p) p.classList.add('hidden');
  const fill = $('#progressPopupFill');
  if (fill) fill.style.width = '0%';
}

function showDownloadToast() {
  const p = $('#dlPopup');
  if (!p) return;
  // garante que o popout de progresso não está visível
  const gp = $('#progressPopup');
  if (gp) gp.classList.add('hidden');
  const logo = $('#pwLogo');
  const pl = $('#dlPopupLogo');
  if (logo && pl && logo.src) pl.src = logo.src;
  p.classList.remove('hidden');
  const fill = p.querySelector('.popout-bar-fill');
  if (fill) {
    fill.style.animation = 'none';
    void fill.offsetWidth;
    fill.style.animation = '';
  }
  setTimeout(() => {
    p.classList.add('hidden');
  }, 5000);
}

function hideDownloadToast() {
  const p = $('#dlPopup');
  if (p) p.classList.add('hidden');
}

function render() {
  const c = $('#itemsContainer');
  c.innerHTML = '';
  state.items.forEach((it, idx) => {
    const card = document.createElement('div');
    card.className = 'item-card';
    card.dataset.id = it.id;
    const header = document.createElement('div');
    header.className = 'item-header';
    const title = document.createElement('div');
    title.className = 'item-title';
    title.textContent = `${it.nome}`;
    if (it.imported) {
      const badge = document.createElement('span');
      badge.className = 'imported-badge';
      badge.textContent = 'importado';
      title.appendChild(badge);
    }
    header.appendChild(title);
    // botão remover será adicionado no rodapé do card
    const cfg = document.createElement('div');
    cfg.className = 'item-config';
    cfg.innerHTML = `
      <div class="row">
        <label class="chk">
          <input type="checkbox" ${it.config.autoSize?'checked':''} data-k="autoSize" />
          tamanho automático
        </label>
      </div>
      <div class="row per-ori">
        <span class="lbl">Horizontal</span>
        <span class="mini-lbl">Largura</span>
        <input type="number" step="0.1" value="${it.config.hLarguraValor}" data-k="hLarguraValor" />
        <select data-k="hLarguraUn">
          <option value="cm"${it.config.hLarguraUn==='cm'?' selected':''}>cm</option>
          <option value="mm"${it.config.hLarguraUn==='mm'?' selected':''}>mm</option>
        </select>
        <span class="mini-lbl">Altura</span>
        <input type="number" step="0.1" value="${it.config.hAlturaValor}" data-k="hAlturaValor" />
        <select data-k="hAlturaUn">
          <option value="cm"${it.config.hAlturaUn==='cm'?' selected':''}>cm</option>
          <option value="mm"${it.config.hAlturaUn==='mm'?' selected':''}>mm</option>
        </select>
      </div>
      <div class="row per-ori">
        <span class="lbl">Vertical</span>
        <span class="mini-lbl">Largura</span>
        <input type="number" step="0.1" value="${it.config.vLarguraValor}" data-k="vLarguraValor" />
        <select data-k="vLarguraUn">
          <option value="cm"${it.config.vLarguraUn==='cm'?' selected':''}>cm</option>
          <option value="mm"${it.config.vLarguraUn==='mm'?' selected':''}>mm</option>
        </select>
        <span class="mini-lbl">Altura</span>
        <input type="number" step="0.1" value="${it.config.vAlturaValor}" data-k="vAlturaValor" />
        <select data-k="vAlturaUn">
          <option value="cm"${it.config.vAlturaUn==='cm'?' selected':''}>cm</option>
          <option value="mm"${it.config.vAlturaUn==='mm'?' selected':''}>mm</option>
        </select>
      </div>
    `;
    cfg.querySelectorAll('input,select').forEach(el => {
      el.addEventListener('change', (e) => {
        const k = e.target.getAttribute('data-k');
        if (!k) return;
        if (e.target.type === 'checkbox') it.config[k] = e.target.checked;
        else if (e.target.type === 'number') it.config[k] = parseFloat(e.target.value || '0') || 0;
        else it.config[k] = e.target.value;
        if (k === 'autoSize') {
          it.config.perOriEnabled = !it.config.autoSize; // interno, sem checkbox
        }
        const ori = cfg.querySelectorAll('[data-k=\"hLarguraValor\"],[data-k=\"hLarguraUn\"],[data-k=\"hAlturaValor\"],[data-k=\"hAlturaUn\"],[data-k=\"vLarguraValor\"],[data-k=\"vLarguraUn\"],[data-k=\"vAlturaValor\"],[data-k=\"vAlturaUn\"]');
        ori.forEach(inp => { inp.disabled = !!it.config.autoSize || !it.config.perOriEnabled; });
        if (k === 'autoSize') {
          render();
        }
      });
    });
    const ori = cfg.querySelectorAll('[data-k=\"hLarguraValor\"],[data-k=\"hLarguraUn\"],[data-k=\"hAlturaValor\"],[data-k=\"hAlturaUn\"],[data-k=\"vLarguraValor\"],[data-k=\"vLarguraUn\"],[data-k=\"vAlturaValor\"],[data-k=\"vAlturaUn\"]');
    ori.forEach(inp => { inp.disabled = !!it.config.autoSize || !it.config.perOriEnabled; });
    const thumbs = document.createElement('div');
    thumbs.className = 'thumbs';
    it.imagens.forEach((img, i) => {
      const th = document.createElement('div');
      th.className = 'thumb';
      th.draggable = true;
      th.dataset.id = img.id;
      const im = document.createElement('img');
      im.src = img.url;
      th.appendChild(im);
      const rm = document.createElement('button');
      rm.className = 'remove btn mini';
      rm.textContent = 'Remover';
      rm.addEventListener('click', () => {
        it.imagens = it.imagens.filter(z => z.id !== img.id);
        reorderImagesByMajority(it);
        render();
      });
      th.appendChild(rm);
      th.addEventListener('dragstart', (ev) => {
        DRAGGING_ID = img.id;
        th.classList.add('dragging');
        ev.dataTransfer.setData('text/plain', img.id);
        ev.dataTransfer.effectAllowed = 'move';
      });
      th.addEventListener('dragend', () => { th.classList.remove('dragging'); DRAGGING_ID = null; });
      th.addEventListener('dragover', (ev) => {
        ev.preventDefault();
        ev.dataTransfer.dropEffect = 'move';
        if (!DRAGGING_ID || DRAGGING_ID === img.id) return;
        const draggedEl = thumbs.querySelector(`[data-id="${DRAGGING_ID}"]`);
        if (!draggedEl) return;
        const rect = th.getBoundingClientRect();
        const insertAfter = ev.clientX > (rect.left + rect.width / 2);
        if (insertAfter) th.insertAdjacentElement('afterend', draggedEl);
        else th.insertAdjacentElement('beforebegin', draggedEl);
      });
      th.addEventListener('drop', (ev) => {
        ev.preventDefault();
        const order = Array.from(thumbs.querySelectorAll('.thumb')).map(el => el.dataset.id);
        const map = {}; it.imagens.forEach(x => { map[x.id] = x; });
        it.imagens = order.filter(id => map[id]).map(id => map[id]);
        render();
      });
      thumbs.appendChild(th);
    });
    const drop = document.createElement('div');
    drop.className = 'dropzone';
    drop.textContent = 'Arraste imagens aqui ou clique para selecionar';
    const file = document.createElement('input');
    file.type = 'file';
    file.accept = 'image/*';
    file.multiple = true;
    file.style.display = 'none';
    drop.addEventListener('click', () => file.click());
    drop.addEventListener('dragover', (ev) => { ev.preventDefault(); drop.classList.add('dragover'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
    drop.addEventListener('drop', (ev) => {
      ev.preventDefault();
      drop.classList.remove('dragover');
      const fl = Array.from(ev.dataTransfer.files || []).filter(f => f.type.startsWith('image/'));
      if (fl.length) void addFilesToItem(it, fl);
    });
    file.addEventListener('change', (ev) => {
      const fl = Array.from(ev.target.files || []);
      if (fl.length) void addFilesToItem(it, fl);
      file.value = '';
    });
    const rowWrap = document.createElement('div');
    rowWrap.className = 'item-row';
    const side = document.createElement('div');
    side.className = 'item-side-controls';
    const upBtn = document.createElement('button');
    upBtn.className = 'btn icon-btn';
    upBtn.textContent = '↑';
    if (idx === 0) {
      upBtn.style.display = 'none';
    } else {
      upBtn.addEventListener('click', () => swapItems(idx, idx - 1));
    }
    const downBtn = document.createElement('button');
    downBtn.className = 'btn icon-btn';
    downBtn.textContent = '↓';
    downBtn.disabled = idx === state.items.length - 1;
    downBtn.addEventListener('click', () => swapItems(idx, idx + 1));
    const sel = document.createElement('select');
    sel.className = 'swap-select';
    const opt0 = document.createElement('option');
    opt0.value = '';
    opt0.textContent = 'Trocar com…';
    opt0.disabled = true;
    opt0.selected = true;
    sel.appendChild(opt0);
    state.items.forEach((other, jdx) => {
      if (jdx === idx) return;
      const opt = document.createElement('option');
      opt.value = String(jdx);
      opt.textContent = other.nome;
      sel.appendChild(opt);
    });
    sel.addEventListener('change', (e) => {
      const j = parseInt(e.target.value, 10);
      if (!Number.isNaN(j)) swapItems(idx, j);
      e.target.value = '';
    });
    side.appendChild(upBtn);
    side.appendChild(downBtn);
    side.appendChild(sel);
    // seletor "Ir para..."
    const gotoSel = document.createElement('select');
    gotoSel.className = 'swap-select';
    const g0 = document.createElement('option');
    g0.value = '';
    g0.textContent = 'Ir para…';
    g0.disabled = true;
    g0.selected = true;
    gotoSel.appendChild(g0);
    state.items.forEach((other, jdx) => {
      if (jdx === idx) return;
      const opt = document.createElement('option');
      opt.value = String(jdx);
      opt.textContent = other.nome;
      gotoSel.appendChild(opt);
    });
    gotoSel.addEventListener('change', (e) => {
      const j = parseInt(e.target.value, 10);
      if (!Number.isNaN(j)) moveItemTo(idx, j);
      e.target.value = '';
    });
    side.appendChild(gotoSel);
    const body = document.createElement('div');
    body.className = 'item-body';
    body.appendChild(header);
    body.appendChild(cfg);
    body.appendChild(thumbs);
    body.appendChild(drop);
    body.appendChild(file);
    const footer = document.createElement('div');
    footer.className = 'item-footer';
    const rmBtn = document.createElement('button');
    rmBtn.className = 'btn ghost mini';
    rmBtn.textContent = 'Remover';
    rmBtn.addEventListener('click', () => {
      state.items = state.items.filter(x => x.id !== it.id);
      renumberItems();
      if (state.items.length === 0) setGenerateEnabled(false);
      render();
    });
    footer.appendChild(rmBtn);
    body.appendChild(footer);
    rowWrap.appendChild(body);
    rowWrap.appendChild(side);
    card.appendChild(rowWrap);
    c.appendChild(card);
  });
  if (state.items.length > 0) {
    const footer = document.createElement('div');
    footer.className = 'add-footer';
    const addBtn = document.createElement('button');
    addBtn.className = 'btn';
    addBtn.textContent = 'Adicionar Item';
    addBtn.addEventListener('click', () => {
      const it = newItem(state.items.length + 1);
      state.items.push(it);
      setGenerateEnabled(true);
      render();
      setTimeout(() => {
        const el = document.querySelector(`.item-card[data-id="${it.id}"]`);
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }, 0);
    });
    footer.appendChild(addBtn);
    const genBtnBottom = document.createElement('button');
    genBtnBottom.id = 'generateBtnBottom';
    genBtnBottom.className = 'btn primary';
    genBtnBottom.textContent = 'Gerar .docx';
    genBtnBottom.disabled = state.items.length === 0;
    genBtnBottom.addEventListener('click', onGenerate);
    footer.appendChild(genBtnBottom);
    c.appendChild(footer);
  }
  scheduleSave();
}

function renumberItems() {
  state.items.forEach((it, i) => {
    it.nome = `Item ${String(i + 1).padStart(2, '0')}`;
  });
}

function swapItems(i, j) {
  if (i < 0 || j < 0 || i >= state.items.length || j >= state.items.length) return;
  const arr = state.items;
  const tmp = arr[i];
  arr[i] = arr[j];
  arr[j] = tmp;
  renumberItems();
  render();
}

function moveItemTo(i, j) {
  if (i < 0 || j < 0 || i >= state.items.length || j >= state.items.length) return;
  if (i === j) return;
  const arr = state.items;
  const [item] = arr.splice(i, 1);
  arr.splice(j, 0, item);
  renumberItems();
  render();
}

function reorderImagesByMajority(item) {
  const imgs = item?.imagens || [];
  if (imgs.length <= 1) return;
  const horizontais = [];
  const verticais = [];
  for (const im of imgs) {
    const w = im.w || 0;
    const h = im.h || 0;
    if (w >= h) horizontais.push(im);
    else verticais.push(im);
  }
  item.imagens = (verticais.length > horizontais.length)
    ? [...verticais, ...horizontais]
    : [...horizontais, ...verticais];
}

async function addFilesToItem(item, files) {
  for (const f of files) {
    const url = URL.createObjectURL(f);
    let w = null, h = null;
    try {
      const bmp = await fileToImageBitmap(f);
      w = bmp.width; h = bmp.height;
    } catch {}
    item.imagens.push({ id: crypto.randomUUID(), file: f, url, w, h });
  }
  // Completa dimensões ausentes das imagens antigas, quando possível
  for (const img of item.imagens) {
    if ((img.w == null || img.h == null) && img.file) {
      try {
        const bmp = await fileToImageBitmap(img.file);
        img.w = bmp.width; img.h = bmp.height;
      } catch {}
    }
  }
  reorderImagesByMajority(item);
  setGenerateEnabled(state.items.length > 0);
  render();
}

function bindBasics() {
  $('#aboutBtn').addEventListener('click', () => $('#aboutModal').classList.remove('hidden'));
  $('#closeAbout').addEventListener('click', () => $('#aboutModal').classList.add('hidden'));
  const importInput = $('#importDocxInput');
  const importBtn   = $('#importDocxBtn');
  if (importBtn && importInput) {
    importBtn.addEventListener('click', () => importInput.click());
    importInput.addEventListener('change', (e) => {
      const f = e.target.files?.[0];
      if (f) void onImportDocx(f);
      importInput.value = '';
    });
  }
  const resetBtn = $('#resetBtn');
  if (resetBtn) resetBtn.addEventListener('click', () => void resetApp());
  $('#addItemBtn').addEventListener('click', () => {
    const it = newItem(state.items.length + 1);
    state.items.push(it);
    setGenerateEnabled(true);
    render();
    setTimeout(() => {
      const el = document.querySelector(`.item-card[data-id="${it.id}"]`);
      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 0);
  });
  $('#orientacaoSelect').addEventListener('change', (e) => { state.params.ORIENTACAO = e.target.value; scheduleSave(); });
  const orgEl = $('#orgaoNome');
  if (orgEl) {
    const autosize = () => {
      const s = document.createElement('span');
      const cs = getComputedStyle(orgEl);
      s.style.visibility = 'hidden';
      s.style.position = 'absolute';
      s.style.whiteSpace = 'pre';
      s.style.font = cs.font;
      s.textContent = orgEl.value || orgEl.placeholder || '';
      document.body.appendChild(s);
      const w = Math.ceil(s.offsetWidth + 24);
      document.body.removeChild(s);
      orgEl.style.width = Math.min(w, (orgEl.parentElement ? orgEl.parentElement.clientWidth : w)) + 'px';
    };
    orgEl.addEventListener('input', (e) => {
      state.params.ORGAO_NOME = e.target.value || '';
      autosize();
      scheduleSave();
    });
    orgEl.addEventListener('change', (e) => { state.params.ORGAO_NOME = e.target.value || ''; scheduleSave(); });
    autosize();
  }
  $('#mSup').addEventListener('change', (e) => { state.params.MARGENS_CM.sup = parseFloat(e.target.value || '0') || 0; scheduleSave(); });
  $('#mInf').addEventListener('change', (e) => { state.params.MARGENS_CM.inf = parseFloat(e.target.value || '0') || 0; scheduleSave(); });
  $('#mEsq').addEventListener('change', (e) => { state.params.MARGENS_CM.esq = parseFloat(e.target.value || '0') || 0; scheduleSave(); });
  $('#mDir').addEventListener('change', (e) => { state.params.MARGENS_CM.dir = parseFloat(e.target.value || '0') || 0; scheduleSave(); });
  $('#nomeArquivo').addEventListener('change', (e) => { state.params.NOME_ARQUIVO = e.target.value || 'itens_imagens.docx'; scheduleSave(); });
  $('#generateBtn').addEventListener('click', onGenerate);
  const el = $('#pwLogo');
  if (el) {
    (async () => {
      const candidates = [
        'logo/logo_two.png','LOGO/logo_two.png',
        'logo/pw.png','logo/pw.jpg','logo/pw.svg',
        'logo/logo.png','logo/logo.jpg','logo/logo.svg',
        'LOGO/pw.png','LOGO/pw.jpg','LOGO/pw.svg',
        'LOGO/logo.png','LOGO/logo.jpg','LOGO/logo.svg'
      ];
      for (const u of candidates) {
        const found = await new Promise(res => {
          const img = new Image();
          img.onload  = () => res(true);
          img.onerror = () => res(false);
          img.src = u;
        });
        if (found) { el.src = u; el.classList.remove('hidden'); break; }
      }
    })();
  }
}

async function fileToImageBitmap(file) {
  const buf = await file.arrayBuffer();
  const blob = new Blob([buf], { type: file.type || 'image/png' });
  try { return await createImageBitmap(blob); } catch {
    return await new Promise((res, rej) => {
      const img = new Image();
      img.onload = () => res(img);
      img.onerror = rej;
      img.src = URL.createObjectURL(blob);
    });
  }
}

function drawToCanvas(img, targetW, targetH, opt) {
  const c = document.createElement('canvas');
  c.width = targetW;
  c.height = targetH;
  const ctx = c.getContext('2d');
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.clearRect(0,0,c.width,c.height);
  let sx=0, sy=0, sw=img.width, sh=img.height;
  if (opt.manterProporcao) {
    const srcR = img.width / img.height;
    const dstR = targetW / targetH;
    if (opt.modoAjuste === 'cover') {
      if (srcR > dstR) {
        const newW = img.height * dstR;
        sx = Math.max(0, Math.floor((img.width - newW) / 2));
        sw = Math.floor(newW);
      } else {
        const newH = img.width / dstR;
        sy = Math.max(0, Math.floor((img.height - newH) / 2));
        sh = Math.floor(newH);
      }
    } else {
      if (srcR > dstR) {
        const drawW = targetW;
        const drawH = Math.round(drawW / srcR);
        const dy = Math.floor((targetH - drawH) / 2);
        ctx.drawImage(img, 0, 0, img.width, img.height, 0, dy, drawW, drawH);
        drawBorder(ctx, targetW, targetH, opt);
        return c;
      } else {
        const drawH = targetH;
        const drawW = Math.round(drawH * srcR);
        const dx = Math.floor((targetW - drawW) / 2);
        ctx.drawImage(img, 0, 0, img.width, img.height, dx, 0, drawW, drawH);
        drawBorder(ctx, targetW, targetH, opt);
        return c;
      }
    }
  }
  ctx.drawImage(img, sx, sy, sw, sh, 0, 0, targetW, targetH);
  drawBorder(ctx, targetW, targetH, opt);
  return c;
}

function drawBorder(ctx, w, h, opt) {
  const espPx = opt.bordaEspPx || 0;
  if (espPx <= 0) return;
  ctx.save();
  ctx.strokeStyle = parseColor(opt.bordaCor || '#000000');
  ctx.lineWidth = espPx;
  const off = espPx / 2;
  ctx.strokeRect(off, off, w - espPx, h - espPx);
  ctx.restore();
}

async function processImage(file, itemCfg) {
  const bordaEspPx = 0;
  const img = await fileToImageBitmap(file);
  let larguraPx;
  let alturaPx;
  if (itemCfg.targetWidthPx != null && itemCfg.targetHeightPx != null) {
    larguraPx = Math.max(1, Math.round(itemCfg.targetWidthPx));
    alturaPx = Math.max(1, Math.round(itemCfg.targetHeightPx));
  } else if (itemCfg.perOriEnabled) {
    const isHoriz = img.width >= img.height;
    const wVal = isHoriz ? itemCfg.hLarguraValor : itemCfg.vLarguraValor;
    const wUn = isHoriz ? itemCfg.hLarguraUn : itemCfg.vLarguraUn;
    const hVal = isHoriz ? itemCfg.hAlturaValor : itemCfg.vAlturaValor;
    const hUn = isHoriz ? itemCfg.hAlturaUn : itemCfg.vAlturaUn;
    larguraPx = wUn === 'cm' ? cmToPx(wVal) : mmToPx(wVal);
    alturaPx = hUn === 'cm' ? cmToPx(hVal) : mmToPx(hVal);
  } else {
    larguraPx = itemCfg.larguraUn === 'cm' ? cmToPx(itemCfg.larguraValor) : mmToPx(itemCfg.larguraValor);
    alturaPx = itemCfg.alturaUn === 'cm' ? cmToPx(itemCfg.alturaValor) : mmToPx(itemCfg.alturaValor);
  }
  const displayW = Math.max(1, larguraPx);
  const displayH = Math.max(1, alturaPx);
  const renderW = Math.max(1, Math.round(displayW * RENDER_SCALE));
  const renderH = Math.max(1, Math.round(displayH * RENDER_SCALE));
  const c = drawToCanvas(img, renderW, renderH, {
    manterProporcao: true,
    modoAjuste: itemCfg.modoAjuste || 'contain',
    bordaEspPx,
    bordaCor: itemCfg.bordaCor || '#000000',
  });
  const blob = await new Promise(res => c.toBlob(res, 'image/png'));
  const buf = await blob.arrayBuffer();
  return { buf, width: displayW, height: displayH };
}

function buildTableForItem(docx, images, cellSpaceTwips, dynamicRows) {
  const rows = dynamicRows || Math.ceil(images.length / 2);
  const cols = 2;
  const cells = [];
  const padTwips = ptToTwips(3); // espaçamento interno fixo de 3pt
  for (let r=0;r<rows;r++) {
    const rowCells = [];
    for (let c=0;c<cols;c++) {
      const idx = r*cols + c;
      const img = images[idx];
      if (img) {
        rowCells.push(new docx.TableCell({
          margins: { top: padTwips, bottom: padTwips, left: padTwips, right: padTwips },
          verticalAlign: docx.VerticalAlign.CENTER,
          borders: { top: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
          children: [
            new docx.Paragraph({
              alignment: docx.AlignmentType.CENTER,
              spacing: { before: 0, after: 0, line: 240 },
              children: [
                new docx.ImageRun({
                  data: img.buf,
                  transformation: { width: img.width, height: img.height },
                }),
              ],
            }),
          ],
        }));
      } else {
        rowCells.push(new docx.TableCell({
          margins: { top: padTwips, bottom: padTwips, left: padTwips, right: padTwips },
          verticalAlign: docx.VerticalAlign.CENTER,
          borders: { top: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
          children: [ new docx.Paragraph({ spacing: { before: 0, after: 0 } }) ],
        }));
      }
    }
    const base = r * cols;
    const rowMaxHeightPx = Math.max(
      images[base]?.height || 0,
      images[base + 1]?.height || 0,
    );
    const height = rowMaxHeightPx
      ? { value: pxToTwips(rowMaxHeightPx) + padTwips * 2, rule: docx.HeightRule.ATLEAST }
      : undefined;
    cells.push(new docx.TableRow({ children: rowCells, height }));
  }
  return new docx.Table({
    width: { size: 0, type: docx.WidthType.AUTO },
    rows: cells,
    alignment: docx.AlignmentType.CENTER,
    layout: docx.TableLayoutType.AUTOFIT,
    cellSpacing: cellSpaceTwips,
    borders: { top: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, insideH: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, insideV: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
  });
}

function buildTableForLayout(docx, images, layoutRows, cellSpaceTwips) {
  const padTwips = ptToTwips(3);
  // Map processed images by id for quick access
  const byId = {};
  images.forEach((im, ix) => { byId[ix] = im; });
  // But our processed images are in same order as layoutRows flatten; create a map by matching order
  // Build a linear list by DOM order in render: 'images' corresponds to layout order, so we can consume sequentially.
  let cursor = 0;
  const rows = [];
  for (const row of layoutRows) {
    const rowCells = [];
    const count = row.length;
    for (let i = 0; i < count; i++) {
      const img = images[cursor++];
      if (img) {
        rowCells.push(new docx.TableCell({
          margins: { top: padTwips, bottom: padTwips, left: padTwips, right: padTwips },
          verticalAlign: docx.VerticalAlign.CENTER,
          borders: { top: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
          children: [
            new docx.Paragraph({
              alignment: docx.AlignmentType.CENTER,
              spacing: { before: 0, after: 0, line: 240 },
              children: [
                new docx.ImageRun({
                  data: img.buf,
                  transformation: { width: img.width, height: img.height },
                }),
              ],
            }),
          ],
        }));
      } else {
        rowCells.push(new docx.TableCell({
          margins: { top: padTwips, bottom: padTwips, left: padTwips, right: padTwips },
          verticalAlign: docx.VerticalAlign.CENTER,
          borders: { top: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
          children: [ new docx.Paragraph({ spacing: { before: 0, after: 0 } }) ],
        }));
      }
    }
    rows.push(new docx.TableRow({ children: rowCells }));
  }
  return new docx.Table({
    width: { size: 0, type: docx.WidthType.AUTO },
    rows,
    alignment: docx.AlignmentType.CENTER,
    layout: docx.TableLayoutType.AUTOFIT,
    cellSpacing: cellSpaceTwips,
    borders: { top: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, insideH: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, insideV: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
  });
}

function buildTablesForLayout(docx, images, layoutRows, cellSpaceTwips) {
  const tables = [];
  let cursor = 0;
  for (const row of layoutRows) {
    const count = row.length;
    const padTwips = ptToTwips(3);
    const rowCells = [];
    for (let i = 0; i < count; i++) {
      const img = images[cursor++];
      if (img) {
        rowCells.push(new docx.TableCell({
          width: { size: Math.floor(100 / count), type: docx.WidthType.PERCENT },
          margins: { top: padTwips, bottom: padTwips, left: padTwips, right: padTwips },
          verticalAlign: docx.VerticalAlign.CENTER,
          borders: { top: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
          children: [
            new docx.Paragraph({
              alignment: docx.AlignmentType.CENTER,
              spacing: { before: 0, after: 0, line: 240 },
              children: [
                new docx.ImageRun({
                  data: img.buf,
                  transformation: { width: img.width, height: img.height },
                }),
              ],
            }),
          ],
        }));
      } else {
        rowCells.push(new docx.TableCell({
          width: { size: Math.floor(100 / count), type: docx.WidthType.PERCENT },
          margins: { top: padTwips, bottom: padTwips, left: padTwips, right: padTwips },
          verticalAlign: docx.VerticalAlign.CENTER,
          borders: { top: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
          children: [ new docx.Paragraph({ spacing: { before: 0, after: 0 } }) ],
        }));
      }
    }
    const table = new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENT },
      rows: [ new docx.TableRow({ children: rowCells }) ],
      alignment: docx.AlignmentType.CENTER,
      layout: docx.TableLayoutType.AUTOFIT,
      cellSpacing: cellSpaceTwips,
      borders: { top: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, insideH: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, insideV: { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
    });
    tables.push(table);
  }
  return tables;
}

function pageSizeCm(orientation) {
  // A4 retrato: 21 x 29.7, paisagem: invertido
  if (orientation === 'paisagem') return { w: 29.7, h: 21.0 };
  return { w: 21.0, h: 29.7 };
}

async function computeAutoLayout(it, params) {
  const OR = params.ORIENTACAO || 'retrato';
  const { w: pageW, h: pageH } = pageSizeCm(OR);
  const marg = params.MARGENS_CM || { sup: 1.27, inf: 1.27, esq: 1.27, dir: 1.27 };
  const usableW = pageW - (marg.esq || 0) - (marg.dir || 0);
  const usableH = pageH - (marg.sup || 0) - (marg.inf || 0);
  const cellSpaceCm = 0;
  const padCm = ptToCm(3);
  const cellW2 = (usableW - cellSpaceCm * 2) / 2;
  const contentW2 = Math.max(0.1, cellW2 - padCm * 2);
  const headerAllowance = 6.0; // cm reservados para títulos/parágrafos (um pouco menor para ampliar imagens)
  const safety = 0.98; // margem de segurança mais suave para ampliar imagens
  const contentW = contentW2;
  const MIN_VERT_H_CM = 6.3;
  const MAX_VERT_H_CM = 6.8;
  const MIN_HORIZ_W_CM = 6.3;
  const MAX_HORIZ_W_CM = 6.85;

  const metas = [];
  for (const img of it.imagens) {
    const bitmap = await fileToImageBitmap(img.file);
    metas.push({ id: img.id, file: img.file, w: bitmap.width, h: bitmap.height, aspect: bitmap.width/bitmap.height });
  }
  const totalRows = Math.max(1, Math.ceil(metas.length / 2));
  const maxH = Math.max(0.1, usableH - headerAllowance);
  const perRowH = Math.max(0.1, ((maxH / totalRows) - padCm * 2 - cellSpaceCm * 2) * safety);
  const layoutRows = [];
  for (let i = 0; i < metas.length; i += 2) {
    const rowMetas = metas.slice(i, i + 2);
    const row = rowMetas.map((m) => {
      const maxHeightByWidth = Math.max(0.1, contentW / Math.max(0.0001, m.aspect));
      const isPortrait = m.aspect < 1;
      let height = Math.max(0.1, Math.min(perRowH, maxHeightByWidth));
      if (isPortrait) {
        height = Math.min(MAX_VERT_H_CM, Math.max(MIN_VERT_H_CM, height));
        const width = Math.max(0.1, Math.min(contentW, height * m.aspect));
        return { id: m.id, file: m.file, widthCm: width, heightCm: height, aspect: m.aspect };
      }
      let width = Math.max(MIN_HORIZ_W_CM, Math.min(Math.max(0.1, Math.min(contentW, MAX_HORIZ_W_CM)), height * m.aspect));
      height = width / Math.max(0.0001, m.aspect);
      return { id: m.id, file: m.file, widthCm: width, heightCm: height, aspect: m.aspect };
    });
    layoutRows.push(row);
  }
  // Expansão: se há espaço sobrando na página, aumenta proporcionalmente
  // respeitando o limite de largura da célula e mantendo proporção.
  (function expandToFill() {
    const maxH = Math.max(0.1, usableH - headerAllowance);
    const totalHeight = () => {
      let sum = 0;
      for (const row of layoutRows) {
        const rowH = Math.max(...row.map(im => im.heightCm || 0)) + padCm * 2 + cellSpaceCm * 2;
        sum += rowH;
      }
      return sum;
    };
    let th = totalHeight();
    const headroom = maxH - th;
    if (headroom <= maxH * 0.05) return; // pouco ganho, ignora
    const target = maxH * 0.98;
    let factor = Math.max(1.0, target / Math.max(0.1, th));
    // limita o fator pelo teto de largura por linha
    const allowedFactors = [];
    for (const row of layoutRows) {
      for (const im of row) {
        if (im.widthCm > 0) allowedFactors.push(contentW / im.widthCm);
      }
    }
    const maxAllowed = Math.max(1.0, Math.min(...allowedFactors.filter(x => Number.isFinite(x) && x > 0)));
    factor = Math.min(factor, maxAllowed);
    if (factor <= 1.001) return;
    for (const row of layoutRows) {
      for (const im of row) {
        const isPortrait = (im.aspect || 1) < 1;
        const aspect = Math.max(0.0001, (im.aspect || 1));
        let newW = Math.min(contentW, im.widthCm * factor);
        if (isPortrait) {
          let newH = newW / aspect;
          if (newH < MIN_VERT_H_CM) {
            newH = MIN_VERT_H_CM;
            newW = Math.min(contentW, newH * aspect);
          }
          if (newH > MAX_VERT_H_CM) {
            newH = MAX_VERT_H_CM;
            newW = Math.min(contentW, newH * aspect);
          }
          im.widthCm = newW;
          im.heightCm = newW / aspect;
        } else {
          const capMaxW = Math.max(0.1, Math.min(contentW, MAX_HORIZ_W_CM));
          newW = Math.max(MIN_HORIZ_W_CM, Math.min(capMaxW, newW));
          im.widthCm = newW;
          im.heightCm = newW / aspect;
        }
      }
    }
  })();
  // Ajuste final: se a soma das alturas das linhas excede a área disponível,
  // escala todas as imagens para caber na mesma página.
  const totalHeight = () => {
    let sum = 0;
    for (const row of layoutRows) {
      const rowH = Math.max(...row.map(im => im.heightCm || 0)) + padCm * 2 + cellSpaceCm * 2;
      sum += rowH;
    }
    return sum;
  };
  for (let iter = 0; iter < 3; iter++) {
    const maxH = Math.max(0.1, usableH - headerAllowance);
    const th = totalHeight();
    if (th <= maxH) break;
    const f = Math.max(0.8, (maxH / th) * 0.98);
    for (const row of layoutRows) {
      for (const im of row) {
        const isPortrait = (im.aspect || 1) < 1;
        const aspect = Math.max(0.0001, (im.aspect || 1));
        let newW = Math.min(contentW, im.widthCm * f);
        if (isPortrait) {
          let newH = newW / aspect;
          if (newH < MIN_VERT_H_CM) {
            newH = MIN_VERT_H_CM;
            newW = Math.min(contentW, newH * aspect);
          }
          if (newH > MAX_VERT_H_CM) {
            newH = MAX_VERT_H_CM;
            newW = Math.min(contentW, newH * aspect);
          }
          im.widthCm = newW;
          im.heightCm = newW / aspect;
        } else {
          const capMaxW = Math.max(0.1, Math.min(contentW, MAX_HORIZ_W_CM));
          newW = Math.max(MIN_HORIZ_W_CM, Math.min(capMaxW, newW));
          im.widthCm = newW;
          im.heightCm = newW / aspect;
        }
      }
    }
  }
  for (const row of layoutRows) {
    for (const im of row) delete im.aspect;
  }
  return { layoutRows, contentW };
}

async function parseDocxAllSections(buf) {
  await ensureJSZip();
  const zip = await window.JSZip.loadAsync(buf);

  // rId -> media filename
  const relsFile = zip.file('word/_rels/document.xml.rels');
  if (!relsFile) return [];
  const relsXml = await relsFile.async('string');
  const rIdToMedia = {};
  for (const m of relsXml.matchAll(/Id="(rId\d+)"[^>]*Target="([^"]+)"/g)) {
    if (m[2].includes('media/')) rIdToMedia[m[1]] = m[2].split('media/').pop();
  }

  const docFile = zip.file('word/document.xml');
  if (!docFile) return [];
  const docXml = await docFile.async('string');

  // All <w:sectPr positions (each marks the end of a section)
  const sectPrPositions = [];
  let sp = 0;
  while (true) {
    const idx = docXml.indexOf('<w:sectPr', sp);
    if (idx === -1) break;
    sectPrPositions.push(idx);
    sp = idx + 1;
  }
  if (sectPrPositions.length === 0) return [];

  // r:embed positions (image references in document order)
  const rIdRefs = [];
  for (const m of docXml.matchAll(/r:embed="(rId\d+)"/g)) {
    rIdRefs.push({ rId: m[1], pos: m.index });
  }

  // Item title positions
  const titleRefs = [];
  for (const m of docXml.matchAll(/Item\s+(\d{1,4})/gi)) {
    titleRefs.push({ num: +m[1], pos: m.index });
  }

  // Group images and titles by section (between consecutive sectPr positions)
  const bodyStart = docXml.indexOf('<w:body>') || 0;
  const sections = [];
  let prevEnd = bodyStart;
  for (const sectPos of sectPrPositions) {
    const sectionRIds = rIdRefs
      .filter(r => r.pos >= prevEnd && r.pos < sectPos)
      .map(r => r.rId);
    const titles = titleRefs
      .filter(t => t.pos >= prevEnd && t.pos < sectPos)
      .sort((a, b) => b.num - a.num);
    const itemNum = titles.length > 0 ? titles[0].num : 0;
    if (sectionRIds.length > 0 && itemNum > 0) {
      sections.push({ itemNum, rIds: sectionRIds });
    }
    prevEnd = sectPos;
  }

  const mimeMap = { jpg:'image/jpeg', jpeg:'image/jpeg', png:'image/png', gif:'image/gif', webp:'image/webp', bmp:'image/bmp' };
  const items = [];
  for (const section of sections) {
    const imagens = [];
    const seenMedia = new Set();
    for (const rId of section.rIds) {
      const filename = rIdToMedia[rId];
      if (!filename || seenMedia.has(filename)) continue;
      seenMedia.add(filename);
      const mediaFile = zip.file(`word/media/${filename}`);
      if (!mediaFile) continue;
      const data  = await mediaFile.async('arraybuffer');
      const ext   = filename.split('.').pop().toLowerCase();
      const mime  = mimeMap[ext] || 'image/jpeg';
      const blob  = new Blob([data], { type: mime });
      const file  = new File([blob], filename, { type: mime });
      const url   = URL.createObjectURL(blob);
      const { w, h } = await new Promise(res => {
        const img = new Image();
        img.onload  = () => res({ w: img.naturalWidth,  h: img.naturalHeight });
        img.onerror = () => res({ w: 800, h: 600 });
        img.src = url;
      });
      imagens.push({ id: crypto.randomUUID(), file, url, w, h });
    }
    if (imagens.length === 0) continue;
    items.push({
      id: crypto.randomUUID(),
      nome: `Item ${String(section.itemNum).padStart(2, '0')}`,
      imported: true,
      config: {
        autoSize: true, perOriEnabled: false,
        larguraValor: 8, larguraUn: 'cm',
        alturaValor: 6,  alturaUn: 'cm',
        bordaCor: '#000000', modoAjuste: 'contain',
        hLarguraValor: 0, hLarguraUn: 'cm',
        hAlturaValor: 0,  hAlturaUn: 'cm',
        vLarguraValor: 0, vLarguraUn: 'cm',
        vAlturaValor: 0,  vAlturaUn: 'cm',
      },
      imagens,
    });
  }
  items.sort((a, b) => {
    const n = s => parseInt(s.nome.replace(/\D/g, '')) || 0;
    return n(a) - n(b);
  });
  return items;
}

async function mergeDocxAppend(origBuf, newBuf) {
  await ensureJSZip();
  const origZip = await window.JSZip.loadAsync(origBuf);
  const newZip  = await window.JSZip.loadAsync(newBuf);

  // 1. Max rId in original relationships
  const origRelsFile = origZip.file('word/_rels/document.xml.rels');
  if (!origRelsFile) throw new Error('Arquivo importado não contém word/_rels/document.xml.rels');
  let origRelsXml = await origRelsFile.async('string');
  const origRIds  = [...origRelsXml.matchAll(/Id="rId(\d+)"/g)].map(m => +m[1]);
  const maxRId    = origRIds.length ? Math.max(...origRIds) : 0;

  // 2. Count original media files (exclude directory entries ending in '/')
  const origMediaKeys = Object.keys(origZip.files).filter(f => f.startsWith('word/media/') && !f.endsWith('/'));
  let imgCounter = origMediaKeys.length;

  // 3. Copy new media files into original zip with renamed paths (exclude directory entries)
  const newMediaKeys = Object.keys(newZip.files).filter(f => f.startsWith('word/media/') && !f.endsWith('/'));
  const mediaRenameMap = {};
  for (const mf of newMediaKeys) {
    const oldName = mf.replace('word/media/', '');
    const ext     = oldName.includes('.') ? oldName.slice(oldName.lastIndexOf('.')) : '';
    imgCounter++;
    const newName = `image${imgCounter}${ext}`;
    mediaRenameMap[oldName] = newName;
    const data = await newZip.file(mf).async('arraybuffer');
    origZip.file(`word/media/${newName}`, data);
  }

  // 4. Update new rels: offset rIds and rename media targets
  const newRelsFile = newZip.file('word/_rels/document.xml.rels');
  if (!newRelsFile) throw new Error('Novo documento não contém word/_rels/document.xml.rels');
  let newRelsXml = await newRelsFile.async('string');
  newRelsXml = newRelsXml.replace(/Id="rId(\d+)"/g, (_, n) => `Id="rId${+n + maxRId}"`);
  for (const [oldName, newName] of Object.entries(mediaRenameMap)) {
    newRelsXml = newRelsXml.split(`media/${oldName}`).join(`media/${newName}`);
  }
  // Handle both self-closing and paired Relationship tags
  const newRelEls = [...newRelsXml.matchAll(/<Relationship\b[^>]*(?:\/>|>[^<]*<\/Relationship>)/g)].map(m => m[0]);
  origRelsXml = origRelsXml.replace('</Relationships>', newRelEls.join('\n') + '\n</Relationships>');
  origZip.file('word/_rels/document.xml.rels', origRelsXml);

  // 5. Update new document.xml: offset all rId references
  const newDocFile = newZip.file('word/document.xml');
  if (!newDocFile) throw new Error('Novo documento não contém word/document.xml');
  let newDocXml = await newDocFile.async('string');
  newDocXml = newDocXml.replace(/r:embed="rId(\d+)"/g, (_, n) => `r:embed="rId${+n + maxRId}"`);
  newDocXml = newDocXml.replace(/r:id="rId(\d+)"/g,    (_, n) => `r:id="rId${+n + maxRId}"`);

  // 6. Extract body content from both files
  const origDocFile = origZip.file('word/document.xml');
  if (!origDocFile) throw new Error('Arquivo importado não contém word/document.xml');
  let origDocXml = await origDocFile.async('string');
  const origInner = (origDocXml.match(/<w:body>([\s\S]*)<\/w:body>/) || ['', ''])[1];
  const newInner  = (newDocXml.match(/<w:body>([\s\S]*)<\/w:body>/)  || ['', ''])[1];

  // 7. Convert trailing body-level <w:sectPr> into a section-break paragraph.
  // Uses lastIndexOf to find only the LAST sectPr — regex with lazy quantifier
  // incorrectly spans from first to last sectPr in multi-section documents.
  let origPrepared = origInner;
  const SECT_OPEN  = '<w:sectPr';
  const SECT_CLOSE = '</w:sectPr>';
  const lastOpen = origInner.lastIndexOf(SECT_OPEN);
  if (lastOpen !== -1) {
    const closeIdx = origInner.indexOf(SECT_CLOSE, lastOpen);
    if (closeIdx !== -1) {
      const lastClose = closeIdx + SECT_CLOSE.length;
      const after = origInner.slice(lastClose).trim();
      if (after === '') {
        // Confirmed body-level sectPr — nothing follows it
        const sectPr = origInner.slice(lastOpen, lastClose);
        origPrepared = origInner.slice(0, lastOpen) + `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`;
      }
    }
  }

  // 8. Merge and write back (use function replacement to avoid $ substitution issues)
  const mergedDocXml = origDocXml.replace(
    /<w:body>[\s\S]*<\/w:body>/,
    () => `<w:body>${origPrepared}${newInner}</w:body>`
  );
  origZip.file('word/document.xml', mergedDocXml);

  return origZip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

async function onGenerate() {
  if (!window.docx) {
    showStatus('Carregando biblioteca DOCX…');
    await new Promise(r => setTimeout(r, 0));
    const ok = await ensureDocx();
    if (!ok || !window.docx) {
      hideStatus();
      alert('Não foi possível carregar a biblioteca DOCX.\n\nVerifique sua conexão com a Internet ou tente novamente. Caso esteja offline, posso embutir a biblioteca localmente se desejar.');
      return;
    }
  }
  const docx = window.docx;
  if (state.items.length === 0) { alert('Adicione pelo menos um Item'); return; }
  showStatus('Gerando documento…');
  updateProgress(0);
  await new Promise(r => setTimeout(r, 0));
  const cellSpaceTwips = 0; // espaçamento externo entre células desativado
  const margens = state.params.MARGENS_CM || { sup: 1.27, inf: 1.27, esq: 1.27, dir: 1.27 };
  const orgInput = $('#orgaoNome');
  if (orgInput) state.params.ORGAO_NOME = (orgInput.value || '').trim();
  const sections = [];
  const summary = [];
  let hadError = false;
  for (let i = 0; i < state.items.length; i++) {
    const it = state.items[i];
    let useImages = it.imagens.slice();
    let layoutRows = undefined;
    if (it.config.autoSize) {
      const res = await computeAutoLayout(it, state.params);
      layoutRows = res.layoutRows;
      const flat = [];
      for (const row of layoutRows) {
        for (const p of row) {
          flat.push({ id: p.id, file: p.file, widthCm: p.widthCm, heightCm: p.heightCm });
        }
      }
      useImages = flat;
    } else {
      if (useImages.length > 6) useImages.length = 6;
    }
    const processed = [];
    for (let k = 0; k < useImages.length; k++) {
      const imgRef = useImages[k];
      const cfg = { ...it.config };
      if (it.config.autoSize) {
        cfg.targetWidthPx = cmToPx(imgRef.widthCm);
        cfg.targetHeightPx = cmToPx(imgRef.heightCm);
        cfg.manterProporcao = true;
        cfg.modoAjuste = 'contain';
      }
      const info = await processImage(imgRef.file, cfg);
      processed.push(info);
      const base = (i / state.items.length) * 98;
      const inc = ((k + 1) / Math.max(1, useImages.length)) * (98 / state.items.length);
      updateProgress(Math.min(98, Math.floor(base + inc)));
    }
    const tables = (layoutRows && it.config.autoSize)
      ? [buildTableForItem(docx, processed, cellSpaceTwips, layoutRows.length)]
      : [buildTableForItem(docx, processed, cellSpaceTwips)];
    const isFirstSection = (i === 0);
    const laudoTitlePara = isFirstSection ? new docx.Paragraph({
      alignment: docx.AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [
        new docx.TextRun({ text: 'LAUDO FOTOGRÁFICO', font: 'Lucida Bright', size: 28 }),
      ],
    }) : null;
    const spacerBetweenTitleAndOrgao = isFirstSection ? new docx.Paragraph({}) : null;
    const orgaoPara = isFirstSection ? new docx.Paragraph({
      alignment: docx.AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [
        new docx.TextRun({ text: state.params.ORGAO_NOME || '', font: 'Lucida Bright', size: 24 }),
      ],
    }) : null;
    const spacerBetweenOrgaoAndTitle = isFirstSection ? new docx.Paragraph({}) : null;
    const spacerBetweenOrgaoAndTitle2 = isFirstSection ? new docx.Paragraph({}) : null;
    const titlePara = new docx.Paragraph({
      alignment: docx.AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: [
        new docx.TextRun({
          text: it.nome,
          font: 'Lucida Bright',
          size: 24,
        }),
      ],
    });
    const spacerAfterTitle1 = new docx.Paragraph({});
    const spacerAfterTitle2 = new docx.Paragraph({});
    sections.push({
      properties: {
        page: {
          margin: {
            top: cmToTwips(margens.sup || 1.27),
            bottom: cmToTwips(margens.inf || 1.27),
            left: cmToTwips(margens.esq || 1.27),
            right: cmToTwips(margens.dir || 1.27),
            footer: cmToTwips(0),
          },
          size: {
            orientation: (state.params.ORIENTACAO === 'paisagem') ? docx.PageOrientation.LANDSCAPE : docx.PageOrientation.PORTRAIT,
          },
        },
      },
      children: isFirstSection
        ? [ laudoTitlePara, spacerBetweenTitleAndOrgao, orgaoPara, spacerBetweenOrgaoAndTitle, spacerBetweenOrgaoAndTitle2, titlePara, spacerAfterTitle1, spacerAfterTitle2, ...tables ].filter(Boolean)
        : [ titlePara, spacerAfterTitle1, spacerAfterTitle2, ...tables ].filter(Boolean),
    });
    summary.push({ nome: it.nome, imagens: processed.length });
  }
  const docxDoc = new docx.Document({ sections });
  const finalBlob = await docx.Packer.toBlob(docxDoc);
  const name = (state.params.NOME_ARQUIVO || 'itens_imagens.docx').replace(/\.docx$/i, '') + '.docx';
  const a = window.document.createElement('a');
  a.href = URL.createObjectURL(finalBlob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1500);
  updateProgress(100);
  hideStatus();
  setTimeout(() => showDownloadToast(), 600);
}

// ─── Importação de .docx existente ───────────────────────────────────────────
async function ensureJSZip() {
  if (window.JSZip) return true;
  const candidates = [
    'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
    'https://unpkg.com/jszip@3.10.1/dist/jszip.min.js',
  ];
  for (const url of candidates) {
    try {
      await new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = url;
        const to = setTimeout(() => { s.remove(); reject(); }, 8000);
        s.onload = () => { clearTimeout(to); resolve(); };
        s.onerror = () => { clearTimeout(to); reject(); };
        document.head.appendChild(s);
      });
      if (window.JSZip) return true;
    } catch {}
  }
  return false;
}

async function parseDocxOrgaoNome(buf) {
  try {
    const zip = await window.JSZip.loadAsync(buf);
    const docXmlFile = zip.file('word/document.xml');
    if (!docXmlFile) return '';
    const xml = await docXmlFile.async('string');
    const text = xml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    const laudoIdx = text.search(/LAUDO\s+FOTOGR[AÁ]FICO/i);
    if (laudoIdx === -1) return '';
    const afterLaudo = text.slice(laudoIdx).replace(/LAUDO\s+FOTOGR[AÁ]FICO\s*/i, '');
    const itemIdx = afterLaudo.search(/\bItem\s+\d{1,4}\b/i);
    if (itemIdx === -1) return '';
    return afterLaudo.slice(0, itemIdx).trim();
  } catch { return ''; }
}

async function parseDocxLastItem(buf) {
  const ok = await ensureJSZip();
  if (!ok) throw new Error('Não foi possível carregar JSZip');
  const zip = await window.JSZip.loadAsync(buf);
  const docXmlFile = zip.file('word/document.xml');
  if (!docXmlFile) throw new Error('Arquivo inválido: word/document.xml não encontrado');
  const xml = await docXmlFile.async('string');
  // Remove tags XML para ter só o texto puro
  const text = xml.replace(/<[^>]+>/g, ' ');
  const matches = [...text.matchAll(/Item\s+(\d{1,4})/gi)];
  const numbers = matches.map(m => parseInt(m[1], 10)).filter(n => !isNaN(n) && n > 0);
  if (numbers.length === 0) return 0;
  return Math.max(...numbers);
}

function setFieldsLocked(locked) {
  const ids = ['orientacaoSelect', 'mSup', 'mInf', 'mEsq', 'mDir'];
  ids.forEach(id => {
    const el = $(`#${id}`);
    if (!el) return;
    el.disabled = locked;
    el.style.opacity = locked ? '0.4' : '';
    el.style.cursor  = locked ? 'not-allowed' : '';
  });
  const hint = $('#lockedHint');
  if (hint) hint.classList.toggle('hidden', !locked);
}

function updateOffsetUI() {
  const offset = state.params.ITEM_OFFSET || 0;
  const btn = $('#importDocxBtn');
  const badge = $('#importOffsetBadge');
  if (!btn) return;
  if (offset > 0) {
    btn.textContent = 'Importar .docx';
    if (badge) {
      badge.textContent = `Continuando a partir do Item ${String(offset + 1).padStart(2, '0')}`;
      badge.classList.remove('hidden');
    }
  } else {
    btn.textContent = 'Importar .docx';
    if (badge) badge.classList.add('hidden');
  }
}

async function resetApp() {
  if (!await showConfirm('Reiniciar o site vai apagar todos os itens e imagens e restaurar as configurações originais.\n\nDeseja continuar?')) return;
  localStorage.removeItem(_LS_KEY);
  try {
    const db = await _openDB();
    await new Promise((resolve) => {
      const tx = db.transaction(_STORE, 'readwrite');
      tx.objectStore(_STORE).clear();
      tx.oncomplete = resolve;
      tx.onerror    = resolve;
    });
  } catch (_) {}
  state.items = [];
  state.params = {
    ORIENTACAO: 'retrato',
    MARGENS_CM: { sup: 1.27, inf: 1.27, esq: 1.27, dir: 1.27 },
    ORGAO_NOME: '',
    NOME_ARQUIVO: 'itens_imagens.docx',
    ITEM_OFFSET: 0,
    PARAMS_LOCKED: false,
  };
  syncParamsToUI();
  setGenerateEnabled(false);
  render();
}

async function onImportDocx(file) {
  try {
    showStatus('Lendo arquivo .docx…');
    const buf = await file.arrayBuffer();
    const lastItem = await parseDocxLastItem(buf);
    hideStatus();
    if (lastItem === 0) {
      alert('Nenhum item encontrado no arquivo.\n\nVerifique se o arquivo foi gerado por este sistema.');
      return;
    }
    const msg = `O arquivo contém ${lastItem} item(s).\n\nEles serão carregados na lista e você poderá reordená-los junto com os novos.\n\nConfirmar?`;
    if (!await showConfirm(msg)) return;
    showStatus('Extraindo itens do arquivo…');
    const [importedItems, orgaoNome] = await Promise.all([
      parseDocxAllSections(buf),
      parseDocxOrgaoNome(buf),
    ]);
    hideStatus();
    if (importedItems.length === 0) {
      alert('Não foi possível extrair os itens do arquivo.');
      return;
    }
    // Replace all items with imported ones; new items added after are numbered from lastItem+1
    state.items = [...importedItems];
    state.params.ITEM_OFFSET = lastItem;
    state.params.PARAMS_LOCKED = true;
    if (orgaoNome) state.params.ORGAO_NOME = orgaoNome;
    state.params.NOME_ARQUIVO = file.name.replace(/\.docx$/i, '') + '.docx';
    updateOffsetUI();
    setFieldsLocked(true);
    syncParamsToUI();
    setGenerateEnabled(true);
    scheduleSave();
    render();
  } catch (e) {
    hideStatus();
    console.error(e);
    alert('Erro ao ler o arquivo: ' + e.message);
  }
}
// ─────────────────────────────────────────────────────────────────────────────

function showConfirm(msg) {
  return new Promise((resolve) => {
    const modal  = $('#confirmModal');
    const msgEl  = $('#confirmModalMsg');
    const okBtn  = $('#confirmModalOk');
    const cancel = $('#confirmModalCancel');
    const logo   = $('#confirmModalLogo');
    const pwLogo = $('#pwLogo');
    if (pwLogo && pwLogo.src) logo.src = pwLogo.src;
    msgEl.textContent = msg;
    modal.classList.remove('hidden');
    function close(result) {
      modal.classList.add('hidden');
      okBtn.removeEventListener('click', onOk);
      cancel.removeEventListener('click', onCancel);
      resolve(result);
    }
    function onOk()     { close(true);  }
    function onCancel() { close(false); }
    okBtn.addEventListener('click', onOk);
    cancel.addEventListener('click', onCancel);
  });
}

async function boot() {
  bindBasics();
  ensureDocx().catch(()=>{});
  _restoring = true;
  const restored = await restoreState();
  _restoring = false;
  if (restored) {
    syncParamsToUI();
    if (state.items.length > 0) setGenerateEnabled(true);
  }
  render();
}

boot();

// (autenticação removida)
