(() => {
  const db = new CloudDB('cloud-word', 'documents');
  const page = document.getElementById('page');
  const docTitle = document.getElementById('docTitle');
  const saveStatus = document.getElementById('saveStatus');
  const docList = document.getElementById('docList');
  const wordCountEl = document.getElementById('wordCount');
  const charCountEl = document.getElementById('charCount');

  let currentId = null;
  let allDocs = [];
  let savedRange = null;

  const FONTS = [
    'Arial', 'Helvetica', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Segoe UI', 'Calibri', 'Century Gothic',
    'Franklin Gothic Medium', 'Lucida Sans Unicode', 'Impact', 'Comic Sans MS',
    'Times New Roman', 'Georgia', 'Garamond', 'Palatino Linotype', 'Book Antiqua', 'Cambria', 'Constantia', 'Rockwell',
    'Courier New', 'Consolas', 'Lucida Console',
  ];

  /* ---------------- helpers ---------------- */
  const setStatus = (text) => { saveStatus.textContent = text; };

  function newDocRecord() {
    return { id: uid(), title: 'Untitled document', html: '<p>Start writing</p>', updatedAt: Date.now() };
  }

  async function refreshDocList() {
    allDocs = (await db.getAll()).sort((a, b) => b.updatedAt - a.updatedAt);
    docList.innerHTML = '';
    allDocs.forEach(d => {
      const item = document.createElement('div');
      item.className = 'doc-item' + (d.id === currentId ? ' active' : '');
      item.innerHTML = `<span class="name"></span><span class="meta"><span class="when"></span><button class="del" title="Delete">Delete</button></span>`;
      item.querySelector('.name').textContent = d.title || 'Untitled document';
      item.querySelector('.when').textContent = formatTime(d.updatedAt);
      item.addEventListener('click', (e) => {
        if (e.target.closest('.del')) return;
        loadDoc(d.id);
      });
      item.querySelector('.del').addEventListener('click', async (e) => {
        e.stopPropagation();
        if (!confirm(`Delete "${d.title}"? This can't be undone.`)) return;
        await db.delete(d.id);
        if (d.id === currentId) {
          const rest = allDocs.filter(x => x.id !== d.id);
          if (rest.length) loadDoc(rest[0].id); else createDoc();
        } else {
          refreshDocList();
        }
      });
      docList.appendChild(item);
    });
  }

  async function loadDoc(id) {
    const rec = await db.get(id);
    if (!rec) return;
    currentId = id;
    docTitle.value = rec.title;
    page.innerHTML = rec.html;
    updateCounts();
    refreshDocList();
    setStatus('Saved');
    history.replaceState(null, '', `word.html?doc=${encodeURIComponent(id)}`);
  }

  async function createDoc() {
    const rec = newDocRecord();
    await db.put(rec);
    await loadDoc(rec.id);
  }

  const saveNow = debounce(async () => {
    if (!currentId) return;
    await db.put({ id: currentId, title: docTitle.value.trim() || 'Untitled document', html: page.innerHTML, updatedAt: Date.now() });
    setStatus('Saved');
    refreshDocList();
  }, 500);

  function markDirty() { setStatus('Saving'); saveNow(); }

  function updateCounts() {
    const text = page.innerText.trim();
    const words = text ? text.split(/\s+/).length : 0;
    wordCountEl.textContent = `${words} word${words === 1 ? '' : 's'}`;
    charCountEl.textContent = `${text.length} characters`;
  }

  /* ---------------- font list ---------------- */
  const fontFamilySelect = document.getElementById('fontFamily');
  FONTS.forEach(f => {
    const opt = document.createElement('option');
    opt.value = /\s/.test(f) ? `'${f}', sans-serif` : `${f}, sans-serif`;
    opt.textContent = f;
    opt.style.fontFamily = opt.value;
    if (f === 'Arial') opt.selected = true;
    fontFamilySelect.appendChild(opt);
  });

  /* ---------------- selection tracking (so toolbar/menu controls that steal
     focus, like <select> dropdowns, still act on the last real selection) ---------------- */
  function captureSelection() {
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      const r = sel.getRangeAt(0);
      if (page.contains(r.commonAncestorContainer)) savedRange = r.cloneRange();
    }
  }
  page.addEventListener('keyup', captureSelection);
  page.addEventListener('mouseup', captureSelection);
  document.addEventListener('selectionchange', () => {
    if (document.activeElement === page) captureSelection();
  });

  function restoreSelection() {
    const sel = window.getSelection();
    const inPage = sel.rangeCount > 0 && page.contains(sel.getRangeAt(0).commonAncestorContainer);
    if (inPage) return sel.getRangeAt(0);
    if (savedRange) {
      sel.removeAllRanges();
      sel.addRange(savedRange);
      return savedRange;
    }
    return null;
  }

  /* ---------------- toolbar commands ---------------- */
  document.querySelectorAll('.tbtn[data-cmd]').forEach(btn => {
    btn.addEventListener('click', () => {
      page.focus();
      document.execCommand(btn.dataset.cmd, false, null);
      markDirty();
      syncToolbarState();
    });
  });

  document.getElementById('blockFormat').addEventListener('change', (e) => {
    page.focus(); restoreSelection();
    document.execCommand('formatBlock', false, e.target.value);
    markDirty();
  });
  document.getElementById('fontFamily').addEventListener('change', (e) => {
    page.focus(); restoreSelection();
    document.execCommand('fontName', false, e.target.value);
    markDirty();
  });
  document.getElementById('fontSize').addEventListener('change', (e) => {
    page.focus(); restoreSelection();
    document.execCommand('fontSize', false, 7);
    page.querySelectorAll('font[size="7"]').forEach(f => {
      f.removeAttribute('size');
      f.style.fontSize = e.target.value + 'pt';
    });
    markDirty();
  });
  document.getElementById('textColor').addEventListener('input', (e) => {
    page.focus(); restoreSelection();
    document.execCommand('foreColor', false, e.target.value); markDirty();
  });
  document.getElementById('hiliteColor').addEventListener('input', (e) => {
    page.focus(); restoreSelection();
    document.execCommand('hiliteColor', false, e.target.value); markDirty();
  });

  function insertLink() {
    const url = prompt('Link URL:', 'https://');
    if (url) { page.focus(); restoreSelection(); document.execCommand('createLink', false, url); markDirty(); }
  }
  document.getElementById('btnLink').addEventListener('click', insertLink);

  const imageInput = document.getElementById('imageInput');
  function triggerInsertImage() { imageInput.click(); }
  document.getElementById('btnImage').addEventListener('click', triggerInsertImage);
  imageInput.addEventListener('change', () => {
    const file = imageInput.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      page.focus(); restoreSelection();
      document.execCommand('insertImage', false, reader.result);
      markDirty();
    };
    reader.readAsDataURL(file);
    imageInput.value = '';
  });

  function insertTable() {
    const rows = parseInt(prompt('Rows:', '3'), 10) || 3;
    const cols = parseInt(prompt('Columns:', '3'), 10) || 3;
    let html = '<table>';
    for (let r = 0; r < rows; r++) html += '<tr>' + '<td>&nbsp;</td>'.repeat(cols) + '</tr>';
    html += '</table><p></p>';
    page.focus(); restoreSelection();
    document.execCommand('insertHTML', false, html);
    markDirty();
  }
  document.getElementById('btnTable').addEventListener('click', insertTable);

  function insertPageBreak() {
    page.focus(); restoreSelection();
    document.execCommand('insertHTML', false, '<div class="page-break" contenteditable="false"></div><p></p>');
    markDirty();
  }
  document.getElementById('btnPageBreak').addEventListener('click', insertPageBreak);

  function applyLineSpacing(value) {
    page.focus();
    const range = restoreSelection();
    if (!range) return;
    const blocks = blocksInRange(range);
    blocks.forEach(b => { b.style.lineHeight = value; });
    markDirty();
  }
  document.getElementById('lineSpacing').addEventListener('mousedown', captureSelection);
  document.getElementById('lineSpacing').addEventListener('change', (e) => applyLineSpacing(e.target.value));

  function blocksInRange(range) {
    const selector = 'p,li,h1,h2,h3,blockquote,pre,div';
    const all = page.querySelectorAll(selector);
    const hit = [];
    all.forEach(el => { if (range.intersectsNode(el)) hit.push(el); });
    if (hit.length === 0) {
      let node = range.startContainer;
      while (node && node !== page && !(node.nodeType === 1 && node.matches(selector))) node = node.parentNode;
      if (node && node !== page) hit.push(node);
    }
    return hit;
  }

  /* ---------------- table row/column editing ---------------- */
  function currentCell() {
    const sel = window.getSelection();
    let node = sel && sel.rangeCount ? sel.getRangeAt(0).startContainer : (savedRange ? savedRange.startContainer : null);
    while (node && node !== page) {
      if (node.nodeType === 1 && (node.tagName === 'TD' || node.tagName === 'TH')) return node;
      node = node.parentNode;
    }
    return null;
  }
  function tableRowBelow() {
    const cell = currentCell();
    if (!cell) { alert('Click inside a table cell first.'); return; }
    const tr = cell.closest('tr');
    const newRow = tr.cloneNode(true);
    newRow.querySelectorAll('td,th').forEach(td => { td.innerHTML = '&nbsp;'; });
    tr.after(newRow);
    markDirty();
  }
  function tableRowDel() {
    const cell = currentCell();
    if (!cell) { alert('Click inside a table cell first.'); return; }
    const tr = cell.closest('tr');
    const table = cell.closest('table');
    if (table.querySelectorAll('tr').length <= 1) table.remove(); else tr.remove();
    markDirty();
  }
  function tableColRight() {
    const cell = currentCell();
    if (!cell) { alert('Click inside a table cell first.'); return; }
    const idx = Array.prototype.indexOf.call(cell.parentNode.children, cell);
    const table = cell.closest('table');
    table.querySelectorAll('tr').forEach(tr => {
      const ref = tr.children[idx];
      const td = document.createElement('td');
      td.innerHTML = '&nbsp;';
      if (ref) ref.after(td); else tr.appendChild(td);
    });
    markDirty();
  }
  function tableColDel() {
    const cell = currentCell();
    if (!cell) { alert('Click inside a table cell first.'); return; }
    const idx = Array.prototype.indexOf.call(cell.parentNode.children, cell);
    const table = cell.closest('table');
    const rows = table.querySelectorAll('tr');
    if (rows[0].children.length <= 1) table.remove();
    else rows.forEach(tr => { if (tr.children[idx]) tr.children[idx].remove(); });
    markDirty();
  }
  document.getElementById('btnRowBelow').addEventListener('click', tableRowBelow);
  document.getElementById('btnRowDel').addEventListener('click', tableRowDel);
  document.getElementById('btnColRight').addEventListener('click', tableColRight);
  document.getElementById('btnColDel').addEventListener('click', tableColDel);

  /* ---------------- zoom ---------------- */
  function setZoom(value) {
    page.style.zoom = value;
    document.getElementById('zoomSelect').value = value;
    document.querySelectorAll('[data-zoom]').forEach(b => b.classList.toggle('checked', b.dataset.zoom === String(value)));
  }
  document.getElementById('zoomSelect').addEventListener('change', (e) => setZoom(e.target.value));
  document.querySelectorAll('[data-zoom]').forEach(btn => {
    btn.addEventListener('click', () => setZoom(btn.dataset.zoom));
  });

  function syncToolbarState() {
    document.querySelectorAll('.tbtn[data-cmd]').forEach(btn => {
      try { btn.classList.toggle('active', document.queryCommandState(btn.dataset.cmd)); } catch (e) {}
    });
  }

  page.addEventListener('input', () => { markDirty(); updateCounts(); });
  page.addEventListener('keyup', syncToolbarState);
  page.addEventListener('mouseup', syncToolbarState);
  docTitle.addEventListener('input', markDirty);

  /* ---------------- sidebar ---------------- */
  document.getElementById('btnNewDoc').addEventListener('click', createDoc);

  /* ---------------- print ---------------- */
  function doPrint() { window.print(); }
  document.getElementById('btnPrint').addEventListener('click', doPrint);

  /* ---------------- export ---------------- */
  function download(blob, filename) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }
  function exportAs(type) {
    const title = (docTitle.value.trim() || 'document').replace(/[^\w\- ]+/g, '');
    if (type === 'docx') {
      const fullHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${page.innerHTML}</body></html>`;
      download(window.htmlDocx.asBlob(fullHtml), `${title}.docx`);
    } else if (type === 'html') {
      const fullHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title></head><body>${page.innerHTML}</body></html>`;
      download(new Blob([fullHtml], { type: 'text/html' }), `${title}.html`);
    } else if (type === 'txt') {
      download(new Blob([page.innerText], { type: 'text/plain' }), `${title}.txt`);
    }
  }
  const exportMenuBtn = document.getElementById('btnExportMenu');
  const exportMenu = document.getElementById('exportMenu');
  exportMenuBtn.addEventListener('click', () => exportMenu.classList.toggle('open'));
  document.addEventListener('click', (e) => { if (!e.target.closest('.menu-wrap')) exportMenu.classList.remove('open'); });
  exportMenu.addEventListener('click', (e) => {
    const type = e.target.dataset.export;
    if (!type) return;
    exportAs(type);
    exportMenu.classList.remove('open');
  });

  /* ---------------- import ---------------- */
  function escapeHtml(s) { return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  async function createDocFrom(title, html) {
    const rec = { id: uid(), title: title || 'Imported document', html, updatedAt: Date.now() };
    await db.put(rec);
    await loadDoc(rec.id);
  }
  const fileOpen = document.getElementById('fileOpen');
  function triggerOpen() { fileOpen.click(); }
  fileOpen.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'docx') {
      const buf = await file.arrayBuffer();
      const result = await window.mammoth.convertToHtml({ arrayBuffer: buf });
      await createDocFrom(file.name.replace(/\.docx$/i, ''), result.value);
    } else if (ext === 'txt') {
      const text = await file.text();
      const html = text.split(/\n{2,}/).map(p => `<p>${escapeHtml(p).replace(/\n/g, '<br>')}</p>`).join('');
      await createDocFrom(file.name.replace(/\.txt$/i, ''), html || '<p></p>');
    } else if (ext === 'html') {
      const text = await file.text();
      await createDocFrom(file.name.replace(/\.html?$/i, ''), text);
    } else {
      alert('Supported formats: .docx, .txt, .html');
    }
    e.target.value = '';
  });

  /* ---------------- find & replace ---------------- */
  const findBar = document.getElementById('findBar');
  const findInput = document.getElementById('findInput');
  const replaceInput = document.getElementById('replaceInput');
  const findCount = document.getElementById('findCount');
  let findMatches = [];
  let findIdx = -1;

  function getTextNodes() {
    const walker = document.createTreeWalker(page, NodeFilter.SHOW_TEXT);
    const nodes = [];
    let n;
    while ((n = walker.nextNode())) nodes.push(n);
    return nodes;
  }
  function findAllMatches(query) {
    if (!query) return [];
    const q = query.toLowerCase();
    const matches = [];
    getTextNodes().forEach(node => {
      const text = node.textContent.toLowerCase();
      let idx = 0;
      while ((idx = text.indexOf(q, idx)) !== -1) {
        matches.push({ node, start: idx, end: idx + q.length });
        idx += q.length;
      }
    });
    return matches;
  }
  function highlightMatch(m) {
    const range = document.createRange();
    range.setStart(m.node, m.start);
    range.setEnd(m.node, m.end);
    const s = window.getSelection();
    s.removeAllRanges();
    s.addRange(range);
    const el = m.node.parentElement;
    if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center' });
  }
  function replaceMatch(m, replacement) {
    const range = document.createRange();
    range.setStart(m.node, m.start);
    range.setEnd(m.node, m.end);
    range.deleteContents();
    range.insertNode(document.createTextNode(replacement));
  }
  function openFindBar() {
    findBar.hidden = false;
    findInput.focus(); findInput.select();
    runFind();
  }
  function closeFindBar() {
    findBar.hidden = true;
    findMatches = []; findIdx = -1; findCount.textContent = '';
    page.focus();
  }
  function runFind() {
    findMatches = findAllMatches(findInput.value);
    findIdx = findMatches.length ? 0 : -1;
    updateFindStatus();
    if (findIdx >= 0) highlightMatch(findMatches[findIdx]);
  }
  function updateFindStatus() {
    findCount.textContent = findMatches.length ? `${findIdx + 1} of ${findMatches.length}` : (findInput.value ? 'No matches' : '');
  }
  document.getElementById('btnFind').addEventListener('click', openFindBar);
  document.getElementById('btnFindClose').addEventListener('click', closeFindBar);
  findInput.addEventListener('input', runFind);
  document.getElementById('btnFindNext').addEventListener('click', () => {
    if (!findMatches.length) return;
    findIdx = (findIdx + 1) % findMatches.length;
    updateFindStatus(); highlightMatch(findMatches[findIdx]);
  });
  document.getElementById('btnReplace').addEventListener('click', () => {
    if (findIdx < 0) return;
    replaceMatch(findMatches[findIdx], replaceInput.value);
    markDirty(); updateCounts();
    runFind();
  });
  document.getElementById('btnReplaceAll').addEventListener('click', () => {
    if (!findMatches.length) return;
    findMatches.slice().reverse().forEach(m => replaceMatch(m, replaceInput.value));
    markDirty(); updateCounts();
    runFind();
  });
  findInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); document.getElementById('btnFindNext').click(); }
    else if (e.key === 'Escape') { e.preventDefault(); closeFindBar(); }
  });

  document.addEventListener('keydown', (e) => {
    const mod = e.ctrlKey || e.metaKey;
    if (mod && e.key.toLowerCase() === 's') { e.preventDefault(); markDirty(); }
    if (mod && e.key.toLowerCase() === 'f') { e.preventDefault(); openFindBar(); }
  });

  /* ---------------- sidebar toggle ---------------- */
  const sidebar = document.getElementById('sidebar');
  function toggleSidebar() { sidebar.classList.toggle('collapsed'); }

  /* ---------------- menu bar wiring ---------------- */
  CloudCommon.bindMenuAction('mNew', createDoc);
  CloudCommon.bindMenuAction('mOpenTrigger', triggerOpen);
  CloudCommon.bindMenuAction('mExportDocx', () => exportAs('docx'));
  CloudCommon.bindMenuAction('mExportHtml', () => exportAs('html'));
  CloudCommon.bindMenuAction('mExportTxt', () => exportAs('txt'));
  CloudCommon.bindMenuAction('mPrint', doPrint);
  CloudCommon.bindMenuAction('mDelete', async () => {
    if (!currentId) return;
    if (!confirm('Delete this document? This can\'t be undone.')) return;
    await db.delete(currentId);
    const rest = (await db.getAll()).sort((a, b) => b.updatedAt - a.updatedAt);
    if (rest.length) loadDoc(rest[0].id); else createDoc();
  });
  CloudCommon.bindMenuAction('mClose', () => { location.href = 'index.html'; });
  CloudCommon.bindMenuAction('mUndo', () => { page.focus(); document.execCommand('undo'); markDirty(); });
  CloudCommon.bindMenuAction('mRedo', () => { page.focus(); document.execCommand('redo'); markDirty(); });
  CloudCommon.bindMenuAction('mFind', openFindBar);
  CloudCommon.bindMenuAction('mSelectAll', () => { page.focus(); document.execCommand('selectAll'); });
  CloudCommon.bindMenuAction('mInsertImage', triggerInsertImage);
  CloudCommon.bindMenuAction('mInsertTable', insertTable);
  CloudCommon.bindMenuAction('mInsertLink', insertLink);
  CloudCommon.bindMenuAction('mInsertLine', () => { page.focus(); restoreSelection(); document.execCommand('insertHorizontalRule'); markDirty(); });
  CloudCommon.bindMenuAction('mInsertBreak', insertPageBreak);
  CloudCommon.bindMenuAction('mBold', () => { page.focus(); restoreSelection(); document.execCommand('bold'); markDirty(); });
  CloudCommon.bindMenuAction('mItalic', () => { page.focus(); restoreSelection(); document.execCommand('italic'); markDirty(); });
  CloudCommon.bindMenuAction('mUnderline', () => { page.focus(); restoreSelection(); document.execCommand('underline'); markDirty(); });
  CloudCommon.bindMenuAction('mLs1', () => applyLineSpacing('1'));
  CloudCommon.bindMenuAction('mLs115', () => applyLineSpacing('1.15'));
  CloudCommon.bindMenuAction('mLs15', () => applyLineSpacing('1.5'));
  CloudCommon.bindMenuAction('mLs2', () => applyLineSpacing('2'));
  CloudCommon.bindMenuAction('mClearFormat', () => { page.focus(); restoreSelection(); document.execCommand('removeFormat'); markDirty(); });
  CloudCommon.bindMenuAction('mRowBelow', tableRowBelow);
  CloudCommon.bindMenuAction('mRowDel', tableRowDel);
  CloudCommon.bindMenuAction('mColRight', tableColRight);
  CloudCommon.bindMenuAction('mColDel', tableColDel);
  CloudCommon.bindMenuAction('mToggleSidebar', toggleSidebar);
  CloudCommon.bindMenuAction('mHelpCenter', () => { location.href = 'help.html#word'; });
  CloudCommon.bindMenuAction('mAbout', () => { location.href = 'about.html'; });
  document.getElementById('btnHelp').addEventListener('click', () => { location.href = 'help.html#word'; });
  document.getElementById('btnSettings').addEventListener('click', () => CloudCommon.openModal('settingsModal'));

  /* ---------------- boot ---------------- */
  (async () => {
    CloudCommon.initTheme();
    CloudCommon.initMenuBar();
    CloudCommon.initModalDismiss();

    const params = new URLSearchParams(location.search);
    const docId = params.get('doc');
    const wantNew = params.has('new');

    allDocs = await db.getAll();
    if (docId && allDocs.some(d => d.id === docId)) {
      await loadDoc(docId);
    } else if (wantNew || allDocs.length === 0) {
      await createDoc();
    } else {
      allDocs.sort((a, b) => b.updatedAt - a.updatedAt);
      await loadDoc(allDocs[0].id);
    }
  })();

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
  }
})();
