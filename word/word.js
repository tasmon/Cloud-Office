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

  /* ---------------- helpers ---------------- */
  const setStatus = (text) => { saveStatus.textContent = text; };

  function newDocRecord() {
    return { id: uid(), title: 'Untitled document', html: '<p>Start writing…</p>', updatedAt: Date.now() };
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

  function markDirty() { setStatus('Saving…'); saveNow(); }

  function updateCounts() {
    const text = page.innerText.trim();
    const words = text ? text.split(/\s+/).length : 0;
    wordCountEl.textContent = `${words} word${words === 1 ? '' : 's'}`;
    charCountEl.textContent = `${text.length} characters`;
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
    page.focus();
    document.execCommand('formatBlock', false, e.target.value);
    markDirty();
  });
  document.getElementById('fontFamily').addEventListener('change', (e) => {
    page.focus();
    document.execCommand('fontName', false, e.target.value);
    markDirty();
  });
  document.getElementById('fontSize').addEventListener('change', (e) => {
    page.focus();
    // execCommand fontSize only takes 1-7, so wrap selection manually for pt sizing
    document.execCommand('fontSize', false, 7);
    page.querySelectorAll('font[size="7"]').forEach(f => {
      f.removeAttribute('size');
      f.style.fontSize = e.target.value + 'pt';
    });
    markDirty();
  });
  document.getElementById('textColor').addEventListener('input', (e) => {
    page.focus(); document.execCommand('foreColor', false, e.target.value); markDirty();
  });
  document.getElementById('hiliteColor').addEventListener('input', (e) => {
    page.focus(); document.execCommand('hiliteColor', false, e.target.value); markDirty();
  });

  document.getElementById('btnLink').addEventListener('click', () => {
    const url = prompt('Link URL:', 'https://');
    if (url) { page.focus(); document.execCommand('createLink', false, url); markDirty(); }
  });

  const imageInput = document.getElementById('imageInput');
  document.getElementById('btnImage').addEventListener('click', () => imageInput.click());
  imageInput.addEventListener('change', () => {
    const file = imageInput.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      page.focus();
      document.execCommand('insertImage', false, reader.result);
      markDirty();
    };
    reader.readAsDataURL(file);
    imageInput.value = '';
  });

  document.getElementById('btnTable').addEventListener('click', () => {
    const rows = parseInt(prompt('Rows:', '3'), 10) || 3;
    const cols = parseInt(prompt('Columns:', '3'), 10) || 3;
    let html = '<table>';
    for (let r = 0; r < rows; r++) {
      html += '<tr>' + '<td>&nbsp;</td>'.repeat(cols) + '</tr>';
    }
    html += '</table><p></p>';
    page.focus();
    document.execCommand('insertHTML', false, html);
    markDirty();
  });

  function syncToolbarState() {
    document.querySelectorAll('.tbtn[data-cmd]').forEach(btn => {
      try {
        btn.classList.toggle('active', document.queryCommandState(btn.dataset.cmd));
      } catch (e) {}
    });
  }

  page.addEventListener('input', () => { markDirty(); updateCounts(); });
  page.addEventListener('keyup', syncToolbarState);
  page.addEventListener('mouseup', syncToolbarState);
  docTitle.addEventListener('input', markDirty);

  document.addEventListener('keydown', (e) => {
    const mod = e.ctrlKey || e.metaKey;
    if (mod && e.key.toLowerCase() === 's') { e.preventDefault(); markDirty(); }
  });

  /* ---------------- sidebar ---------------- */
  document.getElementById('toggleSidebar').addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('collapsed');
  });
  document.getElementById('btnNewDoc').addEventListener('click', createDoc);

  /* ---------------- print ---------------- */
  document.getElementById('btnPrint').addEventListener('click', () => window.print());

  /* ---------------- export ---------------- */
  const exportMenuBtn = document.getElementById('btnExportMenu');
  const exportMenu = document.getElementById('exportMenu');
  exportMenuBtn.addEventListener('click', () => exportMenu.classList.toggle('open'));
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.menu-wrap')) exportMenu.classList.remove('open');
  });

  function download(blob, filename) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  exportMenu.addEventListener('click', (e) => {
    const type = e.target.dataset.export;
    if (!type) return;
    const title = (docTitle.value.trim() || 'document').replace(/[^\w\- ]+/g, '');
    if (type === 'docx') {
      const fullHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${page.innerHTML}</body></html>`;
      const blob = window.htmlDocx.asBlob(fullHtml);
      download(blob, `${title}.docx`);
    } else if (type === 'html') {
      const fullHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title></head><body>${page.innerHTML}</body></html>`;
      download(new Blob([fullHtml], { type: 'text/html' }), `${title}.html`);
    } else if (type === 'txt') {
      download(new Blob([page.innerText], { type: 'text/plain' }), `${title}.txt`);
    }
    exportMenu.classList.remove('open');
  });

  /* ---------------- import ---------------- */
  document.getElementById('fileOpen').addEventListener('change', async (e) => {
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

  function escapeHtml(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  async function createDocFrom(title, html) {
    const rec = { id: uid(), title: title || 'Imported document', html, updatedAt: Date.now() };
    await db.put(rec);
    await loadDoc(rec.id);
  }

  /* ---------------- boot ---------------- */
  (async () => {
    allDocs = await db.getAll();
    if (allDocs.length === 0) {
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
