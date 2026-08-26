(() => {
  const db = new CloudDB('cloud-slides', 'decks');
  const slideEl = document.getElementById('slide');
  const slideList = document.getElementById('slideList');
  const deckTitle = document.getElementById('deckTitle');
  const saveStatus = document.getElementById('saveStatus');
  const textTools = document.getElementById('textTools');
  const slideCountLabel = document.getElementById('slideCountLabel');
  const notesInput = document.getElementById('notesInput');

  const SLIDE_W_IN = 13.333, SLIDE_H_IN = 7.5; // widescreen inches for pptx export

  let deck = null;      // { id, title, slides:[{id,bg,notes,elements:[...]}], active, updatedAt }
  let selectedIds = []; // ordered; last = primary (drives the text-formatting toolbar)
  let dragState = null;
  const undoMgr = new UndoManager(80);
  function snapshotUndo() { undoMgr.snapshot(deck); }

  const newEl = (type, extra) => Object.assign({
    id: uid(), type, x: 10, y: 10, w: 30, h: 12, rotation: 0,
  }, extra);

  const newSlide = (elements = [], bg = '#ffffff') => ({ id: uid(), bg, notes: '', elements });

  function newDeck() {
    return {
      id: uid(), title: 'Untitled presentation',
      slides: [ newSlide([
        newEl('text', { x: 8, y: 34, w: 84, h: 18, text: 'Click to add title', fontSize: 40, bold: true, align: 'center', color: '#14161f' }),
        newEl('text', { x: 8, y: 54, w: 84, h: 10, text: 'Click to add subtitle', fontSize: 18, align: 'center', color: '#5b5d6b' }),
      ]) ],
      active: 0, updatedAt: Date.now(),
    };
  }

  function activeSlide() { return deck.slides[deck.active]; }
  function findEl(id) { return activeSlide().elements.find(e => e.id === id); }
  function isSelected(id) { return selectedIds.includes(id); }
  function primarySelected() { return selectedIds.length ? findEl(selectedIds[selectedIds.length - 1]) : null; }
  function selectedEls() { return selectedIds.map(findEl).filter(Boolean); }

  /* ---------------- rendering ---------------- */
  function renderElementsInto(container, slideData, opts = {}) {
    container.innerHTML = '';
    if (!container.style.position) container.style.position = 'relative';
    if (!container.style.overflow) container.style.overflow = 'hidden';

    const base = document.createElement('div');
    base.style.position = 'absolute';
    base.style.left = '0'; base.style.top = '0';
    base.style.width = '960px'; base.style.height = '540px';
    base.style.transformOrigin = 'top left';
    base.style.background = slideData.bg || '#ffffff';
    container.appendChild(base);

    slideData.elements.forEach(el => {
      const div = document.createElement('div');
      div.className = 'el ' + el.type + (el.type === 'shape' ? ' ' + el.shape : '');
      div.style.left = el.x + '%';
      div.style.top = el.y + '%';
      div.style.width = el.w + '%';
      div.style.height = el.h + '%';
      if (el.type === 'text') {
        div.style.alignItems = el.valign === 'top' ? 'flex-start' : 'center';
        const inner = document.createElement('div');
        inner.className = 'inner';
        inner.textContent = el.text || '';
        inner.style.fontSize = (el.fontSize || 18) + 'px';
        inner.style.fontFamily = el.fontFamily || 'Arial, sans-serif';
        inner.style.fontWeight = el.bold ? '700' : '400';
        inner.style.fontStyle = el.italic ? 'italic' : 'normal';
        inner.style.textAlign = el.align || 'left';
        inner.style.color = el.color || '#14161f';
        inner.style.background = el.fill || 'transparent';
        inner.style.whiteSpace = 'pre-wrap';
        inner.style.wordBreak = 'break-word';
        div.appendChild(inner);
      } else if (el.type === 'shape') {
        div.style.background = el.fill || '#d65d3a';
      } else if (el.type === 'image') {
        const img = document.createElement('img');
        img.src = el.src;
        div.appendChild(img);
      } else if (el.type === 'line') {
        const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        svg.setAttribute('viewBox', '0 0 100 100');
        svg.setAttribute('preserveAspectRatio', 'none');
        svg.style.width = '100%'; svg.style.height = '100%'; svg.style.overflow = 'visible';
        const color = el.color || '#14161f';
        if (el.arrow) {
          const markerId = 'arrow-' + el.id;
          svg.innerHTML = `<defs><marker id="${markerId}" markerWidth="8" markerHeight="8" refX="6" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="${color}"/></marker></defs>
            <line x1="0" y1="0" x2="100" y2="100" stroke="${color}" stroke-width="3" vector-effect="non-scaling-stroke" marker-end="url(#${markerId})"/>`;
        } else {
          svg.innerHTML = `<line x1="0" y1="0" x2="100" y2="100" stroke="${color}" stroke-width="3" vector-effect="non-scaling-stroke"/>`;
        }
        div.appendChild(svg);
      }
      if (!opts.mini) {
        div.dataset.id = el.id;
        if (isSelected(el.id)) div.classList.add('selected');
        div.addEventListener('mousedown', (e) => onElMouseDown(e, el));
        if (el.type === 'text') div.addEventListener('dblclick', () => editText(div, el));
        if (selectedIds.length === 1 && selectedIds[0] === el.id) {
          const handle = document.createElement('div');
          handle.className = 'handle';
          handle.addEventListener('mousedown', (e) => onResizeMouseDown(e, el));
          div.appendChild(handle);
        }
      } else {
        div.style.pointerEvents = 'none';
      }
      base.appendChild(div);
    });

    const w = container.clientWidth;
    const scale = w ? w / 960 : 1;
    base.style.transform = `scale(${scale})`;
    return base;
  }

  function renderSlide() { renderElementsInto(slideEl, activeSlide()); syncTextTools(); }

  function renderPanel() {
    slideList.innerHTML = '';
    deck.slides.forEach((s, idx) => {
      const thumb = document.createElement('div');
      thumb.className = 'slide-thumb' + (idx === deck.active ? ' active' : '');
      const num = document.createElement('span'); num.className = 'num'; num.textContent = idx + 1;
      const mini = document.createElement('div'); mini.className = 'mini';
      mini.style.position = 'absolute'; mini.style.inset = '0';
      thumb.appendChild(mini); thumb.appendChild(num);
      if (deck.slides.length > 1) {
        const del = document.createElement('button');
        del.className = 'del'; del.textContent = '✕';
        del.addEventListener('click', (e) => { e.stopPropagation(); deleteSlide(idx); });
        thumb.appendChild(del);
      }
      thumb.addEventListener('click', () => { deck.active = idx; selectedIds = []; renderAll(); });
      slideList.appendChild(thumb);
      renderElementsInto(mini, s, { mini: true });
    });
    slideCountLabel.textContent = `${deck.active + 1} / ${deck.slides.length}`;
  }

  function renderAll() {
    renderSlide(); renderPanel();
    document.getElementById('bgColor').value = rgbToHex(activeSlide().bg);
    notesInput.value = activeSlide().notes || '';
  }

  function rgbToHex(v) { return /^#/.test(v) ? v : '#ffffff'; }

  /* ---------------- selection & drag ---------------- */
  function selectEl(id, additive) {
    if (additive) {
      selectedIds = isSelected(id) ? selectedIds.filter(x => x !== id) : selectedIds.concat(id);
    } else {
      selectedIds = [id];
    }
    renderSlide();
  }
  function clearSelection() { selectedIds = []; renderSlide(); }

  slideEl.parentElement.addEventListener('mousedown', (e) => {
    if (e.target === slideEl) clearSelection();
  });

  function onElMouseDown(e, el) {
    if (e.target.classList.contains('handle')) return;
    e.stopPropagation();
    if (e.shiftKey) {
      selectEl(el.id, true);
    } else if (!isSelected(el.id)) {
      selectEl(el.id, false);
    }
    // else: already part of a multi-selection and shift isn't held. Keep the group selected, drag them all.
    snapshotUndo();
    const rect = slideEl.getBoundingClientRect();
    const origins = {};
    selectedEls().forEach(se => { origins[se.id] = { x: se.x, y: se.y }; });
    dragState = { mode: 'move', el, startX: e.clientX, startY: e.clientY, origins, rect };
    window.addEventListener('mousemove', onDragMove);
    window.addEventListener('mouseup', onDragEnd);
  }
  function onResizeMouseDown(e, el) {
    e.stopPropagation(); e.preventDefault();
    snapshotUndo();
    const rect = slideEl.getBoundingClientRect();
    dragState = { mode: 'resize', el, startX: e.clientX, startY: e.clientY, origW: el.w, origH: el.h, rect };
    window.addEventListener('mousemove', onDragMove);
    window.addEventListener('mouseup', onDragEnd);
  }
  function onDragMove(e) {
    if (!dragState) return;
    const { el, rect } = dragState;
    const dxPct = ((e.clientX - dragState.startX) / rect.width) * 100;
    const dyPct = ((e.clientY - dragState.startY) / rect.height) * 100;
    if (dragState.mode === 'move') {
      Object.keys(dragState.origins).forEach(id => {
        const se = findEl(id);
        if (!se) return;
        const o = dragState.origins[id];
        se.x = clamp(o.x + dxPct, -5, 100 - 2);
        se.y = clamp(o.y + dyPct, -5, 100 - 2);
      });
    } else {
      el.w = clamp(dragState.origW + dxPct, 3, 100);
      el.h = clamp(dragState.origH + dyPct, 3, 100);
    }
    renderSlide();
  }
  function onDragEnd() {
    window.removeEventListener('mousemove', onDragMove);
    window.removeEventListener('mouseup', onDragEnd);
    if (dragState) { markDirty(); renderPanel(); }
    dragState = null;
  }
  const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));

  function editText(div, el) {
    selectedIds = [el.id];
    const inner = div.querySelector('.inner');
    inner.contentEditable = 'true';
    inner.focus();
    document.execCommand('selectAll', false, null);
    inner.addEventListener('blur', function onBlur() {
      snapshotUndo();
      el.text = inner.textContent;
      inner.contentEditable = 'false';
      inner.removeEventListener('blur', onBlur);
      markDirty(); renderPanel();
    }, { once: true });
  }

  document.addEventListener('keydown', (e) => {
    const mod = e.ctrlKey || e.metaKey;
    if (mod && e.key.toLowerCase() === 'z' && !overlayOpen()) { e.preventDefault(); doUndo(); return; }
    if (mod && (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z')) && !overlayOpen()) { e.preventDefault(); doRedo(); return; }
    if ((e.key === 'Delete' || e.key === 'Backspace') && selectedIds.length && !isEditingText()) {
      e.preventDefault();
      snapshotUndo();
      activeSlide().elements = activeSlide().elements.filter(el => !isSelected(el.id));
      selectedIds = [];
      markDirty(); renderAll();
    }
  });
  function overlayOpen() { return document.getElementById('presentOverlay').classList.contains('open'); }
  function isEditingText() {
    const active = document.activeElement;
    return active && active.classList && active.classList.contains('inner') && active.isContentEditable;
  }

  /* ---------------- toolbar: insert ---------------- */
  document.getElementById('btnAddText').addEventListener('click', () => {
    snapshotUndo();
    const el = newEl('text', { text: 'New text', fontSize: 24, align: 'left', color: '#14161f' });
    activeSlide().elements.push(el);
    selectEl(el.id); markDirty(); renderPanel();
  });
  document.getElementById('btnAddRect').addEventListener('click', () => {
    snapshotUndo();
    const el = newEl('shape', { shape: 'rect', fill: document.getElementById('fillColor').value });
    activeSlide().elements.push(el); selectEl(el.id); markDirty(); renderPanel();
  });
  document.getElementById('btnAddEllipse').addEventListener('click', () => {
    snapshotUndo();
    const el = newEl('shape', { shape: 'ellipse', fill: document.getElementById('fillColor').value });
    activeSlide().elements.push(el); selectEl(el.id); markDirty(); renderPanel();
  });
  document.getElementById('btnAddLine').addEventListener('click', () => {
    snapshotUndo();
    const el = newEl('line', { w: 30, h: 0.5, arrow: false, color: document.getElementById('fillColor').value });
    activeSlide().elements.push(el); selectEl(el.id); markDirty(); renderPanel();
  });
  document.getElementById('btnAddArrow').addEventListener('click', () => {
    snapshotUndo();
    const el = newEl('line', { w: 30, h: 15, arrow: true, color: document.getElementById('fillColor').value });
    activeSlide().elements.push(el); selectEl(el.id); markDirty(); renderPanel();
  });
  const imageInput = document.getElementById('imageInput');
  document.getElementById('btnAddImage').addEventListener('click', () => imageInput.click());
  imageInput.addEventListener('change', () => {
    const file = imageInput.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      snapshotUndo();
      const el = newEl('image', { src: reader.result, w: 40, h: 40 });
      activeSlide().elements.push(el); selectEl(el.id); markDirty(); renderPanel();
    };
    reader.readAsDataURL(file);
    imageInput.value = '';
  });
  document.getElementById('btnDeleteEl').addEventListener('click', () => {
    if (!selectedIds.length) return;
    snapshotUndo();
    activeSlide().elements = activeSlide().elements.filter(el => !isSelected(el.id));
    selectedIds = []; markDirty(); renderAll();
  });

  /* ---------------- toolbar: text formatting ---------------- */
  function withSelectedText(fn) {
    const el = primarySelected();
    if (!el || el.type !== 'text') return;
    snapshotUndo();
    selectedEls().filter(se => se.type === 'text').forEach(fn);
    markDirty(); renderSlide(); renderPanel();
  }
  document.getElementById('fontSize').addEventListener('change', (e) => withSelectedText(el => el.fontSize = +e.target.value));
  document.getElementById('fontFamily').addEventListener('change', (e) => withSelectedText(el => el.fontFamily = e.target.value));
  document.getElementById('tBold').addEventListener('click', () => withSelectedText(el => el.bold = !el.bold));
  document.getElementById('tItalic').addEventListener('click', () => withSelectedText(el => el.italic = !el.italic));
  document.querySelectorAll('#textTools [data-align]').forEach(btn => {
    btn.addEventListener('click', () => withSelectedText(el => el.align = btn.dataset.align));
  });
  document.getElementById('textColor').addEventListener('input', (e) => withSelectedText(el => el.color = e.target.value));
  document.getElementById('fillColor').addEventListener('input', (e) => {
    const els = selectedEls().filter(el => el.type === 'shape' || el.type === 'line');
    if (!els.length) return;
    snapshotUndo();
    els.forEach(el => { if (el.type === 'shape') el.fill = e.target.value; else el.color = e.target.value; });
    markDirty(); renderSlide(); renderPanel();
  });
  document.getElementById('bgColor').addEventListener('input', (e) => {
    snapshotUndo();
    activeSlide().bg = e.target.value; markDirty(); renderSlide(); renderPanel();
  });

  function syncTextTools() {
    const el = primarySelected();
    const isText = el && el.type === 'text';
    textTools.classList.toggle('disabled', !isText);
    if (isText) {
      document.getElementById('fontSize').value = el.fontSize || 24;
      document.getElementById('fontFamily').value = el.fontFamily || "Arial, sans-serif";
      document.getElementById('tBold').classList.toggle('active', !!el.bold);
      document.getElementById('tItalic').classList.toggle('active', !!el.italic);
      document.getElementById('textColor').value = el.color || '#14161f';
    }
  }

  /* ---------------- layering ---------------- */
  document.getElementById('btnToFront').addEventListener('click', () => {
    if (!selectedIds.length) return;
    snapshotUndo();
    const s = activeSlide();
    const selected = s.elements.filter(el => isSelected(el.id));
    s.elements = s.elements.filter(el => !isSelected(el.id)).concat(selected);
    markDirty(); renderSlide(); renderPanel();
  });
  document.getElementById('btnToBack').addEventListener('click', () => {
    if (!selectedIds.length) return;
    snapshotUndo();
    const s = activeSlide();
    const selected = s.elements.filter(el => isSelected(el.id));
    s.elements = selected.concat(s.elements.filter(el => !isSelected(el.id)));
    markDirty(); renderSlide(); renderPanel();
  });

  /* ---------------- align ---------------- */
  document.querySelectorAll('[data-align-obj]').forEach(btn => {
    btn.addEventListener('click', () => {
      const els = selectedEls();
      if (!els.length) return;
      snapshotUndo();
      const mode = btn.dataset.alignObj;
      els.forEach(el => {
        if (mode === 'left') el.x = 0;
        else if (mode === 'right') el.x = 100 - el.w;
        else if (mode === 'hcenter') el.x = 50 - el.w / 2;
        else if (mode === 'top') el.y = 0;
        else if (mode === 'bottom') el.y = 100 - el.h;
        else if (mode === 'vcenter') el.y = 50 - el.h / 2;
      });
      markDirty(); renderSlide(); renderPanel();
    });
  });

  /* ---------------- speaker notes ---------------- */
  document.getElementById('btnNotesToggle').addEventListener('click', () => {
    const panel = document.getElementById('notesPanel');
    panel.hidden = !panel.hidden;
  });
  notesInput.addEventListener('input', () => {
    activeSlide().notes = notesInput.value;
    markDirty();
  });

  /* ---------------- slides mgmt ---------------- */
  document.getElementById('btnAddSlideBlank').addEventListener('click', () => addSlide(newSlide()));
  document.getElementById('btnAddSlideTitle').addEventListener('click', () => addSlide(newSlide([
    newEl('text', { x: 8, y: 38, w: 84, h: 16, text: 'Title slide', fontSize: 36, bold: true, align: 'center' }),
    newEl('text', { x: 8, y: 56, w: 84, h: 10, text: 'Subtitle', fontSize: 18, align: 'center', color: '#5b5d6b' }),
  ])));
  document.getElementById('btnAddSlideContent').addEventListener('click', () => addSlide(newSlide([
    newEl('text', { x: 6, y: 6, w: 88, h: 14, text: 'Slide title', fontSize: 30, bold: true }),
    newEl('text', { x: 6, y: 24, w: 88, h: 68, text: '• First point\n• Second point\n• Third point', fontSize: 18, valign: 'top' }),
  ])));
  document.getElementById('btnDuplicateSlide').addEventListener('click', () => {
    const clone = JSON.parse(JSON.stringify(activeSlide()));
    clone.id = uid();
    clone.elements.forEach(el => { el.id = uid(); });
    addSlide(clone);
  });
  function addSlide(s) {
    snapshotUndo();
    deck.slides.splice(deck.active + 1, 0, s);
    deck.active += 1; selectedIds = [];
    markDirty(); renderAll();
  }
  function deleteSlide(idx) {
    if (!confirm('Delete this slide?')) return;
    snapshotUndo();
    deck.slides.splice(idx, 1);
    deck.active = Math.max(0, Math.min(deck.active, deck.slides.length - 1));
    selectedIds = []; markDirty(); renderAll();
  }

  /* ---------------- undo / redo ---------------- */
  function doUndo() {
    const prev = undoMgr.undo(deck);
    if (!prev) return;
    deck = prev;
    deck.active = Math.min(deck.active, deck.slides.length - 1);
    selectedIds = [];
    renderAll(); markDirty();
  }
  function doRedo() {
    const next = undoMgr.redo(deck);
    if (!next) return;
    deck = next;
    deck.active = Math.min(deck.active, deck.slides.length - 1);
    selectedIds = [];
    renderAll(); markDirty();
  }
  document.getElementById('btnUndo').addEventListener('click', doUndo);
  document.getElementById('btnRedo').addEventListener('click', doRedo);

  /* ---------------- save / load / multi-deck ---------------- */
  const deckList = document.getElementById('deckList');
  let allDecks = [];

  const setStatus = (t) => saveStatus.textContent = t;
  const saveNow = debounce(async () => {
    if (!deck) return;
    deck.title = deckTitle.value.trim() || 'Untitled presentation';
    deck.updatedAt = Date.now();
    await db.put(deck);
    setStatus('Saved');
    refreshDeckList();
  }, 500);
  function markDirty() { setStatus('Saving…'); saveNow(); }
  deckTitle.addEventListener('input', markDirty);

  async function refreshDeckList() {
    allDecks = (await db.getAll()).sort((a, b) => b.updatedAt - a.updatedAt);
    deckList.innerHTML = '';
    allDecks.forEach(d => {
      const item = document.createElement('div');
      item.className = 'doc-item' + (deck && d.id === deck.id ? ' active' : '');
      item.innerHTML = `<span class="name"></span><span class="meta"><span class="when"></span><button class="del" title="Delete">Delete</button></span>`;
      item.querySelector('.name').textContent = d.title || 'Untitled presentation';
      item.querySelector('.when').textContent = formatTime(d.updatedAt);
      item.addEventListener('click', (e) => {
        if (e.target.closest('.del')) return;
        loadDeck(d.id);
      });
      item.querySelector('.del').addEventListener('click', async (e) => {
        e.stopPropagation();
        if (!confirm(`Delete "${d.title}"? This can't be undone.`)) return;
        await db.delete(d.id);
        if (deck && d.id === deck.id) {
          const rest = allDecks.filter(x => x.id !== d.id);
          if (rest.length) loadDeck(rest[0].id); else createDeck();
        } else {
          refreshDeckList();
        }
      });
      deckList.appendChild(item);
    });
  }

  async function loadDeck(id) {
    const rec = await db.get(id);
    if (!rec) return;
    deck = rec;
    deck.slides.forEach(s => { if (s.notes == null) s.notes = ''; });
    undoMgr.clear();
    selectedIds = [];
    deckTitle.value = deck.title;
    renderAll();
    refreshDeckList();
    setStatus('Saved');
    history.replaceState(null, '', `slides.html?doc=${encodeURIComponent(id)}`);
  }

  async function createDeck() {
    deck = newDeck();
    await db.put(deck);
    await loadDeck(deck.id);
  }

  async function deleteCurrentDeck() {
    if (!deck) return;
    if (!confirm('Delete this presentation? This can\'t be undone.')) return;
    await db.delete(deck.id);
    const rest = (await db.getAll()).sort((a, b) => b.updatedAt - a.updatedAt);
    if (rest.length) loadDeck(rest[0].id); else createDeck();
  }

  document.getElementById('btnNewDeck').addEventListener('click', createDeck);

  async function boot() {
    const params = new URLSearchParams(location.search);
    const docId = params.get('doc');
    const wantNew = params.has('new');

    allDecks = await db.getAll();
    if (docId && allDecks.some(d => d.id === docId)) {
      deck = allDecks.find(d => d.id === docId);
    } else if (wantNew || allDecks.length === 0) {
      deck = newDeck();
      await db.put(deck);
    } else {
      allDecks.sort((a, b) => b.updatedAt - a.updatedAt);
      deck = allDecks[0];
    }
    deck.slides.forEach(s => { if (s.notes == null) s.notes = ''; });
    deckTitle.value = deck.title;
    renderAll();
    refreshDeckList();
  }
  boot();

  /* ---------------- present mode ---------------- */
  const overlay = document.getElementById('presentOverlay');
  const presentSlide = document.getElementById('presentSlide');
  const presCounter = document.getElementById('presCounter');
  let presentIdx = 0;

  document.getElementById('btnPresent').addEventListener('click', () => {
    presentIdx = deck.active;
    overlay.classList.add('open');
    renderPresent();
    if (overlay.requestFullscreen) overlay.requestFullscreen().catch(() => {});
  });
  document.getElementById('btnExitPresent').addEventListener('click', exitPresent);
  function exitPresent() {
    overlay.classList.remove('open');
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
  }
  function renderPresent() {
    renderElementsInto(presentSlide, deck.slides[presentIdx], { mini: true });
    presCounter.textContent = `${presentIdx + 1} / ${deck.slides.length}`;
  }
  document.getElementById('presPrev').addEventListener('click', () => { presentIdx = Math.max(0, presentIdx - 1); renderPresent(); });
  document.getElementById('presNext').addEventListener('click', () => { presentIdx = Math.min(deck.slides.length - 1, presentIdx + 1); renderPresent(); });
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) { presentIdx = Math.min(deck.slides.length - 1, presentIdx + 1); renderPresent(); }
  });
  document.addEventListener('keydown', (e) => {
    if (!overlay.classList.contains('open')) return;
    if (e.key === 'Escape') exitPresent();
    else if (e.key === 'ArrowRight' || e.key === ' ') { presentIdx = Math.min(deck.slides.length - 1, presentIdx + 1); renderPresent(); }
    else if (e.key === 'ArrowLeft') { presentIdx = Math.max(0, presentIdx - 1); renderPresent(); }
  });

  /* ---------------- export ---------------- */
  const exportMenuBtn = document.getElementById('btnExportMenu');
  const exportMenu = document.getElementById('exportMenu');
  exportMenuBtn.addEventListener('click', () => exportMenu.classList.toggle('open'));
  document.addEventListener('click', (e) => { if (!e.target.closest('.menu-wrap')) exportMenu.classList.remove('open'); });

  exportMenu.addEventListener('click', (e) => {
    const type = e.target.dataset.export;
    if (!type) return;
    if (type === 'pptx') exportPptx();
    else if (type === 'pdf') exportPdf();
    exportMenu.classList.remove('open');
  });

  function pctToIn(pct, total) { return (pct / 100) * total; }

  function exportPptx() {
    const pres = new window.PptxGenJS();
    pres.defineLayout({ name: 'CLOUD_WIDE', width: SLIDE_W_IN, height: SLIDE_H_IN });
    pres.layout = 'CLOUD_WIDE';
    deck.slides.forEach(s => {
      const slide = pres.addSlide();
      slide.background = { color: (s.bg || '#ffffff').replace('#', '') };
      if (s.notes) slide.addNotes(s.notes);
      s.elements.forEach(el => {
        const x = pctToIn(el.x, SLIDE_W_IN), y = pctToIn(el.y, SLIDE_H_IN);
        const w = pctToIn(el.w, SLIDE_W_IN), h = pctToIn(el.h, SLIDE_H_IN);
        if (el.type === 'text') {
          slide.addText(el.text || '', {
            x, y, w, h, fontSize: Math.max(1, Math.round(el.fontSize || 18)),
            bold: !!el.bold, italic: !!el.italic, align: el.align || 'left',
            color: (el.color || '#14161f').replace('#', ''), valign: el.valign === 'top' ? 'top' : 'middle',
            fontFace: el.fontFamily ? el.fontFamily.split(',')[0].replace(/'/g, '').trim() : undefined,
            wrap: true,
          });
        } else if (el.type === 'shape') {
          slide.addShape(el.shape === 'ellipse' ? 'ellipse' : 'rect', {
            x, y, w, h, fill: { color: (el.fill || '#d65d3a').replace('#', '') },
          });
        } else if (el.type === 'line') {
          slide.addShape('line', {
            x, y, w, h,
            line: { color: (el.color || '#14161f').replace('#', ''), width: 2, endArrowType: el.arrow ? 'triangle' : 'none' },
          });
        } else if (el.type === 'image') {
          slide.addImage({ data: el.src, x, y, w, h });
        }
      });
    });
    const name = (deckTitle.value.trim() || 'presentation').replace(/[^\w\- ]+/g, '');
    pres.writeFile({ fileName: `${name}.pptx` });
  }

  function exportPdf() {
    const style = document.createElement('style');
    style.id = 'printDeckStyle';
    style.textContent = `
      #printDeck{ position:fixed; left:-99999px; top:0; }
      @media print {
        body > *:not(#printDeck){ display:none !important; }
        #printDeck{ position:static; left:auto; top:auto; }
        .print-slide{ position:relative; page-break-after:always; overflow:hidden; margin:0 auto; }
      }
    `;
    document.head.appendChild(style);

    const container = document.createElement('div');
    container.id = 'printDeck';
    document.body.appendChild(container);

    deck.slides.forEach(s => {
      const page = document.createElement('div');
      page.className = 'print-slide';
      page.style.width = '960px';
      page.style.height = '540px';
      container.appendChild(page);
      renderElementsInto(page, s, { mini: true });
    });

    window.print();
    setTimeout(() => { container.remove(); style.remove(); }, 500);
  }

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
  }

  window.addEventListener('resize', debounce(() => {
    renderSlide(); renderPanel();
    if (overlay.classList.contains('open')) renderPresent();
  }, 120));

  /* ---------------- font family ---------------- */
  const FONTS = [
    'Arial', 'Helvetica', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Segoe UI', 'Calibri', 'Century Gothic',
    'Times New Roman', 'Georgia', 'Garamond', 'Cambria', 'Courier New', 'Consolas', 'Impact',
  ];
  const fontFamilySelect = document.getElementById('fontFamily');
  FONTS.forEach(f => {
    const opt = document.createElement('option');
    opt.value = /\s/.test(f) ? `'${f}', sans-serif` : `${f}, sans-serif`;
    opt.textContent = f;
    if (f === 'Arial') opt.selected = true;
    fontFamilySelect.appendChild(opt);
  });

  /* ---------------- deck sidebar toggle ---------------- */
  const deckSidebar = document.getElementById('deckSidebar');
  function toggleDeckSidebar() { deckSidebar.classList.toggle('collapsed'); }
  document.getElementById('toggleDeckSidebar').addEventListener('click', toggleDeckSidebar);

  /* ---------------- menu bar wiring ---------------- */
  CloudCommon.bindMenuAction('mNew', createDeck);
  CloudCommon.bindMenuAction('mExportPptx', () => document.querySelector('[data-export="pptx"]').click());
  CloudCommon.bindMenuAction('mExportPdf', () => document.querySelector('[data-export="pdf"]').click());
  CloudCommon.bindMenuAction('mPrint', exportPdf);
  CloudCommon.bindMenuAction('mDelete', deleteCurrentDeck);
  CloudCommon.bindMenuAction('mClose', () => { location.href = 'index.html'; });
  CloudCommon.bindMenuAction('mUndo', doUndo);
  CloudCommon.bindMenuAction('mRedo', doRedo);
  CloudCommon.bindMenuAction('mDeleteEl', () => document.getElementById('btnDeleteEl').click());
  CloudCommon.bindMenuAction('mAddText', () => document.getElementById('btnAddText').click());
  CloudCommon.bindMenuAction('mAddRect', () => document.getElementById('btnAddRect').click());
  CloudCommon.bindMenuAction('mAddEllipse', () => document.getElementById('btnAddEllipse').click());
  CloudCommon.bindMenuAction('mAddLine', () => document.getElementById('btnAddLine').click());
  CloudCommon.bindMenuAction('mAddArrow', () => document.getElementById('btnAddArrow').click());
  CloudCommon.bindMenuAction('mAddImage', () => document.getElementById('btnAddImage').click());
  CloudCommon.bindMenuAction('mAddSlideBlank', () => document.getElementById('btnAddSlideBlank').click());
  CloudCommon.bindMenuAction('mAddSlideTitle', () => document.getElementById('btnAddSlideTitle').click());
  CloudCommon.bindMenuAction('mAddSlideContent', () => document.getElementById('btnAddSlideContent').click());
  CloudCommon.bindMenuAction('mBold', () => document.getElementById('tBold').click());
  CloudCommon.bindMenuAction('mItalic', () => document.getElementById('tItalic').click());
  CloudCommon.bindMenuAction('mToFront', () => document.getElementById('btnToFront').click());
  CloudCommon.bindMenuAction('mToBack', () => document.getElementById('btnToBack').click());
  CloudCommon.bindMenuAction('mDuplicateSlide', () => document.getElementById('btnDuplicateSlide').click());
  CloudCommon.bindMenuAction('mDeleteSlide', () => {
    const idx = deck.active;
    if (deck.slides.length <= 1) { alert('A presentation needs at least one slide.'); return; }
    if (!confirm('Delete this slide?')) return;
    snapshotUndo();
    deck.slides.splice(idx, 1);
    deck.active = Math.max(0, Math.min(deck.active, deck.slides.length - 1));
    selectedIds = []; markDirty(); renderAll();
  });
  CloudCommon.bindMenuAction('mNotesToggle', () => document.getElementById('btnNotesToggle').click());
  CloudCommon.bindMenuAction('mPresent', () => document.getElementById('btnPresent').click());
  CloudCommon.bindMenuAction('mToggleDeckSidebar', toggleDeckSidebar);
  CloudCommon.bindMenuAction('mHelpCenter', () => { location.href = 'help.html#slides'; });
  CloudCommon.bindMenuAction('mAbout', () => { location.href = 'about.html'; });
  document.getElementById('btnHelp').addEventListener('click', () => { location.href = 'help.html#slides'; });
  document.getElementById('btnSettings').addEventListener('click', () => CloudCommon.openModal('settingsModal'));

  document.addEventListener('keydown', (e) => {
    if (e.key === 'F5' && !overlayOpen()) { e.preventDefault(); document.getElementById('btnPresent').click(); }
  });

  CloudCommon.initTheme();
  CloudCommon.initMenuBar();
  CloudCommon.initModalDismiss();
})();
