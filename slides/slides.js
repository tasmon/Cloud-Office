(() => {
  const db = new CloudDB('cloud-slides', 'decks');
  const slideEl = document.getElementById('slide');
  const slideList = document.getElementById('slideList');
  const deckTitle = document.getElementById('deckTitle');
  const saveStatus = document.getElementById('saveStatus');
  const textTools = document.getElementById('textTools');
  const slideCountLabel = document.getElementById('slideCountLabel');

  const SLIDE_W_IN = 13.333, SLIDE_H_IN = 7.5; // widescreen inches for pptx export

  let deck = null;      // { id, title, slides:[{id,bg,elements:[...]}], active, updatedAt }
  let selectedId = null;
  let dragState = null;

  const newEl = (type, extra) => Object.assign({
    id: uid(), type, x: 10, y: 10, w: 30, h: 12, rotation: 0,
  }, extra);

  const newSlide = (elements = [], bg = '#ffffff') => ({ id: uid(), bg, elements });

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

  /* ---------------- rendering ---------------- */
  function renderElementsInto(container, slideData, opts = {}) {
    container.innerHTML = '';
    if (!container.style.position) container.style.position = 'relative';
    if (!container.style.overflow) container.style.overflow = 'hidden';

    // Elements are authored against a fixed 960x540 canvas (matches the 13.333in x 7.5in
    // export size at 72 "px" per inch, so 1 canvas px === 1 pt). The canvas is then scaled
    // with a CSS transform to fit whatever size the container actually renders at, so
    // font sizes stay proportional at any zoom level — thumbnail, main editor, or present mode.
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
      }
      if (!opts.mini) {
        div.dataset.id = el.id;
        if (el.id === selectedId) div.classList.add('selected');
        div.addEventListener('mousedown', (e) => onElMouseDown(e, el));
        if (el.type === 'text') div.addEventListener('dblclick', () => editText(div, el));
        const handle = document.createElement('div');
        handle.className = 'handle';
        handle.addEventListener('mousedown', (e) => onResizeMouseDown(e, el));
        div.appendChild(handle);
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
      thumb.addEventListener('click', () => { deck.active = idx; selectedId = null; renderAll(); });
      slideList.appendChild(thumb);
      renderElementsInto(mini, s, { mini: true });
    });
    slideCountLabel.textContent = `${deck.active + 1} / ${deck.slides.length}`;
  }

  function renderAll() { renderSlide(); renderPanel(); document.getElementById('bgColor').value = rgbToHex(activeSlide().bg); }

  function rgbToHex(v) { return /^#/.test(v) ? v : '#ffffff'; }

  /* ---------------- selection & drag ---------------- */
  function selectEl(id) { selectedId = id; renderSlide(); }

  slideEl.parentElement.addEventListener('mousedown', (e) => {
    if (e.target === slideEl) { selectedId = null; renderSlide(); }
  });

  function onElMouseDown(e, el) {
    if (e.target.classList.contains('handle')) return;
    e.stopPropagation();
    selectEl(el.id);
    const rect = slideEl.getBoundingClientRect();
    dragState = { mode: 'move', el, startX: e.clientX, startY: e.clientY, origX: el.x, origY: el.y, rect };
    window.addEventListener('mousemove', onDragMove);
    window.addEventListener('mouseup', onDragEnd);
  }
  function onResizeMouseDown(e, el) {
    e.stopPropagation(); e.preventDefault();
    selectEl(el.id);
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
      el.x = clamp(dragState.origX + dxPct, -5, 100 - 2);
      el.y = clamp(dragState.origY + dyPct, -5, 100 - 2);
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
    const inner = div.querySelector('.inner');
    inner.contentEditable = 'true';
    inner.focus();
    document.execCommand('selectAll', false, null);
    inner.addEventListener('blur', function onBlur() {
      el.text = inner.textContent;
      inner.contentEditable = 'false';
      inner.removeEventListener('blur', onBlur);
      markDirty(); renderPanel();
    }, { once: true });
  }

  document.addEventListener('keydown', (e) => {
    if ((e.key === 'Delete' || e.key === 'Backspace') && selectedId && !isEditingText()) {
      e.preventDefault();
      activeSlide().elements = activeSlide().elements.filter(el => el.id !== selectedId);
      selectedId = null;
      markDirty(); renderAll();
    }
  });
  function isEditingText() {
    const active = document.activeElement;
    return active && active.classList && active.classList.contains('inner') && active.isContentEditable;
  }

  /* ---------------- toolbar: insert ---------------- */
  document.getElementById('btnAddText').addEventListener('click', () => {
    const el = newEl('text', { text: 'New text', fontSize: 24, align: 'left', color: '#14161f' });
    activeSlide().elements.push(el);
    selectEl(el.id); markDirty(); renderPanel();
  });
  document.getElementById('btnAddRect').addEventListener('click', () => {
    const el = newEl('shape', { shape: 'rect', fill: document.getElementById('fillColor').value });
    activeSlide().elements.push(el); selectEl(el.id); markDirty(); renderPanel();
  });
  document.getElementById('btnAddEllipse').addEventListener('click', () => {
    const el = newEl('shape', { shape: 'ellipse', fill: document.getElementById('fillColor').value });
    activeSlide().elements.push(el); selectEl(el.id); markDirty(); renderPanel();
  });
  const imageInput = document.getElementById('imageInput');
  document.getElementById('btnAddImage').addEventListener('click', () => imageInput.click());
  imageInput.addEventListener('change', () => {
    const file = imageInput.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const el = newEl('image', { src: reader.result, w: 40, h: 40 });
      activeSlide().elements.push(el); selectEl(el.id); markDirty(); renderPanel();
    };
    reader.readAsDataURL(file);
    imageInput.value = '';
  });
  document.getElementById('btnDeleteEl').addEventListener('click', () => {
    if (!selectedId) return;
    activeSlide().elements = activeSlide().elements.filter(el => el.id !== selectedId);
    selectedId = null; markDirty(); renderAll();
  });

  /* ---------------- toolbar: text formatting ---------------- */
  function withSelectedText(fn) {
    const el = selectedId && findEl(selectedId);
    if (!el || el.type !== 'text') return;
    fn(el); markDirty(); renderSlide(); renderPanel();
  }
  document.getElementById('fontSize').addEventListener('change', (e) => withSelectedText(el => el.fontSize = +e.target.value));
  document.getElementById('tBold').addEventListener('click', () => withSelectedText(el => el.bold = !el.bold));
  document.getElementById('tItalic').addEventListener('click', () => withSelectedText(el => el.italic = !el.italic));
  document.querySelectorAll('#textTools [data-align]').forEach(btn => {
    btn.addEventListener('click', () => withSelectedText(el => el.align = btn.dataset.align));
  });
  document.getElementById('textColor').addEventListener('input', (e) => withSelectedText(el => el.color = e.target.value));
  document.getElementById('fillColor').addEventListener('input', (e) => {
    const el = selectedId && findEl(selectedId);
    if (el && (el.type === 'shape')) { el.fill = e.target.value; markDirty(); renderSlide(); renderPanel(); }
  });
  document.getElementById('bgColor').addEventListener('input', (e) => {
    activeSlide().bg = e.target.value; markDirty(); renderSlide(); renderPanel();
  });

  function syncTextTools() {
    const el = selectedId && findEl(selectedId);
    const isText = el && el.type === 'text';
    textTools.classList.toggle('disabled', !isText);
    if (isText) {
      document.getElementById('fontSize').value = el.fontSize || 24;
      document.getElementById('tBold').classList.toggle('active', !!el.bold);
      document.getElementById('tItalic').classList.toggle('active', !!el.italic);
      document.getElementById('textColor').value = el.color || '#14161f';
    }
  }

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
  function addSlide(s) {
    deck.slides.splice(deck.active + 1, 0, s);
    deck.active += 1; selectedId = null;
    markDirty(); renderAll();
  }
  function deleteSlide(idx) {
    if (!confirm('Delete this slide?')) return;
    deck.slides.splice(idx, 1);
    deck.active = Math.max(0, Math.min(deck.active, deck.slides.length - 1));
    selectedId = null; markDirty(); renderAll();
  }

  /* ---------------- save / load ---------------- */
  const setStatus = (t) => saveStatus.textContent = t;
  const saveNow = debounce(async () => {
    if (!deck) return;
    deck.title = deckTitle.value.trim() || 'Untitled presentation';
    deck.updatedAt = Date.now();
    await db.put(deck);
    setStatus('Saved');
  }, 500);
  function markDirty() { setStatus('Saving…'); saveNow(); }
  deckTitle.addEventListener('input', markDirty);

  async function boot() {
    const all = await db.getAll();
    if (all.length === 0) { deck = newDeck(); await db.put(deck); }
    else { all.sort((a, b) => b.updatedAt - a.updatedAt); deck = all[0]; }
    deckTitle.value = deck.title;
    renderAll();
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
      s.elements.forEach(el => {
        const x = pctToIn(el.x, SLIDE_W_IN), y = pctToIn(el.y, SLIDE_H_IN);
        const w = pctToIn(el.w, SLIDE_W_IN), h = pctToIn(el.h, SLIDE_H_IN);
        if (el.type === 'text') {
          slide.addText(el.text || '', {
            x, y, w, h, fontSize: Math.max(1, Math.round(el.fontSize || 18)),
            bold: !!el.bold, italic: !!el.italic, align: el.align || 'left',
            color: (el.color || '#14161f').replace('#', ''), valign: el.valign === 'top' ? 'top' : 'middle',
            wrap: true,
          });
        } else if (el.type === 'shape') {
          slide.addShape(el.shape === 'ellipse' ? 'ellipse' : 'rect', {
            x, y, w, h, fill: { color: (el.fill || '#d65d3a').replace('#', '') },
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
})();
