(() => {
  const ROWS = 100;
  const COLS = 26;

  const db = new CloudDB('cloud-sheet', 'workbooks');
  const grid = document.getElementById('grid');
  const gridWrap = document.getElementById('gridWrap');
  const bookTitle = document.getElementById('bookTitle');
  const saveStatus = document.getElementById('saveStatus');
  const cellRefEl = document.getElementById('cellRef');
  const formulaInput = document.getElementById('formulaInput');
  const tabsEl = document.getElementById('tabs');

  let book = null;      // { id, title, sheets:[{name, cells:{}}], active, updatedAt }
  let sel = { r: 0, c: 0 }; // 0-indexed
  let editing = false;

  /* ---------------- column / ref helpers ---------------- */
  const colName = (i) => {
    let s = '';
    i++;
    while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  };
  const colIndex = (name) => {
    let n = 0;
    for (const ch of name) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  };
  const ref = (r, c) => `${colName(c)}${r + 1}`;
  const parseRef = (s) => {
    const m = /^([A-Z]+)([0-9]+)$/.exec(s);
    if (!m) return null;
    return { r: parseInt(m[2], 10) - 1, c: colIndex(m[1]) };
  };

  function activeSheet() { return book.sheets[book.active]; }
  function getCell(r, c) { return activeSheet().cells[ref(r, c)]; }

  /* ---------------- formula engine ---------------- */
  const FUNCS = {
    SUM: (a) => flat(a).reduce((s, v) => s + (num(v) || 0), 0),
    AVERAGE: (a) => { const v = flat(a).filter(x => x !== '' && x != null).map(num).filter(x => !isNaN(x)); return v.length ? v.reduce((s, x) => s + x, 0) / v.length : 0; },
    MIN: (a) => { const v = flat(a).map(num).filter(x => !isNaN(x)); return v.length ? Math.min(...v) : 0; },
    MAX: (a) => { const v = flat(a).map(num).filter(x => !isNaN(x)); return v.length ? Math.max(...v) : 0; },
    COUNT: (a) => flat(a).map(num).filter(x => !isNaN(x)).length,
    COUNTA: (a) => flat(a).filter(v => v !== '' && v != null).length,
    IF: (a) => (a[0] ? a[1] : (a.length > 2 ? a[2] : false)),
    AND: (a) => flat(a).every(Boolean),
    OR: (a) => flat(a).some(Boolean),
    NOT: (a) => !a[0],
    CONCATENATE: (a) => flat(a).map(str).join(''),
    ROUND: (a) => { const d = a[1] != null ? num(a[1]) : 0; const f = Math.pow(10, d); return Math.round(num(a[0]) * f) / f; },
    ABS: (a) => Math.abs(num(a[0])),
    SQRT: (a) => Math.sqrt(num(a[0])),
    POWER: (a) => Math.pow(num(a[0]), num(a[1])),
    TODAY: () => new Date().toLocaleDateString(),
    NOW: () => new Date().toLocaleString(),
    LEN: (a) => str(a[0]).length,
    UPPER: (a) => str(a[0]).toUpperCase(),
    LOWER: (a) => str(a[0]).toLowerCase(),
    TRIM: (a) => str(a[0]).trim(),
  };
  function flat(a) { return a.reduce((acc, v) => acc.concat(Array.isArray(v) ? flat(v) : [v]), []); }
  function num(v) { const n = parseFloat(v); return isNaN(n) ? (v === true ? 1 : v === false ? 0 : NaN) : n; }
  function str(v) { return v == null ? '' : String(v); }

  class Tokenizer {
    constructor(src) { this.src = src; this.pos = 0; }
    tokens() {
      const out = [];
      const s = this.src;
      let i = 0;
      while (i < s.length) {
        const ch = s[i];
        if (/\s/.test(ch)) { i++; continue; }
        if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(s[i + 1] || ''))) {
          let j = i; while (j < s.length && /[0-9.]/.test(s[j])) j++;
          out.push({ t: 'num', v: parseFloat(s.slice(i, j)) }); i = j; continue;
        }
        if (ch === '"') {
          let j = i + 1; while (j < s.length && s[j] !== '"') j++;
          out.push({ t: 'str', v: s.slice(i + 1, j) }); i = j + 1; continue;
        }
        if (/[A-Za-z]/.test(ch)) {
          let j = i; while (j < s.length && /[A-Za-z0-9]/.test(s[j])) j++;
          const word = s.slice(i, j).toUpperCase();
          out.push({ t: /^[A-Z]+[0-9]+$/.test(word) ? 'ref' : 'name', v: word });
          i = j; continue;
        }
        if ('+-*/^(),:%'.includes(ch)) { out.push({ t: ch }); i++; continue; }
        if (ch === '=' || ch === '<' || ch === '>') {
          let op = ch;
          if (ch === '<' && s[i + 1] === '>') { op = '<>'; i++; }
          else if (s[i + 1] === '=') { op += '='; i++; }
          out.push({ t: 'cmp', v: op }); i++; continue;
        }
        if (ch === '&') { out.push({ t: '&' }); i++; continue; }
        i++; // skip unknown char
      }
      out.push({ t: 'end' });
      return out;
    }
  }

  class Parser {
    constructor(tokens, ctx) { this.tk = tokens; this.i = 0; this.ctx = ctx; }
    peek() { return this.tk[this.i]; }
    next() { return this.tk[this.i++]; }
    parse() { const v = this.expr(); return v; }
    expr() { return this.concat(); }
    concat() {
      let left = this.compare();
      while (this.peek().t === '&') { this.next(); const right = this.compare(); left = str(left) + str(right); }
      return left;
    }
    compare() {
      let left = this.additive();
      while (this.peek().t === 'cmp') {
        const op = this.next().v; const right = this.additive();
        const l = num(left), r = num(right);
        switch (op) {
          case '=': left = left == right; break;
          case '<': left = l < r; break;
          case '>': left = l > r; break;
          case '<=': left = l <= r; break;
          case '>=': left = l >= r; break;
          case '<>': left = left != right; break;
        }
      }
      return left;
    }
    additive() {
      let left = this.term();
      while (this.peek().t === '+' || this.peek().t === '-') {
        const op = this.next().t; const right = this.term();
        left = op === '+' ? num(left) + num(right) : num(left) - num(right);
      }
      return left;
    }
    term() {
      let left = this.power();
      while (this.peek().t === '*' || this.peek().t === '/' || this.peek().t === '%') {
        const op = this.next().t; const right = this.power();
        if (op === '*') left = num(left) * num(right);
        else if (op === '/') left = num(left) / num(right);
        else left = num(left) % num(right);
      }
      return left;
    }
    power() {
      let left = this.unary();
      if (this.peek().t === '^') { this.next(); const right = this.power(); left = Math.pow(num(left), num(right)); }
      return left;
    }
    unary() {
      if (this.peek().t === '-') { this.next(); return -num(this.unary()); }
      if (this.peek().t === '+') { this.next(); return num(this.unary()); }
      return this.primary();
    }
    primary() {
      const tok = this.peek();
      if (tok.t === 'num') { this.next(); return tok.v; }
      if (tok.t === 'str') { this.next(); return tok.v; }
      if (tok.t === '(') { this.next(); const v = this.expr(); if (this.peek().t === ')') this.next(); return v; }
      if (tok.t === 'name') {
        this.next();
        if (this.peek().t === '(') {
          this.next();
          const args = [];
          if (this.peek().t !== ')') {
            args.push(this.expr());
            while (this.peek().t === ',') { this.next(); args.push(this.expr()); }
          }
          if (this.peek().t === ')') this.next();
          const fn = FUNCS[tok.v];
          if (!fn) throw new Error('#NAME?');
          return fn(args);
        }
        if (tok.v === 'TRUE') return true;
        if (tok.v === 'FALSE') return false;
        throw new Error('#NAME?');
      }
      if (tok.t === 'ref') {
        this.next();
        if (this.peek().t === ':') {
          this.next();
          const tok2 = this.next(); // ref
          return this.ctx.range(tok.v, tok2.v);
        }
        return this.ctx.cellValue(tok.v);
      }
      this.next();
      return '';
    }
  }

  const evalStack = new Set();
  function evalFormula(formula, sheet) {
    const ctx = {
      cellValue(refStr) {
        const p = parseRef(refStr);
        if (!p) return '';
        return computeCell(p.r, p.c, sheet);
      },
      range(a, b) {
        const p1 = parseRef(a), p2 = parseRef(b);
        const r1 = Math.min(p1.r, p2.r), r2 = Math.max(p1.r, p2.r);
        const c1 = Math.min(p1.c, p2.c), c2 = Math.max(p1.c, p2.c);
        const out = [];
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out.push(computeCell(r, c, sheet));
        return out;
      },
    };
    const tokens = new Tokenizer(formula.slice(1)).tokens();
    const parser = new Parser(tokens, ctx);
    return parser.parse();
  }

  function computeCell(r, c, sheet) {
    const key = ref(r, c);
    const cell = sheet.cells[key];
    if (!cell) return '';
    if (cell.f == null) return cell.v ?? '';
    const stackKey = sheet.name + '!' + key;
    if (evalStack.has(stackKey)) return '#CIRC!';
    evalStack.add(stackKey);
    let result;
    try { result = evalFormula(cell.f, sheet); }
    catch (e) { result = typeof e.message === 'string' && e.message.startsWith('#') ? e.message : '#ERROR!'; }
    evalStack.delete(stackKey);
    return result;
  }

  /* ---------------- formatting ---------------- */
  function formatValue(v, fmt) {
    if (v === '' || v == null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'string' && (fmt == null || fmt === 'general')) return v;
    const n = num(v);
    if (isNaN(n)) return str(v);
    switch (fmt) {
      case 'number': return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      case 'integer': return Math.round(n).toLocaleString();
      case 'currency': return '$' + n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      case 'percent': return (n * 100).toLocaleString(undefined, { maximumFractionDigits: 1 }) + '%';
      case 'date': { const d = new Date(n); return isNaN(d) ? str(v) : d.toLocaleDateString(); }
      default: return str(v);
    }
  }

  /* ---------------- grid rendering ---------------- */
  function buildGridSkeleton() {
    grid.innerHTML = '';
    const thead = document.createElement('thead');
    const hr = document.createElement('tr');
    hr.appendChild(Object.assign(document.createElement('th'), { className: 'corner', textContent: '' }));
    for (let c = 0; c < COLS; c++) {
      const th = document.createElement('th');
      th.textContent = colName(c);
      hr.appendChild(th);
    }
    thead.appendChild(hr);
    grid.appendChild(thead);

    const tbody = document.createElement('tbody');
    for (let r = 0; r < ROWS; r++) {
      const tr = document.createElement('tr');
      const rh = document.createElement('th');
      rh.className = 'rowhead';
      rh.textContent = r + 1;
      tr.appendChild(rh);
      for (let c = 0; c < COLS; c++) {
        const td = document.createElement('td');
        td.dataset.r = r; td.dataset.c = c;
        td.addEventListener('mousedown', onCellMouseDown);
        td.addEventListener('dblclick', () => startEdit());
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    grid.appendChild(tbody);
  }

  function renderAll() {
    const sheet = activeSheet();
    for (let r = 0; r < ROWS; r++) {
      for (let c = 0; c < COLS; c++) {
        renderCell(r, c, sheet);
      }
    }
    highlightSelection();
    renderTabs();
    formulaInput.value = currentRaw();
    cellRefEl.textContent = ref(sel.r, sel.c);
  }

  function tdAt(r, c) { return grid.tBodies[0].rows[r].cells[c + 1]; }

  function renderCell(r, c, sheet) {
    const td = tdAt(r, c);
    const cell = sheet.cells[ref(r, c)];
    const val = computeCell(r, c, sheet);
    td.textContent = formatValue(val, cell && cell.style && cell.style.fmt);
    td.className = td.classList.contains('selected') ? 'selected' : '';
    if (cell && cell.style) {
      const st = cell.style;
      td.style.fontWeight = st.bold ? '700' : '400';
      td.style.fontStyle = st.italic ? 'italic' : 'normal';
      td.style.textDecoration = st.underline ? 'underline' : 'none';
      td.style.textAlign = st.align || (typeof val === 'number' ? 'right' : 'left');
      td.style.color = st.color || '';
      td.style.background = st.fill || '';
      if (st.border) td.classList.add('bordered');
    } else {
      td.style.fontWeight = '400'; td.style.fontStyle = 'normal'; td.style.textDecoration = 'none';
      td.style.textAlign = typeof val === 'number' ? 'right' : 'left';
      td.style.color = ''; td.style.background = '';
    }
  }

  function renderColumn(c) { for (let r = 0; r < ROWS; r++) renderCell(r, c, activeSheet()); }
  function renderDependents() {
    // simplest correct approach: recompute + repaint the whole grid (cheap enough at this size)
    const sheet = activeSheet();
    for (let r = 0; r < ROWS; r++) for (let c = 0; c < COLS; c++) renderCell(r, c, sheet);
  }

  function highlightSelection() {
    grid.querySelectorAll('td.selected').forEach(td => td.classList.remove('selected'));
    const td = tdAt(sel.r, sel.c);
    if (td) { td.classList.add('selected'); td.scrollIntoView({ block: 'nearest', inline: 'nearest' }); }
  }

  function currentRaw() {
    const cell = getCell(sel.r, sel.c);
    if (!cell) return '';
    return cell.f != null ? cell.f : (cell.v ?? '');
  }

  /* ---------------- selection / editing ---------------- */
  function onCellMouseDown(e) {
    if (editing) commitEdit();
    sel = { r: +e.currentTarget.dataset.r, c: +e.currentTarget.dataset.c };
    highlightSelection();
    formulaInput.value = currentRaw();
    cellRefEl.textContent = ref(sel.r, sel.c);
    grid.focus();
  }

  function startEdit(prefill) {
    editing = true;
    const td = tdAt(sel.r, sel.c);
    td.classList.add('editing');
    td.contentEditable = 'true';
    td.textContent = prefill != null ? prefill : currentRaw();
    td.focus();
    placeCursorEnd(td);
  }

  function placeCursorEnd(el) {
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(false);
    const s = window.getSelection();
    s.removeAllRanges(); s.addRange(range);
  }

  function commitEdit(moveDelta) {
    const td = tdAt(sel.r, sel.c);
    let raw;
    if (editing) {
      raw = td.textContent.trim();
      td.contentEditable = 'false';
      td.classList.remove('editing');
      editing = false;
    } else {
      raw = formulaInput.value.trim();
    }
    setCellRaw(sel.r, sel.c, raw);
    markDirty();
    renderDependents();
    if (moveDelta) moveSelection(moveDelta[0], moveDelta[1]);
    else { renderCell(sel.r, sel.c, activeSheet()); highlightSelection(); }
    formulaInput.value = currentRaw();
    cellRefEl.textContent = ref(sel.r, sel.c);
  }

  function cancelEdit() {
    if (editing) {
      const td = tdAt(sel.r, sel.c);
      td.contentEditable = 'false';
      td.classList.remove('editing');
      editing = false;
      renderCell(sel.r, sel.c, activeSheet());
    }
    formulaInput.value = currentRaw();
  }

  function setCellRaw(r, c, raw) {
    const sheet = activeSheet();
    const key = ref(r, c);
    const existing = sheet.cells[key];
    if (raw === '') {
      if (existing) { delete existing.f; delete existing.v; if (!existing.style) delete sheet.cells[key]; }
      return;
    }
    const cell = existing || (sheet.cells[key] = { style: {} });
    if (raw.startsWith('=')) { cell.f = raw; cell.v = undefined; }
    else if (/^[-+]?[0-9]*\.?[0-9]+%$/.test(raw)) { cell.f = null; cell.v = parseFloat(raw) / 100; }
    else { cell.f = null; const n = num(raw); cell.v = (raw !== '' && !isNaN(n) && /^[-+]?[0-9]*\.?[0-9]+$/.test(raw)) ? n : raw; }
  }

  function moveSelection(dr, dc) {
    sel.r = Math.min(ROWS - 1, Math.max(0, sel.r + dr));
    sel.c = Math.min(COLS - 1, Math.max(0, sel.c + dc));
    highlightSelection();
    formulaInput.value = currentRaw();
    cellRefEl.textContent = ref(sel.r, sel.c);
  }

  grid.addEventListener('keydown', (e) => {
    if (editing) {
      if (e.key === 'Enter') { e.preventDefault(); commitEdit([1, 0]); }
      else if (e.key === 'Tab') { e.preventDefault(); commitEdit([0, 1]); }
      else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
      return;
    }
    if (e.key === 'ArrowUp') { e.preventDefault(); moveSelection(-1, 0); }
    else if (e.key === 'ArrowDown') { e.preventDefault(); moveSelection(1, 0); }
    else if (e.key === 'ArrowLeft') { e.preventDefault(); moveSelection(0, -1); }
    else if (e.key === 'ArrowRight') { e.preventDefault(); moveSelection(0, 1); }
    else if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); startEdit(); }
    else if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); setCellRaw(sel.r, sel.c, ''); markDirty(); renderDependents(); highlightSelection(); formulaInput.value = ''; }
    else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) { startEdit(e.key); }
  });
  grid.tabIndex = 0;
  gridWrap.addEventListener('click', () => grid.focus());

  formulaInput.addEventListener('focus', () => { if (!editing) { editing = 'bar'; } });
  formulaInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); editing = false; commitEdit([1, 0]); grid.focus(); }
    else if (e.key === 'Escape') { e.preventDefault(); editing = false; formulaInput.value = currentRaw(); grid.focus(); }
  });
  formulaInput.addEventListener('blur', () => {
    if (editing === 'bar') { editing = false; commitEdit(null); }
  });

  /* ---------------- toolbar ---------------- */
  function applyStyle(mutator) {
    const sheet = activeSheet();
    const key = ref(sel.r, sel.c);
    const cell = sheet.cells[key] || (sheet.cells[key] = { f: null, v: '', style: {} });
    cell.style = cell.style || {};
    mutator(cell.style);
    renderCell(sel.r, sel.c, sheet);
    markDirty();
  }
  document.querySelectorAll('.tbtn[data-cmd]').forEach(btn => {
    btn.addEventListener('click', () => {
      applyStyle(st => { st[btn.dataset.cmd] = !st[btn.dataset.cmd]; });
      grid.focus();
    });
  });
  document.querySelectorAll('.tbtn[data-align]').forEach(btn => {
    btn.addEventListener('click', () => { applyStyle(st => { st.align = btn.dataset.align; }); grid.focus(); });
  });
  document.getElementById('textColor').addEventListener('input', (e) => applyStyle(st => { st.color = e.target.value; }));
  document.getElementById('fillColor').addEventListener('input', (e) => applyStyle(st => { st.fill = e.target.value; }));
  document.getElementById('numFormat').addEventListener('change', (e) => applyStyle(st => { st.fmt = e.target.value; }));
  document.getElementById('btnBorder').addEventListener('click', () => applyStyle(st => { st.border = !st.border; }));

  document.getElementById('btnInsertRow').addEventListener('click', () => shiftRows(sel.r, 1));
  document.getElementById('btnDeleteRow').addEventListener('click', () => shiftRows(sel.r, -1));
  document.getElementById('btnInsertCol').addEventListener('click', () => shiftCols(sel.c, 1));
  document.getElementById('btnDeleteCol').addEventListener('click', () => shiftCols(sel.c, -1));

  function shiftRows(atRow, dir) {
    const sheet = activeSheet();
    const newCells = {};
    Object.keys(sheet.cells).forEach(k => {
      const p = parseRef(k);
      let r = p.r;
      if (dir > 0 && r >= atRow) r += 1;
      if (dir < 0) { if (r === atRow) return; if (r > atRow) r -= 1; }
      if (r >= ROWS) return;
      newCells[ref(r, p.c)] = sheet.cells[k];
    });
    sheet.cells = newCells;
    markDirty(); renderDependents(); highlightSelection();
  }
  function shiftCols(atCol, dir) {
    const sheet = activeSheet();
    const newCells = {};
    Object.keys(sheet.cells).forEach(k => {
      const p = parseRef(k);
      let c = p.c;
      if (dir > 0 && c >= atCol) c += 1;
      if (dir < 0) { if (c === atCol) return; if (c > atCol) c -= 1; }
      if (c >= COLS) return;
      newCells[ref(p.r, c)] = sheet.cells[k];
    });
    sheet.cells = newCells;
    markDirty(); renderDependents(); highlightSelection();
  }

  document.getElementById('btnSortAsc').addEventListener('click', () => sortColumn(1));
  document.getElementById('btnSortDesc').addEventListener('click', () => sortColumn(-1));
  function sortColumn(dir) {
    const sheet = activeSheet();
    const c = sel.c;
    const rowsData = [];
    for (let r = 0; r < ROWS; r++) rowsData.push(sheet.cells[ref(r, c)] || null);
    const indexed = rowsData.map((cell, i) => ({ cell, i }));
    indexed.sort((a, b) => {
      const av = a.cell ? (a.cell.f != null ? computeCell(a.i, c, sheet) : a.cell.v) : '';
      const bv = b.cell ? (b.cell.f != null ? computeCell(b.i, c, sheet) : b.cell.v) : '';
      const an = num(av), bn = num(bv);
      let cmp;
      if (!isNaN(an) && !isNaN(bn)) cmp = an - bn;
      else cmp = str(av).localeCompare(str(bv));
      return cmp * dir;
    });
    // rewrite whole column preserving other columns unaffected
    const newVals = indexed.map(x => x.cell);
    for (let r = 0; r < ROWS; r++) {
      const key = ref(r, c);
      if (newVals[r]) sheet.cells[key] = newVals[r]; else delete sheet.cells[key];
    }
    markDirty(); renderDependents();
  }

  /* ---------------- tabs ---------------- */
  function renderTabs() {
    tabsEl.innerHTML = '';
    book.sheets.forEach((s, idx) => {
      const tab = document.createElement('div');
      tab.className = 'tab' + (idx === book.active ? ' active' : '');
      const name = document.createElement('span');
      name.textContent = s.name;
      tab.appendChild(name);
      if (book.sheets.length > 1) {
        const close = document.createElement('span');
        close.className = 'close'; close.textContent = '✕';
        close.addEventListener('click', (e) => { e.stopPropagation(); deleteSheet(idx); });
        tab.appendChild(close);
      }
      tab.addEventListener('click', () => { book.active = idx; sel = { r: 0, c: 0 }; renderAll(); });
      tab.addEventListener('dblclick', () => {
        const n = prompt('Sheet name:', s.name);
        if (n) { s.name = n; markDirty(); renderTabs(); }
      });
      tabsEl.appendChild(tab);
    });
  }
  document.getElementById('btnAddSheet').addEventListener('click', () => {
    book.sheets.push({ name: 'Sheet' + (book.sheets.length + 1), cells: {} });
    book.active = book.sheets.length - 1;
    sel = { r: 0, c: 0 };
    markDirty(); renderAll();
  });
  function deleteSheet(idx) {
    if (!confirm(`Delete "${book.sheets[idx].name}"?`)) return;
    book.sheets.splice(idx, 1);
    book.active = Math.max(0, book.active - (idx <= book.active ? 1 : 0));
    markDirty(); renderAll();
  }

  /* ---------------- save / load ---------------- */
  const setStatus = (t) => saveStatus.textContent = t;
  const saveNow = debounce(async () => {
    if (!book) return;
    book.title = bookTitle.value.trim() || 'Untitled workbook';
    book.updatedAt = Date.now();
    await db.put(book);
    setStatus('Saved');
  }, 500);
  function markDirty() { setStatus('Saving…'); saveNow(); }
  bookTitle.addEventListener('input', markDirty);

  function newBook() {
    return { id: uid(), title: 'Untitled workbook', sheets: [{ name: 'Sheet1', cells: {} }], active: 0, updatedAt: Date.now() };
  }

  async function boot() {
    buildGridSkeleton();
    const all = await db.getAll();
    if (all.length === 0) { book = newBook(); await db.put(book); }
    else { all.sort((a, b) => b.updatedAt - a.updatedAt); book = all[0]; }
    bookTitle.value = book.title;
    renderAll();
  }
  boot();

  /* ---------------- print ---------------- */
  document.getElementById('btnPrint').addEventListener('click', () => window.print());

  /* ---------------- export ---------------- */
  const exportMenuBtn = document.getElementById('btnExportMenu');
  const exportMenu = document.getElementById('exportMenu');
  exportMenuBtn.addEventListener('click', () => exportMenu.classList.toggle('open'));
  document.addEventListener('click', (e) => { if (!e.target.closest('.menu-wrap')) exportMenu.classList.remove('open'); });

  function download(blob, filename) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
  }

  function sheetToAOA(sheet) {
    const aoa = [];
    for (let r = 0; r < ROWS; r++) {
      let rowHasData = false;
      const row = [];
      for (let c = 0; c < COLS; c++) {
        const cell = sheet.cells[ref(r, c)];
        if (cell) rowHasData = true;
        row.push(cell ? (cell.f != null ? computeCell(r, c, sheet) : cell.v) : '');
      }
      if (rowHasData || aoa.length) aoa.push(row);
    }
    // trim trailing empty rows
    while (aoa.length && aoa[aoa.length - 1].every(v => v === '' || v == null)) aoa.pop();
    return aoa;
  }

  exportMenu.addEventListener('click', (e) => {
    const type = e.target.dataset.export;
    if (!type) return;
    const name = (bookTitle.value.trim() || 'workbook').replace(/[^\w\- ]+/g, '');
    if (type === 'xlsx') {
      const wb = XLSX.utils.book_new();
      book.sheets.forEach(s => {
        const ws = XLSX.utils.aoa_to_sheet(sheetToAOA(s));
        XLSX.utils.book_append_sheet(wb, ws, s.name.slice(0, 31));
      });
      XLSX.writeFile(wb, `${name}.xlsx`);
    } else if (type === 'csv') {
      const ws = XLSX.utils.aoa_to_sheet(sheetToAOA(activeSheet()));
      const csv = XLSX.utils.sheet_to_csv(ws);
      download(new Blob([csv], { type: 'text/csv' }), `${name}.csv`);
    }
    exportMenu.classList.remove('open');
  });

  /* ---------------- import ---------------- */
  document.getElementById('fileOpen').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const newSheets = wb.SheetNames.map(name => {
      const ws = wb.Sheets[name];
      const cells = {};
      Object.keys(ws).forEach(addr => {
        if (addr.startsWith('!')) return;
        const p = parseRef(addr);
        if (!p || p.r >= ROWS || p.c >= COLS) return;
        const cd = ws[addr];
        cells[addr] = { f: cd.f ? '=' + cd.f : null, v: cd.v, style: {} };
      });
      return { name, cells };
    });
    book = { id: uid(), title: file.name.replace(/\.[^.]+$/, ''), sheets: newSheets.length ? newSheets : [{ name: 'Sheet1', cells: {} }], active: 0, updatedAt: Date.now() };
    await db.put(book);
    bookTitle.value = book.title;
    sel = { r: 0, c: 0 };
    renderAll();
    e.target.value = '';
  });

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
  }
})();
