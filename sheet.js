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

  let book = null;      // { id, title, sheets:[{name, cells:{}, merges:[], freeze:{row,col}}], active, updatedAt }
  let sel = { r: 0, c: 0 }; // 0-indexed, the active/focus cell
  let selAnchor = { r: 0, c: 0 }; // range anchor (equals sel for a single-cell selection)
  let isSelecting = false;
  let editing = false;
  let clipboard = null; // { rows: [[cellObj,...]], cut: bool, srcR1,srcC1,srcR2,srcC2 }
  const undoMgr = new UndoManager(80);
  function snapshotUndo() { undoMgr.snapshot(book); }

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
  function mergeAt(r, c, sheet) {
    return (sheet.merges || []).find(m => r >= m.r && r < m.r + m.rowSpan && c >= m.c && c < m.c + m.colSpan);
  }
  function rangeBounds() {
    const r1 = Math.min(sel.r, selAnchor.r), r2 = Math.max(sel.r, selAnchor.r);
    const c1 = Math.min(sel.c, selAnchor.c), c2 = Math.max(sel.c, selAnchor.c);
    return { r1, r2, c1, c2 };
  }
  function isRangeSelection() { return sel.r !== selAnchor.r || sel.c !== selAnchor.c; }

  /* ---------------- formula engine ---------------- */
  const FUNCS = {
    SUM: (a) => flat(a).reduce((s, v) => s + (num(v) || 0), 0),
    AVERAGE: (a) => { const v = flat(a).filter(x => x !== '' && x != null).map(num).filter(x => !isNaN(x)); return v.length ? v.reduce((s, x) => s + x, 0) / v.length : 0; },
    MIN: (a) => { const v = flat(a).map(num).filter(x => !isNaN(x)); return v.length ? Math.min(...v) : 0; },
    MAX: (a) => { const v = flat(a).map(num).filter(x => !isNaN(x)); return v.length ? Math.max(...v) : 0; },
    COUNT: (a) => flat(a).map(num).filter(x => !isNaN(x)).length,
    COUNTA: (a) => flat(a).filter(v => v !== '' && v != null).length,
    COUNTIF: (a) => { const [range, crit] = a; const test = matchCriteria(crit); return flat([range]).filter(test).length; },
    SUMIF: (a) => {
      const range = flat([a[0]]); const test = matchCriteria(a[1]);
      const sumRange = a[2] != null ? flat([a[2]]) : range;
      let total = 0;
      range.forEach((v, i) => { if (test(v)) total += num(sumRange[i]) || 0; });
      return total;
    },
    VLOOKUP: (a) => {
      const [needle, table, idx, exact] = a;
      if (Array.isArray(table)) {
        for (const row of table) {
          const cell0 = Array.isArray(row) ? row[0] : row;
          const match = exact === false ? str(cell0) == str(needle) : str(cell0) === str(needle);
          if (match) { const rowArr = Array.isArray(row) ? row : [row]; return rowArr[(idx || 1) - 1]; }
        }
      }
      return '#N/A';
    },
    IFERROR: (a) => { const v = a[0]; return (typeof v === 'string' && v.startsWith('#')) ? a[1] : v; },
    IF: (a) => (a[0] ? a[1] : (a.length > 2 ? a[2] : false)),
    AND: (a) => flat(a).every(Boolean),
    OR: (a) => flat(a).some(Boolean),
    NOT: (a) => !a[0],
    CONCATENATE: (a) => flat(a).map(str).join(''),
    TEXT: (a) => { const n = num(a[0]); if (isNaN(n)) return str(a[0]); const fmt = str(a[1] || ''); if (/0\.00|#,##0\.00/.test(fmt)) return n.toFixed(2); if (/%/.test(fmt)) return (n * 100).toFixed(0) + '%'; if (/0/.test(fmt)) return String(Math.round(n)); return String(n); },
    DATE: (a) => { const d = new Date(num(a[0]), (num(a[1]) || 1) - 1, num(a[2]) || 1); return d.toLocaleDateString(); },
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
  function matchCriteria(crit) {
    if (typeof crit === 'string' && /^[<>]=?/.test(crit)) {
      const m = /^(<=|>=|<>|<|>|=)(.*)$/.exec(crit);
      const op = m[1], rhs = num(m[2]);
      return (v) => { const n = num(v); switch (op) { case '<': return n < rhs; case '>': return n > rhs; case '<=': return n <= rhs; case '>=': return n >= rhs; case '<>': return n != rhs; default: return n === rhs; } };
    }
    return (v) => str(v).toLowerCase() === str(crit).toLowerCase() || num(v) === num(crit) && !isNaN(num(v));
  }
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
        const rows = [];
        for (let r = r1; r <= r2; r++) {
          const row = [];
          for (let c = c1; c <= c2; c++) row.push(computeCell(r, c, sheet));
          rows.push(row);
        }
        return rows;
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
  let colEls = [];
  function buildGridSkeleton() {
    grid.innerHTML = '';
    const colgroup = document.createElement('colgroup');
    colgroup.appendChild(Object.assign(document.createElement('col'), { style: 'width:44px' }));
    colEls = [];
    for (let c = 0; c < COLS; c++) {
      const col = document.createElement('col');
      col.style.width = ((book && book.colWidths && book.colWidths[c]) || 96) + 'px';
      colgroup.appendChild(col);
      colEls.push(col);
    }
    grid.appendChild(colgroup);

    const thead = document.createElement('thead');
    const hr = document.createElement('tr');
    hr.appendChild(Object.assign(document.createElement('th'), { className: 'corner', textContent: '' }));
    for (let c = 0; c < COLS; c++) {
      const th = document.createElement('th');
      th.textContent = colName(c);
      th.style.position = 'relative';
      const resizer = document.createElement('div');
      resizer.className = 'col-resizer';
      resizer.addEventListener('mousedown', (e) => onColResizeStart(e, c));
      th.appendChild(resizer);
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
        td.addEventListener('mouseenter', onCellMouseEnter);
        td.addEventListener('dblclick', () => startEdit());
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    grid.appendChild(tbody);
  }

  let resizingCol = null;
  function onColResizeStart(e, c) {
    e.preventDefault(); e.stopPropagation();
    snapshotUndo();
    const startX = e.clientX;
    const startW = colEls[c].getBoundingClientRect().width;
    resizingCol = c;
    function onMove(ev) {
      const w = Math.max(30, startW + (ev.clientX - startX));
      colEls[c].style.width = w + 'px';
    }
    function onUp() {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      book.colWidths = book.colWidths || {};
      book.colWidths[c] = parseInt(colEls[c].style.width, 10);
      resizingCol = null;
      markDirty();
    }
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
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
    updateCellRefLabel();
    document.getElementById('btnFreeze').classList.toggle('active', !!(sheet.freeze && sheet.freeze.row));
    (book.colWidths ? Object.keys(book.colWidths) : []).forEach(c => { if (colEls[c]) colEls[c].style.width = book.colWidths[c] + 'px'; });
  }

  function tdAt(r, c) { return grid.tBodies[0].rows[r].cells[c + 1]; }

  function renderCell(r, c, sheet) {
    const td = tdAt(r, c);
    const merge = mergeAt(r, c, sheet);
    const isCovered = merge && (merge.r !== r || merge.c !== c);
    const cell = isCovered ? null : sheet.cells[ref(r, c)];
    const val = isCovered ? '' : computeCell(r, c, sheet);
    td.textContent = isCovered ? '' : formatValue(val, cell && cell.style && cell.style.fmt);
    const wasSelected = td.classList.contains('selected');
    const wasInRange = td.classList.contains('in-range');
    td.className = '';
    if (wasSelected) td.classList.add('selected');
    if (wasInRange) td.classList.add('in-range');
    if (cell && cell.style) {
      const st = cell.style;
      td.style.fontWeight = st.bold ? '700' : '400';
      td.style.fontStyle = st.italic ? 'italic' : 'normal';
      td.style.textDecoration = st.underline ? 'underline' : 'none';
      td.style.textAlign = st.align || (typeof val === 'number' ? 'right' : 'left');
      td.style.color = st.color || '';
      td.style.background = st.fill || '';
      td.style.fontFamily = st.fontFamily || '';
      if (st.border) td.classList.add('bordered');
    } else {
      td.style.fontWeight = '400'; td.style.fontStyle = 'normal'; td.style.textDecoration = 'none';
      td.style.textAlign = typeof val === 'number' ? 'right' : 'left';
      td.style.color = ''; td.style.background = ''; td.style.fontFamily = '';
    }
    // Merge visuals: hide the internal border between two cells that belong to the same merge.
    td.style.borderTopColor = ''; td.style.borderRightColor = ''; td.style.borderBottomColor = ''; td.style.borderLeftColor = '';
    if (merge) {
      const bg = (cell && cell.style && cell.style.fill) || (sheet.cells[ref(merge.r, merge.c)] && sheet.cells[ref(merge.r, merge.c)].style && sheet.cells[ref(merge.r, merge.c)].style.fill) || '#fff';
      if (r > merge.r) td.style.borderTopColor = bg;
      if (r < merge.r + merge.rowSpan - 1) td.style.borderBottomColor = bg;
      if (c > merge.c) td.style.borderLeftColor = bg;
      if (c < merge.c + merge.colSpan - 1) td.style.borderRightColor = bg;
      if (isCovered) td.style.background = bg;
    }
    // Frozen row/column visuals.
    td.classList.remove('frozen-row', 'frozen-col');
    if (sheet.freeze) {
      if (sheet.freeze.row && r === 0) td.classList.add('frozen-row');
      if (sheet.freeze.col && c === 0) td.classList.add('frozen-col');
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
    grid.querySelectorAll('td.in-range').forEach(td => td.classList.remove('in-range'));
    const { r1, r2, c1, c2 } = rangeBounds();
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) { const td = tdAt(r, c); if (td) td.classList.add('in-range'); }
    const td = tdAt(sel.r, sel.c);
    if (td) { td.classList.add('selected'); if (td.scrollIntoView) td.scrollIntoView({ block: 'nearest', inline: 'nearest' }); }
  }

  function currentRaw() {
    const cell = getCell(sel.r, sel.c);
    if (!cell) return '';
    return cell.f != null ? cell.f : (cell.v ?? '');
  }

  // If (r,c) is covered by a merge, jump to that merge's anchor cell instead.
  function resolveTarget(r, c) {
    const m = mergeAt(r, c, activeSheet());
    return m ? { r: m.r, c: m.c } : { r, c };
  }

  /* ---------------- selection / editing ---------------- */
  function onCellMouseDown(e) {
    if (editing) commitEdit();
    const t = resolveTarget(+e.currentTarget.dataset.r, +e.currentTarget.dataset.c);
    sel = { r: t.r, c: t.c };
    selAnchor = e.shiftKey ? selAnchor : { r: t.r, c: t.c };
    isSelecting = true;
    highlightSelection();
    formulaInput.value = currentRaw();
    updateCellRefLabel();
    grid.focus();
  }
  function onCellMouseEnter(e) {
    if (!isSelecting || editing) return;
    sel = { r: +e.currentTarget.dataset.r, c: +e.currentTarget.dataset.c };
    highlightSelection();
    updateCellRefLabel();
  }
  window.addEventListener('mouseup', () => { isSelecting = false; });

  function updateCellRefLabel() {
    cellRefEl.textContent = isRangeSelection() ? `${ref(selAnchor.r, selAnchor.c)}:${ref(sel.r, sel.c)}` : ref(sel.r, sel.c);
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
    snapshotUndo();
    setCellRaw(sel.r, sel.c, raw);
    markDirty();
    renderDependents();
    if (moveDelta) moveSelection(moveDelta[0], moveDelta[1]);
    else { renderCell(sel.r, sel.c, activeSheet()); highlightSelection(); }
    formulaInput.value = currentRaw();
    updateCellRefLabel();
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

  function moveSelection(dr, dc, extend) {
    const nr = Math.min(ROWS - 1, Math.max(0, sel.r + dr));
    const nc = Math.min(COLS - 1, Math.max(0, sel.c + dc));
    sel = { r: nr, c: nc };
    if (!extend) selAnchor = { r: nr, c: nc };
    highlightSelection();
    formulaInput.value = currentRaw();
    updateCellRefLabel();
  }

  /* ---------------- copy / cut / paste ---------------- */
  function copyRange(cut) {
    const { r1, r2, c1, c2 } = rangeBounds();
    const rows = [];
    for (let r = r1; r <= r2; r++) {
      const row = [];
      for (let c = c1; c <= c2; c++) {
        const cell = activeSheet().cells[ref(r, c)];
        row.push(cell ? JSON.parse(JSON.stringify(cell)) : null);
      }
      rows.push(row);
    }
    clipboard = { rows, cut, r1, c1, r2, c2 };
    if (cut) {
      grid.querySelectorAll('td.marching').forEach(td => td.classList.remove('marching'));
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) { const td = tdAt(r, c); if (td) td.classList.add('marching'); }
    }
  }
  function pasteRange() {
    if (!clipboard) return;
    snapshotUndo();
    const sheet = activeSheet();
    const destR = sel.r, destC = sel.c;
    clipboard.rows.forEach((row, ri) => {
      row.forEach((cellData, ci) => {
        const r = destR + ri, c = destC + ci;
        if (r >= ROWS || c >= COLS) return;
        const key = ref(r, c);
        if (cellData) sheet.cells[key] = JSON.parse(JSON.stringify(cellData));
        else delete sheet.cells[key];
      });
    });
    if (clipboard.cut) {
      for (let r = clipboard.r1; r <= clipboard.r2; r++) for (let c = clipboard.c1; c <= clipboard.c2; c++) {
        if (r >= destR && r <= destR + (clipboard.r2 - clipboard.r1) && c >= destC && c <= destC + (clipboard.c2 - clipboard.c1)) continue;
        delete sheet.cells[ref(r, c)];
      }
      grid.querySelectorAll('td.marching').forEach(td => td.classList.remove('marching'));
      clipboard = null;
    }
    markDirty(); renderDependents(); highlightSelection();
    formulaInput.value = currentRaw();
  }
  function clearRange() {
    snapshotUndo();
    const { r1, r2, c1, c2 } = rangeBounds();
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) setCellRaw(r, c, '');
    markDirty(); renderDependents(); highlightSelection();
    formulaInput.value = '';
  }

  grid.addEventListener('keydown', (e) => {
    const mod = e.ctrlKey || e.metaKey;
    if (editing) {
      if (e.key === 'Enter') { e.preventDefault(); commitEdit([1, 0]); }
      else if (e.key === 'Tab') { e.preventDefault(); commitEdit([0, 1]); }
      else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
      return;
    }
    if (mod && e.key.toLowerCase() === 'z') { e.preventDefault(); doUndo(); return; }
    if (mod && (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z'))) { e.preventDefault(); doRedo(); return; }
    if (mod && e.key.toLowerCase() === 'c') { e.preventDefault(); copyRange(false); return; }
    if (mod && e.key.toLowerCase() === 'x') { e.preventDefault(); copyRange(true); return; }
    if (mod && e.key.toLowerCase() === 'v') { e.preventDefault(); pasteRange(); return; }
    if (mod && e.key.toLowerCase() === 'f') { e.preventDefault(); openFindBar(); return; }
    if (e.key === 'ArrowUp') { e.preventDefault(); moveSelection(-1, 0, e.shiftKey); }
    else if (e.key === 'ArrowDown') { e.preventDefault(); moveSelection(1, 0, e.shiftKey); }
    else if (e.key === 'ArrowLeft') { e.preventDefault(); moveSelection(0, -1, e.shiftKey); }
    else if (e.key === 'ArrowRight') { e.preventDefault(); moveSelection(0, 1, e.shiftKey); }
    else if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); startEdit(); }
    else if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); clearRange(); }
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

  /* ---------------- undo / redo ---------------- */
  function doUndo() {
    const prev = undoMgr.undo(book);
    if (!prev) return;
    book = prev;
    book.active = Math.min(book.active, book.sheets.length - 1);
    sel = { r: Math.min(sel.r, ROWS - 1), c: Math.min(sel.c, COLS - 1) };
    selAnchor = { ...sel };
    renderAll(); markDirty();
  }
  function doRedo() {
    const next = undoMgr.redo(book);
    if (!next) return;
    book = next;
    book.active = Math.min(book.active, book.sheets.length - 1);
    renderAll(); markDirty();
  }
  document.getElementById('btnUndo').addEventListener('click', doUndo);
  document.getElementById('btnRedo').addEventListener('click', doRedo);

  /* ---------------- merge / freeze ---------------- */
  document.getElementById('btnMerge').addEventListener('click', () => {
    const sheet = activeSheet();
    const { r1, r2, c1, c2 } = rangeBounds();
    const existing = mergeAt(sel.r, sel.c, sheet);
    snapshotUndo();
    sheet.merges = sheet.merges || [];
    if (existing && existing.r === r1 && existing.c === c1 && existing.rowSpan === (r2 - r1 + 1) && existing.colSpan === (c2 - c1 + 1)) {
      sheet.merges = sheet.merges.filter(m => m !== existing);
    } else if (r2 > r1 || c2 > c1) {
      sheet.merges = sheet.merges.filter(m => !(m.r <= r2 && m.r + m.rowSpan - 1 >= r1 && m.c <= c2 && m.c + m.colSpan - 1 >= c1));
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) delete sheet.cells[ref(r, c)];
      sheet.merges.push({ r: r1, c: c1, rowSpan: r2 - r1 + 1, colSpan: c2 - c1 + 1 });
      sel = { r: r1, c: c1 }; selAnchor = { r: r1, c: c1 };
    }
    markDirty(); renderDependents(); highlightSelection();
  });
  document.getElementById('btnFreeze').addEventListener('click', () => {
    const sheet = activeSheet();
    snapshotUndo();
    const on = !(sheet.freeze && sheet.freeze.row);
    sheet.freeze = { row: on, col: on };
    document.getElementById('btnFreeze').classList.toggle('active', on);
    markDirty(); renderDependents();
  });

  /* ---------------- toolbar ---------------- */
  function applyStyle(mutator) {
    snapshotUndo();
    const sheet = activeSheet();
    const { r1, r2, c1, c2 } = rangeBounds();
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const key = ref(r, c);
        const cell = sheet.cells[key] || (sheet.cells[key] = { f: null, v: '', style: {} });
        cell.style = cell.style || {};
        mutator(cell.style);
        renderCell(r, c, sheet);
      }
    }
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
    snapshotUndo();
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
    snapshotUndo();
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
    snapshotUndo();
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
      tab.addEventListener('click', () => { book.active = idx; sel = { r: 0, c: 0 }; selAnchor = { r: 0, c: 0 }; renderAll(); });
      tab.addEventListener('dblclick', () => {
        const n = prompt('Sheet name:', s.name);
        if (n) { snapshotUndo(); s.name = n; markDirty(); renderTabs(); }
      });
      tabsEl.appendChild(tab);
    });
  }
  document.getElementById('btnAddSheet').addEventListener('click', () => {
    snapshotUndo();
    book.sheets.push({ name: 'Sheet' + (book.sheets.length + 1), cells: {}, merges: [], freeze: { row: false, col: false } });
    book.active = book.sheets.length - 1;
    sel = { r: 0, c: 0 }; selAnchor = { r: 0, c: 0 };
    markDirty(); renderAll();
  });
  function deleteSheet(idx) {
    if (!confirm(`Delete "${book.sheets[idx].name}"?`)) return;
    snapshotUndo();
    book.sheets.splice(idx, 1);
    book.active = Math.max(0, book.active - (idx <= book.active ? 1 : 0));
    markDirty(); renderAll();
  }

  /* ---------------- save / load / multi-workbook ---------------- */
  const bookList = document.getElementById('bookList');
  let allBooks = [];

  const setStatus = (t) => saveStatus.textContent = t;
  const saveNow = debounce(async () => {
    if (!book) return;
    book.title = bookTitle.value.trim() || 'Untitled workbook';
    book.updatedAt = Date.now();
    await db.put(book);
    setStatus('Saved');
    refreshBookList();
  }, 500);
  function markDirty() { setStatus('Saving…'); saveNow(); }
  bookTitle.addEventListener('input', markDirty);

  function newBook() {
    return { id: uid(), title: 'Untitled workbook', sheets: [{ name: 'Sheet1', cells: {}, merges: [], freeze: { row: false, col: false } }], active: 0, colWidths: {}, updatedAt: Date.now() };
  }

  async function refreshBookList() {
    allBooks = (await db.getAll()).sort((a, b) => b.updatedAt - a.updatedAt);
    bookList.innerHTML = '';
    allBooks.forEach(b => {
      const item = document.createElement('div');
      item.className = 'doc-item' + (book && b.id === book.id ? ' active' : '');
      item.innerHTML = `<span class="name"></span><span class="meta"><span class="when"></span><button class="del" title="Delete">Delete</button></span>`;
      item.querySelector('.name').textContent = b.title || 'Untitled workbook';
      item.querySelector('.when').textContent = formatTime(b.updatedAt);
      item.addEventListener('click', (e) => {
        if (e.target.closest('.del')) return;
        loadBook(b.id);
      });
      item.querySelector('.del').addEventListener('click', async (e) => {
        e.stopPropagation();
        if (!confirm(`Delete "${b.title}"? This can't be undone.`)) return;
        await db.delete(b.id);
        if (book && b.id === book.id) {
          const rest = allBooks.filter(x => x.id !== b.id);
          if (rest.length) loadBook(rest[0].id); else createBook();
        } else {
          refreshBookList();
        }
      });
      bookList.appendChild(item);
    });
  }

  async function loadBook(id) {
    const rec = await db.get(id);
    if (!rec) return;
    book = rec;
    undoMgr.clear();
    sel = { r: 0, c: 0 }; selAnchor = { r: 0, c: 0 };
    bookTitle.value = book.title;
    renderAll();
    refreshBookList();
    setStatus('Saved');
    history.replaceState(null, '', `sheet.html?doc=${encodeURIComponent(id)}`);
  }

  async function createBook() {
    book = newBook();
    await db.put(book);
    await loadBook(book.id);
  }

  async function deleteCurrentBook() {
    if (!book) return;
    if (!confirm('Delete this workbook? This can\'t be undone.')) return;
    await db.delete(book.id);
    const rest = (await db.getAll()).sort((a, b) => b.updatedAt - a.updatedAt);
    if (rest.length) loadBook(rest[0].id); else createBook();
  }

  document.getElementById('btnNewBook').addEventListener('click', createBook);

  async function boot() {
    buildGridSkeleton();
    const params = new URLSearchParams(location.search);
    const docId = params.get('doc');
    const wantNew = params.has('new');

    allBooks = await db.getAll();
    if (docId && allBooks.some(b => b.id === docId)) {
      book = allBooks.find(b => b.id === docId);
    } else if (wantNew || allBooks.length === 0) {
      book = newBook();
      await db.put(book);
    } else {
      allBooks.sort((a, b) => b.updatedAt - a.updatedAt);
      book = allBooks[0];
    }
    bookTitle.value = book.title;
    renderAll();
    refreshBookList();
  }
  boot();

  /* ---------------- find & replace ---------------- */
  const findBar = document.getElementById('findBar');
  const findInput = document.getElementById('findInput');
  const replaceInput = document.getElementById('replaceInput');
  const findCount = document.getElementById('findCount');
  let findMatches = [];
  let findIdx = -1;

  function openFindBar() {
    findBar.hidden = false;
    findInput.focus(); findInput.select();
    runFind();
  }
  function closeFindBar() {
    findBar.hidden = true;
    findMatches = []; findIdx = -1; findCount.textContent = '';
    grid.focus();
  }
  function runFind() {
    const q = findInput.value.toLowerCase();
    findMatches = [];
    if (q) {
      const sheet = activeSheet();
      for (let r = 0; r < ROWS; r++) for (let c = 0; c < COLS; c++) {
        const val = formatValue(computeCell(r, c, sheet), null);
        if (String(val).toLowerCase().includes(q)) findMatches.push({ r, c });
      }
    }
    findIdx = findMatches.length ? 0 : -1;
    updateFindStatus();
    if (findIdx >= 0) goToMatch();
  }
  function updateFindStatus() {
    findCount.textContent = findMatches.length ? `${findIdx + 1} of ${findMatches.length}` : (findInput.value ? 'No matches' : '');
  }
  function goToMatch() {
    if (findIdx < 0 || findIdx >= findMatches.length) return;
    const m = findMatches[findIdx];
    sel = { r: m.r, c: m.c }; selAnchor = { r: m.r, c: m.c };
    highlightSelection(); formulaInput.value = currentRaw(); updateCellRefLabel();
  }
  document.getElementById('btnFind').addEventListener('click', openFindBar);
  document.getElementById('btnFindClose').addEventListener('click', closeFindBar);
  findInput.addEventListener('input', runFind);
  document.getElementById('btnFindNext').addEventListener('click', () => {
    if (!findMatches.length) return;
    findIdx = (findIdx + 1) % findMatches.length;
    updateFindStatus(); goToMatch();
  });
  document.getElementById('btnReplace').addEventListener('click', () => {
    if (findIdx < 0) return;
    snapshotUndo();
    const m = findMatches[findIdx];
    setCellRaw(m.r, m.c, replaceInput.value);
    markDirty(); renderDependents();
    runFind();
  });
  document.getElementById('btnReplaceAll').addEventListener('click', () => {
    if (!findMatches.length) return;
    snapshotUndo();
    findMatches.forEach(m => setCellRaw(m.r, m.c, replaceInput.value));
    markDirty(); renderDependents();
    runFind();
  });
  findInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); document.getElementById('btnFindNext').click(); }
    else if (e.key === 'Escape') { e.preventDefault(); closeFindBar(); }
  });

  /* ---------------- chart ---------------- */
  const chartModal = document.getElementById('chartModal');
  const chartSvgWrap = document.getElementById('chartSvgWrap');
  let chartType = 'bar';

  function gatherChartData() {
    const sheet = activeSheet();
    const { r1, r2, c1, c2 } = rangeBounds();
    const labels = [];
    const series = [];
    for (let r = r1; r <= r2; r++) {
      const label = formatValue(computeCell(r, c1, sheet), null) || `Row ${r + 1}`;
      const value = c2 > c1 ? num(computeCell(r, c2, sheet)) : num(computeCell(r, c1, sheet));
      labels.push(String(label));
      series.push(isNaN(value) ? 0 : value);
    }
    return { labels, series };
  }

  function renderChart() {
    const { labels, series } = gatherChartData();
    const W = 640, H = 360, PAD = 44;
    const max = Math.max(...series, 0.0001);
    const min = Math.min(...series, 0);
    let svg = `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" font-family="sans-serif">`;
    svg += `<line x1="${PAD}" y1="${H - PAD}" x2="${W - 10}" y2="${H - PAD}" stroke="#c9c4b6"/>`;
    svg += `<line x1="${PAD}" y1="10" x2="${PAD}" y2="${H - PAD}" stroke="#c9c4b6"/>`;
    const plotW = W - PAD - 20, plotH = H - PAD - 20;
    if (chartType === 'bar') {
      const bw = plotW / series.length * 0.6;
      const gap = plotW / series.length;
      series.forEach((v, i) => {
        const h = (v - Math.min(min, 0)) / (max - Math.min(min, 0)) * plotH;
        const x = PAD + i * gap + gap * 0.2;
        const y = H - PAD - h;
        svg += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bw.toFixed(1)}" height="${h.toFixed(1)}" fill="#1e8e5a" rx="2"/>`;
        svg += `<text x="${(x + bw / 2).toFixed(1)}" y="${H - PAD + 16}" font-size="10" fill="#5b5d6b" text-anchor="middle">${escapeXml(labels[i].slice(0, 8))}</text>`;
      });
    } else if (chartType === 'line') {
      const gap = plotW / Math.max(1, series.length - 1);
      const pts = series.map((v, i) => {
        const x = PAD + i * gap;
        const y = H - PAD - ((v - Math.min(min, 0)) / (max - Math.min(min, 0))) * plotH;
        return [x, y];
      });
      svg += `<polyline points="${pts.map(p => p.map(n => n.toFixed(1)).join(',')).join(' ')}" fill="none" stroke="#1e8e5a" stroke-width="2.5"/>`;
      pts.forEach(([x, y], i) => {
        svg += `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="3.5" fill="#1e8e5a"/>`;
        svg += `<text x="${x.toFixed(1)}" y="${H - PAD + 16}" font-size="10" fill="#5b5d6b" text-anchor="middle">${escapeXml(labels[i].slice(0, 8))}</text>`;
      });
    } else if (chartType === 'pie') {
      const total = series.reduce((s, v) => s + Math.abs(v), 0) || 1;
      const cx = W / 2, cy = H / 2, r = Math.min(W, H) / 2 - 50;
      const colors = ['#1e8e5a', '#3457d5', '#d65d3a', '#c9a13b', '#7a5cd6', '#2aa9a0', '#c0392b', '#4b5563'];
      let angle = -Math.PI / 2;
      series.forEach((v, i) => {
        const frac = Math.abs(v) / total;
        const next = angle + frac * Math.PI * 2;
        const x1 = cx + r * Math.cos(angle), y1 = cy + r * Math.sin(angle);
        const x2 = cx + r * Math.cos(next), y2 = cy + r * Math.sin(next);
        const large = frac > 0.5 ? 1 : 0;
        svg += `<path d="M${cx},${cy} L${x1.toFixed(1)},${y1.toFixed(1)} A${r},${r} 0 ${large} 1 ${x2.toFixed(1)},${y2.toFixed(1)} Z" fill="${colors[i % colors.length]}"/>`;
        angle = next;
      });
    }
    svg += '</svg>';
    chartSvgWrap.innerHTML = svg;
  }
  function escapeXml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

  document.getElementById('btnChart').addEventListener('click', () => {
    chartModal.hidden = false;
    renderChart();
  });
  document.getElementById('btnChartClose').addEventListener('click', () => { chartModal.hidden = true; });
  document.querySelectorAll('[data-chart-type]').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('[data-chart-type]').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      chartType = btn.dataset.chartType;
      renderChart();
    });
  });

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
        if (s.merges && s.merges.length) {
          ws['!merges'] = s.merges.map(m => ({ s: { r: m.r, c: m.c }, e: { r: m.r + m.rowSpan - 1, c: m.c + m.colSpan - 1 } }));
        }
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
      const merges = (ws['!merges'] || [])
        .filter(m => m.e.r < ROWS && m.e.c < COLS)
        .map(m => ({ r: m.s.r, c: m.s.c, rowSpan: m.e.r - m.s.r + 1, colSpan: m.e.c - m.s.c + 1 }));
      return { name, cells, merges, freeze: { row: false, col: false } };
    });
    book = { id: uid(), title: file.name.replace(/\.[^.]+$/, ''), sheets: newSheets.length ? newSheets : [{ name: 'Sheet1', cells: {}, merges: [], freeze: { row: false, col: false } }], active: 0, colWidths: {}, updatedAt: Date.now() };
    undoMgr.clear();
    await db.put(book);
    bookTitle.value = book.title;
    sel = { r: 0, c: 0 }; selAnchor = { r: 0, c: 0 };
    renderAll();
    refreshBookList();
    history.replaceState(null, '', `sheet.html?doc=${encodeURIComponent(book.id)}`);
    e.target.value = '';
  });

  /* ---------------- font family ---------------- */
  const FONTS = [
    'Arial', 'Helvetica', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Segoe UI', 'Calibri', 'Century Gothic',
    'Times New Roman', 'Georgia', 'Garamond', 'Palatino Linotype', 'Cambria',
    'Courier New', 'Consolas', 'Lucida Console',
  ];
  const fontFamilySelect = document.getElementById('fontFamily');
  FONTS.forEach(f => {
    const opt = document.createElement('option');
    opt.value = /\s/.test(f) ? `'${f}', sans-serif` : `${f}, sans-serif`;
    opt.textContent = f;
    if (f === 'Arial') opt.selected = true;
    fontFamilySelect.appendChild(opt);
  });
  fontFamilySelect.addEventListener('change', (e) => applyStyle(st => { st.fontFamily = e.target.value; }));

  /* ---------------- sidebar toggle ---------------- */
  const sidebar = document.getElementById('sidebar');
  function toggleSidebar() { sidebar.classList.toggle('collapsed'); }
  document.getElementById('toggleSidebar').addEventListener('click', toggleSidebar);

  /* ---------------- menu bar wiring ---------------- */
  CloudCommon.bindMenuAction('mNew', createBook);
  CloudCommon.bindMenuAction('mOpenTrigger', () => document.getElementById('fileOpen').click());
  CloudCommon.bindMenuAction('mExportXlsx', () => document.querySelector('[data-export="xlsx"]').click());
  CloudCommon.bindMenuAction('mExportCsv', () => document.querySelector('[data-export="csv"]').click());
  CloudCommon.bindMenuAction('mPrint', () => window.print());
  CloudCommon.bindMenuAction('mDelete', deleteCurrentBook);
  CloudCommon.bindMenuAction('mClose', () => { location.href = 'index.html'; });
  CloudCommon.bindMenuAction('mUndo', doUndo);
  CloudCommon.bindMenuAction('mRedo', doRedo);
  CloudCommon.bindMenuAction('mCopy', () => copyRange(false));
  CloudCommon.bindMenuAction('mCut', () => copyRange(true));
  CloudCommon.bindMenuAction('mPaste', pasteRange);
  CloudCommon.bindMenuAction('mDeleteRange', clearRange);
  CloudCommon.bindMenuAction('mFind', openFindBar);
  CloudCommon.bindMenuAction('mInsertRow', () => document.getElementById('btnInsertRow').click());
  CloudCommon.bindMenuAction('mInsertCol', () => document.getElementById('btnInsertCol').click());
  CloudCommon.bindMenuAction('mInsertChart', () => document.getElementById('btnChart').click());
  CloudCommon.bindMenuAction('mInsertSheet', () => document.getElementById('btnAddSheet').click());
  CloudCommon.bindMenuAction('mBold', () => applyStyle(st => { st.bold = !st.bold; }));
  CloudCommon.bindMenuAction('mItalic', () => applyStyle(st => { st.italic = !st.italic; }));
  CloudCommon.bindMenuAction('mUnderline', () => applyStyle(st => { st.underline = !st.underline; }));
  document.querySelectorAll('[data-fmt]').forEach(btn => {
    CloudCommon.bindMenuAction(btn.id || (btn.id = 'mFmt' + btn.dataset.fmt), () => applyStyle(st => { st.fmt = btn.dataset.fmt; }));
  });
  CloudCommon.bindMenuAction('mMerge', () => document.getElementById('btnMerge').click());
  CloudCommon.bindMenuAction('mFreeze', () => document.getElementById('btnFreeze').click());
  CloudCommon.bindMenuAction('mSortAsc', () => sortColumn(1));
  CloudCommon.bindMenuAction('mSortDesc', () => sortColumn(-1));
  CloudCommon.bindMenuAction('mDeleteRow', () => document.getElementById('btnDeleteRow').click());
  CloudCommon.bindMenuAction('mDeleteCol', () => document.getElementById('btnDeleteCol').click());
  CloudCommon.bindMenuAction('mToggleSidebar', toggleSidebar);
  CloudCommon.bindMenuAction('mHelpCenter', () => { location.href = 'help.html#sheet'; });
  CloudCommon.bindMenuAction('mAbout', () => { location.href = 'about.html'; });
  document.getElementById('btnHelp').addEventListener('click', () => { location.href = 'help.html#sheet'; });
  document.getElementById('btnSettings').addEventListener('click', () => CloudCommon.openModal('settingsModal'));

  document.addEventListener('DOMContentLoaded', () => {});
  CloudCommon.initTheme();
  CloudCommon.initMenuBar();
  CloudCommon.initModalDismiss();

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
  }
})();
