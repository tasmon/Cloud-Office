/* Cloud Office 2.0.1 - shared UI behavior: theme, menu bar, modals */
const CloudCommon = (() => {
  const THEME_KEY = 'cloud-office-theme';

  function systemPrefersDark() {
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  }

  function applyTheme(mode) {
    const resolved = mode === 'system' ? (systemPrefersDark() ? 'dark' : 'light') : mode;
    document.documentElement.dataset.theme = resolved;
  }

  function getTheme() {
    try { return localStorage.getItem(THEME_KEY) || 'light'; } catch (e) { return 'light'; }
  }

  function setTheme(mode) {
    try { localStorage.setItem(THEME_KEY, mode); } catch (e) {}
    applyTheme(mode);
    document.querySelectorAll('[data-theme-choice]').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.themeChoice === mode);
    });
  }

  function initTheme() {
    applyTheme(getTheme());
    document.querySelectorAll('[data-theme-choice]').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.themeChoice === getTheme());
      btn.addEventListener('click', () => setTheme(btn.dataset.themeChoice));
    });
    if (window.matchMedia) {
      window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
        if (getTheme() === 'system') applyTheme('system');
      });
    }
  }

  /* ---------------- menu bar ---------------- */
  function initMenuBar() {
    const items = document.querySelectorAll('.menubar-item');
    function closeAll(except) {
      items.forEach(it => {
        if (it === except) return;
        it.classList.remove('open');
        const dd = it.querySelector('.menubar-dropdown');
        if (dd) dd.classList.remove('open');
      });
    }
    items.forEach(item => {
      const dd = item.querySelector('.menubar-dropdown');
      if (!dd) return;
      item.addEventListener('click', (e) => {
        if (e.target.closest('.menubar-dropdown')) return;
        const willOpen = !dd.classList.contains('open');
        closeAll(item);
        dd.classList.toggle('open', willOpen);
        item.classList.toggle('open', willOpen);
      });
      item.addEventListener('mouseenter', () => {
        const anyOpen = Array.from(items).some(it => it.classList.contains('open'));
        if (!anyOpen) return;
        closeAll(item);
        dd.classList.add('open');
        item.classList.add('open');
      });
    });
    document.addEventListener('click', (e) => {
      if (!e.target.closest('.menubar')) closeAll(null);
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') closeAll(null);
    });
  }

  /* Wire a menubar dropdown button to a handler and auto-close the menu after. */
  function bindMenuAction(id, handler) {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener('click', () => {
      handler();
      document.querySelectorAll('.menubar-dropdown.open').forEach(d => d.classList.remove('open'));
      document.querySelectorAll('.menubar-item.open').forEach(d => d.classList.remove('open'));
    });
  }

  /* ---------------- modals ---------------- */
  function openModal(id) {
    const el = document.getElementById(id);
    if (el) el.hidden = false;
  }
  function closeModal(id) {
    const el = document.getElementById(id);
    if (el) el.hidden = true;
  }
  function initModalDismiss() {
    document.querySelectorAll('.modal-backdrop').forEach(backdrop => {
      backdrop.addEventListener('click', (e) => { if (e.target === backdrop) backdrop.hidden = true; });
    });
    document.addEventListener('keydown', (e) => {
      if (e.key !== 'Escape') return;
      document.querySelectorAll('.modal-backdrop').forEach(b => { if (!b.hidden) b.hidden = true; });
    });
  }

  return { initTheme, setTheme, getTheme, initMenuBar, bindMenuAction, openModal, closeModal, initModalDismiss };
})();
