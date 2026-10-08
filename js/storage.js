// שמירת תרגילים: Google Apps Script אם הוגדר, אחרת בדפדפן בלבד.
(function () {
  const LOCAL_KEY = 'sim-builder-saved';
  const endpoint = () => (window.SITE_CONFIG.SAVE_ENDPOINT || '').trim();

  function readLocal() {
    try { return JSON.parse(localStorage.getItem(LOCAL_KEY) || '[]'); } catch (e) { return []; }
  }
  function writeLocal(items) {
    try { localStorage.setItem(LOCAL_KEY, JSON.stringify(items)); } catch (e) { /* אין אחסון */ }
  }

  async function list() {
    if (!endpoint()) return { shared: false, items: readLocal() };
    const res = await fetch(endpoint() + '?action=list', { method: 'GET' });
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || 'שגיאה בטעינה');
    return { shared: true, items: json.items || [] };
  }

  async function save(entry) {
    entry.id = entry.id || ('ex-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6));
    entry.createdAt = entry.createdAt || new Date().toISOString();
    if (!endpoint()) {
      const items = readLocal();
      items.unshift(entry);
      writeLocal(items);
      return { shared: false, entry };
    }
    // text/plain כדי להימנע מבקשת preflight מול Apps Script
    const res = await fetch(endpoint(), { method: 'POST', body: JSON.stringify({ action: 'save', key: '2221', entry }) });
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || 'שגיאה בשמירה');
    return { shared: true, entry };
  }

  window.STORE = { list, save, isShared: () => !!endpoint() };
})();
