(function () {
  const P = window.PROMPTS;
  const PERSONAS = window.PERSONAS;
  const FIELDS = window.PERSONA_FIELDS;
  const DRAFT_KEY = 'sim-builder-draft';
  const MAX_STUDENTS = 5;

  const $ = (s, el) => (el || document).querySelector(s);
  const $$ = (s, el) => Array.from((el || document).querySelectorAll(s));
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const clone = o => JSON.parse(JSON.stringify(o));
  const personaById = id => PERSONAS.find(p => p.id === id);

  // ---------------------------------------------------------------- state
  function applyPersona(st, id, i) {
    st.personaId = id;
    const p = personaById(id);
    if (!st.nameTouched) st.name = p ? p.name : '';
    if (!st.labelTouched) st.label = p ? p.label : 'STUDENT_' + (i + 1);
    if (!st.gTouched) st.g = p ? p.g : (st.g || 'f');
    st.persona = {};
    FIELDS.forEach(([k]) => { st.persona[k] = p ? P.gender(p[k], st.g) : ''; });
    st.personaEdited = false;
    return st;
  }
  function newStudent(i) {
    return applyPersona({ name: '', label: '', g: 'f', error: '', opening: '', nameTouched: false, labelTouched: false, gTouched: false }, PERSONAS[i % PERSONAS.length].id, i);
  }
  function newState() {
    return {
      settings: { mode: 'manual', teacher: 'f', count: 1 },
      exercise: { title: '', grade: '', learned: '', task: '', figure: '', notYet: '', centralIdea: '', solution: '', pitfalls: '' },
      stages: [{ name: '', goal: '' }, { name: '', goal: '' }, { name: '', goal: '' }],
      students: Array.from({ length: MAX_STUDENTS }, (_, i) => newStudent(i))
    };
  }
  function normalize(s) {
    const base = newState();
    s = s || {};
    const out = {
      settings: Object.assign(base.settings, s.settings || {}),
      exercise: Object.assign(base.exercise, s.exercise || {}),
      stages: Array.isArray(s.stages) && s.stages.length ? s.stages : base.stages,
      students: base.students.map((b, i) => Object.assign(b, (s.students || [])[i] || {}))
    };
    out.settings.count = Math.min(MAX_STUDENTS, Math.max(1, Number(out.settings.count) || 1));
    return out;
  }

  let state;
  try { state = normalize(JSON.parse(localStorage.getItem(DRAFT_KEY) || 'null')); } catch (e) { state = newState(); }

  function saveDraft() { try { localStorage.setItem(DRAFT_KEY, JSON.stringify(state)); } catch (e) { /* בלי אחסון */ } }

  function getPath(obj, path) { return path.split('.').reduce((o, k) => (o == null ? o : o[k]), obj); }
  function setPath(obj, path, val) {
    const ks = path.split('.');
    const last = ks.pop();
    const tgt = ks.reduce((o, k) => o[k], obj);
    tgt[last] = val;
  }

  // ---------------------------------------------------------------- helpers
  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toast._t);
    toast._t = setTimeout(() => t.classList.remove('show'), 2600);
  }
  async function copyText(text) {
    try { await navigator.clipboard.writeText(text); }
    catch (e) {
      const ta = document.createElement('textarea');
      ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove();
    }
    toast('הועתק');
  }
  function renderPrompt(text, mark) {
    const safe = esc(text);
    if (!mark) return safe.replace(/[]/g, '');
    return safe.replace(//g, '<mark class="var">').replace(//g, '</mark>');
  }

  // ---------------------------------------------------------------- gate
  async function sha256(s) {
    const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s));
    return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
  }
  function openApp() {
    $('#gate').hidden = true;
    $('#app').hidden = false;
    route();
  }
  $('#gate-form').addEventListener('submit', async e => {
    e.preventDefault();
    const v = $('#gate-pass').value.trim();
    let ok = false;
    try { ok = (await sha256(v)) === window.SITE_CONFIG.PASS_HASH; } catch (err) { ok = v === '2221'; }
    if (ok) {
      try { sessionStorage.setItem('sim-gate', '1'); } catch (err) { /* */ }
      openApp();
    } else {
      $('#gate-err').textContent = 'הסיסמה לא נכונה. נסו שוב.';
      $('#gate-pass').select();
    }
  });

  // ---------------------------------------------------------------- a11y panel
  const A11Y_KEY = 'sim-builder-a11y';
  let a11y = { scale: 100, contrast: false, links: false, motion: false };
  try { a11y = Object.assign(a11y, JSON.parse(localStorage.getItem(A11Y_KEY) || '{}')); } catch (e) { /* */ }
  function applyA11y() {
    const html = document.documentElement;
    html.style.fontSize = a11y.scale + '%';
    html.classList.toggle('hc', a11y.contrast);
    html.classList.toggle('ul-links', a11y.links);
    html.classList.toggle('no-motion', a11y.motion);
    $$('[data-a11y="contrast"]').forEach(c => { c.checked = a11y.contrast; });
    $$('[data-a11y="links"]').forEach(c => { c.checked = a11y.links; });
    $$('[data-a11y="motion"]').forEach(c => { c.checked = a11y.motion; });
    try { localStorage.setItem(A11Y_KEY, JSON.stringify(a11y)); } catch (e) { /* */ }
  }
  $('#a11y-toggle').addEventListener('click', () => {
    const p = $('#a11y-panel');
    p.hidden = !p.hidden;
    $('#a11y-toggle').setAttribute('aria-expanded', String(!p.hidden));
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && !$('#a11y-panel').hidden) { $('#a11y-panel').hidden = true; $('#a11y-toggle').setAttribute('aria-expanded', 'false'); $('#a11y-toggle').focus(); }
  });
  $('#a11y-panel').addEventListener('click', e => {
    const b = e.target.closest('button[data-a11y]');
    if (!b) return;
    const a = b.dataset.a11y;
    if (a === 'bigger') a11y.scale = Math.min(150, a11y.scale + 10);
    if (a === 'smaller') a11y.scale = Math.max(90, a11y.scale - 10);
    if (a === 'reset') a11y.scale = 100;
    applyA11y();
  });
  $('#a11y-panel').addEventListener('change', e => {
    const a = e.target.dataset.a11y;
    if (a) { a11y[a] = e.target.checked; applyA11y(); }
  });
  applyA11y();

  // ---------------------------------------------------------------- routing
  function route() {
    const v = (location.hash || '#build').slice(1).split('/')[0];
    const view = ['build', 'saved', 'fixed'].includes(v) ? v : 'build';
    $$('.view').forEach(s => { s.hidden = s.id !== 'view-' + view; });
    $$('.main-nav a').forEach(a => {
      if (a.dataset.view === view) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
    });
    if (view === 'saved') loadSaved();
    if (view === 'fixed') renderFixed();
  }
  window.addEventListener('hashchange', route);

  // ---------------------------------------------------------------- steps
  let currentStep = 1;
  function goStep(n, focus) {
    currentStep = n;
    $$('[data-step-panel]').forEach(p => { p.hidden = Number(p.dataset.stepPanel) !== n; });
    $$('.stepper button').forEach(b => {
      if (Number(b.dataset.step) === n) b.setAttribute('aria-current', 'step'); else b.removeAttribute('aria-current');
      b.classList.toggle('done', Number(b.dataset.step) < n);
    });
    if (n === 4) renderStudents();
    if (n === 5) renderOutputs();
    refreshPeeks();
    if (focus !== false) {
      const h = $('[data-step-panel="' + n + '"] h2');
      h.setAttribute('tabindex', '-1');
      h.focus({ preventScroll: true });
      window.scrollTo({ top: $('.build-layout').offsetTop - 80, behavior: a11y.motion ? 'auto' : 'smooth' });
    }
  }
  $('.stepper').addEventListener('click', e => { const b = e.target.closest('button[data-step]'); if (b) goStep(Number(b.dataset.step)); });
  document.addEventListener('click', e => { const b = e.target.closest('[data-go]'); if (b) goStep(Number(b.dataset.go)); });

  // ---------------------------------------------------------------- binding
  function fillBound(root) {
    $$('[data-path]', root).forEach(el => {
      const v = getPath(state, el.dataset.path);
      if (el.type === 'radio') el.checked = String(v) === el.value;
      else if (el.type === 'checkbox') el.checked = !!v;
      else el.value = v == null ? '' : v;
    });
  }

  let refreshTimer;
  function scheduleRefresh() {
    clearTimeout(refreshTimer);
    refreshTimer = setTimeout(() => { refreshPeeks(); if (currentStep === 5) renderOutputs(); }, 250);
  }

  document.addEventListener('input', onBoundInput);
  document.addEventListener('change', e => { if (e.target.type === 'radio') onBoundInput(e); });

  function onBoundInput(e) {
    const el = e.target;
    if (!el.dataset || !el.dataset.path) return;
    if (el.type === 'radio' && (!el.checked || e.type === 'input')) return;
    const path = el.dataset.path;
    let val = el.type === 'checkbox' ? el.checked : el.value;
    if (el.hasAttribute('data-num')) val = Number(val);
    setPath(state, path, val);

    const m = path.match(/^students\.(\d+)\.(.+)$/);
    if (m) {
      const i = Number(m[1]);
      const st = state.students[i];
      const key = m[2];
      if (key === 'name') { st.nameTouched = true; const echo = $('[data-name-echo="' + i + '"]'); if (echo) echo.textContent = st.name || 'בלי שם'; }
      if (key === 'label') st.labelTouched = true;
      if (key === 'g') {
        st.gTouched = true;
        const p = personaById(st.personaId);
        if (p && !st.personaEdited) {
          FIELDS.forEach(([k]) => { st.persona[k] = P.gender(p[k], st.g); });
          fillBound($('#student-' + i));
        }
      }
      if (key.startsWith('persona.')) st.personaEdited = true;
    }
    if (path === 'settings.count' && currentStep === 4) renderStudents();
    saveDraft();
    scheduleRefresh();
  }

  // ---------------------------------------------------------------- stages
  function renderStages() {
    const list = $('#stages-list');
    list.innerHTML = state.stages.map((s, i) => `
      <div class="stage-row">
        <div class="stage-num" aria-hidden="true">${i + 1}</div>
        <div class="stage-fields">
          <div class="field">
            <label for="stage-name-${i}">שם שלב ${i + 1}</label>
            <input id="stage-name-${i}" type="text" data-path="stages.${i}.name" placeholder="למשל: בודקים את המקרה שבשרטוט">
          </div>
          <div class="field">
            <label for="stage-goal-${i}">יעד: מה התלמיד/ה אומר/ת בעצמו/ה</label>
            <textarea id="stage-goal-${i}" rows="2" data-path="stages.${i}.goal" placeholder="למשל: שני שטחים שונים, וגם שבמקרה הזה אורן טועה"></textarea>
          </div>
        </div>
        <div class="stage-actions">
          <button type="button" class="icon-btn" data-stage="up" data-i="${i}" aria-label="הזזת שלב ${i + 1} למעלה" ${i === 0 ? 'disabled' : ''}>↑</button>
          <button type="button" class="icon-btn" data-stage="down" data-i="${i}" aria-label="הזזת שלב ${i + 1} למטה" ${i === state.stages.length - 1 ? 'disabled' : ''}>↓</button>
          <button type="button" class="icon-btn danger" data-stage="del" data-i="${i}" aria-label="מחיקת שלב ${i + 1}" ${state.stages.length <= 1 ? 'disabled' : ''}>✕</button>
        </div>
      </div>`).join('');
    fillBound(list);
  }
  $('#add-stage').addEventListener('click', () => {
    if (state.stages.length >= 8) { toast('אפשר עד 8 שלבים'); return; }
    state.stages.push({ name: '', goal: '' });
    renderStages(); saveDraft(); scheduleRefresh();
    $('#stage-name-' + (state.stages.length - 1)).focus();
  });
  $('#stages-list').addEventListener('click', e => {
    const b = e.target.closest('[data-stage]');
    if (!b) return;
    const i = Number(b.dataset.i);
    const s = state.stages;
    if (b.dataset.stage === 'up' && i > 0) [s[i - 1], s[i]] = [s[i], s[i - 1]];
    if (b.dataset.stage === 'down' && i < s.length - 1) [s[i + 1], s[i]] = [s[i], s[i + 1]];
    if (b.dataset.stage === 'del' && s.length > 1) s.splice(i, 1);
    renderStages(); saveDraft(); scheduleRefresh();
  });

  // ---------------------------------------------------------------- students
  function renderStudents() {
    const wrap = $('#students');
    const n = state.settings.count;
    let html = '';
    for (let i = 0; i < n; i++) {
      const st = state.students[i];
      const opts = PERSONAS.map(p => `
        <label class="persona-opt">
          <input type="radio" name="persona-${i}" value="${p.id}" data-action="persona" data-i="${i}" ${st.personaId === p.id ? 'checked' : ''}>
          <span class="persona-body"><strong>${esc(p.name)}</strong><span>${esc(P.gender(p.tagline, p.g))}</span></span>
        </label>`).join('') + `
        <label class="persona-opt custom">
          <input type="radio" name="persona-${i}" value="custom" data-action="persona" data-i="${i}" ${st.personaId === 'custom' ? 'checked' : ''}>
          <span class="persona-body"><strong>בנו לבד</strong><span>שדות ריקים לתיאור דמות חדשה</span></span>
        </label>`;
      const fields = FIELDS.map(([k, label, hint]) => `
        <div class="field">
          <label for="p-${i}-${k}">${esc(label)}</label>
          <textarea id="p-${i}-${k}" rows="${k === 'reactions' ? 7 : 3}" data-path="students.${i}.persona.${k}" aria-describedby="p-${i}-${k}-h"></textarea>
          <p class="hint" id="p-${i}-${k}-h">${esc(hint)}</p>
        </div>`).join('');
      html += `
      <article class="student-card" id="student-${i}" aria-labelledby="st-h-${i}">
        <h3 id="st-h-${i}"><span class="badge">${i + 1}</span> <span data-name-echo="${i}">${esc(st.name || 'בלי שם')}</span></h3>

        <fieldset class="field">
          <legend>אישיות</legend>
          <div class="persona-grid">${opts}</div>
        </fieldset>

        <div class="grid-3">
          <div class="field">
            <label for="st-name-${i}">שם הדמות</label>
            <input id="st-name-${i}" type="text" data-path="students.${i}.name">
          </div>
          <fieldset class="field">
            <legend>לשון פנייה לדמות</legend>
            <div class="seg">
              <label><input type="radio" name="g-${i}" value="f" data-path="students.${i}.g"><span>נקבה</span></label>
              <label><input type="radio" name="g-${i}" value="m" data-path="students.${i}.g"><span>זכר</span></label>
            </div>
          </fieldset>
          <div class="field">
            <label for="st-label-${i}">תווית במנהל השיח</label>
            <input id="st-label-${i}" type="text" dir="ltr" data-path="students.${i}.label" aria-describedby="st-label-h-${i}">
            <p class="hint" id="st-label-h-${i}">המילה שמנהל השיח מחזיר כשהדמות מדברת. לשנות לפי מה שהמערכת מצפה.</p>
          </div>
        </div>

        <details class="persona-details" ${st.personaId === 'custom' ? 'open' : ''}>
          <summary>תיאור האישיות (אפשר לערוך)</summary>
          <div class="persona-fields">${fields}</div>
        </details>

        <div class="field error-field">
          <div class="label-row">
            <label for="st-error-${i}">הטעות של הדמות בתרגיל הזה *</label>
            <button type="button" class="btn btn-ghost sm" data-action="suggest" data-i="${i}">הצעת טעות מ־AI</button>
          </div>
          <textarea id="st-error-${i}" rows="4" data-path="students.${i}.error" aria-describedby="st-error-h-${i}" placeholder="מה הדמות עושה ומה היא חושבת. למשל: מחשבת את שני החלקים 65 כפול 40, כמו בסעיף 2…"></textarea>
          <p class="hint" id="st-error-h-${i}">בגוף שלישי. מה הדמות עשתה ומה היא חושבת, בלי מה יתקן את זה.</p>
        </div>

        <div class="field">
          <label for="st-open-${i}">משפט פתיחה קבוע</label>
          <input id="st-open-${i}" type="text" data-path="students.${i}.opening" aria-describedby="st-open-h-${i}">
          <p class="hint" id="st-open-h-${i}">לא חובה. אם ריק, הדמות פותחת במשפט משלה שמתאים לאישיות.</p>
        </div>
      </article>`;
    }
    wrap.innerHTML = html;
    fillBound(wrap);
  }

  $('#students').addEventListener('change', e => {
    const el = e.target;
    if (el.dataset.action !== 'persona') return;
    const i = Number(el.dataset.i);
    const st = state.students[i];
    if (st.personaEdited && !confirm('השינויים בתיאור האישיות יימחקו. להמשיך?')) {
      el.checked = false;
      const prev = $(`input[name="persona-${i}"][value="${st.personaId}"]`);
      if (prev) prev.checked = true;
      return;
    }
    applyPersona(st, el.value, i);
    saveDraft();
    renderStudents();
    const focusEl = $(`input[name="persona-${i}"][value="${el.value}"]`);
    if (focusEl) focusEl.focus();
    scheduleRefresh();
  });

  // ---------------------------------------------------------------- AI error helper
  let helperIdx = 0;
  $('#students').addEventListener('click', e => {
    const b = e.target.closest('[data-action="suggest"]');
    if (!b) return;
    helperIdx = Number(b.dataset.i);
    $('#helper-text').textContent = P.errorHelper(state, helperIdx);
    $('#helper-paste').value = '';
    $('#dlg-error').showModal();
  });
  $('#copy-helper').addEventListener('click', () => copyText($('#helper-text').textContent));
  $('#helper-apply').addEventListener('click', () => {
    const v = $('#helper-paste').value.trim();
    if (!v) { $('#helper-paste').focus(); return; }
    state.students[helperIdx].error = v.replace(/^להדבקה:\s*/, '');
    saveDraft();
    const ta = $('#st-error-' + helperIdx);
    if (ta) ta.value = state.students[helperIdx].error;
    $('#dlg-error').close();
    toast('הטעות הוכנסה');
    scheduleRefresh();
  });

  // ---------------------------------------------------------------- peeks (fixed prompts toggles)
  function peekText(kind) {
    if (kind === 'orchestrator') return P.orchestrator(state);
    if (kind === 'shared') return P.shared(state);
    if (kind === 'mentor') return P.mentor(state);
    if (kind === 'character') return P.character(state, 0);
    return '';
  }
  function refreshPeeks() {
    $$('details.peek[open]').forEach(d => { $('.prompt-view', d).innerHTML = renderPrompt(peekText(d.dataset.peek), true); });
  }
  $$('details.peek').forEach(d => d.addEventListener('toggle', () => { if (d.open) refreshPeeks(); }));

  // ---------------------------------------------------------------- outputs
  const REQUIRED = [
    ['exercise.title', 'שם התרגיל'], ['exercise.grade', 'כיתה ורמה'], ['exercise.learned', 'מה כבר למדו'],
    ['exercise.task', 'טקסט המשימה']
  ];
  function missingList() {
    const miss = REQUIRED.filter(([p]) => !String(getPath(state, p) || '').trim()).map(([, l]) => l);
    if (!state.stages.some(s => s.name.trim() && s.goal.trim())) miss.push('לפחות שלב אחד עם שם ויעד');
    for (let i = 0; i < state.settings.count; i++) {
      const st = state.students[i];
      if (!st.name.trim()) miss.push('שם לתלמיד/ה ' + (i + 1));
      if (!st.error.trim()) miss.push('הטעות של ' + (st.name || 'תלמיד/ה ' + (i + 1)));
      if (!st.label.trim()) miss.push('תווית ל' + (st.name || 'תלמיד/ה ' + (i + 1)));
    }
    const labels = state.students.slice(0, state.settings.count).map(s => s.label.trim().toUpperCase());
    if (new Set(labels).size !== labels.length) miss.push('תוויות שונות לכל תלמיד/ה (יש כפילות)');
    return miss;
  }

  let outputsCache = [];
  function renderOutputs() {
    const n = state.settings.count;
    outputsCache = [
      { title: 'מנהל השיח', slot: 'בתיבת: פרומפט מנהל שיח', text: P.orchestrator(state) },
      { title: 'המנחה', slot: 'בתיבת: פרומפט מנחה', text: P.mentor(state) },
      { title: 'הנחיות משותפות לכל התלמידים', slot: 'בתיבת: פרומפט על לכל תלמידי ה־AI', text: P.shared(state) }
    ];
    for (let i = 0; i < n; i++) {
      const st = state.students[i];
      outputsCache.push({ title: 'דמות: ' + (st.name || 'תלמיד/ה ' + (i + 1)), slot: 'בתיבת הדמות של ' + (st.name || 'תלמיד/ה ' + (i + 1)) + ' · תווית ' + (st.label || '—'), text: P.character(state, i) });
    }
    const mark = $('#mark-vars').checked;
    $('#outputs').innerHTML = outputsCache.map((o, i) => `
      <article class="out-card">
        <div class="out-head">
          <div>
            <h3>${esc(o.title)}</h3>
            <p class="slot">${esc(o.slot)} · ${P.strip(o.text).length.toLocaleString('he-IL')} תווים</p>
          </div>
          <button type="button" class="btn btn-primary" data-copy="${i}">העתקה</button>
        </div>
        <details>
          <summary>הצגת הפרומפט</summary>
          <pre class="prompt-view" tabindex="0">${renderPrompt(o.text, mark)}</pre>
        </details>
      </article>`).join('');

    const miss = missingList();
    const box = $('#missing');
    if (miss.length) {
      box.hidden = false;
      box.innerHTML = '<strong>חסרים פרטים.</strong> אפשר להעתיק כבר עכשיו, אבל במקומם יופיע ⟦שם השדה⟧: <ul>' + miss.map(m => '<li>' + esc(m) + '</li>').join('') + '</ul>';
    } else box.hidden = true;

    $('#save-scope').textContent = window.STORE.isShared()
      ? 'התרגיל יישמר בעמוד "תרגילים שנשמרו", פתוח לכל המשתמשים באתר.'
      : 'שימו לב: השמירה המשותפת עוד לא חוברה. התרגיל יישמר רק בדפדפן הזה.';
  }
  $('#outputs').addEventListener('click', e => {
    const b = e.target.closest('[data-copy]');
    if (b) copyText(P.strip(outputsCache[Number(b.dataset.copy)].text));
  });
  $('#mark-vars').addEventListener('change', renderOutputs);

  $('#reset-form').addEventListener('click', () => {
    if (!confirm('לנקות את הטופס ולהתחיל סימולציה חדשה? מה שלא נשמר יימחק.')) return;
    state = newState();
    saveDraft();
    initBuild();
    goStep(1);
  });

  // ---------------------------------------------------------------- save
  $('#open-save').addEventListener('click', () => {
    const names = state.students.slice(0, state.settings.count).map(s => s.name).filter(Boolean).join(', ');
    $('#save-name').value = [state.exercise.title, names].filter(Boolean).join(' — ');
    try { $('#save-author').value = localStorage.getItem('sim-builder-author') || ''; } catch (e) { /* */ }
    $('#save-err').textContent = '';
    $('#save-note').textContent = window.STORE.isShared()
      ? 'התרגיל יהיה פתוח לכל מי שנכנס לאתר.'
      : 'השמירה המשותפת עוד לא חוברה, ולכן התרגיל יישמר רק בדפדפן הזה.';
    $('#dlg-save').showModal();
  });
  $('#save-cancel').addEventListener('click', () => $('#dlg-save').close());
  $('#save-form').addEventListener('submit', async e => {
    e.preventDefault();
    const name = $('#save-name').value.trim();
    const author = $('#save-author').value.trim();
    if (!name || !author) { $('#save-err').textContent = 'צריך שם לתרגיל ושם של מי ששומר/ת.'; return; }
    try { localStorage.setItem('sim-builder-author', author); } catch (err) { /* */ }
    const btn = $('#save-submit');
    btn.disabled = true; btn.textContent = 'שומר…';
    try {
      const studs = state.students.slice(0, state.settings.count).map(s => s.name).filter(Boolean);
      await window.STORE.save({ name, author, students: studs.join(', '), mode: state.settings.mode, data: clone(state) });
      $('#dlg-save').close();
      toast('התרגיל נשמר');
    } catch (err) {
      $('#save-err').textContent = 'השמירה נכשלה: ' + err.message;
    } finally {
      btn.disabled = false; btn.textContent = 'שמירה';
    }
  });

  // ---------------------------------------------------------------- saved view
  let savedItems = [];
  async function loadSaved() {
    const status = $('#saved-status');
    status.textContent = 'טוען…';
    try {
      const r = await window.STORE.list();
      savedItems = r.items || [];
      status.textContent = (r.shared ? '' : 'השמירה המשותפת עוד לא חוברה, ומוצגים רק תרגילים מהדפדפן הזה. ') + savedItems.length + ' תרגילים.';
      renderSaved();
    } catch (err) {
      status.textContent = 'לא הצלחנו לטעון את התרגילים: ' + err.message;
    }
  }
  function renderSaved() {
    const q = $('#saved-search').value.trim();
    const items = savedItems.filter(it => !q || [it.name, it.author, it.students].join(' ').includes(q));
    $('#saved-list').innerHTML = items.length ? items.map(it => {
      const d = it.createdAt ? new Date(it.createdAt).toLocaleDateString('he-IL') : '';
      return `
      <article class="saved-card">
        <h3>${esc(it.name || 'בלי שם')}</h3>
        <dl>
          <div><dt>דמויות</dt><dd>${esc(it.students || '—')}</dd></div>
          <div><dt>שמר/ה</dt><dd>${esc(it.author || '—')}</dd></div>
          <div><dt>תאריך</dt><dd>${esc(d)}</dd></div>
          <div><dt>מעבר שלבים</dt><dd>${it.mode === 'auto' ? 'אוטומטי' : 'ידני'}</dd></div>
        </dl>
        <button type="button" class="btn btn-primary" data-load="${esc(it.id)}">פתיחה</button>
      </article>`;
    }).join('') : '<p class="empty">אין עדיין תרגילים שנשמרו.</p>';
  }
  $('#saved-search').addEventListener('input', renderSaved);
  $('#saved-list').addEventListener('click', e => {
    const b = e.target.closest('[data-load]');
    if (!b) return;
    const it = savedItems.find(x => String(x.id) === b.dataset.load);
    if (!it) return;
    let data = it.data;
    if (typeof data === 'string') { try { data = JSON.parse(data); } catch (err) { data = null; } }
    if (!data) { toast('לא הצלחנו לפתוח את התרגיל'); return; }
    state = normalize(data);
    saveDraft();
    initBuild();
    location.hash = '#build';
    goStep(5);
    toast('נפתח: ' + (it.name || ''));
  });

  // ---------------------------------------------------------------- fixed view
  function placeholderState(mode) {
    const s = newState();
    s.settings.mode = mode;
    s.settings.teacher = state.settings.teacher;
    s.stages = [{ name: '', goal: '' }];
    s.students[0] = { personaId: 'custom', name: '', label: '', g: 'f', error: '', opening: '', persona: {} };
    return s;
  }
  function renderFixed() {
    const mode = ($('input[name="fixed-mode"]:checked') || {}).value || 'manual';
    const s = placeholderState(mode);
    const items = [
      ['מנהל השיח', P.orchestrator(s)],
      ['המנחה', P.mentor(s)],
      ['הנחיות משותפות לכל התלמידים', P.shared(s)],
      ['פרומפט הדמות', P.character(s, 0)]
    ];
    $('#fixed-list').innerHTML = items.map(([t, txt]) => `
      <details class="peek fixed">
        <summary>${esc(t)}</summary>
        <div class="peek-legend"><mark class="var">מסומן</mark> = חלק שמשתנה לפי מה שממלאים</div>
        <pre class="prompt-view" tabindex="0">${renderPrompt(txt, true)}</pre>
      </details>`).join('');
  }
  $$('input[name="fixed-mode"]').forEach(r => r.addEventListener('change', renderFixed));

  // ---------------------------------------------------------------- init
  function initBuild() {
    fillBound($('[data-step-panel="1"]'));
    fillBound($('[data-step-panel="2"]'));
    renderStages();
    if (currentStep === 4) renderStudents();
  }
  initBuild();
  goStep(1, false);

  let gateOk = false;
  try { gateOk = sessionStorage.getItem('sim-gate') === '1'; } catch (e) { /* */ }
  if (gateOk) openApp(); else $('#gate-pass').focus();
})();
