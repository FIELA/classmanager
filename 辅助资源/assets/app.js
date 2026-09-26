/**
 * 班级助理 · 界面
 * 数据在项目文件夹“数据/”下的记录文件里（见 store.js）；浏览器只暂存“还没写进文件夹”的修改。
 */
(function () {
  'use strict';

  const core = window.CMCore;
  const store = window.CMStore;
  const xl = window.CMExcel;
  if (!core || !store || !xl || !window.ExcelJS) {
    alert('程序文件加载失败，请检查“辅助资源/assets/”是否完整');
    return;
  }

  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const today = () => core.localISODate();
  const pad2 = n => String(n).padStart(2, '0');
  const HAS_FS = typeof window.showDirectoryPicker === 'function';
  const UA = navigator.userAgent || '';
  const PLACE = /Windows/i.test(UA) ? 'win' : (/Mac OS X|Macintosh/i.test(UA) ? 'mac' : 'web');
  const LS = { outbox: 'classmanager.v1.outbox', ui: 'classmanager.v1.ui', seen: 'classmanager.v1.exportSeen' };
  const IDB_NAME = 'classmanager-v1';
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const MODE_TEXT = { random: '随机排座', select: '按成绩选座', score: '按成绩自动排', rotate: '轮换', hold: '沿用', adjust: '人工微调' };

  const S = {
    root: null,
    pendingRoot: null,
    rootName: '',
    projectId: null,
    sid: core.randomId(8),
    seq: 0,
    clock: null,
    writer: null,
    ops: [],
    st: store.emptyState(),
    readResult: null,
    cache: new Map(),
    exportSeen: new Set(),
    notices: [],
    saving: false,
    busy: false,
    ui: { tab: 'seating', classId: null, view: 'student', version: null, mode: null, scoresWeek: null, scoresStudent: null, trajStudent: null },
    draw: { n: 1, result: [], resultClass: null, noRepeat: false, drawn: new Map() },
    scoring: null,
    adjust: null,
    select: null,
    gender: null,
    layoutDraft: null,
    itemsDraft: null,
    importCands: []
  };

  /* =========================================================
     小工具
     ========================================================= */

  function toast(msg, type, ms) {
    if (!msg) return;
    const box = $('toasts');
    const el = document.createElement('div');
    el.className = `toast ${type || ''}`;
    const span = document.createElement('span');
    span.textContent = msg;
    const x = document.createElement('button');
    x.textContent = '✕';
    x.setAttribute('aria-label', '关闭');
    const close = () => { el.style.opacity = '0'; setTimeout(() => el.remove(), 200); };
    x.onclick = close;
    el.appendChild(span);
    el.appendChild(x);
    box.appendChild(el);
    setTimeout(close, ms || (type === 'error' ? 9000 : 4500));
  }
  const showError = msg => toast(msg, 'error');

  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 300);
  }

  function loadUi() {
    try {
      const u = JSON.parse(localStorage.getItem(LS.ui) || 'null');
      if (u) {
        if (u.view === 'teacher' || u.view === 'student') S.ui.view = u.view;
        if (u.classId) S.ui.classId = u.classId;
        if (['seating', 'scores', 'student', 'settings'].indexOf(u.tab) >= 0) S.ui.tab = u.tab;
      }
    } catch (e) { /* 忽略 */ }
  }

  function saveUi() {
    try { localStorage.setItem(LS.ui, JSON.stringify({ view: S.ui.view, classId: S.ui.classId, tab: S.ui.tab })); } catch (e) { /* 忽略 */ }
  }

  function openDB() {
    return new Promise(resolve => {
      if (!window.indexedDB) return resolve(null);
      const req = indexedDB.open(IDB_NAME, 1);
      req.onupgradeneeded = e => { e.target.result.createObjectStore('handles'); };
      req.onsuccess = e => resolve(e.target.result);
      req.onerror = () => resolve(null);
    });
  }

  async function idbPut(key, value) {
    const db = await openDB();
    if (!db) return;
    try { db.transaction('handles', 'readwrite').objectStore('handles').put(value, key); } catch (e) { /* 忽略 */ }
  }

  async function idbGet(key) {
    const db = await openDB();
    if (!db) return null;
    return new Promise(resolve => {
      try {
        const req = db.transaction('handles', 'readonly').objectStore('handles').get(key);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => resolve(null);
      } catch (e) { resolve(null); }
    });
  }

  async function permitted(handle, ask) {
    try {
      if ((await handle.queryPermission({ mode: 'readwrite' })) === 'granted') return true;
      if (ask) return (await handle.requestPermission({ mode: 'readwrite' })) === 'granted';
    } catch (e) { /* 忽略 */ }
    return false;
  }

  /* =========================================================
     数据：连接、读取、提交、保存
     ========================================================= */

  function newSession() {
    S.sid = core.randomId(8);
    S.seq = 0;
    S.clock = store.createClock(S.sid);
    S.writer = new store.SessionWriter({
      root: S.root,
      sid: S.sid,
      place: PLACE,
      header: store.makeHeader(S.sid, S.clock, { place: PLACE, started: core.localDateTime(), ua: shortUA() })
    });
  }

  function shortUA() {
    const m = UA.match(/(Edg|Chrome|Firefox|Version)\/(\d+)/);
    const os = /Windows NT 6\.1/.test(UA) ? 'Win7' : /Windows NT 10/.test(UA) ? 'Win10/11' : /Mac OS X/.test(UA) ? 'macOS' : '';
    return `${m ? `${m[1] === 'Edg' ? 'Edge' : m[1]} ${m[2]}` : '浏览器'} ${os}`.trim();
  }

  const ENTRY = store.ENTRY_FILES[0];

  async function resolveRoot(handle, initial) {
    const r = await store.findProjectRoot(handle);
    if (r.root) return r.root;
    if (r.reason === 'old-project') {
      alert(`这是旧版座次项目的文件夹。请选择本项目所在的文件夹（里面有“${ENTRY}”）。`);
      return null;
    }
    if (initial) return null;
    if (r.reason === 'multiple') {
      alert(`“${handle.name}”里有不止一个项目文件夹，请直接选择要用的那一个（里面有“${ENTRY}”）。`);
      return null;
    }
    return confirm(`“${handle.name}”里没有找到本项目（没有“${ENTRY}”）。\n确定要把数据保存在这个文件夹里吗？`) ? handle : null;
  }

  async function loadData() {
    const read = await store.readProject(S.root, { cache: S.cache });
    S.readResult = read;
    S.projectId = read.project && read.project.projectId ? read.project.projectId : null;
    read.ops.forEach(o => S.clock.observe(o.t));
    S.ops = read.ops.concat(S.writer ? S.writer.ops : []);
    S.st = store.reduce(S.ops);
  }

  async function connect(handle, opts) {
    opts = opts || {};
    const root = await resolveRoot(handle, opts.initial);
    if (!root) return false;
    const sameRoot = S.root && root.isSameEntry ? await S.root.isSameEntry(root) : false;
    const carry = S.writer && !sameRoot ? S.writer.ops.slice(S.writer.written) : [];
    const oldProject = S.projectId;
    S.root = root;
    S.rootName = root.name;
    S.pendingRoot = null;
    if (!sameRoot) newSession();
    await idbPut('root', root);
    S.busy = true;
    try {
      await loadData();
    } catch (e) {
      S.busy = false;
      showError('读取文件夹失败：' + store.describeError(e));
      return false;
    }
    S.busy = false;
    if (carry.length) await adoptOps(carry, oldProject, '上一个文件夹');
    await recoverOutbox();
    ensureClass();
    renderAll();
    if (!opts.initial) toast(`已连接文件夹【${S.rootName}】`);
    setTimeout(backgroundTasks, 50);
    return true;
  }

  // 把别处未保存的修改并入本次会话（保留原编号与时间戳）
  async function adoptOps(list, project, from) {
    const missing = list.filter(o => o && o.id && !S.st.ids.has(o.id));
    if (!missing.length) return 0;
    if (project && S.projectId && project !== S.projectId &&
      !confirm(`${from}里有 ${missing.length} 条没保存成功的修改，但它们属于另一个数据文件夹。\n要写入当前文件夹吗？`)) return 0;
    missing.forEach(o => { S.clock.observe(o.t); S.writer.add(o); S.ops.push(o); });
    S.st = store.reduce(S.ops);
    saveOutbox();
    await flushNow();
    toast(`已补存 ${missing.length} 条之前没保存成功的修改`);
    return missing.length;
  }

  async function recoverOutbox() {
    let box = null;
    try { box = JSON.parse(localStorage.getItem(LS.outbox) || 'null'); } catch (e) { box = null; }
    if (!box || box.sid === S.sid || !Array.isArray(box.ops)) return;
    const n = await adoptOps(box.ops, box.project, `浏览器暂存区（${box.folder || '未知文件夹'}）`);
    if (!n) {
      const still = box.ops.filter(o => o && o.id && !S.st.ids.has(o.id));
      if (!still.length) localStorage.removeItem(LS.outbox);
    }
  }

  function saveOutbox() {
    if (!S.writer) return;
    const pending = S.writer.ops.slice(S.writer.written);
    try {
      if (!pending.length) {
        const box = JSON.parse(localStorage.getItem(LS.outbox) || 'null');
        if (!box || box.sid === S.sid) localStorage.removeItem(LS.outbox);
      } else {
        localStorage.setItem(LS.outbox, JSON.stringify({ project: S.projectId, folder: S.rootName, sid: S.sid, ops: pending }));
      }
    } catch (e) { /* 存储不可用时只能依赖文件夹 */ }
  }

  function applyNew(op) {
    S.writer.add(op);
    S.ops.push(op);
    if (!S.st.ids.has(op.id)) {
      S.st.ids.add(op.id);
      if (store.applyOp(S.st, op)) S.st.opCount++;
      if (op.t > S.st.maxT) S.st.maxT = op.t;
    }
  }

  function commit(k, payload, opts) {
    if (!S.root || !S.writer) {
      showError('还没有连接文件夹，修改无法保存');
      return null;
    }
    const op = Object.assign({ k, id: `${S.sid}.${++S.seq}`, t: S.clock.now(), s: S.sid }, payload);
    applyNew(op);
    saveOutbox();
    if (!opts || !opts.noFlush) scheduleFlush();
    return op;
  }

  let flushTimer = null;
  function scheduleFlush(delay) {
    clearTimeout(flushTimer);
    flushTimer = setTimeout(flushNow, delay == null ? 200 : delay);
    renderSaveStatus();
  }

  async function flushNow() {
    clearTimeout(flushTimer);
    flushTimer = null;
    if (!S.writer) return { ok: false };
    S.saving = true;
    renderSaveStatus();
    const res = await S.writer.flush();
    S.saving = false;
    saveOutbox();
    renderSaveStatus();
    renderAlerts();
    return res;
  }

  async function backgroundTasks() {
    if (!S.root) return;
    try {
      const w = await store.maybeWriteSnapshot(S.root, S.readResult, S.writer.fileName, S.sid, false);
      if (w && w.ok && !w.skipped) S.readResult = await store.readProject(S.root, { cache: S.cache });
    } catch (e) { /* 快照只是加速，失败不影响使用 */ }
    await scanExports();
  }

  // 检查“导出/”里被老师改过的 Excel，把改动写成记录。
  // 每个文件版本（路径 + 大小 + 修改时间）只读一次；本网页刚写出的版本直接记为已读。
  function loadSeen() {
    try { (JSON.parse(localStorage.getItem(LS.seen) || '[]') || []).forEach(k => S.exportSeen.add(k)); } catch (e) { /* 忽略 */ }
  }

  function markSeen(key) {
    S.exportSeen.add(key);
    try { localStorage.setItem(LS.seen, JSON.stringify(Array.from(S.exportSeen).slice(-500))); } catch (e) { /* 忽略 */ }
  }

  function classCtx(classId) {
    const cls = S.st.classes.get(classId);
    if (!cls || cls.deleted) return null;
    return { cls, students: store.studentsOf(cls, true), items: store.itemsOf(S.st) };
  }

  async function scanExports() {
    if (!S.root) return;
    let dir;
    try { dir = await store.getSubDir(S.root, store.EXPORT_DIR, false); } catch (e) { return; }
    if (!dir) return;
    const files = [];
    try {
      for (const e of await store.listEntries(dir)) {
        if (e.kind === 'directory') {
          for (const f of await store.listEntries(e)) {
            if (f.kind === 'file' && /\.xlsx$/i.test(f.name) && !store.isJunkName(f.name)) files.push({ handle: f, path: `${e.name}/${f.name}` });
          }
        } else if (/\.xlsx$/i.test(e.name) && !store.isJunkName(e.name)) {
          files.push({ handle: e, path: e.name });
        }
      }
    } catch (e) { return; }
    let changed = false;
    for (const f of files) {
      let file;
      try { file = await f.handle.getFile(); } catch (e) { continue; }
      const key = `${f.path}|${file.size}|${file.lastModified}`;
      if (S.exportSeen.has(key)) continue;
      markSeen(key);
      try {
        const res = await xl.readExportEdits(await file.arrayBuffer(), f.handle.name, classCtx);
        if (!res) continue;
        const fresh = res.ops.filter(o => !S.st.ids.has(o.id));
        if (fresh.length) {
          fresh.forEach(o => applyNew(Object.assign({}, o, { t: S.clock.now(), s: S.sid })));
          changed = true;
          S.notices.push({ level: 'info', text: `ℹ 已从 Excel“${f.handle.name}”读入修改：${res.parts.join('；') || `${fresh.length} 处`}` });
        }
        if (res.warnings.length) S.notices.push({ level: 'warn', text: `⚠ Excel“${f.handle.name}”：${res.warnings.join('；')}` });
      } catch (e) {
        S.notices.push({ level: 'warn', text: `⚠ 读取“${f.path}”失败：${e.message}` });
      }
    }
    if (changed) {
      saveOutbox();
      scheduleFlush();
    }
    if (changed || S.notices.length) renderAll();
  }

  let lastReturn = 0;
  async function onReturn() {
    if (!S.root || S.busy || S.ui.mode === 'adjust' || S.ui.mode === 'select') return;
    if (Date.now() - lastReturn < 2500) return;
    lastReturn = Date.now();
    if (S.writer && S.writer.pendingCount()) await flushNow();
    const before = `${S.st.maxT}|${S.st.opCount}`;
    try {
      S.busy = true;
      await loadData();
    } catch (e) {
      S.busy = false;
      renderAlerts();
      return;
    }
    S.busy = false;
    await scanExports();
    if (`${S.st.maxT}|${S.st.opCount}` !== before) renderAll();
  }

  /* =========================================================
     派生数据
     ========================================================= */

  function classes() {
    return store.listClasses(S.st);
  }

  function ensureClass() {
    const list = classes();
    if (!list.some(c => c.id === S.ui.classId)) S.ui.classId = list.length ? list[0].id : null;
  }

  function currentClass() {
    ensureClass();
    return S.ui.classId ? S.st.classes.get(S.ui.classId) : null;
  }

  function studentMaps(cls) {
    const all = store.studentsOf(cls, true);
    return { all, byId: new Map(all.map(s => [s.id, s])), names: core.displayNames(all) };
  }

  const genderClass = s => (s && s.gender === 'M' ? 'g-m' : s && s.gender === 'F' ? 'g-f' : 'g-u');

  function defaultGenDate(cls) {
    const wk = core.mondayOf(today());
    return cls && cls.finals.has(wk) ? core.addDays(wk, 7) : today();
  }

  // 当前显示的那张座次表
  function displayed(cls) {
    if (!cls) return null;
    if (S.ui.mode === 'select' && S.select) return { kind: 'select', date: S.select.date, seats: core.mapToSeats(S.select.placed), layout: S.select.layout, mode: 'select' };
    if (S.ui.mode === 'adjust' && S.adjust) return { kind: S.adjust.kind, finalId: S.adjust.finalId, date: S.adjust.date, seats: core.mapToSeats(S.adjust.work), layout: S.adjust.layout, mode: S.adjust.mode, meta: S.adjust.meta };
    const v = S.ui.version;
    if (v === 'draft' && cls.draft) return Object.assign({ kind: 'draft' }, cls.draft);
    if (v && v !== 'draft') {
      const f = store.finalsOf(cls).find(x => x.finalId === v);
      if (f) return Object.assign({ kind: 'final' }, f);
    }
    const eff = store.effectiveFinal(cls, today());
    if (eff) return Object.assign({ kind: 'final' }, eff);
    if (cls.draft) return Object.assign({ kind: 'draft' }, cls.draft);
    const all = store.finalsOf(cls);
    return all.length ? Object.assign({ kind: 'final' }, all[all.length - 1]) : null;
  }

  function clockBehind() {
    return S.clock && S.clock.maxWall() - Date.now() > 12 * 3600 * 1000;
  }

  function weekTotals(cls, week) {
    const m = new Map();
    store.scoresOf(S.st, cls.id).forEach(r => {
      if (r.week !== week) return;
      const x = m.get(r.studentId) || { plus: 0, minus: 0 };
      if (r.value > 0) x.plus += r.value; else x.minus -= r.value;
      m.set(r.studentId, x);
    });
    return m;
  }

  /* =========================================================
     座次图（学生版 / 教师版共用；教师版旋转 180°）
     ========================================================= */

  function seatFont(cols) {
    return cols <= 6 ? 20 : cols <= 8 ? 18 : cols <= 10 ? 16 : cols <= 12 ? 15 : 13;
  }

  function roomHTML(layout, seatMap, view, cellFn) {
    const L = core.normalizeLayout(layout);
    const cols = core.displayColumns(L, view);
    const tmpl = ['var(--row-label-w)'].concat(cols.map(x => (x.type === 'aisle' ? 'var(--aisle-w)' : 'minmax(0, 1fr)'))).join(' ');
    const dis = new Set(L.disabled);
    const stage = `<div class="corner"></div><div class="stage" style="grid-column: 2 / -1">讲　台</div>`;
    let html = `<div class="room ${view}" style="grid-template-columns:${tmpl}; --seat-font:${seatFont(L.cols)}px; min-width:${Math.max(560, L.cols * 76)}px">`;
    if (view !== 'teacher') html += stage;
    html += '<div></div>' + cols.map(x => (x.type === 'aisle' ? '<div></div>' : `<div class="col-head">第${x.c}列</div>`)).join('');
    for (const r of core.displayRows(L, view)) {
      const zi = core.rowZoneIndex(L.rows, r);
      html += `<div class="row-label zone-${zi}"><b>第${r}排</b><span>${core.ROW_ZONES[zi]}区</span></div>`;
      for (const x of cols) {
        if (x.type === 'aisle') { html += '<div class="aisle"></div>'; continue; }
        const k = core.seatKey(r, x.c);
        html += dis.has(k) ? '<div class="seat void"></div>' : cellFn(r, x.c, seatMap.get(k) || null);
      }
    }
    if (view === 'teacher') html += stage;
    return html + '</div>';
  }

  function seatHTML(r, c, sid, ctx) {
    const k = core.seatKey(r, c);
    if (!sid) {
      const cls = ['seat', 'empty'];
      if (ctx.mode === 'select') cls.push('pick-target');
      if (ctx.swapSel === k) cls.push('swap-selected');
      return `<div class="${cls.join(' ')}" data-r="${r}" data-c="${c}" title="${core.seatLabel(r, c)} · 空座"><span class="empty-label">${ctx.mode === 'select' ? '可选' : '空'}</span></div>`;
    }
    const s = ctx.byId.get(sid);
    const name = ctx.names.get(sid) || (s ? s.name : '（已删除）');
    const cls = ['seat', genderClass(s)];
    if (Array.from(name).length >= 4) cls.push('long');
    const drawIdx = ctx.drawn.get(sid);
    if (drawIdx !== undefined) cls.push('drawn');
    if (ctx.scoreSel && ctx.scoreSel.has(sid)) cls.push('score-selected');
    if (ctx.swapSel === k) cls.push('swap-selected');
    else if (ctx.swapSel) cls.push('swap-target');
    const pinned = ctx.pins.get(sid) === k;
    let tags = '';
    if (ctx.week) {
      const w = ctx.week.get(sid);
      if (w && w.plus) tags += `<span class="tag-plus">+${w.plus}</span>`;
      if (w && w.minus) tags += `<span class="tag-minus">-${w.minus}</span>`;
    }
    const title = `${name}${r ? ` · ${core.seatLabel(r, c)}` : ''}${pinned ? ' · 已固定' : ''}`;
    return `<div class="${cls.join(' ')}" data-r="${r}" data-c="${c}" data-sid="${esc(sid)}" title="${esc(title)}">` +
      `${drawIdx !== undefined ? `<span class="draw-no">${drawIdx + 1}</span>` : ''}${pinned ? '<span class="pin" aria-label="已固定">📌</span>' : ''}` +
      `<b class="seat-name">${esc(name)}</b>${tags ? `<span class="tags">${tags}</span>` : ''}</div>`;
  }

  /* =========================================================
     渲染：总入口
     ========================================================= */

  function renderAll() {
    $('today-tag').textContent = today();
    renderHeader();
    renderAlerts();
    const connected = !!S.root;
    $('welcome').hidden = connected;
    ['seating', 'scores', 'student', 'settings'].forEach(t => { $('panel-' + t).hidden = !connected || S.ui.tab !== t; });
    document.querySelectorAll('.tabs [data-tab]').forEach(b => b.classList.toggle('active', b.dataset.tab === S.ui.tab));
    if (!connected) { renderWelcome(); return; }
    if (S.ui.tab === 'seating') renderSeating();
    else if (S.ui.tab === 'scores') renderScores();
    else if (S.ui.tab === 'student') renderStudent();
    else if (S.ui.tab === 'settings') renderSettings();
    saveUi();
  }

  function renderHeader() {
    const sel = $('class-select');
    const list = classes();
    sel.innerHTML = list.length
      ? list.map(c => `<option value="${esc(c.id)}" ${c.id === S.ui.classId ? 'selected' : ''}>${esc(c.name)}</option>`).join('')
      : '<option value="">（还没有班级）</option>';
    sel.disabled = !list.length || !!S.ui.mode;
    const btn = $('btn-folder');
    btn.hidden = !HAS_FS;
    if (S.root) {
      btn.innerHTML = `📁 已连接：<b>${esc(S.rootName)}</b>`;
      btn.title = '已连接：修改会直接保存到这个文件夹。点击可换一个文件夹';
    } else if (S.pendingRoot) {
      btn.innerHTML = `📂 重新连接：<b>${esc(S.pendingRoot.name)}</b>`;
      btn.title = '浏览器需要再授权一次才能读写这个文件夹';
    } else {
      btn.textContent = '📂 连接文件夹';
      btn.title = `选择项目所在的文件夹（里面有“${ENTRY}”）`;
    }
    renderSaveStatus();
  }

  function renderSaveStatus() {
    const el = $('save-status');
    if (!S.root || !S.writer) { el.textContent = ''; el.className = 'save-status'; return; }
    const n = S.writer.pendingCount();
    if (n && S.writer.lastError && !S.saving) {
      el.textContent = `⚠ 未保存 ${n} 条（点此重试）`;
      el.className = 'save-status bad';
    } else if (n || S.saving) {
      el.textContent = '保存中…';
      el.className = 'save-status busy';
    } else {
      el.textContent = S.writer.written ? '✓ 已保存到文件夹' : '✓ 已连接';
      el.className = 'save-status ok';
    }
  }

  function renderAlerts() {
    const items = [];
    if (!HAS_FS) {
      items.push({ level: 'error', html: '⛔ 这个浏览器不能读写文件夹。请用 Chrome 或 Edge（版本 86 以上）打开“' + ENTRY + '”；Win7 电脑可以用 Chrome 109 或 Edge 109。' });
    }
    if (S.root && S.writer && S.writer.lastError && S.writer.pendingCount()) {
      items.push({ level: 'error', html: `⚠ 有 <b>${S.writer.pendingCount()}</b> 条修改还没保存到文件夹：${esc(S.writer.lastError)}。修改暂存在浏览器里；插好 U 盘或关闭占用文件的程序后点“重试”。`, act: 'retry', actText: '重试保存' });
    }
    if (S.root && clockBehind()) {
      items.push({ level: 'warn', html: `⚠ 这台电脑的时间（${esc(core.localDateTime())}）比已有记录还早，可能不准。记分前请核对“记分日期”。` });
    }
    const cls = S.root ? currentClass() : null;
    if (cls) {
      cls.conflicts.forEach(cf => {
        const where = cf.kind === 'final' ? `${core.weekSpan(cf.key)} 的定版` : '座位布局';
        items.push({
          level: 'warn',
          html: `⚠ ${esc(cls.name)} ${where}有两份各自保存的版本（可能是在两台电脑上分别做的），已采用较晚的一份。`,
          act: 'conflict-other', actText: '改用另一份', act2: 'conflict-keep', act2Text: '保留当前', data: `${cf.kind}|${cf.key}`
        });
      });
    }
    S.notices.forEach((n, i) => items.push({ level: n.level, html: esc(n.text), act: 'dismiss', actText: '知道了', data: String(i) }));
    if (S.readResult && S.readResult.warnings.length) items.push({ level: 'info', html: 'ℹ ' + S.readResult.warnings.map(esc).join('；') });
    $('alerts').innerHTML = items.map(it => `
      <div class="alert ${it.level}"><span>${it.html}</span>
        ${it.act ? `<button type="button" data-alert="${it.act}" data-arg="${esc(it.data || '')}">${esc(it.actText)}</button>` : ''}
        ${it.act2 ? `<button type="button" class="quiet" data-alert="${it.act2}" data-arg="${esc(it.data || '')}">${esc(it.act2Text)}</button>` : ''}
      </div>`).join('');
  }

  function renderWelcome() {
    const body = $('welcome-body');
    if (!HAS_FS) {
      body.innerHTML = '<p class="warn-text">这个浏览器不能读写文件夹。请用 Chrome 或 Edge（86 以上）打开本页面；Win7 电脑可用 Chrome 109 / Edge 109。</p>';
      $('btn-welcome-connect').hidden = true;
      return;
    }
    $('btn-welcome-connect').hidden = false;
    $('btn-welcome-connect').textContent = S.pendingRoot ? `📂 重新连接：${S.pendingRoot.name}` : '📂 连接文件夹';
    body.innerHTML = S.pendingRoot ? `<p>上次使用的是【<b>${esc(S.pendingRoot.name)}</b>】，点下面的按钮再授权一次即可继续。</p>` : '';
  }

  /* =========================================================
     座次页
     ========================================================= */

  function renderSeating() {
    const cls = currentClass();
    const rec = displayed(cls);
    renderSeatHead(cls, rec);
    renderDrawBar(cls);
    renderModeBar(cls);
    renderScoringBar(cls);
    renderLegend(cls, rec);
    renderSeatMap(cls, rec);
    renderTools(cls, rec);
  }

  function renderSeatHead(cls, rec) {
    const title = $('seat-title');
    const badges = [];
    let meta = '';
    if (!cls) {
      title.textContent = '还没有班级';
      $('seat-badges').innerHTML = '';
      $('seat-meta').innerHTML = '请到“班级设置”新建班级或导入名单。';
    } else if (!rec) {
      title.textContent = `${cls.name} · 还没有座次表`;
      $('seat-badges').innerHTML = '';
      $('seat-meta').textContent = '打开下方“排座工具”，选择方式后点“生成候选”。';
    } else {
      title.textContent = `座次表-${cls.name}-${core.weekSpan(rec.date)}`;
      if (rec.kind === 'final') badges.push('<span class="badge final">定版</span>');
      if (rec.kind === 'draft') badges.push('<span class="badge draft">候选稿</span>');
      if (rec.kind === 'select') badges.push('<span class="badge draft">选座中</span>');
      if (S.ui.mode === 'adjust') badges.push('<span class="badge edit">微调中</span>');
      const eff = store.effectiveFinal(cls, today());
      if (rec.kind === 'final' && eff && eff.finalId === rec.finalId) badges.push('<span class="badge today">今天生效</span>');
      if (rec.kind === 'final' && rec.date > today()) badges.push('<span class="badge today">尚未生效</span>');
      $('seat-badges').innerHTML = badges.join('');
      const count = rec.seats.length;
      const mode = MODE_TEXT[rec.mode] || '';
      const extra = rec.meta && rec.meta.genderPair ? '（男女搭配）' : '';
      meta = `${core.shortDate(rec.date)}起生效 · ${count} 人 · ${mode}${extra}`;
      if (rec.kind === 'draft') meta += ' · 确认定版后才生效';
      $('seat-meta').textContent = meta;
    }
    $('view-student').classList.toggle('active', S.ui.view === 'student');
    $('view-teacher').classList.toggle('active', S.ui.view === 'teacher');
    $('view-student').setAttribute('aria-pressed', String(S.ui.view === 'student'));
    $('view-teacher').setAttribute('aria-pressed', String(S.ui.view === 'teacher'));
    const sb = $('btn-scoring');
    sb.classList.toggle('on', S.ui.mode === 'scoring');
    sb.textContent = S.ui.mode === 'scoring' ? '★ 退出记分' : '★ 课堂记分';
    sb.disabled = !cls || (!!S.ui.mode && S.ui.mode !== 'scoring');
  }

  function renderDrawBar(cls) {
    const chips = [1, 2, 3, 4, 5, 6].map(n => `<button type="button" data-n="${n}" class="${S.draw.n === n && !$('draw-custom').value ? 'active' : ''}">${n}</button>`).join('');
    $('draw-chips').innerHTML = chips;
    $('draw-norepeat').checked = S.draw.noRepeat;
    const disabled = !cls || S.ui.mode === 'select' || S.ui.mode === 'adjust';
    $('btn-draw').disabled = disabled;
    let info = '';
    if (cls && S.draw.noRepeat) {
      const pool = drawPool(cls);
      const drawn = S.draw.drawn.get(cls.id) || new Set();
      const done = pool.filter(id => drawn.has(id)).length;
      info = `本轮已抽 ${done}/${pool.length} 人`;
    }
    $('draw-info').textContent = info;
    const box = $('draw-result');
    if (!cls || S.draw.resultClass !== cls.id || !S.draw.result.length) { box.hidden = true; box.innerHTML = ''; return; }
    const { byId, names } = studentMaps(cls);
    const rec = displayed(cls);
    const pos = new Map(rec ? rec.seats.map(t => [t[2], core.seatLabel(t[0], t[1])]) : []);
    box.hidden = false;
    box.innerHTML = `<span class="dr-title">抽中 ${S.draw.result.length} 人</span>` + S.draw.result.map((sid, i) => {
      const s = byId.get(sid);
      return `<span class="dr-chip ${genderClass(s)}"><i>${i + 1}</i><b class="${genderClass(s)}" style="color:${s && s.gender === 'M' ? 'var(--male)' : s && s.gender === 'F' ? 'var(--female)' : 'var(--ink)'}">${esc(names.get(sid) || '')}</b>${pos.has(sid) ? `<small>${pos.get(sid)}</small>` : ''}</span>`;
    }).join('');
  }

  function renderModeBar(cls) {
    const bar = $('mode-bar');
    const m = S.ui.mode;
    if (!cls || !m || m === 'scoring') { bar.hidden = true; bar.innerHTML = ''; return; }
    bar.hidden = false;
    const { byId, names } = studentMaps(cls);
    let html = '';
    if (m === 'adjust') {
      const A = S.adjust;
      const msg = A.sel
        ? `已选中 <b>${esc(names.get(A.work.get(A.sel)) || core.seatLabel(core.parseSeatKey(A.sel).r, core.parseSeatKey(A.sel).c))}</b>，再点另一个座位完成对调（可以是空座）`
        : (A.count ? `已调整 ${A.count} 次，可以继续；满意后点右侧保存` : '点一名学生，再点另一个座位，两人对调；点空座则挪过去');
      html = `<span class="mode-tag">微调中</span><span class="mode-msg">${msg}</span>
        <button type="button" data-act="adjust-reset">重置</button>
        <button type="button" data-act="adjust-cancel">取消</button>
        <button type="button" data-act="adjust-save" class="primary">${A.kind === 'final' ? '★ 确认微调并定版' : '保存到候选稿'}</button>`;
    } else if (m === 'pin') {
      html = `<span class="mode-tag">固定学生</span><span class="mode-msg">点学生即固定在当前座位（📌），再点取消。已固定 ${cls.pins.length} 人。</span>
        <button type="button" data-act="pin-clear">全部取消</button>
        <button type="button" data-act="mode-done" class="primary">完成</button>`;
    } else if (m === 'gender') {
      const all = store.studentsOf(cls);
      const cm = all.filter(s => s.gender === 'M').length;
      const cf = all.filter(s => s.gender === 'F').length;
      const G = S.gender;
      const rest = core.fillRemainingGender(all, G.brush).length;
      const other = G.brush === 'F' ? '男生' : '女生';
      html = `<span class="mode-tag">标注性别</span>
        <span>点学生标为</span>
        <div class="segmented brush" role="group" aria-label="标为">
          <button type="button" data-act="brush-F" class="b-f ${G.brush === 'F' ? 'active' : ''}" aria-pressed="${G.brush === 'F'}">♀ 女生</button>
          <button type="button" data-act="brush-M" class="b-m ${G.brush === 'M' ? 'active' : ''}" aria-pressed="${G.brush === 'M'}">♂ 男生</button>
        </div>
        <span class="mode-msg muted">再点一次可取消 · 男 <b class="g-m">${cm}</b> · 女 <b class="g-f">${cf}</b> · 未标注 <b>${all.length - cm - cf}</b></span>
        ${rest ? `<button type="button" data-act="fill-rest" class="primary">剩余 ${rest} 人标为${other}</button>` : ''}
        <button type="button" data-act="gender-undo" ${G.undo.length ? '' : 'disabled'}>↶ 撤销</button>
        <button type="button" data-act="mode-done">完成</button>`;
    } else if (m === 'select') {
      const X = S.select;
      if (X.idx < X.order.length) {
        const sid = X.order[X.idx];
        const s = byId.get(sid);
        const scoreText = s && typeof s.score === 'number' ? `成绩 ${s.score}` : (s && typeof s.rank === 'number' ? `名次 ${s.rank}` : '无成绩');
        const next = X.order.slice(X.idx + 1, X.idx + 6).map(id => esc(names.get(id))).join('、');
        html = `<span class="mode-tag">按成绩选座</span>
          <span class="mode-msg">第 ${X.idx + 1}/${X.order.length} 位：<span class="current-pick" style="color:${s && s.gender === 'M' ? 'var(--male)' : s && s.gender === 'F' ? 'var(--female)' : 'var(--ink)'}">${esc(names.get(sid))}</span>（${scoreText}）请点一个空座
          ${next ? `<span class="queue">之后：${next}${X.order.length - X.idx - 1 > 5 ? '…' : ''}</span>` : ''}</span>
          <button type="button" data-act="select-skip">跳过（最后再选）</button>
          <button type="button" data-act="select-undo" ${X.history.length ? '' : 'disabled'}>撤销上一步</button>
          <button type="button" data-act="select-auto">剩余按名次自动排</button>
          <button type="button" data-act="select-cancel">取消</button>`;
      } else {
        html = `<span class="mode-tag">按成绩选座</span><span class="mode-msg">全部选完。检查一下，然后生成候选稿。</span>
          <button type="button" data-act="select-undo">撤销上一步</button>
          <button type="button" data-act="select-cancel">取消</button>
          <button type="button" data-act="select-finish" class="primary">完成：生成候选稿</button>`;
      }
    }
    bar.innerHTML = html;
  }

  function renderScoringBar(cls) {
    const bar = $('scoring-bar');
    if (!cls || S.ui.mode !== 'scoring') { bar.hidden = true; return; }
    bar.hidden = false;
    const items = store.itemsOf(S.st);
    const cur = S.scoring.item;
    const btn = it => `<button type="button" class="item-btn ${it.value > 0 ? 'bonus' : 'penalty'} ${cur && cur.name === it.name ? 'active' : ''}" data-item="${esc(it.name)}">${esc(it.name)}<b>${core.fmtScore(it.value)}</b></button>`;
    $('items-bonus').innerHTML = items.filter(i => i.value > 0).map(btn).join('');
    $('items-penalty-1').innerHTML = items.filter(i => i.value < 0 && i.row !== 2).map(btn).join('');
    $('items-penalty-2').innerHTML = items.filter(i => i.value < 0 && i.row === 2).map(btn).join('');
    const dateInput = $('score-date');
    if (!dateInput.value) dateInput.value = S.scoring.date || today();
    $('score-date-warn').textContent = clockBehind() ? '电脑时间可能不准，请核对日期' : (dateInput.value !== today() ? '注意：不是今天' : '');
    $('score-item-name').textContent = cur ? `${cur.name}（${core.fmtScore(cur.value)}）` : '未选择';
    const { names } = studentMaps(cls);
    const sel = Array.from(S.scoring.sel);
    $('score-count').textContent = String(sel.length);
    $('score-names').textContent = sel.length ? `（${sel.slice(0, 10).map(id => names.get(id)).join('、')}${sel.length > 10 ? '…' : ''}）` : '';
    $('btn-score-undo').disabled = !S.scoring.lastBatch;
  }

  function renderLegend(cls, rec) {
    const el = $('seat-legend');
    if (!cls) { el.innerHTML = ''; return; }
    const all = store.studentsOf(cls);
    const cm = all.filter(s => s.gender === 'M').length;
    const cf = all.filter(s => s.gender === 'F').length;
    const cap = rec ? core.capacity(rec.layout) : 0;
    el.innerHTML = [
      `<span class="lg g-m">男 ${cm}</span>`,
      `<span class="lg g-f">女 ${cf}</span>`,
      cm + cf < all.length
        ? (S.ui.mode ? `<span class="lg">未标注性别 ${all.length - cm - cf}</span>` : `<button type="button" class="link" data-act="gender-start">未标注性别 ${all.length - cm - cf} · 去标注</button>`)
        : '',
      rec ? `<span>空座 ${cap - rec.seats.length}</span>` : '',
      cls.pins.length ? '<span>📌 固定座位</span>' : '',
      '<span>第 1 排靠近讲台</span>',
      `<span>${S.ui.view === 'teacher' ? '教师版：面向学生' : '学生版：面向讲台'}</span>`
    ].filter(Boolean).join('');
  }

  function drawPool(cls) {
    const rec = displayed(cls);
    const active = new Set(store.studentsOf(cls).map(s => s.id));
    if (rec && rec.seats.length) return rec.seats.map(t => t[2]).filter(id => active.has(id));
    return Array.from(active);
  }

  function renderSeatMap(cls, rec) {
    const map = $('seat-map');
    map.classList.toggle('clickable', !!S.ui.mode);
    map.classList.toggle('gender-mode', S.ui.mode === 'gender');
    if (!cls) { map.innerHTML = ''; $('seat-notes').innerHTML = ''; return; }
    const { byId, names } = studentMaps(cls);
    if (!rec) {
      $('seat-notes').innerHTML = '';
      if (S.ui.mode === 'gender') {
        const ctx0 = { byId, names, drawn: new Map(), pins: new Map(), mode: 'gender' };
        map.innerHTML = '<p class="hint">还没有座次表，按名单顺序显示：</p><div class="roster-grid">' +
          store.studentsOf(cls).map(s => seatHTML(0, 0, s.id, ctx0)).join('') + '</div>';
      } else {
        map.innerHTML = '<div class="notes"><p>还没有座次表。</p></div>';
      }
      return;
    }
    const drawn = new Map();
    if (S.draw.resultClass === cls.id) S.draw.result.forEach((id, i) => drawn.set(id, i));
    const pins = new Map(cls.pins.map(p => [p.sid, core.seatKey(p.r, p.c)]));
    const ctx = {
      byId, names, drawn, pins,
      mode: S.ui.mode,
      scoreSel: S.ui.mode === 'scoring' ? S.scoring.sel : null,
      swapSel: S.ui.mode === 'adjust' ? S.adjust.sel : null,
      week: S.ui.mode === 'scoring' ? weekTotals(cls, core.mondayOf($('score-date').value || today())) : null
    };
    map.innerHTML = roomHTML(rec.layout, core.seatsToMap(rec.seats), S.ui.view, (r, c, sid) => seatHTML(r, c, sid, ctx));
    const notes = [];
    if (rec.meta && Array.isArray(rec.meta.notes)) rec.meta.notes.forEach(n => notes.push(n));
    if (rec.kind === 'draft' && rec.meta && rec.meta.source) notes.push(`基准：${core.shortDate(rec.meta.sourceDate || '')}的定版`);
    const inactiveSeated = rec.seats.filter(t => byId.get(t[2]) && byId.get(t[2]).active === false).length;
    if (inactiveSeated) notes.push(`有 ${inactiveSeated} 名已离班学生仍在这张表上`);
    const unseated = store.studentsOf(cls).filter(s => !rec.seats.some(t => t[2] === s.id));
    if (unseated.length && rec.kind !== 'select') notes.push(`${unseated.length} 名在班学生不在这张表上：${unseated.slice(0, 6).map(s => names.get(s.id)).join('、')}${unseated.length > 6 ? '…' : ''}（重新生成或微调即可安排）`);
    $('seat-notes').innerHTML = notes.map(n => `<p>• ${esc(n)}</p>`).join('');
  }

  function renderTools(cls, rec) {
    const method = $('gen-method').value;
    const dateInput = $('gen-date');
    if (!dateInput.value || dateInput.dataset.classId !== String(cls ? cls.id : '')) {
      dateInput.value = defaultGenDate(cls);
      dateInput.dataset.classId = cls ? cls.id : '';
    }
    renderGenOptions(cls, method);
    const vs = $('version-select');
    if (cls) {
      const eff = store.effectiveFinal(cls, today());
      const finals = store.finalsOf(cls).slice().reverse();
      let opts = '<option value="">今天生效的（自动）</option>';
      if (cls.draft) opts += `<option value="draft">候选稿 · ${core.weekSpan(cls.draft.date)}（${esc(MODE_TEXT[cls.draft.mode] || '')}）</option>`;
      opts += finals.map(f => `<option value="${esc(f.finalId)}">定版 · ${core.weekSpan(f.date)}（${core.shortDate(f.date)}生效${eff && eff.finalId === f.finalId ? ' · 今天生效' : ''}）</option>`).join('');
      vs.innerHTML = opts;
      const want = S.ui.version || '';
      vs.value = Array.from(vs.options).some(o => o.value === want) ? want : '';
    } else {
      vs.innerHTML = '';
    }
    const busy = !!S.ui.mode;
    vs.disabled = busy || !cls;
    ['btn-generate', 'btn-today', 'btn-pin-mode', 'btn-gender-mode'].forEach(id => { $(id).disabled = busy || !cls; });
    $('btn-adjust').disabled = busy || !rec;
    $('btn-finalize').disabled = busy || !rec || rec.kind !== 'draft';
    $('btn-export-seat').disabled = !rec || rec.kind === 'select';
    $('btn-pin-mode').disabled = !rec || (busy && S.ui.mode !== 'pin');
    $('btn-gender-mode').disabled = !cls || !cls.students.size || (busy && S.ui.mode !== 'gender');
    $('pin-summary').textContent = cls ? (cls.pins.length ? `已固定 ${cls.pins.length} 人。` : '还没有固定的学生。') : '';
  }

  function renderGenOptions(cls, method) {
    const box = $('gen-options');
    const hint = $('gen-hint');
    hint.textContent = '';
    if (!cls) { box.innerHTML = ''; return; }
    const students = store.studentsOf(cls);
    const withScore = students.filter(s => typeof s.score === 'number' || typeof s.rank === 'number').length;
    const genders = students.filter(s => s.gender).length;
    const date = $('gen-date').value || today();
    const source = store.sourceFinalFor(cls, date);
    if (method === 'random') {
      const was = $('opt-gender-pair') ? $('opt-gender-pair').checked : false;
      box.innerHTML = `<label class="check"><input type="checkbox" id="opt-gender-pair" ${was ? 'checked' : ''} ${genders ? '' : 'disabled'}> 同桌尽量男女搭配</label>`;
      if (!genders) hint.textContent = '还没有标注性别，男女搭配暂不可用（可用“标注性别”或在班级设置里填写）。';
    } else if (method === 'select' || method === 'score') {
      box.innerHTML = `<p class="hint">本班 ${withScore}/${students.length} 人有成绩；没有成绩的排在最后。${method === 'select' ? '按名次从高到低，每人点一个空座。' : '名次靠前的优先坐前排中间。'}</p>`;
    } else if (method === 'rotate') {
      const layout = store.layoutOf(cls);
      const r = Object.assign(core.defaultRotation(layout), cls.rotation || {});
      box.innerHTML = `
        <div class="opt-line">前后：<select id="opt-row-dir"><option value="back" ${r.rowDir !== 'front' ? 'selected' : ''}>往后移</option><option value="front" ${r.rowDir === 'front' ? 'selected' : ''}>往前移</option></select>
          <input id="opt-row-shift" type="number" min="0" max="19" value="${r.rowShift}"> 排</div>
        <div class="opt-line">左右：<select id="opt-col-dir"><option value="right" ${r.colDir !== 'left' ? 'selected' : ''}>往右移</option><option value="left" ${r.colDir === 'left' ? 'selected' : ''}>往左移</option></select>
          <input id="opt-col-shift" type="number" min="0" max="19" value="${r.colShift}"> 列</div>`;
      hint.textContent = source
        ? `以 ${core.shortDate(source.date)} 的定版为基准；最后几排移到最前、最右几列移到最左；空座和固定学生不动；填 0 表示这个方向不动。`
        : `${core.weekSpan(date)} 之前还没有定版，无法轮换。`;
    } else if (method === 'hold') {
      box.innerHTML = '';
      hint.textContent = source ? `沿用 ${core.shortDate(source.date)} 的定版，座位不变（只为这一周另存一份）。` : `${core.weekSpan(date)} 之前还没有定版，无法沿用。`;
    }
  }

  /* =========================================================
     座次页：操作
     ========================================================= */

  function exitMode(silent) {
    S.ui.mode = null;
    S.adjust = null;
    S.select = null;
    S.gender = null;
    if (!silent) renderAll();
  }

  function onSeatClick(e) {
    const el = e.target.closest('.seat');
    if (!el || el.classList.contains('void') || !S.ui.mode) return;
    const r = Number(el.dataset.r);
    const c = Number(el.dataset.c);
    const sid = el.dataset.sid || null;
    if (S.ui.mode === 'adjust') onSwapClick(r, c);
    else if (S.ui.mode === 'pin' && sid) togglePin(sid, r, c);
    else if (S.ui.mode === 'gender' && sid) paintGender(sid);
    else if (S.ui.mode === 'select' && !sid) placePick(r, c);
    else if (S.ui.mode === 'scoring' && sid) toggleScoreSel(sid);
  }

  // —— 随机抽取 ——
  function drawCount() {
    const custom = parseInt($('draw-custom').value, 10);
    return Number.isFinite(custom) && custom > 0 ? custom : S.draw.n;
  }

  function doDraw() {
    const cls = currentClass();
    if (!cls) return;
    const pool = drawPool(cls);
    if (!pool.length) { showError('本班还没有学生'); return; }
    const n = drawCount();
    let drawn = S.draw.drawn.get(cls.id);
    if (!drawn) { drawn = new Set(); S.draw.drawn.set(cls.id, drawn); }
    let exclude = S.draw.noRepeat ? Array.from(drawn) : [];
    if (S.draw.noRepeat && pool.every(id => drawn.has(id))) {
      drawn.clear();
      exclude = [];
      toast('本轮所有人都已抽过，重新开始');
    }
    const result = core.drawStudents(pool, n, core.defaultRng, exclude);
    if (result.length < n) toast(S.draw.noRepeat ? `本轮只剩 ${result.length} 人没抽过，已全部抽出` : `本班只有 ${pool.length} 人`);
    if (S.draw.noRepeat) result.forEach(id => drawn.add(id));
    S.draw.result = result;
    S.draw.resultClass = cls.id;
    renderSeating();
  }

  function clearDraw() {
    S.draw.result = [];
    S.draw.resultClass = null;
  }

  // —— 生成 ——
  function readRotation() {
    return {
      rowShift: Math.max(0, parseInt($('opt-row-shift').value, 10) || 0),
      colShift: Math.max(0, parseInt($('opt-col-shift').value, 10) || 0),
      rowDir: $('opt-row-dir').value === 'front' ? 'front' : 'back',
      colDir: $('opt-col-dir').value === 'left' ? 'left' : 'right'
    };
  }

  function onGenerate() {
    const cls = currentClass();
    if (!cls) return;
    const method = $('gen-method').value;
    const date = $('gen-date').value;
    if (!core.isISODate(date)) { showError('请先选择生效日期'); return; }
    const students = store.studentsOf(cls);
    if (!students.length) { showError('本班还没有学生，请先在“班级设置”里导入名单'); return; }
    const layout = store.layoutOf(cls);
    if (method === 'select') { startSelect(cls, date); return; }
    let res;
    const meta = {};
    try {
      if (method === 'random') {
        const gp = !!($('opt-gender-pair') && $('opt-gender-pair').checked);
        res = core.randomSeating(students, layout, cls.pins, { genderPair: gp });
        meta.genderPair = gp;
      } else if (method === 'score') {
        if (!students.some(s => typeof s.score === 'number' || typeof s.rank === 'number')) throw new Error('本班还没有成绩或名次，无法按成绩排座');
        res = core.scoreSeating(students, layout, cls.pins);
      } else {
        const source = store.sourceFinalFor(cls, date);
        if (!source) throw new Error(`${core.weekSpan(date)} 之前还没有定版，无法${method === 'rotate' ? '轮换' : '沿用'}；可以先用随机或按成绩生成`);
        if (method === 'rotate') {
          const rot = readRotation();
          res = core.rotateSeating(source, students, layout, cls.pins, rot);
          meta.rotation = res.rotation;
          if (JSON.stringify(rot) !== JSON.stringify(cls.rotation)) commit('class.update', { classId: cls.id, patch: { rotation: rot } }, { noFlush: true });
        } else {
          res = core.holdSeating(source, students, layout, cls.pins);
        }
        meta.source = source.finalId;
        meta.sourceDate = source.date;
      }
    } catch (e) {
      showError(e.message);
      return;
    }
    if (cls.draft && !confirm('已有一份候选稿，要用新生成的替换它吗？')) return;
    meta.notes = res.notes;
    commit('draft.set', { classId: cls.id, draft: { draftId: 'd_' + core.randomId(8), date, mode: method, seats: res.seats, layout, meta } });
    S.ui.version = 'draft';
    clearDraw();
    renderAll();
    toast(`已生成候选稿（${MODE_TEXT[method]}）。检查后点“确认定版”才生效`);
  }

  // —— 按成绩选座 ——
  function startSelect(cls, date) {
    const students = store.studentsOf(cls);
    if (!students.some(s => typeof s.score === 'number' || typeof s.rank === 'number')) {
      showError('本班还没有成绩或名次，无法按成绩选座。可在“班级设置”的名单里填写成绩，或导入带成绩的名单');
      return;
    }
    const layout = store.layoutOf(cls);
    const pinInfo = core.sanitizePins(cls.pins, layout, students);
    const placed = new Map(pinInfo.valid.map(p => [core.seatKey(p.r, p.c), p.sid]));
    const pinned = new Set(pinInfo.valid.map(p => p.sid));
    const order = core.rankStudents(students.filter(s => !pinned.has(s.id))).map(s => s.id);
    if (order.length > core.capacity(layout) - placed.size) { showError('座位不够，请先在“班级设置”里增加排数或列数'); return; }
    S.select = { date, layout, placed, order, idx: 0, history: [] };
    S.ui.mode = 'select';
    S.ui.version = null;
    clearDraw();
    renderAll();
    $('mode-bar').scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function placePick(r, c) {
    const X = S.select;
    if (X.idx >= X.order.length) return;
    const k = core.seatKey(r, c);
    if (X.placed.has(k)) return;
    X.placed.set(k, X.order[X.idx]);
    X.history.push(k);
    X.idx++;
    renderSeating();
  }

  function onModeAction(act) {
    const cls = currentClass();
    if (!cls) return;
    const X = S.select;
    if (act === 'select-skip') {
      if (X.idx < X.order.length - 1) {
        X.order.push(X.order.splice(X.idx, 1)[0]);
        renderSeating();
      }
    } else if (act === 'select-undo') {
      const k = X.history.pop();
      if (k) { X.placed.delete(k); X.idx--; renderSeating(); }
    } else if (act === 'select-auto') {
      try {
        core.placeLeftovers(X.placed, X.order.slice(X.idx), X.layout);
        X.order.slice(X.idx).forEach(() => X.history.push(null));
        X.idx = X.order.length;
      } catch (e) { showError(e.message); }
      renderSeating();
    } else if (act === 'select-cancel') {
      if (confirm('放弃这次选座吗？')) exitMode();
    } else if (act === 'select-finish') {
      if (cls.draft && !confirm('已有一份候选稿，要用这次选座的结果替换它吗？')) return;
      commit('draft.set', { classId: cls.id, draft: { draftId: 'd_' + core.randomId(8), date: X.date, mode: 'select', seats: core.mapToSeats(X.placed), layout: X.layout, meta: { notes: [] } } });
      S.ui.mode = null;
      S.select = null;
      S.ui.version = 'draft';
      renderAll();
      toast('已生成候选稿。检查后点“确认定版”才生效');
    } else if (act === 'adjust-reset') {
      S.adjust.work = core.seatsToMap(S.adjust.snapshot);
      S.adjust.sel = null;
      S.adjust.count = 0;
      renderSeating();
    } else if (act === 'adjust-cancel') {
      exitMode();
    } else if (act === 'adjust-save') {
      saveAdjust();
    } else if (act === 'pin-clear') {
      if (cls.pins.length && confirm(`取消全部 ${cls.pins.length} 名学生的固定吗？`)) {
        commit('pins.set', { classId: cls.id, pins: [] });
        renderSeating();
      }
    } else if (act === 'brush-F' || act === 'brush-M') {
      S.gender.brush = act === 'brush-F' ? 'F' : 'M';
      renderModeBar(cls);
    } else if (act === 'fill-rest') {
      fillRestGender();
    } else if (act === 'gender-undo') {
      undoGender();
    } else if (act === 'mode-done') {
      exitMode();
    }
  }

  // —— 微调 ——
  function startAdjust() {
    const cls = currentClass();
    const rec = displayed(cls);
    if (!rec) { showError('当前没有座次表可以微调'); return; }
    S.adjust = {
      kind: rec.kind, finalId: rec.finalId || null, date: rec.date, layout: rec.layout, mode: rec.mode, meta: rec.meta || {},
      snapshot: rec.seats.map(t => t.slice()), work: core.seatsToMap(rec.seats), sel: null, count: 0
    };
    S.ui.mode = 'adjust';
    clearDraw();
    renderAll();
    toast(rec.kind === 'final' ? '已进入微调：这是定版，保存时会重新定版（旧版保留在记录里）' : '已进入微调：点两个座位即可对调');
  }

  function onSwapClick(r, c) {
    const A = S.adjust;
    const k = core.seatKey(r, c);
    if (!A.sel) {
      if (!A.work.has(k)) { toast('先点一名学生，再点要换到的座位'); return; }
      A.sel = k;
    } else if (A.sel === k) {
      A.sel = null;
    } else {
      const a = A.work.get(A.sel);
      const b = A.work.get(k);
      if (b) A.work.set(A.sel, b); else A.work.delete(A.sel);
      A.work.set(k, a);
      A.sel = null;
      A.count++;
    }
    renderSeating();
  }

  async function saveAdjust() {
    const cls = currentClass();
    const A = S.adjust;
    if (!A.count) { toast('还没有调整任何座位'); return; }
    const seats = core.mapToSeats(A.work);
    const v = core.validateSeats(seats, A.layout, store.studentsOf(cls, true), { requireAll: false });
    if (!v.ok) { showError('无法保存：' + v.errors.join('；')); return; }
    const pos = new Map(seats.map(t => [t[2], t]));
    const pins = cls.pins.map(p => (pos.has(p.sid) ? { sid: p.sid, r: pos.get(p.sid)[0], c: pos.get(p.sid)[1] } : p));
    let exportRec = null;
    if (A.kind === 'draft') {
      commit('draft.set', { classId: cls.id, draft: { draftId: cls.draft.draftId, date: A.date, mode: A.mode, seats, layout: A.layout, meta: Object.assign({}, A.meta, { adjusted: true }) } }, { noFlush: true });
    } else {
      if (!confirm(`确定用微调后的座次重新定版 ${core.weekSpan(A.date)} 吗？\n旧版仍保留在记录里。`)) return;
      const finalId = 'f_' + core.randomId(10);
      commit('final.set', { classId: cls.id, finalId, date: A.date, seats, layout: A.layout, mode: 'adjust', meta: { adjustedFrom: A.finalId }, base: A.finalId }, { noFlush: true });
      S.ui.version = finalId;
      exportRec = { finalId, date: A.date, seats, layout: A.layout, mode: 'adjust' };
    }
    if (JSON.stringify(pins) !== JSON.stringify(cls.pins)) commit('pins.set', { classId: cls.id, pins }, { noFlush: true });
    S.ui.mode = null;
    S.adjust = null;
    renderAll();
    const res = await flushNow();
    toast(A.kind === 'draft' ? '已保存到候选稿' : `已重新定版 ${core.weekSpan(A.date)}`);
    if (exportRec && res.ok) exportSeating(cls, exportRec);
  }

  // —— 固定学生 / 标注性别 ——
  function togglePin(sid, r, c) {
    const cls = currentClass();
    let pins = cls.pins.slice();
    const cur = pins.find(p => p.sid === sid);
    if (cur && cur.r === r && cur.c === c) pins = pins.filter(p => p.sid !== sid);
    else {
      pins = pins.filter(p => p.sid !== sid && !(p.r === r && p.c === c));
      pins.push({ sid, r, c });
    }
    commit('pins.set', { classId: cls.id, pins });
    renderSeating();
  }

  function startGender() {
    const cls = currentClass();
    if (!cls || !cls.students.size) { showError('本班还没有学生'); return; }
    if (S.ui.mode && S.ui.mode !== 'gender') exitMode(true);
    S.ui.tab = 'seating';
    S.ui.mode = 'gender';
    S.gender = { brush: 'F', undo: [] };
    clearDraw();
    renderAll();
    $('mode-bar').scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  // 批量改性别，并记下原来的值以便撤销
  function setGenders(cls, changes, label) {
    const undo = [];
    changes.forEach(ch => {
      const s = cls.students.get(ch.sid);
      if (!s || s.gender === ch.gender) return;
      undo.push({ sid: ch.sid, gender: s.gender });
      commit('student.upsert', { classId: cls.id, studentId: ch.sid, fields: { gender: ch.gender } }, { noFlush: true });
    });
    if (undo.length) {
      S.gender.undo.push({ label, changes: undo });
      scheduleFlush();
    }
    return undo.length;
  }

  function paintGender(sid) {
    const cls = currentClass();
    const s = cls.students.get(sid);
    if (!s) return;
    setGenders(cls, [{ sid, gender: core.applyGenderBrush(s.gender, S.gender.brush) }], s.name);
    renderSeating();
  }

  function fillRestGender() {
    const cls = currentClass();
    const list = core.fillRemainingGender(store.studentsOf(cls), S.gender.brush);
    if (!list.length) return;
    const text = list[0].gender === 'M' ? '男生' : '女生';
    if (!confirm(`把剩余 ${list.length} 名未标注性别的学生都标为${text}吗？（可以撤销）`)) return;
    const n = setGenders(cls, list, `剩余 ${list.length} 人标为${text}`);
    renderSeating();
    toast(`已把 ${n} 人标为${text}`);
  }

  function undoGender() {
    const cls = currentClass();
    const last = S.gender.undo.pop();
    if (!last) return;
    last.changes.forEach(ch => commit('student.upsert', { classId: cls.id, studentId: ch.sid, fields: { gender: ch.gender } }, { noFlush: true }));
    scheduleFlush();
    renderSeating();
    toast(`已撤销：${last.label}`);
  }

  // —— 定版与导出 ——
  async function onFinalize() {
    const cls = currentClass();
    const rec = displayed(cls);
    if (!rec || rec.kind !== 'draft') return;
    const students = store.studentsOf(cls, true);
    const v = core.validateSeats(rec.seats, rec.layout, students, { requireAll: true });
    if (!v.ok) { showError('还不能定版：' + v.errors.join('；')); return; }
    const week = core.mondayOf(rec.date);
    const existing = cls.finals.get(week);
    if (existing && !confirm(`${core.weekSpan(rec.date)} 已有定版（${core.shortDate(existing.date)}生效）。\n确定用这张候选稿覆盖吗？旧版仍保留在记录里。`)) return;
    const finalId = 'f_' + core.randomId(10);
    commit('final.set', { classId: cls.id, finalId, date: rec.date, seats: rec.seats, layout: rec.layout, mode: rec.mode, meta: rec.meta || {}, base: existing ? existing.finalId : null }, { noFlush: true });
    commit('draft.set', { classId: cls.id, draft: null }, { noFlush: true });
    S.ui.version = finalId;
    renderAll();
    const res = await flushNow();
    if (!res.ok) return;
    toast(`已定版：${core.weekSpan(rec.date)}，${core.shortDate(rec.date)}起生效`);
    exportSeating(cls, { finalId, date: rec.date, seats: rec.seats, layout: rec.layout, mode: rec.mode });
  }

  async function writeExport(dirName, base, buf) {
    const res = await store.writeNewFile(S.root, [store.EXPORT_DIR, dirName], base, '.xlsx', buf);
    if (res.ok) {
      markSeen(`${dirName}/${res.name}|${res.size}|${res.lastModified}`);
      toast(`已导出：${res.path}`, '', 6000);
    }
    else {
      downloadBlob(new Blob([buf], { type: XLSX_MIME }), base + '.xlsx');
      showError(`没能写入文件夹（${res.error}），已改为下载`);
    }
    return res;
  }

  async function exportSeating(cls, rec) {
    try {
      const students = store.studentsOf(cls, true);
      const dir = xl.safeFileName(cls.name);
      const tag = rec.finalId ? '定版' : '候选';
      const wb = await xl.createSeatingWorkbook(cls, rec, students, {
        exportId: 'e_' + core.randomId(10),
        title: `座次表-${cls.name}-${core.weekSpan(rec.date)}${rec.finalId ? '' : '（候选）'}`
      });
      await writeExport(dir, `座次表-${dir}-${core.weekSpan(rec.date)}-${tag}-${xl.exportStamp()}`, await wb.xlsx.writeBuffer());
    } catch (e) {
      showError('导出座次表失败：' + e.message);
    }
  }

  // —— 课堂记分 ——
  function startScoring() {
    const cls = currentClass();
    if (!cls) return;
    const eff = store.effectiveFinal(cls, today());
    if (!eff) { showError('今天还没有生效的定版座次。请先生成座次并“确认定版”'); return; }
    S.ui.version = null;
    S.ui.mode = 'scoring';
    S.scoring = { item: null, sel: new Set(), lastBatch: S.scoring ? S.scoring.lastBatch : null, date: today() };
    $('score-date').value = today();
    clearDraw();
    renderAll();
    $('scoring-bar').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function toggleScoreSel(sid) {
    if (S.scoring.sel.has(sid)) S.scoring.sel.delete(sid); else S.scoring.sel.add(sid);
    renderSeating();
  }

  async function confirmScoring() {
    const cls = currentClass();
    const sc = S.scoring;
    if (!sc.item) { showError('请先点上方的一个表现项'); return; }
    if (!sc.sel.size) { showError('请在座次表中点选至少 1 名学生'); return; }
    const date = $('score-date').value;
    if (!core.isISODate(date)) { showError('记分日期不对'); return; }
    const now = new Date();
    const time = `${date} ${pad2(now.getHours())}:${pad2(now.getMinutes())}:${pad2(now.getSeconds())}`;
    const recIds = [];
    const names = [];
    sc.sel.forEach(sid => {
      const s = cls.students.get(sid);
      if (!s) return;
      const recId = 'r_' + core.randomId(10);
      commit('score.add', { rec: { recId, classId: cls.id, studentId: sid, name: s.name, item: sc.item.name, value: sc.item.value, date, time } }, { noFlush: true });
      recIds.push(recId);
      names.push(s.name);
    });
    sc.lastBatch = { recIds, text: `${names.join('、')} · ${sc.item.name}` };
    sc.sel.clear();
    renderSeating();
    const res = await flushNow();
    if (res.ok) toast(`已记录：${names.join('、')} ${sc.item.name}（${core.fmtScore(sc.item.value)}）`);
    else showError(`已记录，但还没保存到文件夹：${res.error}`);
  }

  async function undoScoring() {
    const b = S.scoring && S.scoring.lastBatch;
    if (!b) return;
    if (!confirm(`撤销上一次记分（${b.text}）吗？`)) return;
    b.recIds.forEach(id => commit('score.delete', { recId: id }, { noFlush: true }));
    S.scoring.lastBatch = null;
    renderSeating();
    await flushNow();
    toast('已撤销上一次记分');
  }

  /* =========================================================
     课堂表现页
     ========================================================= */

  function scoreWeeks(cls) {
    const set = new Set(store.scoresOf(S.st, cls.id).map(r => r.week));
    set.add(core.mondayOf(today()));
    return Array.from(set).sort().reverse();
  }

  function renderScores() {
    const cls = currentClass();
    if (!cls) { $('scores-title').textContent = '还没有班级'; $('scores-map').innerHTML = ''; $('scores-detail').innerHTML = ''; return; }
    const weeks = scoreWeeks(cls);
    if (!S.ui.scoresWeek || (S.ui.scoresWeek !== 'all' && weeks.indexOf(S.ui.scoresWeek) < 0)) S.ui.scoresWeek = weeks[0];
    const cur = core.mondayOf(today());
    $('scores-week').innerHTML = weeks.map(w => `<option value="${w}" ${w === S.ui.scoresWeek ? 'selected' : ''}>${core.weekSpan(w)}${w === cur ? '（本周）' : ''}</option>`).join('') +
      `<option value="all" ${S.ui.scoresWeek === 'all' ? 'selected' : ''}>全部记录</option>`;
    const sel = S.ui.scoresWeek;
    const recs = store.scoresOf(S.st, cls.id).filter(r => sel === 'all' || r.week === sel);
    $('scores-title').textContent = `课堂表现 · ${cls.name} · ${sel === 'all' ? '全部' : core.weekSpan(sel)}`;
    $('scores-meta').textContent = `${recs.length} 条记录 · 可导出成 Excel，在 Excel 里删改后保存，回到网页会自动读入`;
    $('btn-scores-clear').disabled = sel === 'all' || !recs.length;
    renderScoresDetail(cls, recs);
    const sum = core.summarizeScores(recs);
    let rec = sel === 'all' ? store.effectiveFinal(cls, today()) : store.effectiveFinal(cls, core.addDays(sel, 6));
    if (!rec) rec = store.effectiveFinal(cls, today()) || store.finalsOf(cls).slice(-1)[0] || null;
    const { byId, names } = studentMaps(cls);
    const cell = (r, c, sid) => {
      if (!sid) return '<div class="seat empty"><span class="empty-label">空</span></div>';
      const s = byId.get(sid);
      const x = sum.get(sid);
      const classes = ['seat', genderClass(s)];
      if (x && x.net > 0) classes.push('pos-score');
      if (x && x.net < 0) classes.push('neg-score');
      if (S.ui.scoresStudent === sid) classes.push('focus');
      const tags = x ? `${x.bonusSum ? `<span class="tag-plus">+${x.bonusSum}</span>` : ''}${x.penaltySum ? `<span class="tag-minus">${x.penaltySum}</span>` : ''}` : '';
      return `<div class="${classes.join(' ')}" data-sid="${esc(sid)}" data-r="${r}" data-c="${c}" title="${esc(names.get(sid))}${x ? `：净分 ${core.fmtScore(x.net)}` : ''}"><b class="seat-name">${esc(names.get(sid) || '')}</b>${tags ? `<span class="tags">${tags}</span>` : ''}</div>`;
    };
    const box = $('scores-map');
    box.classList.add('clickable');
    if (rec) {
      box.innerHTML = roomHTML(rec.layout, core.seatsToMap(rec.seats), 'student', cell);
    } else {
      const list = store.studentsOf(cls);
      box.innerHTML = `<p class="hint">还没有定版座次，按名单显示：</p><div class="room" style="grid-template-columns: repeat(8, minmax(0, 1fr))">${list.map(s => cell(0, 0, s.id)).join('')}</div>`;
    }
  }

  function renderScoresDetail(cls, recs) {
    const box = $('scores-detail');
    const sid = S.ui.scoresStudent;
    const s = sid ? cls.students.get(sid) : null;
    if (!s) {
      box.innerHTML = '<span class="muted">💡 点下方座次表中的学生，可在这里查看他的加分、扣分明细，并删除记错的记录。</span>';
      return;
    }
    const mine = recs.filter(r => r.studentId === sid).sort((a, b) => (b.time || b.date).localeCompare(a.time || a.date));
    const x = core.summarizeScores(mine).get(sid) || { bonusSum: 0, bonusCount: 0, penaltySum: 0, penaltyCount: 0, net: 0 };
    box.innerHTML = `<div class="detail-head"><span class="name" style="color:${s.gender === 'M' ? 'var(--male)' : s.gender === 'F' ? 'var(--female)' : 'var(--ink)'}">${esc(s.name)}</span>
      <span>加分 <b class="plus">+${x.bonusSum}</b>（${x.bonusCount} 次）</span><span>扣分 <b>${x.penaltySum}</b>（${x.penaltyCount} 次）</span><span>净分 <b>${core.fmtScore(x.net)}</b></span>
      <button type="button" class="quiet small" data-act="close-detail">✕ 关闭</button></div>
      <div class="chips-list">${mine.length ? mine.map(r => `<span class="score-chip ${r.value > 0 ? 'bonus' : 'penalty'}">${esc(r.item)} <b>${core.fmtScore(r.value)}</b><time>${esc(r.time || r.date)}</time><button type="button" data-del="${esc(r.recId)}" title="删除这条记录">删除</button></span>`).join('') : '<span class="muted">所选范围内没有记录</span>'}</div>`;
  }

  async function deleteScore(recId) {
    const r = S.st.scores.get(recId);
    if (!r || !confirm(`删除这条记录吗？\n${r.name} · ${r.item}（${core.fmtScore(r.value)}）· ${r.time || r.date}`)) return;
    commit('score.delete', { recId });
    renderAll();
  }

  async function exportScores() {
    const cls = currentClass();
    if (!cls) return;
    const sel = S.ui.scoresWeek;
    const recs = store.scoresOf(S.st, cls.id).filter(r => sel === 'all' || r.week === sel);
    try {
      const dir = xl.safeFileName(cls.name);
      const label = sel === 'all' ? '全部' : core.weekSpan(sel);
      const wb = await xl.createScoresWorkbook(cls, recs, store.studentsOf(cls, true), {
        exportId: 'e_' + core.randomId(10),
        title: `课堂表现-${cls.name}-${label}`,
        defaultDate: sel === 'all' ? today() : sel
      });
      await writeExport(dir, `课堂表现-${dir}-${label}-${xl.exportStamp()}`, await wb.xlsx.writeBuffer());
    } catch (e) {
      showError('导出失败：' + e.message);
    }
  }

  async function clearWeek() {
    const cls = currentClass();
    const sel = S.ui.scoresWeek;
    const recs = store.scoresOf(S.st, cls.id).filter(r => r.week === sel);
    if (!recs.length || !confirm(`清空 ${cls.name} ${core.weekSpan(sel)} 的全部 ${recs.length} 条记分吗？\n（记录里会保留“已删除”的标记，不会影响其他周）`)) return;
    recs.forEach(r => commit('score.delete', { recId: r.recId }, { noFlush: true }));
    await flushNow();
    renderAll();
    toast(`已清空 ${core.weekSpan(sel)} 的 ${recs.length} 条记分`);
  }

  /* =========================================================
     学生轨迹
     ========================================================= */

  function renderStudent() {
    const cls = currentClass();
    const sel = $('traj-student');
    if (!cls) { sel.innerHTML = ''; return; }
    const all = store.studentsOf(cls, true).sort((a, b) => ((a.active === false) - (b.active === false)) || core.compareRoster(a, b));
    if (!all.length) {
      sel.innerHTML = '<option value="">（没有学生）</option>';
      $('traj-title').textContent = '学生轨迹';
      $('traj-stats').innerHTML = '';
      $('traj-seats').innerHTML = '';
      $('traj-scores').innerHTML = '';
      $('traj-score-stats').innerHTML = '';
      return;
    }
    if (!all.some(s => s.id === S.ui.trajStudent)) S.ui.trajStudent = all[0].id;
    const names = core.displayNames(all);
    sel.innerHTML = all.map(s => `<option value="${esc(s.id)}" ${s.id === S.ui.trajStudent ? 'selected' : ''}>${esc(names.get(s.id))}${s.active === false ? '（离班）' : ''}</option>`).join('');
    const s = cls.students.get(S.ui.trajStudent);
    $('traj-title').innerHTML = `学生轨迹 · <span style="color:${s.gender === 'M' ? 'var(--male)' : s.gender === 'F' ? 'var(--female)' : 'var(--ink)'}">${esc(s.name)}</span>`;
    const t = today();
    const rows = store.finalsOf(cls).slice().sort((a, b) => (a.date < b.date ? -1 : 1)).map(f => {
      const seat = f.seats.find(x => x[2] === s.id);
      return { f, seat, future: f.date > t };
    });
    const used = rows.filter(x => x.seat && !x.future);
    const cnt = (list, fn) => list.reduce((m, x) => { const k = fn(x); m[k] = (m[k] || 0) + 1; return m; }, {});
    const rz = cnt(used, x => core.rowZone(x.f.layout, x.seat[0]));
    const cz = cnt(used, x => core.colZone(x.f.layout, x.seat[1]));
    $('traj-stats').innerHTML = `<span><b>${used.length}</b>次定版座位</span>
      <span>前区 <b>${rz['前'] || 0}</b> · 中区 <b>${rz['中'] || 0}</b> · 后区 <b>${rz['后'] || 0}</b></span>
      <span>左侧 <b>${cz['左'] || 0}</b> · 中间 <b>${cz['中'] || 0}</b> · 右侧 <b>${cz['右'] || 0}</b></span>
      <span>前中后覆盖 <b>${Object.keys(rz).length}/3</b> · 左中右覆盖 <b>${Object.keys(cz).length}/3</b></span>`;
    $('traj-seats').innerHTML = rows.length ? rows.slice().reverse().map(x => `<tr class="${x.future ? 'future' : ''}">
      <td>${esc(x.f.date)}${x.future ? '（未生效）' : ''}</td><td>${core.weekSpan(x.f.date)}</td><td>${esc(MODE_TEXT[x.f.mode] || x.f.mode)}</td>
      <td>${x.seat ? core.seatLabel(x.seat[0], x.seat[1]) : '（不在这张表上）'}</td>
      <td>${x.seat ? core.rowZone(x.f.layout, x.seat[0]) + '区' : ''}</td><td>${x.seat ? core.colZone(x.f.layout, x.seat[1]) : ''}</td></tr>`).join('')
      : '<tr><td colspan="6" class="muted">还没有定版座次</td></tr>';
    const scores = store.scoresOf(S.st, cls.id).filter(r => r.studentId === s.id).sort((a, b) => (b.time || b.date).localeCompare(a.time || a.date));
    const x = core.summarizeScores(scores).get(s.id) || { bonusSum: 0, bonusCount: 0, penaltySum: 0, penaltyCount: 0, net: 0 };
    $('traj-score-stats').innerHTML = `<span>加分 <b>+${x.bonusSum}</b>（${x.bonusCount} 次）</span><span>扣分 <b>${x.penaltySum}</b>（${x.penaltyCount} 次）</span><span>净分 <b>${core.fmtScore(x.net)}</b></span>`;
    $('traj-scores').innerHTML = scores.length ? scores.map(r => `<tr><td>${esc(r.time || r.date)}</td><td>${core.weekSpan(r.date)}</td><td>${esc(r.item)}</td>
      <td class="${r.value > 0 ? 'plus' : 'minus'}">${core.fmtScore(r.value)}</td><td><button type="button" class="quiet small" data-del="${esc(r.recId)}">删除</button></td></tr>`).join('')
      : '<tr><td colspan="5" class="muted">还没有课堂表现记录</td></tr>';
  }

  /* =========================================================
     班级设置
     ========================================================= */

  function renderSettings() {
    const cls = currentClass();
    $('class-name').value = cls ? cls.name : '';
    ['class-name', 'btn-rename', 'btn-export-roster', 'btn-delete-class', 'btn-add-student', 'btn-save-layout', 'btn-reset-layout'].forEach(id => { $(id).disabled = !cls; });
    renderRoster(cls);
    renderLayoutEditor(cls);
    renderItemsEditor();
    renderDataInfo();
  }

  function renderRoster(cls) {
    const body = $('roster-body');
    if (!cls) { body.innerHTML = '<tr><td colspan="6" class="muted">还没有班级</td></tr>'; $('roster-count').textContent = ''; return; }
    const all = store.studentsOf(cls, true);
    const act = all.filter(s => s.active !== false);
    $('roster-count').textContent = `在班 ${act.length} 人${all.length > act.length ? ` · 离班 ${all.length - act.length} 人` : ''} · 男 ${act.filter(s => s.gender === 'M').length} · 女 ${act.filter(s => s.gender === 'F').length}`;
    body.innerHTML = all.length ? all.map(s => `<tr class="${s.active === false ? 'inactive' : ''}" data-sid="${esc(s.id)}">
      <td><input type="number" class="score" data-f="seq" value="${typeof s.seq === 'number' ? s.seq : ''}" aria-label="序号"></td>
      <td class="name-cell ${genderClass(s)}"><input type="text" data-f="name" value="${esc(s.name)}" aria-label="姓名"></td>
      <td><select data-f="gender" aria-label="性别"><option value="" ${!s.gender ? 'selected' : ''}>未标注</option><option value="M" ${s.gender === 'M' ? 'selected' : ''}>男</option><option value="F" ${s.gender === 'F' ? 'selected' : ''}>女</option></select></td>
      <td><input type="number" step="any" class="score" data-f="score" value="${typeof s.score === 'number' ? s.score : ''}" aria-label="成绩"></td>
      <td><input type="text" class="no" data-f="no" value="${esc(s.no || '')}" aria-label="学号"></td>
      <td><select data-f="active" aria-label="状态"><option value="1" ${s.active !== false ? 'selected' : ''}>在班</option><option value="0" ${s.active === false ? 'selected' : ''}>离班</option></select></td>
    </tr>`).join('') : '<tr><td colspan="6" class="muted">还没有学生：点“导入名单”或“添加学生”</td></tr>';
  }

  function onRosterChange(e) {
    const el = e.target.closest('[data-f]');
    const tr = e.target.closest('tr[data-sid]');
    if (!el || !tr) return;
    const cls = currentClass();
    const sid = tr.dataset.sid;
    const f = el.dataset.f;
    let v = el.value;
    if (f === 'name') {
      v = core.normalizeName(v);
      if (!v) { showError('姓名不能为空'); renderRoster(cls); return; }
    } else if (f === 'score' || f === 'seq') {
      v = v.trim() === '' ? null : Number(v);
      if (v !== null && !Number.isFinite(v)) { showError('请输入数字'); return; }
    } else if (f === 'active') {
      v = v === '1';
      if (!v && !confirm('把这名学生标为“离班”吗？之后生成的座次不再安排他；历史记录保留。')) { renderRoster(cls); return; }
    }
    const fields = {};
    fields[f] = v;
    commit('student.upsert', { classId: cls.id, studentId: sid, fields });
    renderRoster(cls);
  }

  function addStudent() {
    const cls = currentClass();
    const name = core.normalizeName(prompt('新学生的姓名：') || '');
    if (!name) return;
    const all = store.studentsOf(cls, true);
    const seq = all.reduce((m, s) => Math.max(m, typeof s.seq === 'number' ? s.seq : 0), 0) + 1;
    commit('student.upsert', { classId: cls.id, studentId: 's_' + core.randomId(8), fields: { name, gender: '', score: null, seq, active: true } });
    renderAll();
    toast(`已添加 ${name}。下次生成座次时会安排座位`);
  }

  // —— 布局编辑 ——
  function layoutDraft(cls) {
    if (!S.layoutDraft || S.layoutDraft.classId !== cls.id) S.layoutDraft = { classId: cls.id, layout: core.normalizeLayout(store.layoutOf(cls)) };
    return S.layoutDraft.layout;
  }

  function renderLayoutEditor(cls) {
    const box = $('layout-preview');
    if (!cls) { box.innerHTML = ''; $('layout-capacity').textContent = ''; return; }
    const L = layoutDraft(cls);
    $('layout-rows').value = L.rows;
    $('layout-cols').value = L.cols;
    const cap = core.capacity(L);
    const n = store.studentsOf(cls).length;
    const changed = core.layoutKey(L) !== core.layoutKey(store.layoutOf(cls));
    $('layout-capacity').innerHTML = `可用座位 <b>${cap}</b> · 在班学生 <b>${n}</b>${cap < n ? ' <span class="warn-text">座位不够</span>' : ''}${changed ? ' · <span class="warn-text">有未保存的修改</span>' : ''}`;
    const tmpl = ['44px'];
    for (let c = 1; c <= L.cols; c++) {
      tmpl.push('minmax(40px, 1fr)');
      if (c < L.cols) tmpl.push(L.aisles.indexOf(c) >= 0 ? '22px' : '10px');
    }
    const col = c => 2 + (c - 1) * 2;
    let html = `<div class="lp-grid" style="grid-template-columns:${tmpl.join(' ')}"><div class="lp-stage" style="grid-row:1; grid-column: 2 / -1">讲台</div>`;
    for (let c = 1; c < L.cols; c++) {
      const on = L.aisles.indexOf(c) >= 0;
      html += `<div class="lp-gap ${on ? 'on' : ''}" data-gap="${c}" style="grid-row: 2 / span ${L.rows}; grid-column:${col(c) + 1}" title="第${c}列与第${c + 1}列之间：点此${on ? '取消' : '设为'}走廊"></div>`;
    }
    for (let r = 1; r <= L.rows; r++) {
      html += `<div class="lp-row" style="grid-row:${r + 1}; grid-column:1">第${r}排</div>`;
      for (let c = 1; c <= L.cols; c++) {
        const off = L.disabled.indexOf(core.seatKey(r, c)) >= 0;
        html += `<div class="lp-seat ${off ? 'off' : ''}" data-r="${r}" data-c="${c}" style="grid-row:${r + 1}; grid-column:${col(c)}" title="第${r}排第${c}列：点此${off ? '恢复' : '设为不可用'}">${off ? '×' : `${c}`}</div>`;
      }
    }
    box.innerHTML = html + '</div>';
  }

  function onLayoutClick(e) {
    const cls = currentClass();
    if (!cls) return;
    const L = layoutDraft(cls);
    const gap = e.target.closest('[data-gap]');
    const seat = e.target.closest('.lp-seat');
    if (gap) {
      const a = Number(gap.dataset.gap);
      L.aisles = L.aisles.indexOf(a) >= 0 ? L.aisles.filter(x => x !== a) : L.aisles.concat([a]);
    } else if (seat) {
      const k = core.seatKey(Number(seat.dataset.r), Number(seat.dataset.c));
      L.disabled = L.disabled.indexOf(k) >= 0 ? L.disabled.filter(x => x !== k) : L.disabled.concat([k]);
    } else return;
    S.layoutDraft.layout = core.normalizeLayout(L);
    renderLayoutEditor(cls);
  }

  function onLayoutSize() {
    const cls = currentClass();
    if (!cls) return;
    const L = layoutDraft(cls);
    S.layoutDraft.layout = core.normalizeLayout({ rows: $('layout-rows').value, cols: $('layout-cols').value, aisles: L.aisles, disabled: L.disabled });
    renderLayoutEditor(cls);
  }

  function saveLayout() {
    const cls = currentClass();
    const L = layoutDraft(cls);
    const n = store.studentsOf(cls).length;
    if (core.capacity(L) < n && !confirm(`可用座位（${core.capacity(L)}）少于在班学生（${n}），生成座次时会提示座位不够。仍然保存吗？`)) return;
    commit('layout.set', { classId: cls.id, layout: L, layoutId: 'l_' + core.randomId(8), base: cls.layoutId });
    renderAll();
    toast('已保存布局。之后生成的座次按新布局安排；已定版的座次不受影响');
  }

  // —— 表现项 ——
  function itemsDraft() {
    if (!S.itemsDraft) S.itemsDraft = store.itemsOf(S.st).map(i => Object.assign({}, i));
    return S.itemsDraft;
  }

  function renderItemsEditor() {
    const list = itemsDraft();
    $('items-body').innerHTML = list.map((it, i) => `<tr data-i="${i}">
      <td><input type="text" data-f="name" value="${esc(it.name)}" maxlength="12"></td>
      <td><input type="number" data-f="value" value="${it.value}" step="1"></td>
      <td>${it.value < 0 ? `<select data-f="row"><option value="1" ${it.row !== 2 ? 'selected' : ''}>扣分第 1 行</option><option value="2" ${it.row === 2 ? 'selected' : ''}>扣分第 2 行</option></select>` : '加分行'}</td>
      <td><button type="button" class="quiet small" data-del-item="${i}">删除</button></td></tr>`).join('');
  }

  function onItemsInput(e) {
    const el = e.target.closest('[data-f]');
    const tr = e.target.closest('tr[data-i]');
    if (!el || !tr) return;
    const it = itemsDraft()[Number(tr.dataset.i)];
    if (el.dataset.f === 'name') it.name = el.value.trim();
    if (el.dataset.f === 'value') it.value = Number(el.value);
    if (el.dataset.f === 'row') it.row = Number(el.value);
    if (el.dataset.f === 'value') renderItemsEditor();
  }

  function saveItems() {
    const items = core.normalizeItems(itemsDraft());
    if (!items.length) { showError('至少要有一个表现项'); return; }
    if (items.length < itemsDraft().length) toast('已去掉名称为空、分值为 0 或重名的项');
    commit('items.set', { items });
    S.itemsDraft = null;
    renderAll();
    toast('已保存表现项（所有班级共用）');
  }

  function renderDataInfo() {
    const r = S.readResult;
    const files = r ? r.files.length : 0;
    const own = S.writer && S.writer.fileName ? S.writer.fileName : '（本次还没有修改）';
    $('data-info').innerHTML = `
      <p>数据文件夹：<code>${esc(S.rootName)}/${store.DATA_DIR}/</code></p>
      <p>记录文件 ${files} 个${r && r.snapshot ? `（其中 ${r.snapshot.covers} 个已合并进快照）` : ''}；本次会话写入：<code>${esc(own)}</code></p>
      <p>Excel 导出位置：<code>${esc(S.rootName)}/${store.EXPORT_DIR}/班级名/</code>（每次导出都是新文件，旧的可随时删除）</p>
      <p class="hint">“数据”文件夹里的文件不要手动修改或删除；网页只新建文件，从不改动已有文件，所以 SyncTime、OneDrive 同步不会冲突。</p>`;
  }

  // —— 班级 ——
  function newClass() {
    const name = xl.normalizeClassName(prompt('新班级的名称（例如：高二4班）：') || '');
    if (!name) return;
    if (classes().some(c => c.name === name) && !confirm(`已经有“${name}”，仍要再建一个同名班级吗？`)) return;
    const classId = 'c_' + core.randomId(8);
    const order = classes().reduce((m, c) => Math.max(m, c.order === 999 ? 0 : c.order), 0) + 1;
    commit('class.create', { classId, name, order }, { noFlush: true });
    commit('layout.set', { classId, layout: core.normalizeLayout({ rows: 6, cols: 8, aisles: [2, 4, 6] }), layoutId: 'l_' + core.randomId(8), base: null });
    switchClass(classId);
    toast(`已新建 ${name}。接下来导入名单或添加学生`);
  }

  function renameClass() {
    const cls = currentClass();
    const name = xl.normalizeClassName($('class-name').value);
    if (!name || name === cls.name) return;
    commit('class.update', { classId: cls.id, patch: { name } });
    renderAll();
    toast(`已改名为 ${name}`);
  }

  function deleteClass() {
    const cls = currentClass();
    const typed = prompt(`删除后，这个班级不再显示（记录文件里仍保留，不会影响其他班）。\n请输入班级名称“${cls.name}”确认删除：`);
    if (typed === null) return;
    if (typed.trim() !== cls.name) { showError('名称不一致，没有删除'); return; }
    commit('class.delete', { classId: cls.id });
    S.ui.classId = null;
    renderAll();
    toast(`已删除 ${cls.name}`);
  }

  async function exportRoster() {
    const cls = currentClass();
    try {
      const dir = xl.safeFileName(cls.name);
      const wb = await xl.createRosterWorkbook(cls, store.studentsOf(cls, true), { exportId: 'e_' + core.randomId(10) });
      await writeExport(dir, `名单-${dir}-${xl.exportStamp()}`, await wb.xlsx.writeBuffer());
    } catch (e) {
      showError('导出名单失败：' + e.message);
    }
  }

  function switchClass(id) {
    if (S.ui.mode) exitMode(true);
    S.ui.classId = id;
    S.ui.version = null;
    S.ui.scoresStudent = null;
    S.ui.trajStudent = null;
    S.layoutDraft = null;
    clearDraw();
    renderAll();
  }

  /* =========================================================
     导入名单
     ========================================================= */

  function openImport() {
    S.importCands = [];
    $('import-file').value = '';
    $('import-text').value = '';
    renderImportPreview();
    $('import-dialog').showModal();
  }

  async function onImportFile() {
    const f = $('import-file').files[0];
    if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      const res = /\.xlsx$/i.test(f.name)
        ? await xl.parseRosterWorkbook(buf, f.name)
        : xl.parseRosterText(xl.decodeText(buf), xl.classNameFromText(f.name));
      setImportCands(res);
    } catch (e) {
      showError(`无法读取“${f.name}”：${e.message}（旧版 .xls 请先在 Excel/WPS 里另存为 .xlsx）`);
    }
  }

  function onImportParse() {
    const text = $('import-text').value;
    if (!text.trim()) { showError('请先粘贴名单'); return; }
    setImportCands(xl.parseRosterText(text, ''));
  }

  function setImportCands(res) {
    const cur = currentClass();
    S.importCands = res.candidates.map(c => {
      const same = classes().find(x => x.name === xl.normalizeClassName(c.className));
      return Object.assign({}, c, {
        name: xl.normalizeClassName(c.className) || (cur ? cur.name : '新班级'),
        target: same ? same.id : 'new',
        markLeft: false
      });
    });
    if (!S.importCands.length) showError('没有识别到学生姓名。' + (res.warnings || []).join('；'));
    renderImportPreview();
  }

  function renderImportPreview() {
    const box = $('import-preview');
    const list = classes();
    box.innerHTML = S.importCands.map((c, i) => {
      const cols = [c.columns.gender !== undefined ? '性别' : '', c.columns.score !== undefined ? '成绩' : '', c.columns.rank !== undefined ? '名次' : '', c.columns.no !== undefined ? '学号' : '', c.columns.cls !== undefined ? '班级' : ''].filter(Boolean);
      const withScore = c.students.filter(s => typeof s.score === 'number').length;
      const withGender = c.students.filter(s => s.gender).length;
      return `<div class="cand" data-i="${i}">
        <div class="cand-head">
          <span class="muted">${esc(c.source)}</span>
          <label>班级名 <input type="text" data-role="name" value="${esc(c.name)}"></label>
          <label>导入到 <select data-role="target"><option value="new" ${c.target === 'new' ? 'selected' : ''}>新建班级</option>${list.map(x => `<option value="${esc(x.id)}" ${c.target === x.id ? 'selected' : ''}>合并到 ${esc(x.name)}</option>`).join('')}</select></label>
        </div>
        <div class="sample">共 ${c.students.length} 人 · 识别到的列：姓名${cols.length ? '、' + cols.join('、') : ''} · 有成绩 ${withScore} 人 · 有性别 ${withGender} 人<br>前几位：${c.students.slice(0, 6).map(s => esc(s.name)).join('、')}${c.students.length > 6 ? '…' : ''}</div>
        ${c.warnings.length ? `<div class="warns">${c.warnings.map(esc).join('；')}</div>` : ''}
        ${c.target !== 'new' ? `<label class="check"><input type="checkbox" data-role="markLeft" ${c.markLeft ? 'checked' : ''}> 名单里没有的学生标为离班</label><p class="hint">合并时按姓名对应：更新性别、成绩等，新增名单里多出的学生。</p>` : ''}
      </div>`;
    }).join('');
    $('btn-import-confirm').disabled = !S.importCands.length;
  }

  function onImportPreviewChange(e) {
    const el = e.target.closest('[data-role]');
    const box = e.target.closest('.cand');
    if (!el || !box) return;
    const c = S.importCands[Number(box.dataset.i)];
    if (el.dataset.role === 'name') c.name = xl.normalizeClassName(el.value);
    if (el.dataset.role === 'target') { c.target = el.value; renderImportPreview(); }
    if (el.dataset.role === 'markLeft') c.markLeft = el.checked;
  }

  function confirmImport() {
    if (!S.importCands.length) return;
    let firstId = null;
    let total = 0;
    for (const c of S.importCands) {
      let classId = c.target;
      let current = [];
      if (classId === 'new') {
        if (!c.name) { showError('请填写班级名称'); return; }
        classId = 'c_' + core.randomId(8);
        const order = classes().reduce((m, x) => Math.max(m, x.order === 999 ? 0 : x.order), 0) + 1;
        commit('class.create', { classId, name: c.name, order }, { noFlush: true });
        commit('layout.set', { classId, layout: core.defaultLayoutFor(c.students.length), layoutId: 'l_' + core.randomId(8), base: null }, { noFlush: true });
      } else {
        current = store.studentsOf(S.st.classes.get(classId), true);
      }
      const byName = new Map(current.map(s => [s.name, s]));
      const seen = new Set();
      c.students.forEach((s, i) => {
        const fields = { name: s.name };
        if (s.gender) fields.gender = s.gender;
        if (s.score !== undefined) fields.score = s.score;
        if (s.rank !== undefined) fields.rank = s.rank;
        if (s.no) fields.no = s.no;
        if (typeof s.seq === 'number') fields.seq = s.seq;
        const found = byName.get(s.name);
        if (found) {
          seen.add(found.id);
          if (s.active !== false && found.active === false) fields.active = true;
          commit('student.upsert', { classId, studentId: found.id, fields }, { noFlush: true });
        } else {
          if (!('seq' in fields)) fields.seq = current.length + i + 1;
          fields.gender = fields.gender || '';
          fields.active = s.active !== false;
          commit('student.upsert', { classId, studentId: 's_' + core.randomId(8), fields }, { noFlush: true });
        }
      });
      if (c.target !== 'new' && c.markLeft) {
        current.filter(s => s.active !== false && !seen.has(s.id)).forEach(s => commit('student.upsert', { classId, studentId: s.id, fields: { active: false } }, { noFlush: true }));
      }
      total += c.students.length;
      if (!firstId) firstId = classId;
    }
    flushNow();
    $('import-dialog').close();
    switchClass(firstId);
    toast(`已导入 ${S.importCands.length} 个班、共 ${total} 人`);
    S.importCands = [];
  }

  /* =========================================================
     冲突处理与提示操作
     ========================================================= */

  function onAlert(act, arg) {
    const cls = currentClass();
    if (act === 'retry') { flushNow().then(r => { if (r.ok) toast('已保存'); else showError('仍然没保存成功：' + r.error); }); return; }
    if (act === 'dismiss') { S.notices.splice(Number(arg), 1); renderAlerts(); return; }
    if (!cls) return;
    const parts = String(arg).split('|');
    const cf = cls.conflicts.find(x => x.kind === parts[0] && x.key === parts[1]);
    if (!cf) return;
    if (act === 'conflict-keep') {
      commit('conflict.ack', { classId: cls.id, kind2: cf.kind, key: cf.key, winnerId: cf.winnerId });
    } else if (act === 'conflict-other') {
      if (cf.kind === 'final') {
        const l = cf.loser;
        commit('final.set', { classId: cls.id, finalId: 'f_' + core.randomId(10), date: l.date, seats: l.seats, layout: l.layout, mode: l.mode, meta: l.meta || {}, base: cf.winnerId });
      } else {
        commit('layout.set', { classId: cls.id, layout: cf.loser.layout, layoutId: 'l_' + core.randomId(8), base: cf.winnerId });
      }
      S.ui.version = null;
    }
    renderAll();
  }

  /* =========================================================
     事件绑定
     ========================================================= */

  async function onFolderButton() {
    try {
      if (S.pendingRoot && !S.root) {
        if (await permitted(S.pendingRoot, true)) { await connect(S.pendingRoot, {}); return; }
      }
      if (S.root && S.writer && S.writer.pendingCount() && !confirm('当前文件夹还有没保存的修改，换文件夹后会尝试写入新文件夹。继续吗？')) return;
      const handle = await window.showDirectoryPicker({ id: 'classmanager', mode: 'readwrite' });
      await connect(handle, {});
    } catch (e) {
      if (e && e.name !== 'AbortError') showError('连接文件夹失败：' + store.describeError(e));
    }
  }

  function bind() {
    document.querySelectorAll('.tabs [data-tab]').forEach(b => b.addEventListener('click', () => {
      if (S.ui.mode === 'adjust' || S.ui.mode === 'select') { toast('请先完成或取消当前的微调/选座'); return; }
      if (S.ui.mode) exitMode(true);
      S.ui.tab = b.dataset.tab;
      renderAll();
    }));
    $('class-select').addEventListener('change', e => switchClass(e.target.value));
    $('btn-folder').addEventListener('click', onFolderButton);
    $('btn-welcome-connect').addEventListener('click', onFolderButton);
    $('btn-refresh').addEventListener('click', async () => {
      if (!S.root) { location.reload(); return; }
      lastReturn = 0;
      await onReturn();
      renderAll();
      toast('已重新读取文件夹');
    });
    $('save-status').addEventListener('click', () => { if (S.writer && S.writer.pendingCount()) onAlert('retry'); });
    $('alerts').addEventListener('click', e => {
      const b = e.target.closest('[data-alert]');
      if (b) onAlert(b.dataset.alert, b.dataset.arg);
    });

    // 座次页
    $('view-student').addEventListener('click', () => { S.ui.view = 'student'; renderSeating(); saveUi(); });
    $('view-teacher').addEventListener('click', () => { S.ui.view = 'teacher'; renderSeating(); saveUi(); });
    $('seat-map').addEventListener('click', onSeatClick);
    $('mode-bar').addEventListener('click', e => { const b = e.target.closest('[data-act]'); if (b && !b.disabled) onModeAction(b.dataset.act); });
    $('draw-chips').addEventListener('click', e => {
      const b = e.target.closest('[data-n]');
      if (!b) return;
      S.draw.n = Number(b.dataset.n);
      $('draw-custom').value = '';
      renderDrawBar(currentClass());
    });
    $('draw-custom').addEventListener('input', () => renderDrawBar(currentClass()));
    $('draw-custom').addEventListener('keydown', e => { if (e.key === 'Enter') doDraw(); });
    $('btn-draw').addEventListener('click', doDraw);
    $('btn-draw-clear').addEventListener('click', () => { clearDraw(); renderSeating(); });
    $('draw-norepeat').addEventListener('change', e => {
      S.draw.noRepeat = e.target.checked;
      const cls = currentClass();
      if (cls) S.draw.drawn.set(cls.id, new Set(e.target.checked ? S.draw.result : []));
      renderDrawBar(cls);
    });
    $('btn-scoring').addEventListener('click', () => { if (S.ui.mode === 'scoring') exitMode(); else startScoring(); });
    $('score-rows').addEventListener('click', e => {
      const b = e.target.closest('[data-item]');
      if (!b) return;
      S.scoring.item = store.itemsOf(S.st).find(i => i.name === b.dataset.item) || null;
      renderScoringBar(currentClass());
    });
    $('score-date').addEventListener('change', () => { S.scoring.date = $('score-date').value; renderSeating(); });
    $('btn-score-all').addEventListener('click', () => {
      const rec = displayed(currentClass());
      if (rec) rec.seats.forEach(t => S.scoring.sel.add(t[2]));
      renderSeating();
    });
    $('btn-score-none').addEventListener('click', () => { S.scoring.sel.clear(); renderSeating(); });
    $('btn-score-confirm').addEventListener('click', confirmScoring);
    $('btn-score-undo').addEventListener('click', undoScoring);
    $('btn-score-exit').addEventListener('click', () => exitMode());
    $('gen-method').addEventListener('change', () => renderGenOptions(currentClass(), $('gen-method').value));
    $('gen-date').addEventListener('change', () => renderGenOptions(currentClass(), $('gen-method').value));
    $('btn-generate').addEventListener('click', onGenerate);
    $('version-select').addEventListener('change', e => { S.ui.version = e.target.value || null; clearDraw(); renderSeating(); });
    $('btn-today').addEventListener('click', () => { S.ui.version = null; clearDraw(); renderSeating(); });
    $('btn-adjust').addEventListener('click', startAdjust);
    $('btn-finalize').addEventListener('click', onFinalize);
    $('btn-export-seat').addEventListener('click', () => {
      const cls = currentClass();
      const rec = displayed(cls);
      if (rec) exportSeating(cls, rec);
    });
    $('btn-pin-mode').addEventListener('click', () => { if (S.ui.mode === 'pin') exitMode(); else { S.ui.mode = 'pin'; clearDraw(); renderAll(); } });
    $('btn-gender-mode').addEventListener('click', () => { if (S.ui.mode === 'gender') exitMode(); else startGender(); });
    $('seat-legend').addEventListener('click', e => { if (e.target.closest('[data-act="gender-start"]')) startGender(); });
    $('btn-gender-quick').addEventListener('click', startGender);

    // 课堂表现
    $('scores-week').addEventListener('change', e => { S.ui.scoresWeek = e.target.value; S.ui.scoresStudent = null; renderScores(); });
    $('scores-map').addEventListener('click', e => {
      const el = e.target.closest('.seat[data-sid]');
      if (!el) return;
      S.ui.scoresStudent = S.ui.scoresStudent === el.dataset.sid ? null : el.dataset.sid;
      renderScores();
    });
    $('scores-detail').addEventListener('click', e => {
      const d = e.target.closest('[data-del]');
      if (d) deleteScore(d.dataset.del);
      if (e.target.closest('[data-act="close-detail"]')) { S.ui.scoresStudent = null; renderScores(); }
    });
    $('btn-scores-export').addEventListener('click', exportScores);
    $('btn-scores-clear').addEventListener('click', clearWeek);

    // 学生轨迹
    $('traj-student').addEventListener('change', e => { S.ui.trajStudent = e.target.value; renderStudent(); });
    $('traj-scores').addEventListener('click', e => { const d = e.target.closest('[data-del]'); if (d) deleteScore(d.dataset.del); });

    // 班级设置
    $('btn-rename').addEventListener('click', renameClass);
    $('btn-new-class').addEventListener('click', newClass);
    $('btn-delete-class').addEventListener('click', deleteClass);
    $('btn-export-roster').addEventListener('click', exportRoster);
    $('btn-import').addEventListener('click', openImport);
    $('btn-add-student').addEventListener('click', addStudent);
    $('roster-body').addEventListener('change', onRosterChange);
    $('layout-preview').addEventListener('click', onLayoutClick);
    $('layout-rows').addEventListener('change', onLayoutSize);
    $('layout-cols').addEventListener('change', onLayoutSize);
    $('btn-save-layout').addEventListener('click', saveLayout);
    $('btn-reset-layout').addEventListener('click', () => { S.layoutDraft = null; renderSettings(); });
    $('items-body').addEventListener('change', onItemsInput);
    $('items-body').addEventListener('click', e => {
      const b = e.target.closest('[data-del-item]');
      if (!b) return;
      itemsDraft().splice(Number(b.dataset.delItem), 1);
      renderItemsEditor();
    });
    $('btn-add-bonus').addEventListener('click', () => { itemsDraft().push({ id: 'i_' + core.randomId(6), name: '新加分项', value: 1 }); renderItemsEditor(); });
    $('btn-add-penalty').addEventListener('click', () => { itemsDraft().push({ id: 'i_' + core.randomId(6), name: '新扣分项', value: -1, row: 2 }); renderItemsEditor(); });
    $('btn-save-items').addEventListener('click', saveItems);
    $('btn-reset-items').addEventListener('click', () => {
      if (!confirm('恢复为默认的 14 个表现项吗？')) return;
      S.itemsDraft = core.normalizeItems(core.DEFAULT_SCORE_ITEMS);
      saveItems();
    });

    // 导入
    $('import-file').addEventListener('change', onImportFile);
    $('btn-import-parse').addEventListener('click', onImportParse);
    $('import-preview').addEventListener('change', onImportPreviewChange);
    $('import-preview').addEventListener('input', e => { if (e.target.dataset.role === 'name') onImportPreviewChange(e); });
    $('btn-import-cancel').addEventListener('click', () => $('import-dialog').close());
    $('btn-import-confirm').addEventListener('click', confirmImport);

    window.addEventListener('focus', onReturn);
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') onReturn(); });
    window.addEventListener('beforeunload', e => {
      if (S.writer && S.writer.pendingCount()) {
        e.preventDefault();
        e.returnValue = '还有修改没保存到文件夹';
      }
    });
  }

  async function init() {
    loadUi();
    loadSeen();
    bind();
    renderAll();
    if (!HAS_FS) return;
    const h = await idbGet('root');
    if (!h) return;
    if (await permitted(h, false)) await connect(h, { initial: true });
    else { S.pendingRoot = h; renderAll(); }
  }

  // 供测试与排查使用：在控制台里连接一个文件夹句柄（例如浏览器私有目录）
  window.ClassManagerApp = {
    connectHandle: h => connect(h, {}),
    state: () => S,
    reload: () => onReturn()
  };

  init().catch(e => showError('初始化失败：' + e.message));
})();
