/**
 * CMStore — 数据层：记录文件、合并、读写
 *
 * 同步安全的核心约定（配合 SyncTime 双向同步、OneDrive）：
 *   1. 网页只“新建”文件，从不覆盖、改名、移动或删除已有文件。
 *   2. 每次打开网页开始一个“会话”，第一次修改时在 数据/ 下新建一个 .jsonl 记录文件，
 *      本次的所有操作都写进它；页面关闭后它不再变化。别的会话、别的电脑从不写它。
 *   3. 打开网页时读取 数据/ 下全部记录文件，按时间戳合并（同一条操作只算一次）。
 * 于是任何一个文件都只可能在一边被修改，SyncTime 只需把新文件拷到另一边，不会产生冲突。
 */
(function (global) {
  'use strict';

  const core = global.CMCore || (typeof require === 'function' ? require('./core.js') : null);

  const APP = 'classmanager';
  const FORMAT = 1;
  const DATA_DIR = '数据';
  const SNAP_DIR = '快照';
  const EXPORT_DIR = '导出';
  const PROJECT_FILE = 'project.json';
  const SNAPSHOT_AFTER = 30; // 未被快照覆盖的记录文件达到这么多时，写一个新快照
  const CLOCK_TOLERANCE = 5 * 60 * 1000; // 时间误差容忍：超过这么多才算“超前”或“系统时间异常”

  // 系统与临时文件：一律忽略（._ 是 Mac 写 U 盘时的附属文件，~$ 是 Office 锁文件，.crswap 是浏览器写入中的临时文件）
  function isJunkName(n) {
    return !n || n.charAt(0) === '.' || n.indexOf('~$') === 0 || /\.crswap$/i.test(n) || /^(thumbs\.db|desktop\.ini)$/i.test(n);
  }

  /* =========================================================
     混合逻辑时钟：时间戳 = 13 位毫秒:4 位计数:会话号，可直接按字符串排序。
     读到别处的记录后时钟不会倒退，所以即使教室电脑时间不准，因果顺序也不会乱。
     ========================================================= */

  function tsMake(wall, counter, sid) {
    return String(wall).padStart(13, '0') + ':' + String(counter).padStart(4, '0') + ':' + sid;
  }

  function tsParse(t) {
    const m = /^(\d{13}):(\d{4}):([a-z0-9]+)$/.exec(t || '');
    return m ? { wall: Number(m[1]), counter: Number(m[2]), sid: m[3] } : null;
  }

  function createClock(sid, nowFn) {
    nowFn = nowFn || Date.now;
    let lastWall = 0;
    let lastCounter = 0;
    return {
      observe(t) {
        const p = tsParse(t);
        if (!p) return;
        if (p.wall > lastWall || (p.wall === lastWall && p.counter > lastCounter)) {
          lastWall = p.wall;
          lastCounter = p.counter;
        }
      },
      now() {
        const w = nowFn();
        if (w > lastWall) {
          lastWall = w;
          lastCounter = 0;
        } else {
          lastCounter++;
          if (lastCounter > 9999) {
            lastWall++;
            lastCounter = 0;
          }
        }
        return tsMake(lastWall, lastCounter, sid);
      },
      maxWall() {
        return lastWall;
      }
    };
  }

  /* =========================================================
     网络时间：项目以网络时间为基准。依次尝试几个公开的时间服务（只发一个 GET 请求，不带任何数据），
     第一个成功的结果即为基准；都失败（教室电脑不联网）时返回 null，退回本机时间。
     ========================================================= */

  // ISO 时间 → 毫秒；小数秒截到 3 位；没有时区的按 UTC
  function parseIsoUtc(s) {
    const m = /^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})(?:\.(\d+))?(Z|[+-]\d{2}:?\d{2})?$/.exec(String(s || '').trim());
    if (!m) return NaN;
    const frac = m[2] ? '.' + (m[2] + '00').slice(0, 3) : '';
    let zone = m[3] || 'Z';
    if (/^[+-]\d{4}$/.test(zone)) zone = zone.slice(0, 3) + ':' + zone.slice(3);
    return Date.parse(m[1] + frac + zone);
  }

  const TIME_SOURCES = [
    { name: 'worldtimeapi.org', url: 'https://worldtimeapi.org/api/timezone/Etc/UTC', parse: t => parseIsoUtc(JSON.parse(t).utc_datetime) },
    { name: 'timeapi.io', url: 'https://timeapi.io/api/Time/current/zone?timeZone=UTC', parse: t => parseIsoUtc(JSON.parse(t).dateTime) },
    {
      name: 'cloudflare.com',
      url: 'https://www.cloudflare.com/cdn-cgi/trace',
      parse: t => {
        const m = /(?:^|\n)ts=(\d+(?:\.\d+)?)/.exec(t);
        return m ? Math.round(Number(m[1]) * 1000) : NaN;
      }
    }
  ];

  const plausibleTime = ms => Number.isFinite(ms) && ms > Date.UTC(2020, 0, 1) && ms < Date.UTC(2100, 0, 1);

  // 返回 { offset: 网络时间 − 本机时间（毫秒）, source, rtt }；全部失败返回 null
  async function fetchNetTime(fetchFn, opts) {
    opts = opts || {};
    const nowFn = opts.now || Date.now;
    const timeout = opts.timeout || 4000;
    const sources = opts.sources || TIME_SOURCES;
    if (typeof fetchFn !== 'function') return null;
    for (const src of sources) {
      const ctrl = typeof AbortController === 'function' ? new AbortController() : null;
      const timer = ctrl ? setTimeout(() => ctrl.abort(), timeout) : null;
      try {
        const t0 = nowFn();
        const resp = await fetchFn(src.url, { cache: 'no-store', credentials: 'omit', signal: ctrl ? ctrl.signal : undefined });
        if (!resp || !resp.ok) throw new Error('HTTP ' + (resp && resp.status));
        const text = await resp.text();
        const t1 = nowFn();
        const net = src.parse(text);
        if (!plausibleTime(net) || t1 - t0 > timeout) throw new Error('时间无效');
        return { offset: Math.round(net + (t1 - t0) / 2 - t1), source: src.name, rtt: t1 - t0 };
      } catch (e) {
        /* 换下一个来源 */
      } finally {
        if (timer) clearTimeout(timer);
      }
    }
    return null;
  }

  /* =========================================================
     合并：按时间戳排序、按编号去重，逐条应用
     ========================================================= */

  function emptyState() {
    return {
      classes: new Map(),
      scores: new Map(),
      tombstones: new Set(),
      items: null,
      settings: {},
      sessions: new Map(),
      imports: new Map(),
      ids: new Set(),
      maxT: '',
      opCount: 0,
      skipped: 0,
      broken: 0
    };
  }

  const str = v => (typeof v === 'string' ? v : '');
  const num = v => (typeof v === 'number' && Number.isFinite(v) ? v : null);

  function sanitizeSeats(seats) {
    const out = [];
    for (const t of Array.isArray(seats) ? seats : []) {
      if (Array.isArray(t) && Number.isInteger(t[0]) && Number.isInteger(t[1]) && typeof t[2] === 'string') out.push([t[0], t[1], t[2]]);
    }
    return out;
  }

  function newClass(op) {
    return {
      id: op.classId,
      name: str(op.name) || '未命名班级',
      order: num(op.order) === null ? 999 : op.order,
      deleted: false,
      createdT: op.t,
      students: new Map(),
      layout: null,
      layoutId: null,
      rotation: null,
      pins: [],
      draft: null,
      finals: new Map(),
      passageConfig: null,
      passageFinals: new Map(),
      passageHistory: [],
      conflicts: []
    };
  }

  const STUDENT_FIELDS = ['name', 'gender', 'score', 'rank', 'seq', 'no', 'active'];

  function applyStudentFields(stu, fields) {
    for (const k of STUDENT_FIELDS) {
      if (!(k in fields)) continue;
      const v = fields[k];
      if (k === 'name') {
        const n = core.normalizeName(v);
        if (n) stu.name = n;
      } else if (k === 'gender') {
        const g = core.normalizeGender(v);
        stu.gender = g === null ? '' : g;
      } else if (k === 'score' || k === 'rank' || k === 'seq') {
        stu[k] = num(v);
      } else if (k === 'no') {
        stu.no = v === null || v === undefined ? '' : String(v).trim();
      } else if (k === 'active') {
        stu.active = v !== false;
      }
    }
  }

  function dropConflict(cls, kind, key) {
    cls.conflicts = cls.conflicts.filter(c => !(c.kind === kind && c.key === key));
  }

  function applyOp(st, op) {
    const cls = op.classId ? st.classes.get(op.classId) : null;
    switch (op.k) {
      case 'session':
        st.sessions.set(op.s, { place: str(op.place), started: str(op.started), t: op.t });
        return true;
      case 'class.create':
        if (!op.classId || st.classes.has(op.classId)) return false;
        st.classes.set(op.classId, newClass(op));
        return true;
      case 'class.update': {
        if (!cls || !op.patch) return false;
        const p = op.patch;
        if (str(p.name).trim()) cls.name = p.name.trim();
        if (num(p.order) !== null) cls.order = p.order;
        if (p.rotation && typeof p.rotation === 'object') {
          cls.rotation = {
            rowShift: Math.max(0, Math.round(num(p.rotation.rowShift) || 0)),
            colShift: Math.max(0, Math.round(num(p.rotation.colShift) || 0)),
            rowDir: p.rotation.rowDir === 'front' ? 'front' : 'back',
            colDir: p.rotation.colDir === 'left' ? 'left' : 'right'
          };
        }
        return true;
      }
      case 'class.delete':
        if (!cls) return false;
        cls.deleted = true;
        return true;
      case 'student.upsert': {
        if (!cls || !op.studentId || !op.fields) return false;
        let stu = cls.students.get(op.studentId);
        if (!stu) {
          stu = { id: op.studentId, name: '', gender: '', score: null, rank: null, seq: null, no: '', active: true, createdT: op.t };
          cls.students.set(op.studentId, stu);
        }
        applyStudentFields(stu, op.fields);
        if (!stu.name) stu.name = '（未命名）';
        return true;
      }
      case 'layout.set': {
        if (!cls || !op.layout) return false;
        const layout = core.normalizeLayout(op.layout); // 先整理，格式异常时在改动状态之前就出错
        if (cls.layoutId && op.base !== cls.layoutId && op.layoutId !== cls.layoutId) {
          dropConflict(cls, 'layout', 'layout');
          cls.conflicts.push({ kind: 'layout', key: 'layout', winnerId: op.layoutId, loser: { layoutId: cls.layoutId, layout: cls.layout }, t: op.t });
        } else {
          dropConflict(cls, 'layout', 'layout');
        }
        cls.layout = layout;
        cls.layoutId = str(op.layoutId) || null;
        return true;
      }
      case 'pins.set':
        if (!cls) return false;
        cls.pins = (Array.isArray(op.pins) ? op.pins : [])
          .filter(p => p && typeof p.sid === 'string' && Number.isInteger(p.r) && Number.isInteger(p.c))
          .map(p => ({ sid: p.sid, r: p.r, c: p.c }));
        return true;
      case 'draft.set':
        if (!cls) return false;
        if (!op.draft) {
          cls.draft = null;
          return true;
        }
        if (!core.isISODate(op.draft.date)) return false;
        cls.draft = {
          draftId: str(op.draft.draftId) || op.id,
          date: op.draft.date,
          mode: str(op.draft.mode) || 'random',
          seats: sanitizeSeats(op.draft.seats),
          layout: core.normalizeLayout(op.draft.layout),
          meta: op.draft.meta && typeof op.draft.meta === 'object' ? op.draft.meta : {},
          t: op.t
        };
        return true;
      case 'final.set': {
        if (!cls || !op.finalId || !core.isISODate(op.date)) return false;
        const week = core.mondayOf(op.date);
        const rec = {
          finalId: op.finalId,
          week,
          date: op.date,
          seats: sanitizeSeats(op.seats),
          layout: core.normalizeLayout(op.layout),
          mode: str(op.mode) || 'random',
          meta: op.meta && typeof op.meta === 'object' ? op.meta : {},
          base: op.base || null,
          t: op.t,
          session: op.s
        };
        const prev = cls.finals.get(week);
        if (prev && op.finalId !== prev.finalId && op.base !== prev.finalId && !sameExcelExport(prev, op)) {
          dropConflict(cls, 'final', week);
          cls.conflicts.push({ kind: 'final', key: week, winnerId: op.finalId, loser: prev, t: op.t });
        } else if (prev) {
          dropConflict(cls, 'final', week);
        }
        cls.finals.set(week, rec);
        return true;
      }
      case 'passage.config.set': {
        if (!cls || !op.config || !op.configId) return false;
        const ids = op.config.sampleIds;
        if (!str(op.config.representativeId) || !Array.isArray(ids) || ids.length < 5 || ids.length > 10 ||
          new Set(ids).size !== ids.length || ids.some(id => !str(id)) || ids.indexOf(op.config.representativeId) >= 0) return false;
        const prev = cls.passageConfig;
        if (prev && op.base !== prev.configId && op.configId !== prev.configId) {
          dropConflict(cls, 'passage-config', 'config');
          cls.conflicts.push({ kind: 'passage-config', key: 'config', winnerId: op.configId, loser: prev, t: op.t });
        } else if (prev) dropConflict(cls, 'passage-config', 'config');
        cls.passageConfig = { configId: op.configId, representativeId: op.config.representativeId, sampleIds: ids.slice(), t: op.t };
        return true;
      }
      case 'passage.final.set': {
        if (!cls || !op.finalId || !core.isISODate(op.week) || core.mondayOf(op.week) !== op.week || !op.result) return false;
        const r = op.result;
        if (!Array.isArray(r.teacher) || !Array.isArray(r.representative) || !Array.isArray(r.groups) ||
          r.teacher.length !== 5 || r.representative.length !== 5 || r.groups.length !== 5 ||
          r.groups.some(g => !g || !str(g.leader) || !Array.isArray(g.members) || g.members.length < 1)) return false;
        const leaders = r.groups.map(g => g.leader);
        const members = r.groups.reduce((a, g) => a.concat(g.members), []);
        const all = r.teacher.concat(r.representative, members);
        const unique = new Set(all);
        if (unique.size < 15 || all.some(id => !str(id)) ||
          new Set(r.teacher).size !== 5 || new Set(r.representative).size !== 5 ||
          new Set(members).size !== members.length || members.some(id => r.teacher.indexOf(id) >= 0 || r.representative.indexOf(id) >= 0) ||
          r.representative.indexOf(r.teacher[0]) >= 0 ||
          new Set(leaders).size !== 5 || leaders.some(id => r.representative.indexOf(id) < 0)) return false;
        const prev = cls.passageFinals.get(op.week);
        if (prev && op.finalId !== prev.finalId && op.base !== prev.finalId) {
          dropConflict(cls, 'passage', op.week);
          cls.conflicts.push({ kind: 'passage', key: op.week, winnerId: op.finalId, loser: prev, t: op.t });
        } else if (prev) dropConflict(cls, 'passage', op.week);
        const rec = {
          finalId: op.finalId, week: op.week, result: {
            teacher: r.teacher.slice(), representative: r.representative.slice(),
            groups: r.groups.map(g => ({ leader: g.leader, members: g.members.slice() }))
          }, base: op.base || null, t: op.t, session: op.s
        };
        cls.passageFinals.set(op.week, rec);
        cls.passageHistory.push(rec);
        return true;
      }
      case 'conflict.ack':
        if (!cls) return false;
        cls.conflicts = cls.conflicts.filter(c => !(c.kind === op.kind2 && c.key === op.key && c.winnerId === op.winnerId));
        return true;
      case 'score.add': {
        const r = op.rec;
        if (!r || !r.recId || st.scores.has(r.recId) || !core.isISODate(r.date) || !Number.isFinite(r.value) || !r.value) return false;
        st.scores.set(r.recId, {
          recId: r.recId,
          classId: str(r.classId),
          studentId: str(r.studentId),
          name: str(r.name),
          item: str(r.item),
          value: r.value,
          date: r.date,
          week: core.mondayOf(r.date),
          time: str(r.time),
          deleted: st.tombstones.has(r.recId),
          t: op.t
        });
        return true;
      }
      case 'score.delete': {
        if (!op.recId) return false;
        st.tombstones.add(op.recId);
        const rec = st.scores.get(op.recId);
        if (rec) rec.deleted = true;
        return true;
      }
      case 'score.edit': {
        const rec = op.recId ? st.scores.get(op.recId) : null;
        if (!rec || !op.patch) return false;
        const p = op.patch;
        if (str(p.item)) rec.item = p.item;
        if (Number.isFinite(p.value) && p.value) rec.value = p.value;
        if (core.isISODate(p.date)) {
          rec.date = p.date;
          rec.week = core.mondayOf(p.date);
        }
        return true;
      }
      case 'items.set': {
        const items = core.normalizeItems(op.items);
        if (!items.length) return false;
        st.items = items;
        return true;
      }
      case 'settings.set':
        if (!op.key) return false;
        st.settings[op.key] = op.value;
        return true;
      case 'excel.import':
        st.imports.set(op.id, { file: str(op.file), summary: str(op.summary), t: op.t });
        return true;
      default:
        return false; // 未知类型（更新版本写的）：保留在文件里，这里忽略
    }
  }

  // 同一张导出表先后几次读回的修改是顺序修改，不算两份各自保存的版本
  function sameExcelExport(prev, op) {
    const a = prev.meta || {};
    const b = op.meta || {};
    return a.from === 'excel' && b.from === 'excel' && !!a.exportId && a.exportId === b.exportId;
  }

  // 应用一条记录；格式异常（手改、残缺副本、新版本改了字段格式）时跳过，不影响其余记录
  function safeApply(st, op) {
    try {
      return { ok: applyOp(st, op), broken: false };
    } catch (e) {
      return { ok: false, broken: true };
    }
  }

  function compareOps(a, b) {
    if (a.t !== b.t) return a.t < b.t ? -1 : 1;
    if (a.id !== b.id) return a.id < b.id ? -1 : 1;
    return 0;
  }

  function reduce(ops) {
    const sorted = (ops || []).filter(o => o && typeof o.id === 'string' && typeof o.t === 'string' && typeof o.k === 'string').sort(compareOps);
    const st = emptyState();
    for (const op of sorted) {
      if (st.ids.has(op.id)) continue;
      st.ids.add(op.id);
      const r = safeApply(st, op);
      if (r.ok) st.opCount++;
      else st.skipped++;
      if (r.broken) st.broken++;
      if (op.t > st.maxT) st.maxT = op.t;
    }
    return st;
  }

  /* =========================================================
     查询
     ========================================================= */

  function listClasses(st) {
    return Array.from(st.classes.values())
      .filter(c => !c.deleted)
      .sort((a, b) => (a.order - b.order) || a.name.localeCompare(b.name, 'zh'));
  }

  function studentsOf(cls, includeInactive) {
    if (!cls) return [];
    return Array.from(cls.students.values())
      .filter(s => includeInactive || s.active !== false)
      .sort(core.compareRoster);
  }

  function layoutOf(cls) {
    if (!cls) return core.normalizeLayout({});
    return cls.layout || core.defaultLayoutFor(studentsOf(cls).length);
  }

  function finalsOf(cls) {
    if (!cls) return [];
    return Array.from(cls.finals.values()).sort((a, b) => (a.week < b.week ? -1 : a.week > b.week ? 1 : 0));
  }

  // 某天正在使用的座次：生效日期不晚于这一天的最近一次定版
  function effectiveFinal(cls, iso) {
    const list = finalsOf(cls).filter(f => f.date <= iso).sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : (a.t < b.t ? -1 : 1)));
    return list.length ? list[list.length - 1] : null;
  }

  // 生成基准：目标日期所在周之前最近的一次定版（同一周的定版不作为自己的基准）
  function sourceFinalFor(cls, targetIso) {
    const wk = core.mondayOf(targetIso);
    const list = finalsOf(cls).filter(f => f.week < wk);
    return list.length ? list[list.length - 1] : null;
  }

  function scoresOf(st, classId) {
    return Array.from(st.scores.values()).filter(r => !r.deleted && r.classId === classId);
  }

  function itemsOf(st) {
    return st.items || core.normalizeItems(core.DEFAULT_SCORE_ITEMS);
  }

  /* =========================================================
     记录文件（.jsonl：一行一条操作；中途被打断最多损失最后一行）
     ========================================================= */

  function parseJsonl(text) {
    const ops = [];
    let bad = 0;
    const lines = String(text || '').split('\n');
    for (const line of lines) {
      const s = line.trim();
      if (!s) continue;
      try {
        const op = JSON.parse(s);
        if (op && typeof op === 'object' && typeof op.id === 'string' && typeof op.t === 'string' && typeof op.k === 'string') ops.push(op);
        else bad++;
      } catch (e) {
        bad++;
      }
    }
    return { ops, bad };
  }

  function serializeOps(ops) {
    return ops.map(o => JSON.stringify(o)).join('\n') + '\n';
  }

  /* =========================================================
     文件夹读写（浏览器 File System Access 句柄；测试中用内存伪目录，工具中用 Node 适配器）
     ========================================================= */

  function describeError(e) {
    const n = e && e.name;
    if (n === 'NotAllowedError' || n === 'SecurityError') return '没有写入权限，请点“连接文件夹”重新授权';
    if (n === 'NoModificationAllowedError' || n === 'InvalidStateError' || n === 'InvalidModificationError') return '文件正被其他程序占用（例如 Excel/WPS 打开着），请关闭后重试';
    if (n === 'NotFoundError') return '找不到文件夹（U 盘是否已拔出？）';
    if (n === 'QuotaExceededError') return '磁盘空间不足';
    return (e && e.message) || String(e);
  }

  async function getSubDir(dir, name, create) {
    try {
      return await dir.getDirectoryHandle(name, { create: !!create });
    } catch (e) {
      if (!create && e && (e.name === 'NotFoundError' || e.name === 'TypeMismatchError')) return null;
      throw e;
    }
  }

  async function getDirPath(root, parts, create) {
    let d = root;
    for (const p of parts) {
      d = await getSubDir(d, p, create);
      if (!d) return null;
    }
    return d;
  }

  async function fileExists(dir, name) {
    try {
      await dir.getFileHandle(name);
      return true;
    } catch (e) {
      return false;
    }
  }

  async function listEntries(dir) {
    const out = [];
    for await (const e of dir.values()) out.push(e);
    return out;
  }

  function byteLength(data) {
    if (typeof data === 'string') return new TextEncoder().encode(data).length;
    if (data && typeof data.byteLength === 'number') return data.byteLength;
    if (data && typeof data.size === 'number') return data.size;
    return null;
  }

  // 写入并核对大小；失败时如实返回错误
  async function writeFileChecked(dir, name, data) {
    try {
      const fh = await dir.getFileHandle(name, { create: true });
      const w = await fh.createWritable();
      await w.write(data);
      await w.close();
      const f = await fh.getFile();
      const expected = byteLength(data);
      if (expected !== null && f.size !== expected) throw new Error(`写入后文件大小不符（${f.size}/${expected}）`);
      return { ok: true, lastModified: f.lastModified, size: f.size };
    } catch (e) {
      return { ok: false, error: describeError(e) };
    }
  }

  async function uniqueName(dir, base, ext) {
    let name = base + ext;
    let i = 2;
    while (await fileExists(dir, name)) {
      name = `${base}-${i}${ext}`;
      i++;
      if (i > 999) throw new Error('无法生成不重复的文件名');
    }
    return name;
  }

  // 入口网页的文件名（第一个是现在的名字，后面是旧名字，仍然认）
  const ENTRY_FILES = ['班级助理.html', '座次管理.html'];

  async function hasDir(dir, name) {
    try {
      await dir.getDirectoryHandle(name);
      return true;
    } catch (e) {
      return false;
    }
  }

  async function isProjectDir(dir, strict) {
    for (const f of ENTRY_FILES) {
      if (await fileExists(dir, f)) return true;
    }
    return !strict && hasDir(dir, DATA_DIR);
  }

  // 老师选中的文件夹 → 项目文件夹。选中项目本身，或选中它的上一级都可以；不依赖文件夹的名字。
  // reason：self 本身就是；child 在子文件夹中找到唯一一个；multiple 子文件夹里有多个；old-project 旧版座次项目；none 没找到
  async function findProjectRoot(handle) {
    if (await isProjectDir(handle, false)) return { root: handle, reason: 'self' };
    const found = [];
    try {
      for await (const e of handle.values()) {
        if (e.kind !== 'directory' || isJunkName(e.name)) continue;
        if (await isProjectDir(e, true)) {
          found.push(e);
          if (found.length > 1) break;
        }
      }
    } catch (e) { /* 列不出子文件夹时按“没找到”处理 */ }
    if (found.length === 1) return { root: found[0], reason: 'child' };
    if (found.length > 1) return { root: null, reason: 'multiple' };
    if (await fileExists(handle, '教室管理.html')) return { root: null, reason: 'old-project' };
    return { root: null, reason: 'none' };
  }

  // 新建文件（绝不覆盖同名文件）
  async function writeNewFile(root, dirParts, base, ext, data) {
    try {
      const dir = await getDirPath(root, dirParts, true);
      const name = await uniqueName(dir, base, ext);
      const res = await writeFileChecked(dir, name, data);
      return Object.assign({ name, path: dirParts.concat(name).join('/') }, res);
    } catch (e) {
      return { ok: false, error: describeError(e) };
    }
  }

  const pad2 = n => String(n).padStart(2, '0');
  function stamp(d) {
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}_${pad2(d.getHours())}${pad2(d.getMinutes())}${pad2(d.getSeconds())}`;
  }

  // 快照文件名里的时间（本地时间）：snapshot_2026-09-29_153000_会话号.json；认不出返回 null
  function snapshotTime(name) {
    const m = /^snapshot_(\d{4})-(\d{2})-(\d{2})_(\d{2})(\d{2})(\d{2})_/.exec(name || '');
    if (!m) return null;
    const ms = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]).getTime();
    return Number.isFinite(ms) ? ms : null;
  }

  // 读取整个项目：快照 + 未被快照覆盖的记录文件
  // opts.now：作为基准的当前时间（毫秒，最好是网络时间）；时间晚于它的快照算“超前”，只在没有正常快照时才用
  async function readProject(root, opts) {
    opts = opts || {};
    const cache = opts.cache || new Map();
    const now = typeof opts.now === 'number' ? opts.now : Date.now();
    const res = { ops: [], files: [], snapshot: null, project: null, hasDataDir: false, badLines: 0, aheadSnapshots: [], warnings: [] };
    const dataDir = await getSubDir(root, DATA_DIR, false);
    if (!dataDir) return res;
    res.hasDataDir = true;

    try {
      const f = await (await dataDir.getFileHandle(PROJECT_FILE)).getFile();
      res.project = JSON.parse(await f.text());
    } catch (e) { /* 没有项目信息文件也能用 */ }

    let covers = new Map();
    const snapDir = await getSubDir(dataDir, SNAP_DIR, false);
    if (snapDir) {
      const all = (await listEntries(snapDir))
        .filter(e => e.kind === 'file' && !isJunkName(e.name) && /\.json$/i.test(e.name))
        .map(e => ({ e, time: snapshotTime(e.name) }));
      const newestFirst = (a, b) => ((b.time || 0) - (a.time || 0)) || (a.e.name < b.e.name ? 1 : -1);
      const ahead = all.filter(x => x.time !== null && x.time > now + CLOCK_TOLERANCE).sort(newestFirst);
      const normal = all.filter(x => ahead.indexOf(x) < 0).sort(newestFirst);
      res.aheadSnapshots = ahead.map(x => x.e.name);
      if (ahead.length) {
        res.warnings.push(`有 ${ahead.length} 个快照（如“${ahead[0].e.name}”）的时间晚于当前时间 ${core.localDateTime(new Date(now))}，` +
          `可能由系统时间超前的电脑写入，或本机时间偏慢；${normal.length ? '已改用时间正常的快照' : '暂时使用其中最新的一个'}。快照只用于加速，可以手动删除`);
      }
      for (const x of normal.concat(ahead)) {
        const e = x.e;
        try {
          const f = await e.getFile();
          const key = `snap|${e.name}|${f.size}|${f.lastModified}`;
          let snap = cache.get(key);
          if (!snap) {
            snap = JSON.parse(await f.text());
            cache.set(key, snap);
          }
          if (!snap || snap.app !== APP || !Array.isArray(snap.ops) || !snap.covers) throw new Error('格式不对');
          for (const op of snap.ops) res.ops.push(op);
          covers = new Map(Object.keys(snap.covers).map(k => [k, snap.covers[k]]));
          res.snapshot = { name: e.name, covers: covers.size, time: x.time };
          break;
        } catch (err) {
          res.warnings.push(`快照“${e.name}”无法读取，已改用原始记录`);
        }
      }
    }

    const entries = (await listEntries(dataDir)).filter(e => e.kind === 'file' && !isJunkName(e.name) && /\.jsonl$/i.test(e.name));
    for (const e of entries) {
      let file;
      try {
        file = await e.getFile();
      } catch (err) {
        res.warnings.push(`无法读取“${e.name}”：${describeError(err)}`);
        continue;
      }
      const info = { name: e.name, size: file.size, lastModified: file.lastModified, covered: false, ops: 0, bad: 0 };
      if (covers.has(e.name) && covers.get(e.name) === file.size) {
        info.covered = true;
        res.files.push(info);
        continue;
      }
      const key = `${e.name}|${file.size}|${file.lastModified}`;
      let parsed = cache.get(key);
      if (!parsed) {
        parsed = parseJsonl(await file.text());
        cache.set(key, parsed);
      }
      info.ops = parsed.ops.length;
      info.bad = parsed.bad;
      res.badLines += parsed.bad;
      for (const op of parsed.ops) res.ops.push(op);
      res.files.push(info);
    }
    return res;
  }

  // 写快照（新文件）：覆盖除本会话外的全部记录文件。
  // 文件名按基准时间 nowMs（最好是网络时间），且一定晚于当前采用的快照，下次读取时会选中它，不会反复重写
  async function maybeWriteSnapshot(root, readResult, ownFileName, sid, force, nowMs) {
    const uncovered = readResult.files.filter(f => !f.covered && f.name !== ownFileName);
    if (!force && uncovered.length < SNAPSHOT_AFTER) return { ok: true, skipped: true };
    const covers = {};
    readResult.files.forEach(f => { if (f.name !== ownFileName) covers[f.name] = f.size; });
    const ownSid = sid;
    const ops = readResult.ops.filter(o => o.s !== ownSid);
    const prev = readResult.snapshot && typeof readResult.snapshot.time === 'number' ? readResult.snapshot.time + 1000 : 0;
    const when = new Date(Math.max(typeof nowMs === 'number' ? nowMs : Date.now(), prev));
    const body = JSON.stringify({ app: APP, v: FORMAT, created: core.localDateTime(when), covers, ops });
    return writeNewFile(root, [DATA_DIR, SNAP_DIR], `snapshot_${stamp(when)}_${sid}`, '.json', body);
  }

  /* =========================================================
     会话写入器：本次打开网页的全部操作写进同一个新文件
     ========================================================= */

  class SessionWriter {
    constructor(opts) {
      this.root = opts.root;
      this.sid = opts.sid;
      this.place = opts.place || 'web';
      this.header = opts.header; // 记录头；可以是函数，第一次写入时才生成（那时多半已校准网络时间）
      this.startedAt = opts.startedAt || new Date();
      this.fileName = null;
      this.ops = [];
      this.written = 0;
      this.lastError = null;
      this.chain = Promise.resolve();
    }

    pendingCount() {
      return this.ops.length - this.written;
    }

    add(op) {
      this.ops.push(op);
    }

    flush() {
      const run = () => this._flush();
      this.chain = this.chain.then(run, run);
      return this.chain;
    }

    async _flush() {
      if (this.written === this.ops.length) return { ok: true, nothing: true };
      const count = this.ops.length;
      try {
        const dataDir = await getSubDir(this.root, DATA_DIR, true);
        if (typeof this.header === 'function') this.header = this.header();
        if (!this.fileName) this.fileName = await uniqueName(dataDir, `${stamp(this.startedAt)}_${this.place}_${this.sid}`, '.jsonl');
        const res = await writeFileChecked(dataDir, this.fileName, serializeOps([this.header].concat(this.ops.slice(0, count))));
        if (!res.ok) throw new Error(res.error);
        this.written = count;
        this.lastError = null;
        return { ok: true };
      } catch (e) {
        this.lastError = describeError(e);
        return { ok: false, error: this.lastError };
      }
    }
  }

  function makeHeader(sid, clock, info) {
    return Object.assign({ k: 'session', id: `${sid}.0`, t: clock.now(), s: sid, app: APP, v: FORMAT }, info || {});
  }

  const CMStore = {
    APP, FORMAT, DATA_DIR, SNAP_DIR, EXPORT_DIR, PROJECT_FILE, SNAPSHOT_AFTER, CLOCK_TOLERANCE,
    isJunkName, tsMake, tsParse, createClock,
    TIME_SOURCES, parseIsoUtc, fetchNetTime,
    emptyState, applyOp, safeApply, reduce, compareOps,
    listClasses, studentsOf, layoutOf, finalsOf, effectiveFinal, sourceFinalFor, scoresOf, itemsOf,
    parseJsonl, serializeOps,
    describeError, getSubDir, getDirPath, fileExists, listEntries, writeFileChecked, uniqueName, writeNewFile, stamp,
    ENTRY_FILES, findProjectRoot,
    snapshotTime, readProject, maybeWriteSnapshot, SessionWriter, makeHeader
  };

  global.CMStore = CMStore;
  if (typeof module !== 'undefined' && module.exports) module.exports = CMStore;
})(typeof window !== 'undefined' ? window : globalThis);
