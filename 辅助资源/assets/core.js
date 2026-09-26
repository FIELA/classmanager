/**
 * CMCore — 班级座次管理 · 纯规则（不读写任何文件）
 * 网页与 Node 测试共用。只用 Chrome 86 已支持的语法和接口（Win7 上 Chrome/Edge 最高 109）。
 * 坐标约定：第 r 排第 c 列，按学生视角；第 1 排靠近讲台，第 1 列在学生左手边。
 * 教师版只在显示时旋转 180°，数据只有这一套坐标。
 */
(function (global) {
  'use strict';

  /* =========================================================
     日期（一律按本地时区）
     ========================================================= */

  const pad2 = n => String(n).padStart(2, '0');

  function localISODate(d) {
    d = d || new Date();
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  }

  function localDateTime(d) {
    d = d || new Date();
    return `${localISODate(d)} ${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
  }

  function parseISODate(iso) {
    const m = String(iso || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!m) return null;
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12);
    return d.getMonth() === Number(m[2]) - 1 ? d : null;
  }

  function isISODate(s) {
    return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')) && !!parseISODate(s);
  }

  function addDays(iso, n) {
    const d = parseISODate(iso);
    d.setDate(d.getDate() + n);
    return localISODate(d);
  }

  function mondayOf(iso) {
    const d = parseISODate(iso);
    if (!d) return '';
    const day = d.getDay();
    d.setDate(d.getDate() + (day === 0 ? -6 : 1 - day));
    return localISODate(d);
  }

  // 教学周跨度（周一～周五），如 “9.28～10.2”
  function weekSpan(iso) {
    const mon = parseISODate(mondayOf(iso));
    if (!mon) return '';
    const fri = new Date(mon);
    fri.setDate(mon.getDate() + 4);
    return `${mon.getMonth() + 1}.${mon.getDate()}～${fri.getMonth() + 1}.${fri.getDate()}`;
  }

  function shortDate(iso) {
    const d = parseISODate(iso);
    return d ? `${d.getMonth() + 1}月${d.getDate()}日` : '';
  }

  /* =========================================================
     随机数、编号、哈希
     ========================================================= */

  // 可复现的伪随机（测试用）
  function mulberry32(seed) {
    let a = seed >>> 0;
    return function () {
      a = (a + 0x6D2B79F5) >>> 0;
      let t = a;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  // 真随机（浏览器与新版 Node 都有 crypto.getRandomValues）
  function cryptoRng() {
    const c = (typeof crypto !== 'undefined' && crypto && typeof crypto.getRandomValues === 'function') ? crypto : null;
    if (!c) return Math.random;
    const buf = new Uint32Array(1);
    return function () {
      c.getRandomValues(buf);
      return buf[0] / 4294967296;
    };
  }

  const defaultRng = cryptoRng();

  const ID_CHARS = 'abcdefghijkmnpqrstuvwxyz23456789';
  function randomId(len, rng) {
    rng = rng || defaultRng;
    let s = '';
    for (let i = 0; i < len; i++) s += ID_CHARS[Math.floor(rng() * ID_CHARS.length)];
    return s;
  }

  // 53 位字符串哈希（cyrb53），用于生成确定性的编号
  function hash53(str, seed) {
    seed = seed || 0;
    let h1 = 0xdeadbeef ^ seed;
    let h2 = 0x41c6ce57 ^ seed;
    for (let i = 0; i < str.length; i++) {
      const ch = str.charCodeAt(i);
      h1 = Math.imul(h1 ^ ch, 2654435761);
      h2 = Math.imul(h2 ^ ch, 1597334677);
    }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507);
    h1 ^= Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507);
    h2 ^= Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
  }

  function shuffle(arr, rng) {
    rng = rng || defaultRng;
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(rng() * (i + 1));
      const t = a[i];
      a[i] = a[j];
      a[j] = t;
    }
    return a;
  }

  const mod = (a, n) => ((a % n) + n) % n;

  /* =========================================================
     座位布局：rows 排 × cols 列；aisles = 走廊所在位置（第 a 列右侧）；disabled = 不可用座位
     ========================================================= */

  const seatKey = (r, c) => `${r},${c}`;

  function parseSeatKey(k) {
    const p = String(k).split(',');
    return { r: Number(p[0]), c: Number(p[1]) };
  }

  function clampInt(v, min, max, def) {
    const n = Math.round(Number(v));
    if (!Number.isFinite(n)) return def;
    return Math.max(min, Math.min(max, n));
  }

  const LIMITS = { rows: 20, cols: 20 };

  function normalizeLayout(l) {
    l = l || {};
    const rows = clampInt(l.rows, 1, LIMITS.rows, 6);
    const cols = clampInt(l.cols, 1, LIMITS.cols, 8);
    const aisles = Array.from(new Set((l.aisles || []).map(Number)))
      .filter(a => Number.isInteger(a) && a >= 1 && a < cols)
      .sort((a, b) => a - b);
    const disabled = Array.from(new Set((l.disabled || []).map(String)))
      .filter(k => {
        const p = parseSeatKey(k);
        return Number.isInteger(p.r) && Number.isInteger(p.c) && p.r >= 1 && p.r <= rows && p.c >= 1 && p.c <= cols;
      })
      .sort((a, b) => {
        const x = parseSeatKey(a);
        const y = parseSeatKey(b);
        return (x.r - y.r) || (x.c - y.c);
      });
    return { rows, cols, aisles, disabled };
  }

  // 布局签名：用于判断两张座次表是否同一布局
  function layoutKey(l) {
    const n = normalizeLayout(l);
    return `${n.rows}x${n.cols}|${n.aisles.join('.')}|${n.disabled.join(';')}`;
  }

  function isSeatAvailable(l, r, c) {
    const n = normalizeLayout(l);
    if (typeof r === 'string') {
      const p = parseSeatKey(r);
      r = p.r;
      c = p.c;
    }
    return r >= 1 && r <= n.rows && c >= 1 && c <= n.cols && n.disabled.indexOf(seatKey(r, c)) < 0;
  }

  function allSeats(l) {
    const n = normalizeLayout(l);
    const dis = new Set(n.disabled);
    const out = [];
    for (let r = 1; r <= n.rows; r++) {
      for (let c = 1; c <= n.cols; c++) {
        if (!dis.has(seatKey(r, c))) out.push({ r, c });
      }
    }
    return out;
  }

  function capacity(l) {
    return allSeats(l).length;
  }

  // 走廊分隔出的“大组”（列号数组）
  function groupsOf(l) {
    const n = normalizeLayout(l);
    const groups = [];
    let cur = [];
    for (let c = 1; c <= n.cols; c++) {
      cur.push(c);
      if (n.aisles.indexOf(c) >= 0 || c === n.cols) {
        groups.push(cur);
        cur = [];
      }
    }
    return groups;
  }

  const ROW_ZONES = ['前', '中', '后'];

  // 前/中/后区：按排数三等分（多出的排归前面），两排时为前/后
  function rowZoneIndex(rows, r) {
    if (rows <= 1) return 0;
    if (rows === 2) return r === 1 ? 0 : 2;
    return Math.min(2, Math.floor((r - 1) * 3 / rows));
  }

  function rowZone(l, r) {
    return ROW_ZONES[rowZoneIndex(normalizeLayout(l).rows, r)];
  }

  // 左/中/右：有两个以上大组时，最左一组为左、最右一组为右、其余为中；没有走廊时按列三等分
  function colZone(l, c) {
    const n = normalizeLayout(l);
    const groups = groupsOf(n);
    if (groups.length >= 2) {
      const gi = groups.findIndex(g => g.indexOf(c) >= 0);
      if (gi === 0) return '左';
      if (gi === groups.length - 1) return '右';
      return '中';
    }
    const side = Math.round(n.cols / 3);
    if (c <= side) return '左';
    if (c > n.cols - side) return '右';
    return '中';
  }

  function zoneOf(l, r, c) {
    return { row: rowZone(l, r), col: colZone(l, c) };
  }

  // 填座顺序：前排优先，同一排从中间向两侧（空座因此留在最后一排两侧）
  function fillOrder(l) {
    const n = normalizeLayout(l);
    const center = (n.cols + 1) / 2;
    return allSeats(n).sort((a, b) =>
      (a.r - b.r) || (Math.abs(a.c - center) - Math.abs(b.c - center)) || (a.c - b.c)
    );
  }

  // 同桌：同一大组内从左起两两成对
  function deskmatePairs(l) {
    const n = normalizeLayout(l);
    const dis = new Set(n.disabled);
    const pairs = [];
    const groups = groupsOf(n);
    for (let r = 1; r <= n.rows; r++) {
      for (const g of groups) {
        for (let i = 0; i + 1 < g.length; i += 2) {
          const a = { r, c: g[i] };
          const b = { r, c: g[i + 1] };
          if (!dis.has(seatKey(a.r, a.c)) && !dis.has(seatKey(b.r, b.c))) pairs.push([a, b]);
        }
      }
    }
    return pairs;
  }

  // 显示用的列序列（含走廊）：学生版从左到右为第 1…N 列；教师版旋转 180°
  function displayColumns(l, view) {
    const n = normalizeLayout(l);
    const aisles = new Set(n.aisles);
    const out = [];
    if (view === 'teacher') {
      for (let c = n.cols; c >= 1; c--) {
        out.push({ type: 'seat', c });
        if (c > 1 && aisles.has(c - 1)) out.push({ type: 'aisle', after: c - 1 });
      }
    } else {
      for (let c = 1; c <= n.cols; c++) {
        out.push({ type: 'seat', c });
        if (c < n.cols && aisles.has(c)) out.push({ type: 'aisle', after: c });
      }
    }
    return out;
  }

  function displayRows(l, view) {
    const n = normalizeLayout(l);
    const rows = [];
    for (let r = 1; r <= n.rows; r++) rows.push(r);
    return view === 'teacher' ? rows.reverse() : rows;
  }

  // 新班级的默认布局：每排 8 列、4 个大组（走廊在第 2、4、6 列右侧），排数刚好坐下
  function defaultLayoutFor(count) {
    return normalizeLayout({ rows: Math.max(1, Math.ceil((count || 0) / 8)), cols: 8, aisles: [2, 4, 6], disabled: [] });
  }

  function seatLabel(r, c) {
    return `第${r}排第${c}列`;
  }

  /* =========================================================
     学生
     ========================================================= */

  // 姓名规范化：去掉首尾空白与零宽字符；纯中文姓名去掉中间的对齐空格（如“张　三”）
  function normalizeName(v) {
    let s = String(v == null ? '' : v).replace(/[\u200b-\u200d\ufeff]/g, '').trim();
    if (!s) return '';
    const compact = s.replace(/[\s\u3000]+/g, '');
    if (/^[\u3400-\u9fff\u00b7\u2022\uff0e]+$/.test(compact)) return compact.replace(/[\u2022\uff0e]/g, '\u00b7');
    return s.replace(/[\s\u3000]+/g, ' ');
  }

  // 性别：返回 'M' / 'F'；空值返回 ''；无法识别返回 null
  function normalizeGender(v) {
    const s = String(v == null ? '' : v).trim().toLowerCase();
    if (!s) return '';
    if (/^(男|男生|男性|m|male|boy|♂)$/.test(s)) return 'M';
    if (/^(女|女生|女性|f|female|girl|♀)$/.test(s)) return 'F';
    return null;
  }

  const GENDER_TEXT = { M: '男', F: '女', '': '' };

  function activeStudents(students) {
    return (students || []).filter(s => s && s.active !== false);
  }

  const seqOf = s => (typeof s.seq === 'number' && Number.isFinite(s.seq) ? s.seq : Infinity);

  function compareRoster(a, b) {
    const x = seqOf(a);
    const y = seqOf(b);
    if (x !== y) return x < y ? -1 : 1;
    return String(a.name || '').localeCompare(String(b.name || ''), 'zh');
  }

  // 按成绩排名：有成绩按成绩从高到低；全班都没有成绩时按名次；没有的排在最后；并列按序号
  function rankStudents(list) {
    const useScore = (list || []).some(s => typeof s.score === 'number');
    const key = s => {
      if (useScore) return typeof s.score === 'number' ? s.score : null;
      return typeof s.rank === 'number' ? -s.rank : null;
    };
    return (list || []).slice().sort((a, b) => {
      const ka = key(a);
      const kb = key(b);
      if (ka === null && kb !== null) return 1;
      if (kb === null && ka !== null) return -1;
      if (ka !== null && kb !== null && ka !== kb) return kb - ka;
      return compareRoster(a, b);
    });
  }

  // 同名学生在界面上加区分（序号/学号）
  function displayNames(students) {
    const count = new Map();
    (students || []).forEach(s => count.set(s.name, (count.get(s.name) || 0) + 1));
    const out = new Map();
    (students || []).forEach(s => {
      if (count.get(s.name) > 1) {
        const tag = s.no ? s.no : (typeof s.seq === 'number' ? `${s.seq}号` : s.id.slice(-3));
        out.set(s.id, `${s.name}(${tag})`);
      } else {
        out.set(s.id, s.name);
      }
    });
    return out;
  }

  /* =========================================================
     排座：seats 统一为 [[排, 列, 学生编号], …]
     ========================================================= */

  function seatsToMap(seats) {
    const m = new Map();
    (seats || []).forEach(t => m.set(seatKey(t[0], t[1]), t[2]));
    return m;
  }

  function mapToSeats(map) {
    const out = [];
    map.forEach((sid, k) => {
      const p = parseSeatKey(k);
      out.push([p.r, p.c, sid]);
    });
    return out.sort((a, b) => (a[0] - b[0]) || (a[1] - b[1]));
  }

  // 固定学生：去掉离班学生、无效座位、重复项
  function sanitizePins(pins, layout, students) {
    const L = normalizeLayout(layout);
    const act = new Set(activeStudents(students).map(s => s.id));
    const usedSeat = new Set();
    const usedSid = new Set();
    const valid = [];
    const dropped = [];
    for (const p of pins || []) {
      const k = seatKey(p.r, p.c);
      if (!p || !act.has(p.sid) || !isSeatAvailable(L, p.r, p.c) || usedSeat.has(k) || usedSid.has(p.sid)) {
        if (p) dropped.push(p);
        continue;
      }
      usedSeat.add(k);
      usedSid.add(p.sid);
      valid.push({ sid: p.sid, r: p.r, c: p.c });
    }
    return { valid, dropped };
  }

  function capacityError(need, have) {
    return new Error(`座位不够：还有 ${need} 名学生要安排，只剩 ${have} 个空座。请在“班级设置”里增加排数或列数`);
  }

  // 把剩余学生按“前排中间优先”依次放进空座
  function placeLeftovers(map, sids, layout) {
    if (!sids.length) return;
    const free = fillOrder(layout).filter(s => !map.has(seatKey(s.r, s.c)));
    if (sids.length > free.length) throw capacityError(sids.length, free.length);
    sids.forEach((sid, i) => map.set(seatKey(free[i].r, free[i].c), sid));
  }

  function startWithPins(students, layout, pins) {
    const L = normalizeLayout(layout);
    const act = activeStudents(students);
    const pinInfo = sanitizePins(pins, L, act);
    const map = new Map();
    pinInfo.valid.forEach(p => map.set(seatKey(p.r, p.c), p.sid));
    const pinned = new Set(pinInfo.valid.map(p => p.sid));
    return { L, act, map, pinned, pinInfo };
  }

  function pinNotes(pinInfo) {
    return pinInfo.dropped.length ? [`有 ${pinInfo.dropped.length} 个固定座位已失效（学生离班或座位已不存在），本次未使用`] : [];
  }

  // 随机排座；opts.genderPair = 同桌尽量一男一女
  function randomSeating(students, layout, pins, opts, rng) {
    opts = opts || {};
    rng = rng || defaultRng;
    const st = startWithPins(students, layout, pins);
    const toPlace = st.act.filter(s => !st.pinned.has(s.id));
    const free = fillOrder(st.L).filter(s => !st.map.has(seatKey(s.r, s.c)));
    if (toPlace.length > free.length) throw capacityError(toPlace.length, free.length);
    const used = free.slice(0, toPlace.length);
    const map = st.map;
    const notes = pinNotes(st.pinInfo);
    if (opts.genderPair) {
      const usedSet = new Set(used.map(s => seatKey(s.r, s.c)));
      const pairs = shuffle(deskmatePairs(st.L).filter(pr =>
        usedSet.has(seatKey(pr[0].r, pr[0].c)) && usedSet.has(seatKey(pr[1].r, pr[1].c))
      ), rng);
      const males = shuffle(toPlace.filter(s => s.gender === 'M'), rng);
      const females = shuffle(toPlace.filter(s => s.gender === 'F'), rng);
      const done = new Set();
      let mixed = 0;
      for (const pr of pairs) {
        if (!males.length || !females.length) break;
        const m = males.pop();
        const f = females.pop();
        const two = rng() < 0.5 ? [m, f] : [f, m];
        map.set(seatKey(pr[0].r, pr[0].c), two[0].id);
        map.set(seatKey(pr[1].r, pr[1].c), two[1].id);
        done.add(m.id);
        done.add(f.id);
        mixed++;
      }
      const restStudents = shuffle(toPlace.filter(s => !done.has(s.id)), rng);
      const restSeats = shuffle(used.filter(s => !map.has(seatKey(s.r, s.c))), rng);
      restStudents.forEach((s, i) => map.set(seatKey(restSeats[i].r, restSeats[i].c), s.id));
      const unknown = toPlace.filter(s => s.gender !== 'M' && s.gender !== 'F').length;
      notes.push(`男女搭配：${mixed} 对同桌为一男一女${unknown ? `；另有 ${unknown} 人未标注性别，按随机安排` : ''}`);
    } else {
      const ss = shuffle(toPlace, rng);
      const seats = shuffle(used, rng);
      ss.forEach((s, i) => map.set(seatKey(seats[i].r, seats[i].c), s.id));
    }
    return { seats: mapToSeats(map), notes };
  }

  // 按成绩自动排：名次靠前的优先坐前排中间
  function scoreSeating(students, layout, pins) {
    const st = startWithPins(students, layout, pins);
    const ranked = rankStudents(st.act.filter(s => !st.pinned.has(s.id)));
    const free = fillOrder(st.L).filter(s => !st.map.has(seatKey(s.r, s.c)));
    if (ranked.length > free.length) throw capacityError(ranked.length, free.length);
    ranked.forEach((s, i) => st.map.set(seatKey(free[i].r, free[i].c), s.id));
    const noScore = ranked.filter(s => typeof s.score !== 'number' && typeof s.rank !== 'number').length;
    const notes = pinNotes(st.pinInfo);
    if (noScore) notes.push(`${noScore} 人没有成绩，排在最后`);
    return { seats: mapToSeats(st.map), notes };
  }

  function defaultRotation(layout) {
    const n = normalizeLayout(layout);
    const rowShift = n.rows >= 3 ? Math.max(1, Math.round(n.rows / 3)) : (n.rows === 2 ? 1 : 0);
    const groups = groupsOf(n);
    const colShift = groups.length >= 2 ? groups[0].length : (n.cols >= 4 ? 2 : (n.cols >= 2 ? 1 : 0));
    return { rowShift, colShift, rowDir: 'back', colDir: 'right' };
  }

  // 从上一张座次表出发的通用准备：在班且未固定的学生保留原座；被固定座位占用或座位已不存在的学生待重新安排
  function carryOver(source, st) {
    const src = seatsToMap(source && source.seats);
    const actIds = new Set(st.act.map(s => s.id));
    const pinnedSeats = new Set(st.pinInfo.valid.map(p => seatKey(p.r, p.c)));
    const kept = new Map();
    const displaced = [];
    const seen = new Set();
    src.forEach((sid, k) => {
      if (!actIds.has(sid) || st.pinned.has(sid) || seen.has(sid)) return;
      seen.add(sid);
      if (pinnedSeats.has(k) || !isSeatAvailable(st.L, k)) displaced.push(sid);
      else kept.set(k, sid);
    });
    const newcomers = st.act.filter(s => !seen.has(s.id) && !st.pinned.has(s.id)).map(s => s.id);
    return { kept, displaced, newcomers };
  }

  function leftoverNotes(displaced, newcomers) {
    const notes = [];
    if (newcomers.length) notes.push(`${newcomers.length} 名学生上次没有座位（新加入），已安排到空座`);
    if (displaced.length) notes.push(`${displaced.length} 名学生的原座位被固定学生占用或已不存在，已安排到空座`);
    return notes;
  }

  // 轮换：前后方向每列内循环移动 rowShift 位，左右方向每排内循环移动 colShift 位；空座与固定学生不动
  function rotateSeating(source, students, layout, pins, opts) {
    const st = startWithPins(students, layout, pins);
    if (!source || !source.seats) throw new Error('没有可作为基准的座次表');
    if (layoutKey(source.layout) !== layoutKey(st.L)) {
      throw new Error('座位布局在上一次定版之后改过，无法按原位置轮换。请改用随机、按成绩或沿用');
    }
    const o = Object.assign(defaultRotation(st.L), opts || {});
    const rowShift = Math.max(0, Math.round(Number(o.rowShift) || 0));
    const colShift = Math.max(0, Math.round(Number(o.colShift) || 0));
    const co = carryOver(source, st);
    const positions = [];
    co.kept.forEach((sid, k) => positions.push(parseSeatKey(k)));

    const step1 = new Map();
    const byCol = new Map();
    positions.forEach(p => {
      if (!byCol.has(p.c)) byCol.set(p.c, []);
      byCol.get(p.c).push(p);
    });
    byCol.forEach(list => {
      list.sort((a, b) => a.r - b.r);
      const n = list.length;
      const k = mod(o.rowDir === 'front' ? -rowShift : rowShift, n);
      list.forEach((p, i) => step1.set(seatKey(p.r, p.c), seatKey(list[(i + k) % n].r, p.c)));
    });

    const step2 = new Map();
    const byRow = new Map();
    positions.forEach(p => {
      if (!byRow.has(p.r)) byRow.set(p.r, []);
      byRow.get(p.r).push(p);
    });
    byRow.forEach(list => {
      list.sort((a, b) => a.c - b.c);
      const n = list.length;
      const k = mod(o.colDir === 'left' ? -colShift : colShift, n);
      list.forEach((p, i) => step2.set(seatKey(p.r, p.c), seatKey(p.r, list[(i + k) % n].c)));
    });

    const map = st.map;
    co.kept.forEach((sid, k) => map.set(step2.get(step1.get(k)), sid));
    placeLeftovers(map, co.displaced.concat(co.newcomers), st.L);
    const notes = pinNotes(st.pinInfo).concat(leftoverNotes(co.displaced, co.newcomers));
    if (!rowShift && !colShift) notes.push('前后、左右的移动距离都是 0，座位没有变化');
    return { seats: mapToSeats(map), notes, rotation: { rowShift, colShift, rowDir: o.rowDir, colDir: o.colDir } };
  }

  // 沿用：座位原样保留（布局改过时，已不存在的座位上的学生补到空座）
  function holdSeating(source, students, layout, pins) {
    const st = startWithPins(students, layout, pins);
    if (!source || !source.seats) throw new Error('没有可沿用的座次表');
    const co = carryOver(source, st);
    co.kept.forEach((sid, k) => st.map.set(k, sid));
    placeLeftovers(st.map, co.displaced.concat(co.newcomers), st.L);
    return { seats: mapToSeats(st.map), notes: pinNotes(st.pinInfo).concat(leftoverNotes(co.displaced, co.newcomers)) };
  }

  // 校验：座位在布局内、不重复；学生在名单内、不重复；requireAll 时在班学生必须都有座位
  function validateSeats(seats, layout, students, opts) {
    opts = opts || {};
    const L = normalizeLayout(layout);
    const byId = new Map((students || []).map(s => [s.id, s]));
    const errors = [];
    const seenSeat = new Set();
    const seenSid = new Set();
    for (const t of seats || []) {
      const k = seatKey(t[0], t[1]);
      if (!isSeatAvailable(L, t[0], t[1])) errors.push(`${seatLabel(t[0], t[1])}不是可用座位`);
      if (seenSeat.has(k)) errors.push(`${seatLabel(t[0], t[1])}重复安排`);
      seenSeat.add(k);
      const s = byId.get(t[2]);
      if (!s) errors.push(`${seatLabel(t[0], t[1])}上的学生不在名单中`);
      else if (s.active === false) errors.push(`${s.name}已离班，却仍有座位`);
      if (seenSid.has(t[2])) errors.push(`${s ? s.name : '某学生'}被安排了两个座位`);
      seenSid.add(t[2]);
    }
    const unseated = activeStudents(students).filter(s => !seenSid.has(s.id));
    if (opts.requireAll && unseated.length) {
      errors.push(`还有 ${unseated.length} 名在班学生没有座位：${unseated.slice(0, 5).map(s => s.name).join('、')}${unseated.length > 5 ? '等' : ''}`);
    }
    return { ok: !errors.length, errors, unseated: unseated.map(s => s.id) };
  }

  /* =========================================================
     随机抽取
     ========================================================= */

  // 从 pool（学生编号）中不重复地抽 n 人；exclude 中的人不参与
  function drawStudents(pool, n, rng, exclude) {
    const ex = new Set(exclude || []);
    const cand = (pool || []).filter(id => !ex.has(id));
    const k = Math.max(0, Math.min(Math.floor(Number(n) || 0), cand.length));
    return shuffle(cand, rng).slice(0, k);
  }

  /* =========================================================
     课堂表现
     ========================================================= */

  // 默认表现项（与旧项目一致）；老师可在“班级设置”中修改
  const DEFAULT_SCORE_ITEMS = [
    { id: 'b1', name: '优秀作业', value: 1 },
    { id: 'b2', name: '优秀课堂表现', value: 1 },
    { id: 'p1', name: '作业未交', value: -3, row: 1 },
    { id: 'p2', name: '作业未完成', value: -1, row: 1 },
    { id: 'p3', name: '不按座次表就坐', value: -1, row: 1 },
    { id: 'p4', name: '迟到', value: -3, row: 1 },
    { id: 'p5', name: '早退', value: -3, row: 1 },
    { id: 'p6', name: '旷课', value: -10, row: 1 },
    { id: 'p7', name: '睡觉', value: -3, row: 2 },
    { id: 'p8', name: '说话', value: -3, row: 2 },
    { id: 'p9', name: '卫生问题', value: -1, row: 2 },
    { id: 'p10', name: '损坏公物', value: -5, row: 2 },
    { id: 'p11', name: '脏话', value: -2, row: 2 },
    { id: 'p12', name: '其他', value: -3, row: 2 }
  ];

  function normalizeItems(items) {
    const out = [];
    const seen = new Set();
    for (const it of items || []) {
      const name = String((it && it.name) || '').trim();
      const value = Number(it && it.value);
      if (!name || !Number.isFinite(value) || value === 0 || seen.has(name)) continue;
      seen.add(name);
      out.push({ id: String(it.id || 'i_' + hash53(name)), name, value, row: value < 0 ? (it.row === 2 ? 2 : 1) : undefined });
    }
    return out;
  }

  function summarizeScores(records) {
    const map = new Map();
    for (const r of records || []) {
      if (!r || r.deleted || !Number.isFinite(r.value) || !r.value) continue;
      if (!map.has(r.studentId)) map.set(r.studentId, { bonusCount: 0, bonusSum: 0, penaltyCount: 0, penaltySum: 0, net: 0 });
      const s = map.get(r.studentId);
      if (r.value > 0) {
        s.bonusCount++;
        s.bonusSum += r.value;
      } else {
        s.penaltyCount++;
        s.penaltySum += r.value;
      }
      s.net = s.bonusSum + s.penaltySum;
    }
    return map;
  }

  const fmtScore = v => (v > 0 ? `+${v}` : `${v}`);

  const CMCore = {
    // 日期
    localISODate, localDateTime, parseISODate, isISODate, addDays, mondayOf, weekSpan, shortDate,
    // 随机与编号
    mulberry32, cryptoRng, defaultRng, randomId, hash53, shuffle,
    // 布局
    LIMITS, seatKey, parseSeatKey, normalizeLayout, layoutKey, isSeatAvailable, allSeats, capacity, groupsOf,
    ROW_ZONES, rowZoneIndex, rowZone, colZone, zoneOf, fillOrder, deskmatePairs, displayColumns, displayRows,
    defaultLayoutFor, seatLabel,
    // 学生
    normalizeName, normalizeGender, GENDER_TEXT, activeStudents, compareRoster, rankStudents, displayNames,
    // 排座
    seatsToMap, mapToSeats, sanitizePins, randomSeating, scoreSeating, defaultRotation, rotateSeating, holdSeating,
    placeLeftovers, validateSeats,
    // 抽取
    drawStudents,
    // 课堂表现
    DEFAULT_SCORE_ITEMS, normalizeItems, summarizeScores, fmtScore
  };

  global.CMCore = CMCore;
  if (typeof module !== 'undefined' && module.exports) module.exports = CMCore;
})(typeof window !== 'undefined' ? window : globalThis);
