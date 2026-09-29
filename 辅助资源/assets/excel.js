/**
 * CMExcel — 名单导入、Excel 导出、读回老师在导出表里的修改
 *
 * 导出的表格都是“新文件”，网页从不覆盖任何 Excel。
 * 每个导出表里有一张隐藏表 _classmanager，保存导出那一刻的原始内容；
 * 读回时拿表格现在的内容与它比对，只把“老师改动的部分”变成记录，所以旧表不会把新数据改回去。
 * 读回得到的记录编号由文件内容确定：同一处修改在 Mac 和教室电脑上读到，编号相同，只算一次。
 */
(function (global) {
  'use strict';

  const core = global.CMCore || (typeof require === 'function' ? require('./core.js') : null);

  function getExcelJS() {
    if (typeof window !== 'undefined' && window.ExcelJS) return window.ExcelJS;
    if (typeof globalThis !== 'undefined' && globalThis.ExcelJS) return globalThis.ExcelJS;
    if (typeof require === 'function') {
      try { return require('./exceljs.min.js'); } catch (e) { /* ignore */ }
    }
    throw new Error('未加载 ExcelJS 引擎');
  }

  const META_SHEET = '_classmanager';
  const APP = 'classmanager';
  const COLOR = { ink: 'FF0F172A', M: 'FF1D4ED8', F: 'FFDC2626', title: 'FF173D34', muted: 'FF64748B', bonus: 'FF16A34A', penalty: 'FFDC2626' };
  const FONT = '微软雅黑';
  const pad2 = n => String(n).padStart(2, '0');

  /* =========================================================
     单元格读取
     ========================================================= */

  function excelDateText(d) {
    const date = `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
    const hasTime = d.getUTCHours() || d.getUTCMinutes() || d.getUTCSeconds();
    return hasTime ? `${date} ${pad2(d.getUTCHours())}:${pad2(d.getUTCMinutes())}:${pad2(d.getUTCSeconds())}` : date;
  }

  function cellValue(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v;
    if (typeof v === 'object') {
      if ('result' in v) return cellValue(v.result);
      if (Array.isArray(v.richText)) return v.richText.map(t => t.text).join('');
      if ('text' in v) return cellValue(v.text);
      if ('error' in v) return '';
      if ('formula' in v || 'sharedFormula' in v) return '';
    }
    return v;
  }

  function cellText(v) {
    const x = cellValue(v);
    if (x instanceof Date) return excelDateText(x);
    return String(x).trim();
  }

  function parseNumber(v) {
    const x = cellValue(v);
    if (typeof x === 'number') return Number.isFinite(x) ? x : null;
    const s = String(x || '').replace(/＋/g, '+').replace(/[－—–]/g, '-').replace(/[分\s]/g, '');
    return /^[+-]?\d+(\.\d+)?$/.test(s) ? Number(s) : null;
  }

  function normalizeDate(v) {
    const x = cellValue(v);
    if (x instanceof Date) return excelDateText(x).slice(0, 10);
    const s = String(x || '').trim();
    const m = s.match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})/);
    return m ? `${m[1]}-${pad2(m[2])}-${pad2(m[3])}` : '';
  }

  /* =========================================================
     名单导入：表格（xlsx）、CSV、直接粘贴
     ========================================================= */

  const HEADER = {
    name: /^(学生)?姓名$|^名字$/,
    gender: /^(性别|男女|gender|sex)$/i,
    score: /(成绩|分数|总分|得分)$|^(语文|数学|英语|物理|化学|生物|历史|政治|地理)$/,
    rank: /(名次|排名)$/,
    no: /^(学号|考号|学籍号|准考证号)$/,
    seq: /^(序号|座号|编号)$/,
    cls: /^(班级|班别|班)$/,
    status: /^(状态|在班状态)$/
  };

  const cleanHeader = s => String(s || '').replace(/[\s\u3000]/g, '').replace(/[（]/g, '(').replace(/[）]/g, ')').toLowerCase();

  function normalizeClassName(s) {
    return String(s || '').replace(/[\s\u3000]/g, '').replace(/[（(]\s*(\d+)\s*[)）]/g, '$1').trim();
  }

  function classNameFromText(s) {
    const t = normalizeClassName(s);
    const m = t.match(/((?:高|初)[一二三]|[一二三四五六七八九]年级)?(\d{1,2})班/);
    return m ? `${m[1] || ''}${m[2]}班` : '';
  }

  const SKIP_ROW = /^(合计|总计|平均|平均分|备注|说明|注[:：]?|统计)/;
  const looksLikeName = s => /^[\u3400-\u9fff\u00b7]{2,6}$/.test(core.normalizeName(s));

  // rows：二维文本数组。返回按班级分组的学生
  function parseRosterRows(rows, opts) {
    opts = opts || {};
    const warnings = [];
    let h = -1;
    const cols = {};
    for (let i = 0; i < Math.min(rows.length, 12) && h < 0; i++) {
      const row = rows[i] || [];
      for (let j = 0; j < row.length; j++) {
        if (HEADER.name.test(cleanHeader(row[j]))) { h = i; break; }
      }
    }
    if (h >= 0) {
      const row = rows[h];
      for (let j = 0; j < row.length; j++) {
        const t = cleanHeader(row[j]);
        if (!t) continue;
        for (const key of ['name', 'gender', 'rank', 'no', 'seq', 'cls', 'status', 'score']) {
          if (cols[key] === undefined && HEADER[key].test(t)) {
            if (key === 'score' && HEADER.rank.test(t)) continue;
            cols[key] = j;
            break;
          }
        }
      }
    } else {
      // 没有表头：找“最像姓名”的一列
      let best = -1;
      let bestCount = 0;
      const width = Math.max(0, ...rows.slice(0, 200).map(r => (r || []).length));
      for (let j = 0; j < width; j++) {
        const n = rows.filter(r => r && looksLikeName(r[j])).length;
        if (n > bestCount) { best = j; bestCount = n; }
      }
      if (best < 0 || !bestCount) return { groups: [], columns: {}, headerRow: -1, warnings: ['没有找到“姓名”列'] };
      cols.name = best;
      for (let j = 0; j < width; j++) {
        if (j === best) continue;
        const vals = rows.map(r => (r ? r[j] : '')).filter(v => String(v || '').trim());
        if (vals.length && vals.every(v => core.normalizeGender(v))) { cols.gender = j; break; }
      }
      warnings.push('没有找到表头，已把最像姓名的一列当作姓名');
    }

    // 标题里的班级名（如“高二1班开学考地理成绩表”）
    let titleClass = '';
    for (let i = 0; i < (h >= 0 ? h : 0) && !titleClass; i++) {
      for (const cell of rows[i] || []) {
        titleClass = classNameFromText(cell);
        if (titleClass) break;
      }
    }

    // 每个班单独统计（有“班级”列时一次会导入多个班）
    const groups = new Map();
    const start = h >= 0 ? h + 1 : 0;
    for (let i = start; i < rows.length; i++) {
      const row = rows[i] || [];
      const name = core.normalizeName(row[cols.name]);
      if (!name || SKIP_ROW.test(name) || (h < 0 && !looksLikeName(name))) continue;
      // “班级”列的内容原样作为班级名（如“测试1班”“高二(1)班”→“高二1班”）；只有数字时补“班”
      const raw = cols.cls !== undefined ? normalizeClassName(row[cols.cls]) : '';
      const key = (/^\d{1,2}$/.test(raw) ? raw + '班' : raw) || titleClass || opts.className || '';
      if (!groups.has(key)) groups.set(key, { students: [], badGender: 0, noScore: 0 });
      const g = groups.get(key);
      const stu = { name };
      if (cols.gender !== undefined) {
        const gd = core.normalizeGender(row[cols.gender]);
        if (gd === null) { g.badGender++; stu.gender = ''; } else stu.gender = gd;
      }
      if (cols.score !== undefined) {
        stu.score = parseNumber(row[cols.score]);
        if (stu.score === null) g.noScore++;
      }
      if (cols.rank !== undefined) stu.rank = parseNumber(row[cols.rank]);
      if (cols.no !== undefined) stu.no = String(row[cols.no] == null ? '' : row[cols.no]).trim();
      if (cols.seq !== undefined) stu.seq = parseNumber(row[cols.seq]);
      if (cols.status !== undefined && /离/.test(String(row[cols.status] || ''))) stu.active = false;
      g.students.push(stu);
    }
    const out = [];
    groups.forEach((g, key) => {
      const merged = mergeSameName(g.students);
      const w = [];
      if (g.badGender) w.push(`${g.badGender} 人的性别无法识别，已留空`);
      if (g.noScore) w.push(`${g.noScore} 人没有成绩（缺考或空白）`);
      if (merged.dup) w.push(`有 ${merged.dup} 组同名学生，已按同一名学生合并`);
      out.push({ className: key, students: merged.students, warnings: w });
    });
    return { groups: out, columns: cols, headerRow: h, warnings };
  }

  // 同一班级里的同名学生按同一人合并：后面的行补充/更新前面的字段；只要有一行在班，就算在班
  function mergeSameName(list) {
    const byName = new Map();
    const students = [];
    let dup = 0;
    for (const s of list) {
      const prev = byName.get(s.name);
      if (!prev) {
        const copy = Object.assign({}, s);
        byName.set(s.name, copy);
        students.push(copy);
        continue;
      }
      if (!prev.dupCounted) { dup++; prev.dupCounted = true; }
      Object.keys(s).forEach(k => {
        if (k === 'active') return;
        const v = s[k];
        if (v !== undefined && v !== null && v !== '') prev[k] = v;
      });
      if (s.active !== false || prev.active !== false) delete prev.active;
    }
    students.forEach(s => { delete s.dupCounted; });
    return { students, dup };
  }

  function sheetRows(ws, maxRows, maxCols) {
    const rows = [];
    const R = Math.min(ws.rowCount || 0, maxRows || 3000);
    const C = Math.min(ws.columnCount || 0, maxCols || 40);
    for (let r = 1; r <= R; r++) {
      const row = [];
      for (let c = 1; c <= C; c++) row.push(cellText(ws.getCell(r, c).value));
      rows.push(row);
    }
    return rows;
  }

  // 解析名单工作簿：每张可见工作表按班级分组；返回候选班级列表
  async function parseRosterWorkbook(buffer, fileName) {
    const Excel = getExcelJS();
    const wb = new Excel.Workbook();
    await wb.xlsx.load(buffer);
    const candidates = [];
    const warnings = [];
    for (const ws of wb.worksheets) {
      if (ws.state && ws.state !== 'visible') continue;
      if (ws.name === META_SHEET) continue;
      const parsed = parseRosterRows(sheetRows(ws), { className: classNameFromText(ws.name) || classNameFromText(fileName) || '' });
      parsed.groups.forEach(g => {
        if (!g.students.length) return;
        candidates.push({
          source: `${fileName || ''}${wb.worksheets.length > 1 ? ` · ${ws.name}` : ''}`,
          className: g.className || classNameFromText(ws.name) || classNameFromText(fileName) || String(ws.name || '新班级'),
          students: g.students,
          columns: parsed.columns,
          warnings: parsed.warnings.concat(g.warnings)
        });
      });
      if (!parsed.groups.length && parsed.warnings.length) warnings.push(`${ws.name}：${parsed.warnings.join('；')}`);
    }
    return { candidates, warnings };
  }

  // 文本解码：先按 UTF-8，不合法再按 GBK（Excel 在中文 Windows 上另存的 CSV 多为 GBK）
  function decodeText(buffer) {
    const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    let text;
    try {
      text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    } catch (e) {
      text = new TextDecoder('gbk').decode(bytes);
    }
    return text.replace(/^\ufeff/, '');
  }

  function parseCsvLine(line) {
    const out = [];
    let cur = '';
    let q = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (q) {
        if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',' || ch === '，') { out.push(cur); cur = ''; } else cur += ch;
    }
    out.push(cur);
    return out;
  }

  function splitTable(text) {
    const lines = String(text || '').replace(/\r\n?/g, '\n').split('\n').filter(l => l.trim());
    if (lines.some(l => l.indexOf('\t') >= 0)) return lines.map(l => l.split('\t'));
    if (lines.some(l => /[,，]/.test(l))) return lines.map(parseCsvLine);
    return lines.map(l => l.trim().split(/ +/));
  }

  function parseRosterText(text, className) {
    const parsed = parseRosterRows(splitTable(text), { className: className || '' });
    return {
      candidates: parsed.groups.filter(g => g.students.length).map(g => ({
        source: '粘贴/文本',
        className: g.className || className || '新班级',
        students: g.students,
        columns: parsed.columns,
        warnings: parsed.warnings.concat(g.warnings)
      })),
      warnings: parsed.groups.length ? [] : parsed.warnings
    };
  }

  /* =========================================================
     隐藏的原始内容（分段存放，单元格上限约 3.2 万字）
     ========================================================= */

  function writeMeta(wb, meta, base) {
    const ws = wb.addWorksheet(META_SHEET);
    ws.state = 'hidden';
    const text = JSON.stringify(base);
    const chunks = [];
    for (let i = 0; i < text.length; i += 30000) chunks.push(text.slice(i, i + 30000));
    ws.getCell(1, 1).value = JSON.stringify(Object.assign({ app: APP, v: 1, chunks: chunks.length }, meta));
    chunks.forEach((c, i) => { ws.getCell(i + 2, 1).value = c; });
  }

  function readMeta(wb) {
    const ws = wb.getWorksheet(META_SHEET);
    if (!ws) return null;
    try {
      const head = JSON.parse(cellText(ws.getCell(1, 1).value));
      if (!head || head.app !== APP) return null;
      let text = '';
      for (let i = 0; i < (head.chunks || 0); i++) text += cellText(ws.getCell(i + 2, 1).value);
      head.base = text ? JSON.parse(text) : null;
      return head;
    } catch (e) {
      return { broken: true };
    }
  }

  /* =========================================================
     导出通用样式
     ========================================================= */

  const thin = { style: 'thin', color: { argb: 'FFB4C7BA' } };
  const BORDER = { top: thin, left: thin, bottom: thin, right: thin };
  const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const PAGE = { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 1, margins: { left: 0.3, right: 0.3, top: 0.3, bottom: 0.3, header: 0.2, footer: 0.2 } };

  function newWorkbook() {
    const Excel = getExcelJS();
    const wb = new Excel.Workbook();
    wb.creator = '班级助理';
    wb.lastModifiedBy = '班级助理';
    wb.created = new Date();
    wb.modified = new Date();
    return wb;
  }

  function titleCell(ws, r, c1, c2, text, size) {
    if (c2 > c1) ws.mergeCells(r, c1, r, c2);
    const cell = ws.getCell(r, c1);
    cell.value = text;
    cell.font = { name: FONT, size: size || 16, bold: true, color: { argb: COLOR.title } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    return cell;
  }

  function noteCell(ws, r, c1, c2, text) {
    if (c2 > c1) ws.mergeCells(r, c1, r, c2);
    const cell = ws.getCell(r, c1);
    cell.value = text;
    cell.font = { name: FONT, size: 10, color: { argb: COLOR.muted } };
    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    return cell;
  }

  function headerRow(ws, r, values) {
    const row = ws.getRow(r);
    row.values = values;
    row.height = 26;
    row.eachCell(cell => {
      cell.font = { name: FONT, size: 11, bold: true, color: { argb: COLOR.title } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = fill('FFEBF2EE');
      cell.border = BORDER;
    });
  }

  function findHeader(ws, required) {
    const limit = Math.min(ws.rowCount || 0, 10);
    for (let r = 1; r <= limit; r++) {
      const found = {};
      ws.getRow(r).eachCell((cell, c) => {
        const t = cleanHeader(cellText(cell.value));
        if (t) found[t] = c;
      });
      if (required.every(h => found[cleanHeader(h)])) return { row: r, col: h => found[cleanHeader(h)] };
    }
    return null;
  }

  /* =========================================================
     1. 名单导出（可在 Excel 里改姓名、性别、成绩、学号，增删学生）
     ========================================================= */

  const ROSTER_ID_HEADER = '系统编号（勿改）';

  function rosterBase(students) {
    return students.map(s => ({ id: s.id, name: s.name, gender: s.gender || '', score: s.score, rank: s.rank, no: s.no || '', seq: s.seq, active: s.active !== false }));
  }

  async function createRosterWorkbook(cls, students, opts) {
    opts = opts || {};
    const wb = newWorkbook();
    const ws = wb.addWorksheet('名单', { views: [{ state: 'frozen', ySplit: 3 }] });
    const hasRank = students.some(s => typeof s.rank === 'number');
    const headers = ['序号', '姓名', '性别', '成绩'].concat(hasRank ? ['名次'] : []).concat(['学号', '状态', ROSTER_ID_HEADER]);
    const widths = [8, 14, 8, 10].concat(hasRank ? [8] : []).concat([16, 8, 18]);
    widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    titleCell(ws, 1, 1, headers.length - 1, `${cls.name} 名单`, 16);
    ws.getRow(1).height = 32;
    noteCell(ws, 2, 1, headers.length - 1, '可直接修改姓名、性别（男/女）、成绩、学号、状态（在班/离班），可增删行；保存后回到网页会自动读入。最后一列是系统编号，请勿修改。');
    ws.getRow(2).height = 30;
    headerRow(ws, 3, headers);
    const idCol = headers.length;
    students.forEach((s, i) => {
      const r = i + 4;
      const values = [typeof s.seq === 'number' ? s.seq : i + 1, s.name, core.GENDER_TEXT[s.gender || ''] || '', typeof s.score === 'number' ? s.score : null]
        .concat(hasRank ? [typeof s.rank === 'number' ? s.rank : null] : [])
        .concat([s.no || '', s.active === false ? '离班' : '在班', s.id]);
      const row = ws.getRow(r);
      row.values = values;
      row.height = 22;
      row.eachCell({ includeEmpty: true }, (cell, c) => {
        if (c > headers.length) return;
        cell.font = { name: FONT, size: 11, color: { argb: COLOR.ink } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = BORDER;
      });
      row.getCell(2).font = { name: FONT, size: 11, bold: true, color: { argb: s.gender ? COLOR[s.gender] : COLOR.ink } };
      row.getCell(3).dataValidation = { type: 'list', allowBlank: true, formulae: ['"男,女"'] };
      row.getCell(headers.indexOf('状态') + 1).dataValidation = { type: 'list', allowBlank: true, formulae: ['"在班,离班"'] };
      row.getCell(headers.indexOf('学号') + 1).numFmt = '@';
    });
    ws.getColumn(idCol).font = { name: FONT, size: 9, color: { argb: 'FF94A3B8' } };
    writeMeta(wb, { type: 'roster', exportId: opts.exportId, classId: cls.id, className: cls.name, createdAt: opts.createdAt || core.localDateTime() }, rosterBase(students));
    return wb;
  }

  /* =========================================================
     2. 座次表导出（学生版 + 教师版；学生版可对调姓名）
     ========================================================= */

  // 某个座位在工作表中的位置（导出与读回共用）
  function seatCell(view, layout, r, c) {
    const cols = core.displayColumns(layout, view);
    const rows = core.displayRows(layout, view);
    const ci = cols.findIndex(x => x.type === 'seat' && x.c === c);
    const ri = rows.indexOf(r);
    return { row: (view === 'teacher' ? 3 : 4) + ri, col: 2 + ci };
  }

  function drawSeatingSheet(ws, view, cls, seatsInfo, layout, byId, title) {
    const L = core.normalizeLayout(layout);
    const cols = core.displayColumns(L, view);
    const lastCol = 1 + cols.length;
    ws.getColumn(1).width = 12;
    cols.forEach((x, i) => { ws.getColumn(2 + i).width = x.type === 'aisle' ? 3 : 12; });
    titleCell(ws, 1, 1, lastCol, title + (view === 'teacher' ? '（教师视角）' : ''), 18);
    ws.getRow(1).height = 36;
    const stageRow = view === 'teacher' ? 3 + L.rows : 2;
    const colHeadRow = view === 'teacher' ? 2 : 3;
    ws.mergeCells(stageRow, 2, stageRow, lastCol);
    const stage = ws.getCell(stageRow, 2);
    stage.value = view === 'teacher' ? '讲    台（教师面向学生）' : '讲    台';
    stage.font = { name: FONT, size: 13, bold: true, color: { argb: 'FF3C5A51' } };
    stage.alignment = { horizontal: 'center', vertical: 'middle' };
    stage.fill = fill('FFEBF2EE');
    ws.getRow(stageRow).height = 26;
    const head = ws.getCell(colHeadRow, 1);
    head.value = view === 'teacher' ? '教师面向学生' : '学生面向讲台';
    head.font = { name: FONT, size: 9, color: { argb: COLOR.muted } };
    head.alignment = { horizontal: 'center', vertical: 'middle' };
    cols.forEach((x, i) => {
      if (x.type !== 'seat') return;
      const cell = ws.getCell(colHeadRow, 2 + i);
      cell.value = `第${x.c}列`;
      cell.font = { name: FONT, size: 9, color: { argb: COLOR.muted } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    const dis = new Set(L.disabled);
    for (let r = 1; r <= L.rows; r++) {
      const pos = seatCell(view, L, r, 1);
      const lbl = ws.getCell(pos.row, 1);
      lbl.value = `第${r}排 · ${core.rowZone(L, r)}区`;
      lbl.font = { name: FONT, size: 10, bold: true, color: { argb: COLOR.title } };
      lbl.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      lbl.fill = fill('FFF5F8F6');
      lbl.border = BORDER;
      ws.getRow(pos.row).height = 36;
      const rowAisles = core.aislesForRow(L, r);
      cols.forEach((x, i) => {
        if (x.type !== 'aisle' || rowAisles.indexOf(x.after) < 0) return;
        ws.getCell(pos.row, 2 + i).fill = fill('FFE8F3ED');
      });
      for (let c = 1; c <= L.cols; c++) {
        if (dis.has(core.seatKey(r, c))) continue;
        const p = seatCell(view, L, r, c);
        const cell = ws.getCell(p.row, p.col);
        const sid = seatsInfo.get(core.seatKey(r, c));
        const s = sid ? byId.get(sid) : null;
        cell.value = s ? s.name : '';
        cell.font = { name: FONT, size: 14, bold: true, color: { argb: s && s.gender ? COLOR[s.gender] : COLOR.ink } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = BORDER;
        if (!s) cell.fill = fill('FFF8FAFC');
      }
    }
    const noteRow = (view === 'teacher' ? 4 : 5) + L.rows;
    noteCell(ws, noteRow, 1, lastCol, view === 'teacher'
      ? '教师版由学生版自动旋转得到；如需调整请改“学生版”。'
      : '可在本表中对调或挪动姓名，保存后回到网页会自动读入（只能换座位，不能增删学生）。蓝色为男生，红色为女生。');
    ws.pageSetup = PAGE;
  }

  async function createSeatingWorkbook(cls, record, students, opts) {
    opts = opts || {};
    const wb = newWorkbook();
    const byId = new Map(students.map(s => [s.id, s]));
    const seatsInfo = core.seatsToMap(record.seats);
    const title = opts.title || `座次表-${cls.name}-${core.weekSpan(record.date)}`;
    drawSeatingSheet(wb.addWorksheet('学生版', { views: [{ showGridLines: false }] }), 'student', cls, seatsInfo, record.layout, byId, title);
    drawSeatingSheet(wb.addWorksheet('教师版', { views: [{ showGridLines: false }] }), 'teacher', cls, seatsInfo, record.layout, byId, title);
    writeMeta(wb, {
      type: 'seating', exportId: opts.exportId, classId: cls.id, className: cls.name, createdAt: opts.createdAt || core.localDateTime(),
      finalId: record.finalId || null, draft: !record.finalId, date: record.date, layout: core.normalizeLayout(record.layout), mode: record.mode || ''
    }, { seats: record.seats, names: record.seats.map(t => (byId.get(t[2]) || {}).name || '') });
    return wb;
  }

  /* =========================================================
     3. 课堂表现导出（明细可增删改；汇总用公式）
     ========================================================= */

  const SCORE_SHEET = '课堂表现明细';
  const SCORE_HEADERS = ['序号', '日期', '教学周', '姓名', '类别', '表现项目', '分值变动', '记录时间', '记录编号'];

  function scoresBase(records) {
    return records.map(r => ({ recId: r.recId, studentId: r.studentId, name: r.name, item: r.item, value: r.value, date: r.date, time: r.time }));
  }

  async function createScoresWorkbook(cls, records, students, opts) {
    opts = opts || {};
    const wb = newWorkbook();
    wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });
    const title = opts.title || `课堂表现-${cls.name}`;
    const byId = new Map(students.map(s => [s.id, s]));
    const recs = records.slice().sort((a, b) => (a.time || a.date).localeCompare(b.time || b.date) || a.name.localeCompare(b.name, 'zh'));

    const ws1 = wb.addWorksheet(SCORE_SHEET, { views: [{ state: 'frozen', ySplit: 3 }] });
    [7, 12, 13, 12, 8, 16, 10, 20, 14].forEach((w, i) => { ws1.getColumn(i + 1).width = w; });
    titleCell(ws1, 1, 1, SCORE_HEADERS.length - 1, title, 16);
    ws1.getRow(1).height = 32;
    noteCell(ws1, 2, 1, SCORE_HEADERS.length - 1, '可直接删除、修改或新增记录（新增时填姓名、表现项目、分值、日期即可）；保存后回到网页会自动读入。“汇总统计”由公式自动更新。');
    ws1.getRow(2).height = 30;
    headerRow(ws1, 3, SCORE_HEADERS);
    const vCol = SCORE_HEADERS.indexOf('分值变动') + 1;
    const tCol = SCORE_HEADERS.indexOf('类别') + 1;
    recs.forEach((r, i) => {
      const row = ws1.getRow(i + 4);
      row.values = [{ formula: 'ROW()-3', result: i + 1 }, r.date, core.weekSpan(r.date), r.name, r.value > 0 ? '加分' : '扣分', r.item, r.value, r.time || '', r.recId];
      row.height = 22;
      row.eachCell((cell, c) => {
        cell.font = { name: FONT, size: 11, color: { argb: COLOR.ink } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = BORDER;
        if (c === vCol || c === tCol) {
          cell.font = { name: FONT, size: 11, bold: true, color: { argb: r.value > 0 ? COLOR.bonus : COLOR.penalty } };
          cell.fill = fill(r.value > 0 ? 'FFE6F4EA' : 'FFFCE8E6');
        }
      });
      const s = byId.get(r.studentId);
      if (s && s.gender) row.getCell(4).font = { name: FONT, size: 11, bold: true, color: { argb: COLOR[s.gender] } };
      row.getCell(vCol).numFmt = '+0;-0;0';
    });
    ws1.getColumn(SCORE_HEADERS.length).hidden = true;

    const ws2 = wb.addWorksheet('汇总统计', { views: [{ state: 'frozen', ySplit: 2 }] });
    [8, 14, 12, 12, 12, 12, 14].forEach((w, i) => { ws2.getColumn(i + 1).width = w; });
    titleCell(ws2, 1, 1, 7, `${title} 汇总统计`, 16);
    ws2.getRow(1).height = 32;
    headerRow(ws2, 2, ['序号', '姓名', '加分次数', '加分合计', '扣分次数', '扣分合计', '净得分']);
    const NAME = `'${SCORE_SHEET}'!$D:$D`;
    const SC = `'${SCORE_SHEET}'!$G:$G`;
    const sum = core.summarizeScores(recs);
    const roster = students.filter(s => s.active !== false || sum.has(s.id));
    roster.forEach((s, i) => {
      const r = i + 3;
      const x = sum.get(s.id) || { bonusCount: 0, bonusSum: 0, penaltyCount: 0, penaltySum: 0, net: 0 };
      const row = ws2.getRow(r);
      row.values = [
        i + 1, s.name,
        { formula: `COUNTIFS(${NAME},B${r},${SC},">0")`, result: x.bonusCount },
        { formula: `SUMIFS(${SC},${NAME},B${r},${SC},">0")`, result: x.bonusSum },
        { formula: `COUNTIFS(${NAME},B${r},${SC},"<0")`, result: x.penaltyCount },
        { formula: `SUMIFS(${SC},${NAME},B${r},${SC},"<0")`, result: x.penaltySum },
        { formula: `D${r}+F${r}`, result: x.net }
      ];
      row.height = 22;
      row.eachCell((cell, c) => {
        cell.font = { name: FONT, size: 11, color: { argb: COLOR.ink } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = BORDER;
        if (c === 4) cell.font = { name: FONT, size: 11, bold: true, color: { argb: COLOR.bonus } };
        if (c === 6) cell.font = { name: FONT, size: 11, bold: true, color: { argb: COLOR.penalty } };
        if (c === 7) cell.font = { name: FONT, size: 12, bold: true, color: { argb: COLOR.title } };
      });
      if (s.gender) row.getCell(2).font = { name: FONT, size: 11, bold: true, color: { argb: COLOR[s.gender] } };
      [4, 6, 7].forEach(c => { row.getCell(c).numFmt = '+0;-0;0'; });
    });
    writeMeta(wb, {
      type: 'scores', exportId: opts.exportId, classId: cls.id, className: cls.name, createdAt: opts.createdAt || core.localDateTime(),
      defaultDate: opts.defaultDate || core.localISODate()
    }, scoresBase(recs));
    return wb;
  }

  /* =========================================================
     读回：老师在导出表里的修改 → 记录（编号由内容决定）
     ========================================================= */

  const detId = (prefix, exportId, payload) => prefix + core.hash53(`${exportId}|${JSON.stringify(payload)}`);

  function nameIndex(students) {
    const map = new Map();
    students.forEach(s => {
      if (!map.has(s.name)) map.set(s.name, []);
      map.get(s.name).push(s);
    });
    return map;
  }

  function uniqueByName(index, name, activeOnly) {
    const list = (index.get(name) || []).filter(s => !activeOnly || s.active !== false);
    return list.length === 1 ? list[0] : null;
  }

  function diffRoster(wb, meta, cls, students) {
    const ws = wb.getWorksheet('名单');
    const res = { ops: [], parts: [], warnings: [] };
    if (!ws || !Array.isArray(meta.base)) { res.warnings.push('名单表结构已改变，无法读入'); return res; }
    const hd = findHeader(ws, ['姓名']);
    if (!hd) { res.warnings.push('名单表找不到“姓名”表头，无法读入'); return res; }
    const col = h => hd.col(h);
    const rows = [];
    for (let r = hd.row + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const get = h => (col(h) ? row.getCell(col(h)).value : null);
      const name = core.normalizeName(cellText(get('姓名')));
      if (!name) continue;
      const g = core.normalizeGender(cellText(get('性别')));
      rows.push({
        id: cellText(get(ROSTER_ID_HEADER)),
        name,
        gender: g,
        score: col('成绩') ? parseNumber(get('成绩')) : undefined,
        rank: col('名次') ? parseNumber(get('名次')) : undefined,
        no: col('学号') ? cellText(get('学号')) : undefined,
        seq: col('序号') ? parseNumber(get('序号')) : undefined,
        active: col('状态') ? !/离/.test(cellText(get('状态'))) : undefined
      });
    }
    const base = new Map(meta.base.map(b => [b.id, b]));
    const baseByName = nameIndex(meta.base);
    // 只选了部分列排序会让编号与姓名错位：发现错位就改为按姓名对应
    let byName = false;
    for (const row of rows) {
      const b = base.get(row.id);
      if (b && b.name !== row.name && (baseByName.get(row.name) || []).length) { byName = true; break; }
    }
    if (byName) res.warnings.push('名单里的编号与姓名对不上（可能只对部分列排了序），已按姓名对应');
    const matched = new Set();
    const edits = { name: 0, gender: 0, score: 0, other: 0 };
    let added = 0;
    const occurrence = new Map();
    for (const row of rows) {
      let b = null;
      if (!byName) b = base.get(row.id) || null;
      else {
        const cand = (baseByName.get(row.name) || []).filter(x => !matched.has(x.id));
        b = cand.length ? cand[0] : null;
      }
      if (b && matched.has(b.id)) b = null;
      if (b) {
        matched.add(b.id);
        const fields = {};
        if (row.name !== b.name) { fields.name = row.name; edits.name++; }
        if (row.gender !== null && row.gender !== (b.gender || '')) { fields.gender = row.gender; edits.gender++; }
        if (row.score !== undefined && row.score !== (typeof b.score === 'number' ? b.score : null)) { fields.score = row.score; edits.score++; }
        if (row.rank !== undefined && row.rank !== (typeof b.rank === 'number' ? b.rank : null)) { fields.rank = row.rank; edits.other++; }
        if (row.no !== undefined && row.no !== (b.no || '')) { fields.no = row.no; edits.other++; }
        if (row.seq !== undefined && row.seq !== null && row.seq !== b.seq) { fields.seq = row.seq; edits.other++; }
        if (row.active !== undefined && row.active !== (b.active !== false)) { fields.active = row.active; edits.other++; }
        if (Object.keys(fields).length) {
          const payload = { k: 'student.upsert', classId: cls.id, studentId: b.id, fields };
          res.ops.push(Object.assign({ id: detId('x_', meta.exportId, payload) }, payload));
        }
      } else {
        const n = (occurrence.get(row.name) || 0) + 1;
        occurrence.set(row.name, n);
        const studentId = 's_x' + core.hash53(`${meta.exportId}|new|${row.name}|${n}`);
        const fields = { name: row.name, gender: row.gender || '', score: row.score === undefined ? null : row.score, active: row.active !== false };
        if (row.no) fields.no = row.no;
        if (typeof row.seq === 'number') fields.seq = row.seq;
        const payload = { k: 'student.upsert', classId: cls.id, studentId, fields };
        res.ops.push(Object.assign({ id: detId('x_', meta.exportId, payload) }, payload));
        added++;
      }
    }
    let left = 0;
    meta.base.forEach(b => {
      if (matched.has(b.id) || b.active === false) return;
      const payload = { k: 'student.upsert', classId: cls.id, studentId: b.id, fields: { active: false } };
      res.ops.push(Object.assign({ id: detId('x_', meta.exportId, payload) }, payload));
      left++;
    });
    const edited = edits.name + edits.gender + edits.score + edits.other;
    if (edited) res.parts.push(`修改 ${edited} 处（${[edits.name ? `姓名 ${edits.name}` : '', edits.gender ? `性别 ${edits.gender}` : '', edits.score ? `成绩 ${edits.score}` : '', edits.other ? `其他 ${edits.other}` : ''].filter(Boolean).join('、')}）`);
    if (added) res.parts.push(`新增 ${added} 人`);
    if (left) res.parts.push(`${left} 人标为离班（表里删掉了这些行）`);
    return res;
  }

  function diffSeating(wb, meta, cls, students) {
    const ws = wb.getWorksheet('学生版');
    const res = { ops: [], parts: [], warnings: [] };
    if (!ws || !meta.base || !Array.isArray(meta.base.seats) || !meta.layout) { res.warnings.push('座次表结构已改变，无法读入'); return res; }
    const L = core.normalizeLayout(meta.layout);
    const index = nameIndex(students);
    const seats = [];
    const unknown = [];
    for (const seat of core.allSeats(L)) {
      const p = seatCell('student', L, seat.r, seat.c);
      const name = core.normalizeName(cellText(ws.getCell(p.row, p.col).value));
      if (!name) continue;
      const s = uniqueByName(index, name, true);
      if (!s) { unknown.push(name); continue; }
      seats.push([seat.r, seat.c, s.id]);
    }
    const key = arr => arr.map(t => t.join(',')).sort().join(';');
    if (key(seats) === key(meta.base.seats) && !unknown.length) return res;
    if (unknown.length) { res.warnings.push(`座次表中有无法对应的姓名（${unknown.slice(0, 3).join('、')}${unknown.length > 3 ? '等' : ''}），本次未读入`); return res; }
    const ids = arr => arr.map(t => t[2]).sort().join(',');
    if (ids(seats) !== ids(meta.base.seats)) { res.warnings.push('座次表里只能对调或挪动座位，不能增删学生；本次未读入'); return res; }
    if (meta.draft) { res.warnings.push('候选稿导出的座次表修改不会读入，请在网页里微调后再定版'); return res; }
    const sorted = core.mapToSeats(core.seatsToMap(seats));
    const payload = { k: 'final.set', classId: cls.id, date: meta.date, seats: sorted, layout: L, mode: 'adjust', meta: { from: 'excel', exportId: meta.exportId || '' }, base: meta.finalId || null };
    const finalId = 'f_x' + core.hash53(`${meta.exportId}|${JSON.stringify(sorted)}`);
    res.ops.push(Object.assign({ id: detId('x_', meta.exportId, payload), finalId }, payload));
    let moved = 0;
    const before = core.seatsToMap(meta.base.seats);
    sorted.forEach(t => { if (before.get(core.seatKey(t[0], t[1])) !== t[2]) moved++; });
    res.parts.push(`${core.weekSpan(meta.date)} 的座次调整了 ${moved} 个座位`);
    return res;
  }

  function diffScores(wb, meta, cls, students, items) {
    const ws = wb.getWorksheet(SCORE_SHEET);
    const res = { ops: [], parts: [], warnings: [] };
    if (!ws || !Array.isArray(meta.base)) { res.warnings.push('课堂表现表结构已改变，无法读入'); return res; }
    const hd = findHeader(ws, ['姓名', '分值变动']);
    if (!hd) { res.warnings.push('课堂表现表找不到“姓名”“分值变动”表头，无法读入'); return res; }
    const col = h => hd.col(h);
    const index = nameIndex(students);
    const base = new Map(meta.base.map(b => [b.recId, b]));
    const seen = new Set();
    const itemValue = new Map((items || []).map(it => [it.name, it.value]));
    let edited = 0;
    let added = 0;
    let removed = 0;
    const occurrence = new Map();
    for (let r = hd.row + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const get = h => (col(h) ? row.getCell(col(h)).value : null);
      const name = core.normalizeName(cellText(get('姓名')));
      if (!name) continue;
      const item = cellText(get('表现项目'));
      let value = parseNumber(get('分值变动'));
      if (value === null && itemValue.has(item)) value = itemValue.get(item);
      const date = normalizeDate(get('日期')) || normalizeDate(get('记录时间'));
      const recId = cellText(get('记录编号'));
      const b = recId ? base.get(recId) : null;
      if (!value) { res.warnings.push(`第${r}行分值无法识别，已跳过`); if (b) seen.add(b.recId); continue; }
      if (b && !seen.has(b.recId)) {
        seen.add(b.recId);
        if (name !== b.name) {
          const s = uniqueByName(index, name, false);
          if (!s) { res.warnings.push(`第${r}行“${name}”不在本班名单中，已跳过`); continue; }
          const del = { k: 'score.delete', recId: b.recId };
          res.ops.push(Object.assign({ id: detId('x_', meta.exportId, del) }, del));
          const rec = { recId: 'r_x' + core.hash53(`${meta.exportId}|re|${b.recId}|${s.id}`), classId: cls.id, studentId: s.id, name: s.name, item: item || b.item, value, date: date || b.date, time: b.time };
          const add = { k: 'score.add', rec };
          res.ops.push(Object.assign({ id: detId('x_', meta.exportId, add) }, add));
          edited++;
          continue;
        }
        const patch = {};
        if (item && item !== b.item) patch.item = item;
        if (value !== b.value) patch.value = value;
        if (date && date !== b.date) patch.date = date;
        if (Object.keys(patch).length) {
          const payload = { k: 'score.edit', recId: b.recId, patch };
          res.ops.push(Object.assign({ id: detId('x_', meta.exportId, payload) }, payload));
          edited++;
        }
        continue;
      }
      const s = uniqueByName(index, name, false);
      if (!s) { res.warnings.push(`第${r}行“${name}”不在本班名单中（或有同名），已跳过`); continue; }
      const d = date || meta.defaultDate || core.localISODate();
      const sig = `${s.id}|${item}|${value}|${d}`;
      const n = (occurrence.get(sig) || 0) + 1;
      occurrence.set(sig, n);
      const rec = { recId: 'r_x' + core.hash53(`${meta.exportId}|new|${sig}|${n}`), classId: cls.id, studentId: s.id, name: s.name, item: item || '其他', value, date: d, time: `${d} 00:00:00` };
      const payload = { k: 'score.add', rec };
      res.ops.push(Object.assign({ id: detId('x_', meta.exportId, payload) }, payload));
      added++;
    }
    meta.base.forEach(b => {
      if (seen.has(b.recId)) return;
      const payload = { k: 'score.delete', recId: b.recId };
      res.ops.push(Object.assign({ id: detId('x_', meta.exportId, payload) }, payload));
      removed++;
    });
    if (edited) res.parts.push(`修改 ${edited} 条`);
    if (added) res.parts.push(`新增 ${added} 条`);
    if (removed) res.parts.push(`删除 ${removed} 条`);
    return res;
  }

  // 读取一个导出文件，返回需要写入的记录；getClass(classId) 返回 { cls, students, items }
  async function readExportEdits(buffer, fileName, getClass) {
    const Excel = getExcelJS();
    const wb = new Excel.Workbook();
    await wb.xlsx.load(buffer);
    const meta = readMeta(wb);
    if (!meta) return null; // 不是本系统导出的表
    if (meta.broken) return { file: fileName, ops: [], parts: [], warnings: [`“${fileName}”里的隐藏信息已损坏，无法读入修改`] };
    const ctx = getClass(meta.classId);
    if (!ctx) return { file: fileName, ops: [], parts: [], warnings: [`“${fileName}”对应的班级已不存在`] };
    let res;
    if (meta.type === 'roster') res = diffRoster(wb, meta, ctx.cls, ctx.students);
    else if (meta.type === 'seating') res = diffSeating(wb, meta, ctx.cls, ctx.students);
    else if (meta.type === 'scores') res = diffScores(wb, meta, ctx.cls, ctx.students, ctx.items);
    else return null;
    return Object.assign({ file: fileName, type: meta.type, exportId: meta.exportId, className: meta.className }, res);
  }

  function safeFileName(s) {
    return String(s || '').replace(/[\\/:*?"<>|]/g, '').replace(/[\s.]+$/g, '').trim() || '未命名';
  }

  function exportStamp(d) {
    d = d || new Date();
    return `${pad2(d.getMonth() + 1)}${pad2(d.getDate())}-${pad2(d.getHours())}${pad2(d.getMinutes())}${pad2(d.getSeconds())}`;
  }

  const CMExcel = {
    META_SHEET, SCORE_SHEET, SCORE_HEADERS, ROSTER_ID_HEADER, COLOR,
    cellText, parseNumber, normalizeDate, normalizeClassName, classNameFromText,
    parseRosterRows, mergeSameName, parseRosterWorkbook, parseRosterText, decodeText, splitTable,
    readMeta, seatCell, createRosterWorkbook, createSeatingWorkbook, createScoresWorkbook,
    readExportEdits, safeFileName, exportStamp
  };

  global.CMExcel = CMExcel;
  if (typeof module !== 'undefined' && module.exports) module.exports = CMExcel;
})(typeof window !== 'undefined' ? window : globalThis);
