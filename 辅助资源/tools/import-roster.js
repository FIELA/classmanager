#!/usr/bin/env node
/**
 * 从名单/成绩表导入班级（与网页“导入名单”使用同一套识别规则）。
 *
 *   node 辅助资源/tools/import-roster.js [--root 项目目录] [--dry-run] 名单1.xlsx 名单2.xlsx …
 *
 * - 同名班级已存在时合并：按姓名对应，更新成绩/性别，新增名单里多出的学生（不会把人标为离班）。
 * - 新班级使用默认布局（每排 8 列、4 个大组），之后可在网页“班级设置”里修改。
 * - 写入一个新的会话记录文件，并在 导出/<班级>/ 下新建整理好的名单表。只新建文件，不改动已有文件。
 * - 输出只包含班级与人数，不打印学生姓名。
 */
'use strict';

process.env.TZ = process.env.TZ || 'Asia/Shanghai';

const fs = require('fs');
const path = require('path');
const ASSETS = path.join(__dirname, '..', 'assets');
global.ExcelJS = require(path.join(ASSETS, 'exceljs.min.js'));
const core = require(path.join(ASSETS, 'core.js'));
const store = require(path.join(ASSETS, 'store.js'));
const excel = require(path.join(ASSETS, 'excel.js'));
const { NodeDirHandle } = require('./node-fs-handle.js');

function parseArgs(argv) {
  const args = { root: path.join(__dirname, '..', '..'), dryRun: false, files: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--root') args.root = argv[++i];
    else if (a === '--dry-run') args.dryRun = true;
    else args.files.push(a);
  }
  return args;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.files.length) {
    console.error('用法：node 辅助资源/tools/import-roster.js [--root 项目目录] [--dry-run] 名单.xlsx …');
    process.exit(2);
  }
  const root = new NodeDirHandle(path.resolve(args.root));
  const read = await store.readProject(root);
  const st = store.reduce(read.ops);
  const sid = core.randomId(8);
  const clock = store.createClock(sid);
  read.ops.forEach(o => clock.observe(o.t));
  const ops = [];
  let seq = 0;
  const op = (k, payload) => ops.push(Object.assign({ k, id: `${sid}.${++seq}`, t: clock.now(), s: sid }, payload));

  const existing = store.listClasses(st);
  let order = existing.reduce((m, c) => Math.max(m, c.order === 999 ? 0 : c.order), 0);
  const plans = [];

  for (const file of args.files) {
    const buf = fs.readFileSync(file);
    const parsed = await excel.parseRosterWorkbook(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength), path.basename(file));
    parsed.warnings.forEach(w => console.warn(`  提示（${path.basename(file)}）：${w}`));
    for (const cand of parsed.candidates) {
      const name = excel.normalizeClassName(cand.className);
      let cls = existing.find(c => c.name === name) || null;
      const plan = { className: name, source: path.basename(file), added: 0, updated: 0, total: cand.students.length, withScore: 0, withGender: 0, warnings: cand.warnings };
      let classId;
      if (cls) {
        classId = cls.id;
      } else {
        classId = 'c_' + core.randomId(8);
        op('class.create', { classId, name, order: ++order });
      }
      const current = cls ? store.studentsOf(cls, true) : [];
      const byName = new Map(current.map(s => [s.name, s]));
      cand.students.forEach((s, i) => {
        if (typeof s.score === 'number') plan.withScore++;
        if (s.gender) plan.withGender++;
        const fields = { name: s.name };
        if (s.gender) fields.gender = s.gender;
        if (s.score !== undefined) fields.score = s.score;
        if (s.rank !== undefined) fields.rank = s.rank;
        if (s.no) fields.no = s.no;
        fields.seq = typeof s.seq === 'number' ? s.seq : i + 1;
        const found = byName.get(s.name);
        if (found) {
          op('student.upsert', { classId, studentId: found.id, fields });
          plan.updated++;
        } else {
          fields.active = s.active !== false;
          if (!fields.gender) fields.gender = '';
          op('student.upsert', { classId, studentId: 's_' + core.randomId(8), fields });
          plan.added++;
        }
      });
      if (!cls) {
        op('layout.set', { classId, layout: core.defaultLayoutFor(cand.students.length), layoutId: 'l_' + core.randomId(8), base: null });
      }
      plans.push(Object.assign(plan, { classId }));
    }
  }

  for (const p of plans) {
    console.log(`${p.className}（来自 ${p.source}）：共 ${p.total} 人，新增 ${p.added}、更新 ${p.updated}；有成绩 ${p.withScore} 人，有性别 ${p.withGender} 人`);
    p.warnings.forEach(w => console.log(`  提示：${w}`));
  }
  if (args.dryRun) {
    console.log(`（试运行，未写入；将写入 ${ops.length} 条记录）`);
    return;
  }
  if (!ops.length) {
    console.log('没有需要写入的内容');
    return;
  }

  // 项目信息文件只在第一次创建（以后不再改动）
  const dataDir = await root.getDirectoryHandle(store.DATA_DIR, { create: true });
  if (!(await store.fileExists(dataDir, store.PROJECT_FILE))) {
    const info = { app: store.APP, v: store.FORMAT, projectId: 'p_' + core.randomId(10), createdAt: core.localDateTime() };
    const r = await store.writeFileChecked(dataDir, store.PROJECT_FILE, JSON.stringify(info, null, 2) + '\n');
    if (!r.ok) throw new Error('写入项目信息失败：' + r.error);
  }

  const writer = new store.SessionWriter({ root, sid, place: 'tool', header: store.makeHeader(sid, clock, { place: 'tool', started: core.localDateTime() }) });
  ops.forEach(o => writer.add(o));
  const w = await writer.flush();
  if (!w.ok) throw new Error('写入记录失败：' + w.error);
  console.log(`已写入记录文件：${store.DATA_DIR}/${writer.fileName}（${ops.length} 条）`);

  // 导出整理好的名单表
  const st2 = store.reduce((await store.readProject(root)).ops);
  for (const p of plans) {
    const cls = st2.classes.get(p.classId);
    const students = store.studentsOf(cls, true);
    const exportId = 'e_' + core.randomId(10);
    const wb = await excel.createRosterWorkbook(cls, students, { exportId });
    const buf = await wb.xlsx.writeBuffer();
    const dir = excel.safeFileName(cls.name);
    const res = await store.writeNewFile(root, [store.EXPORT_DIR, dir], `名单-${dir}-${excel.exportStamp()}`, '.xlsx', buf);
    if (!res.ok) console.warn(`  ${cls.name} 名单表写入失败：${res.error}`);
    else console.log(`已导出：${res.path}`);
  }
}

main().catch(e => {
  console.error('导入失败：' + (e && e.message ? e.message : e));
  process.exit(1);
});
