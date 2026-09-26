// 数据层测试：时钟、合并、会话文件、快照，以及按 SyncTime 双向同步规则模拟 Mac ↔ U 盘
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { core, store, FakeDir, SyncTimeSim, fakeStudents } = require('./helpers.js');

// 模拟“打开一次网页”：读取项目、创建会话写入器、提交操作
async function openSession(root, sid, opts = {}) {
  const read = await store.readProject(root);
  const clock = store.createClock(sid, opts.now);
  read.ops.forEach(o => clock.observe(o.t));
  const writer = new store.SessionWriter({ root, sid, place: opts.place || 'mac', header: store.makeHeader(sid, clock, { place: opts.place || 'mac' }), startedAt: new Date(2026, 8, 26, 8, 0, 0) });
  let seq = 0;
  const commit = async (k, payload) => {
    const op = Object.assign({ k, id: `${sid}.${++seq}`, t: clock.now(), s: sid }, payload);
    writer.add(op);
    return writer.flush();
  };
  return { read, clock, writer, commit };
}

async function stateOf(root) {
  return store.reduce((await store.readProject(root)).ops);
}

function summary(st) {
  const classes = store.listClasses(st).map(c => ({
    name: c.name,
    students: store.studentsOf(c, true).map(s => `${s.id}:${s.name}:${s.gender}:${s.score}:${s.active}`),
    finals: store.finalsOf(c).map(f => `${f.week}:${f.finalId}:${JSON.stringify(f.seats)}`),
    layout: c.layout ? core.layoutKey(c.layout) : null
  }));
  const scores = Array.from(st.scores.values()).filter(r => !r.deleted).map(r => `${r.recId}:${r.value}`).sort();
  return JSON.stringify({ classes, scores });
}

test('时钟：单调递增；读到别处更晚的记录后不会倒退（电脑时间不准也不乱序）', () => {
  let now = 1000;
  const a = store.createClock('aaaaaaaa', () => now);
  const t1 = a.now();
  const t2 = a.now();
  assert.ok(t2 > t1);
  const b = store.createClock('bbbbbbbb', () => 5000);
  const tb = b.now();
  a.observe(tb);
  now = 10; // 本机时间倒退（例如教室电脑时间不准）
  const t3 = a.now();
  assert.ok(t3 > tb, '观察到更晚的记录后，新记录必须排在它之后');
  assert.deepEqual(store.tsParse(t3).sid, 'aaaaaaaa');
});

test('合并：同编号只算一次；学生字段后写为准；删除标记先于新增也有效', () => {
  const ops = [
    { k: 'class.create', id: 'o1', t: store.tsMake(1, 0, 'a'), classId: 'c1', name: '测试1班' },
    { k: 'student.upsert', id: 'o2', t: store.tsMake(2, 0, 'a'), classId: 'c1', studentId: 's1', fields: { name: '赵测一', gender: '男', score: 80 } },
    { k: 'student.upsert', id: 'o3', t: store.tsMake(3, 0, 'b'), classId: 'c1', studentId: 's1', fields: { gender: 'F' } },
    { k: 'student.upsert', id: 'o3', t: store.tsMake(9, 0, 'b'), classId: 'c1', studentId: 's1', fields: { gender: 'M' } }, // 重复编号，忽略
    { k: 'score.delete', id: 'o4', t: store.tsMake(4, 0, 'a'), recId: 'r1' },
    { k: 'score.add', id: 'o5', t: store.tsMake(5, 0, 'a'), rec: { recId: 'r1', classId: 'c1', studentId: 's1', name: '赵测一', item: '迟到', value: -3, date: '2026-09-28', time: '2026-09-28 08:00:00' } },
    { k: 'score.add', id: 'o6', t: store.tsMake(6, 0, 'a'), rec: { recId: 'r2', classId: 'c1', studentId: 's1', name: '赵测一', item: '优秀作业', value: 1, date: '2026-09-28' } },
    { k: 'score.edit', id: 'o7', t: store.tsMake(7, 0, 'a'), recId: 'r2', patch: { value: 2 } },
    { k: 'future.kind', id: 'o8', t: store.tsMake(8, 0, 'a') }
  ];
  const st = store.reduce(ops.slice().reverse());
  const s1 = st.classes.get('c1').students.get('s1');
  assert.equal(s1.gender, 'F');
  assert.equal(s1.score, 80);
  assert.equal(st.scores.get('r1').deleted, true);
  assert.equal(st.scores.get('r2').value, 2);
  assert.equal(st.skipped, 1, '未知类型应被跳过而不是报错');
});

test('定版冲突：两边各自为同一周定版会被发现；顺序修改不算冲突；确认后消失', () => {
  const base = [{ k: 'class.create', id: 'o1', t: store.tsMake(1, 0, 'a'), classId: 'c1', name: '测试1班' }];
  const f = (id, t, finalId, baseId) => ({ k: 'final.set', id, t: store.tsMake(t, 0, 'a'), classId: 'c1', finalId, date: '2026-09-28', seats: [], layout: { rows: 1, cols: 1 }, base: baseId });
  let st = store.reduce(base.concat([f('o2', 2, 'fa', null), f('o3', 3, 'fb', 'fa')]));
  assert.equal(st.classes.get('c1').conflicts.length, 0);
  st = store.reduce(base.concat([f('o2', 2, 'fa', null), f('o3', 3, 'fb', null)]));
  const c = st.classes.get('c1');
  assert.equal(c.conflicts.length, 1);
  assert.equal(c.conflicts[0].loser.finalId, 'fa');
  assert.equal(c.finals.get('2026-09-28').finalId, 'fb', '较晚的一份生效');
  st = store.reduce(base.concat([f('o2', 2, 'fa', null), f('o3', 3, 'fb', null), { k: 'conflict.ack', id: 'o4', t: store.tsMake(4, 0, 'a'), classId: 'c1', kind2: 'final', key: '2026-09-28', winnerId: 'fb' }]));
  assert.equal(st.classes.get('c1').conflicts.length, 0);
});

test('生效座次与生成基准', () => {
  const cls = { finals: new Map() };
  const put = (date, id) => cls.finals.set(core.mondayOf(date), { finalId: id, week: core.mondayOf(date), date, t: id });
  put('2026-09-21', 'a');
  put('2026-09-30', 'b'); // 周三生效
  assert.equal(store.effectiveFinal(cls, '2026-09-29').finalId, 'a');
  assert.equal(store.effectiveFinal(cls, '2026-09-30').finalId, 'b');
  assert.equal(store.effectiveFinal(cls, '2026-09-20'), null);
  assert.equal(store.sourceFinalFor(cls, '2026-10-01').finalId, 'a', '同一周的定版不作为自己的基准');
  assert.equal(store.sourceFinalFor(cls, '2026-10-05').finalId, 'b');
});

test('记录文件：最后一行写到一半也能读出其余内容', () => {
  const good = JSON.stringify({ k: 'class.create', id: 'x1', t: store.tsMake(1, 0, 'a'), classId: 'c', name: 'n' });
  const res = store.parseJsonl(`${good}\n{"k":"student.ups`);
  assert.equal(res.ops.length, 1);
  assert.equal(res.bad, 1);
});

test('会话写入：第一次修改才建文件；只写自己的文件；写入失败如实报告并可补写', async () => {
  const root = new FakeDir('座次2');
  const s = await openSession(root, 'aaaaaaaa');
  assert.equal((await store.readProject(root)).hasDataDir, false, '只打开不修改，不应创建任何文件');
  await s.commit('class.create', { classId: 'c1', name: '测试1班' });
  const data = root.dirs.get('数据');
  assert.equal(data.files.size, 1);
  const name = Array.from(data.files.keys())[0];
  assert.match(name, /^2026-09-26_080000_mac_aaaaaaaa\.jsonl$/);

  data.setFailWrites(true);
  const r = await s.commit('class.update', { classId: 'c1', patch: { name: '测试一班' } });
  assert.equal(r.ok, false);
  assert.match(r.error, /U 盘/);
  assert.equal(s.writer.pendingCount(), 1);
  data.setFailWrites(false);
  const r2 = await s.writer.flush();
  assert.equal(r2.ok, true);
  assert.equal(s.writer.pendingCount(), 0);
  assert.deepEqual(new Set(data.writeLog), new Set([name]), '只允许写本会话自己的文件');
  const st = await stateOf(root);
  assert.equal(store.listClasses(st)[0].name, '测试一班');
});

test('读取：忽略 ._、~$、.crswap 等系统文件；冲突副本里的重复记录只算一次', async () => {
  const root = new FakeDir('座次2');
  const s = await openSession(root, 'aaaaaaaa');
  await s.commit('class.create', { classId: 'c1', name: '测试1班' });
  const data = root.dirs.get('数据');
  const [name] = Array.from(data.files.keys());
  const text = await root.readText(`数据/${name}`);
  await root.writeText(`数据/${name.replace('.jsonl', '-MacBook-Air.jsonl')}`, text);
  await root.writeText(`数据/._${name}`, 'garbage');
  await root.writeText(`数据/${name}.crswap`, 'garbage');
  await root.writeText('数据/.DS_Store', 'garbage');
  const read = await store.readProject(root);
  assert.equal(read.files.length, 2);
  const st = store.reduce(read.ops);
  assert.equal(store.listClasses(st).length, 1);
  assert.equal(st.opCount, 2, '会话头 + 建班，各一次');
});

test('新建文件从不覆盖同名文件', async () => {
  const root = new FakeDir('座次2');
  const a = await store.writeNewFile(root, ['导出', '测试1班'], '名单-测试1班', '.xlsx', 'A');
  const b = await store.writeNewFile(root, ['导出', '测试1班'], '名单-测试1班', '.xlsx', 'B');
  assert.equal(a.path, '导出/测试1班/名单-测试1班.xlsx');
  assert.equal(b.path, '导出/测试1班/名单-测试1班-2.xlsx');
  assert.equal(await root.readText(a.path), 'A');
});

test('快照：写成新文件；之后读取结果不变，被覆盖的记录文件不再逐个读取', async () => {
  const root = new FakeDir('座次2');
  for (let i = 0; i < 32; i++) {
    const s = await openSession(root, `s${String(i).padStart(7, '0')}`, { now: () => 1000 + i * 10 });
    if (i === 0) await s.commit('class.create', { classId: 'c1', name: '测试1班' });
    else await s.commit('student.upsert', { classId: 'c1', studentId: `st${i}`, fields: { name: `赵测${i}` } });
  }
  const before = await store.readProject(root);
  const sum1 = summary(store.reduce(before.ops));
  const w = await store.maybeWriteSnapshot(root, before, null, 'zzzzzzzz', false);
  assert.equal(w.ok, true);
  assert.match(w.path, /^数据\/快照\/snapshot_.*_zzzzzzzz\.json$/);
  const after = await store.readProject(root);
  assert.equal(after.snapshot.covers, 32);
  assert.equal(after.files.filter(f => !f.covered).length, 0);
  assert.equal(summary(store.reduce(after.ops)), sum1);
});

test('SyncTime 双向同步模拟：Mac 与 U 盘各自使用、交替同步，始终没有冲突，两边数据一致', async () => {
  const mac = new FakeDir('Mac');
  const usb = new FakeDir('U盘');
  const sync = new SyncTimeSim(mac, usb);
  const students = fakeStudents(12);

  // 1. Mac 上建班、导入名单
  const m1 = await openSession(mac, 'mac00001', { now: () => 1000 });
  await m1.commit('class.create', { classId: 'c1', name: '测试1班' });
  for (const s of students) await m1.commit('student.upsert', { classId: 'c1', studentId: s.id, fields: { name: s.name, gender: s.gender, score: s.score } });
  const layout = { rows: 2, cols: 8, aisles: [2, 4, 6] };
  await m1.commit('layout.set', { classId: 'c1', layout, layoutId: 'l1', base: null });
  assert.deepEqual(sync.sync().conflicts, []);

  // 2. 教室（U 盘）：定版、记分；同时 Mac 上这次会话还开着、继续改性别
  const u1 = await openSession(usb, 'win00001', { place: 'win', now: () => 2000 });
  const seats = core.randomSeating(students, layout, [], {}, core.mulberry32(1)).seats;
  await u1.commit('final.set', { classId: 'c1', finalId: 'f1', date: '2026-09-28', seats, layout, mode: 'random', base: null });
  await u1.commit('score.add', { rec: { recId: 'r1', classId: 'c1', studentId: 's001', name: students[0].name, item: '迟到', value: -3, date: '2026-09-28', time: '2026-09-28 08:05:00' } });
  await m1.commit('student.upsert', { classId: 'c1', studentId: 's002', fields: { gender: 'M' } });
  assert.deepEqual(sync.sync().conflicts, []);

  // 3. 回家后 Mac 新开一次：撤销那条记分、调整座位并重新定版；教室里又记一次分；再同步两次
  const m2 = await openSession(mac, 'mac00002', { now: () => 3000 });
  await m2.commit('score.delete', { recId: 'r1' });
  await m2.commit('final.set', { classId: 'c1', finalId: 'f2', date: '2026-09-28', seats: seats.slice().reverse().map((t, i) => [seats[i][0], seats[i][1], t[2]]), layout, mode: 'adjust', base: 'f1' });
  const u2 = await openSession(usb, 'win00002', { place: 'win', now: () => 2500 });
  await u2.commit('score.add', { rec: { recId: 'r2', classId: 'c1', studentId: 's003', name: students[2].name, item: '优秀作业', value: 1, date: '2026-09-29', time: '2026-09-29 10:00:00' } });
  assert.deepEqual(sync.sync().conflicts, []);
  assert.deepEqual(sync.sync().conflicts, []);

  // 两边读出的数据完全一致
  const a = await stateOf(mac);
  const b = await stateOf(usb);
  assert.equal(summary(a), summary(b));
  const c = store.listClasses(a)[0];
  assert.equal(store.finalsOf(c)[0].finalId, 'f2');
  assert.equal(c.conflicts.length, 0);
  assert.equal(c.students.get('s002').gender, 'M');
  assert.deepEqual(Array.from(a.scores.values()).filter(r => !r.deleted).map(r => r.recId), ['r2']);
  // 没有任何文件被删除（SyncTime 会把删除同步到另一边）
  assert.equal(mac.flatten().size, usb.flatten().size);
});

test('模拟器自检：旧做法（两边都改同一个表格文件）会被 SyncTime 判为冲突', async () => {
  const mac = new FakeDir('Mac');
  const usb = new FakeDir('U盘');
  const sync = new SyncTimeSim(mac, usb);
  await mac.writeText('课堂表现-9.28～10.2.xlsx', 'v1');
  sync.sync();
  await mac.writeText('课堂表现-9.28～10.2.xlsx', 'v2-mac');
  await usb.writeText('课堂表现-9.28～10.2.xlsx', 'v2-classroom');
  assert.deepEqual(sync.sync().conflicts, ['课堂表现-9.28～10.2.xlsx']);
});
