// Excel 测试：名单导入（与真实成绩表同样的结构，但全部是虚构姓名）、导出、读回老师的修改
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { core, store, excel, fakeStudents, fakeName } = require('./helpers.js');

// 生成与“高二X班开学考地理成绩表”同结构的工作簿：合并标题行 + 表头“序号/姓名/地理成绩”
async function scoreSheetLike(className, names, scores) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(className);
  ws.mergeCells('A1:C1');
  ws.getCell('A1').value = `${className}开学考地理成绩表`;
  ws.getRow(2).values = ['序号', '姓名', '地理成绩'];
  names.forEach((n, i) => { ws.getRow(i + 3).values = [i + 1, n, scores[i]]; });
  return wb.xlsx.writeBuffer();
}

async function load(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

async function roundTrip(wb) {
  return load(await wb.xlsx.writeBuffer());
}

function ctxFor(cls, students, items) {
  return id => (id === cls.id ? { cls, students, items: items || core.normalizeItems(core.DEFAULT_SCORE_ITEMS) } : null);
}

test('名单导入：跳过标题行，识别“序号/姓名/地理成绩”，班级名取自标题；缺考成绩为空', async () => {
  const names = [0, 1, 2, 3].map(fakeName);
  const buf = await scoreSheetLike('高二9班', names, [60, null, 45.5, 88]);
  const res = await excel.parseRosterWorkbook(buf, '高二9班开学考地理成绩表.xlsx');
  assert.equal(res.candidates.length, 1);
  const c = res.candidates[0];
  assert.equal(c.className, '高二9班');
  assert.deepEqual(c.students.map(s => s.name), names);
  assert.deepEqual(c.students.map(s => s.score), [60, null, 45.5, 88]);
  assert.deepEqual(c.students.map(s => s.seq), [1, 2, 3, 4]);
  assert.equal(c.students[0].gender, undefined, '没有性别列时不设置性别');
  assert.match(c.warnings.join(''), /1 人没有成绩/);
});

test('名单导入：性别列多种写法、姓名中的对齐空格、同名提示、按“班级”列拆成多个班', () => {
  const rows = [
    ['学生名单'],
    ['班级', '姓 名', '性别', '学号'],
    ['高二(1)班', '赵　测一', '男', '001'],
    ['高二(1)班', '钱测一', 'F', '002'],
    ['高二（2）班', '孙测一', '女生', '003'],
    ['高二（2）班', '孙测一', '？', '004'],
    ['', '合计', '', '']
  ];
  const res = excel.parseRosterRows(rows);
  assert.equal(res.groups.length, 2);
  const [g1, g2] = res.groups;
  assert.equal(g1.className, '高二1班');
  assert.deepEqual(g1.students.map(s => [s.name, s.gender, s.no]), [['赵测一', 'M', '001'], ['钱测一', 'F', '002']]);
  assert.equal(g2.className, '高二2班');
  assert.equal(g1.warnings.length, 0, '提示按班分别统计，1 班没有问题');
  assert.match(g2.warnings.join(''), /1 人的性别无法识别/);
  assert.match(g2.warnings.join(''), /同名/);
  assert.deepEqual(g2.students.map(s => [s.name, s.gender, s.no]), [['孙测一', 'F', '004']], '同班同名合并为一人');
});

test('名单导入：“班级”列的名称原样保留（任意前缀），只有数字时补“班”；缺成绩按班统计', () => {
  const rows = [
    ['班级', '姓名', '成绩'],
    ['测试1班', '赵测一', '90'],
    ['测试1班', '钱测一', ''],
    ['物化生2班', '孙测一', '80'],
    ['3', '李测一', ''],
    ['3', '周测一', '']
  ];
  const groups = excel.parseRosterRows(rows).groups;
  assert.deepEqual(groups.map(g => g.className), ['测试1班', '物化生2班', '3班']);
  assert.deepEqual(groups.map(g => g.warnings.join('')), ['1 人没有成绩（缺考或空白）', '', '2 人没有成绩（缺考或空白）']);
});

test('粘贴与 CSV：制表符、逗号带引号、GBK 编码都能读', () => {
  const tsv = '序号\t姓名\t性别\n1\t赵测一\t男\n2\t钱测一\t女\n';
  assert.deepEqual(excel.parseRosterText(tsv, '测试班').candidates[0].students.map(s => s.gender), ['M', 'F']);
  const onlyNames = '赵测一\n钱测一\n孙测一';
  const r2 = excel.parseRosterText(onlyNames, '测试班');
  assert.equal(r2.candidates[0].students.length, 3);
  const csv = '姓名,成绩\n"赵测一",90\n钱测一,"85"\n';
  assert.deepEqual(excel.parseRosterText(csv).candidates[0].students.map(s => s.score), [90, 85]);
  const gbk = Buffer.from([0xd0, 0xd5, 0xc3, 0xfb, 0x0a, 0xd5, 0xd4, 0xb2, 0xe2, 0xd2, 0xbb]); // “姓名\n赵测一” 的 GBK 编码
  assert.equal(excel.decodeText(gbk), '姓名\n赵测一');
});

function makeClass(n = 6) {
  const students = fakeStudents(n);
  const cls = { id: 'c1', name: '测试1班' };
  return { cls, students };
}

test('名单导出：不改动时读回为空；改性别/成绩/姓名、增删行都能读回；同一文件读两次编号相同', async () => {
  const { cls, students } = makeClass(5);
  students[4].gender = '';
  const wb = await excel.createRosterWorkbook(cls, students, { exportId: 'e1' });
  let wb2 = await roundTrip(wb);
  const none = await excel.readExportEdits(await wb2.xlsx.writeBuffer(), 'a.xlsx', ctxFor(cls, students));
  assert.equal(none.ops.length, 0);

  const ws = wb2.getWorksheet('名单');
  ws.getCell('C8').value = '女'; // 第 5 人 性别
  ws.getCell('D4').value = 99; // 第 1 人 成绩
  ws.getCell('B5').value = '钱测改'; // 第 2 人 改名
  ws.spliceRows(6, 1); // 删掉第 3 人
  ws.addRow([9, '新测一', '男', 70, '', '在班', '']);
  const buf = await wb2.xlsx.writeBuffer();
  const res = await excel.readExportEdits(buf, 'a.xlsx', ctxFor(cls, students));
  const byStudent = new Map(res.ops.map(o => [o.studentId, o.fields]));
  assert.equal(byStudent.get('s005').gender, 'F');
  assert.equal(byStudent.get('s001').score, 99);
  assert.equal(byStudent.get('s002').name, '钱测改');
  assert.equal(byStudent.get('s003').active, false);
  const added = res.ops.find(o => o.fields.name === '新测一');
  assert.ok(added && added.fields.gender === 'M' && added.fields.score === 70);
  assert.match(res.parts.join('；'), /新增 1 人/);
  const again = await excel.readExportEdits(buf, 'a-副本.xlsx', ctxFor(cls, students));
  assert.deepEqual(again.ops.map(o => o.id), res.ops.map(o => o.id), 'Mac 与教室读到同一处修改，编号应相同');
  assert.ok(res.ops.every(o => o.id.startsWith('x_')));
});

test('名单导出：只对部分列排序导致编号错位时，改按姓名对应，不会误改名', async () => {
  const { cls, students } = makeClass(3);
  const wb2 = await roundTrip(await excel.createRosterWorkbook(cls, students, { exportId: 'e2' }));
  const ws = wb2.getWorksheet('名单');
  // 只交换了姓名列（编号列没动）
  const n4 = ws.getCell('B4').value;
  ws.getCell('B4').value = ws.getCell('B6').value;
  ws.getCell('B6').value = n4;
  const res = await excel.readExportEdits(await wb2.xlsx.writeBuffer(), 'b.xlsx', ctxFor(cls, students));
  assert.ok(!res.ops.some(o => o.fields && o.fields.name), '不应产生改名');
  assert.match(res.warnings.join(''), /按姓名对应/);
});

test('座次表导出：学生版与教师版互为 180° 旋转；学生版对调两人能读回为定版；增删学生被拒绝', async () => {
  const { cls, students } = makeClass(10);
  const layout = { rows: 2, cols: 6, aisles: [2, 4], disabled: ['2,6'] };
  const seats = core.randomSeating(students, layout, [], {}, core.mulberry32(2)).seats;
  const rec = { finalId: 'f1', date: '2026-09-28', seats, layout, mode: 'random' };
  const wb = await roundTrip(await excel.createSeatingWorkbook(cls, rec, students, { exportId: 'e3' }));
  const byId = new Map(students.map(s => [s.id, s]));
  const wsS = wb.getWorksheet('学生版');
  const wsT = wb.getWorksheet('教师版');
  for (const [r, c, sid] of seats) {
    const ps = excel.seatCell('student', layout, r, c);
    const pt = excel.seatCell('teacher', layout, r, c);
    assert.equal(excel.cellText(wsS.getCell(ps.row, ps.col).value), byId.get(sid).name);
    assert.equal(excel.cellText(wsT.getCell(pt.row, pt.col).value), byId.get(sid).name);
  }
  // 蓝男红女
  const [r0, c0, sid0] = seats[0];
  const p0 = excel.seatCell('student', layout, r0, c0);
  assert.equal(wsS.getCell(p0.row, p0.col).font.color.argb, excel.COLOR[byId.get(sid0).gender]);

  const a = excel.seatCell('student', layout, seats[0][0], seats[0][1]);
  const b = excel.seatCell('student', layout, seats[1][0], seats[1][1]);
  const va = wsS.getCell(a.row, a.col).value;
  wsS.getCell(a.row, a.col).value = wsS.getCell(b.row, b.col).value;
  wsS.getCell(b.row, b.col).value = va;
  const res = await excel.readExportEdits(await wb.xlsx.writeBuffer(), 'c.xlsx', ctxFor(cls, students));
  assert.equal(res.ops.length, 1);
  const op = res.ops[0];
  assert.equal(op.k, 'final.set');
  assert.equal(op.base, 'f1');
  const m = core.seatsToMap(op.seats);
  assert.equal(m.get(core.seatKey(seats[0][0], seats[0][1])), seats[1][2]);
  assert.match(res.parts.join(''), /调整了 2 个座位/);

  wsS.getCell(a.row, a.col).value = '';
  const bad = await excel.readExportEdits(await wb.xlsx.writeBuffer(), 'c.xlsx', ctxFor(cls, students));
  assert.equal(bad.ops.length, 0);
  assert.match(bad.warnings.join(''), /不能增删学生/);
});

test('混合排布局导出后座位对应不变，修改可读回', async () => {
  const { cls, students } = makeClass(14);
  let layout = core.applyRowPattern({ rows: 2, cols: 8, aisles: [2, 4, 6] }, 1, 1, [2, 2, 2]);
  layout = core.applyRowPattern(layout, 2, 2, [3, 2, 3]);
  const seats = core.randomSeating(students, layout, [], {}, core.mulberry32(7)).seats;
  const wb = await roundTrip(await excel.createSeatingWorkbook(cls, { finalId: 'f2', date: '2026-09-28', seats, layout, mode: 'random' }, students, { exportId: 'e-mixed' }));
  const ws = wb.getWorksheet('学生版');
  const byId = new Map(students.map(s => [s.id, s]));
  seats.forEach(([r, c, sid]) => {
    const p = excel.seatCell('student', layout, r, c);
    assert.equal(excel.cellText(ws.getCell(p.row, p.col).value), byId.get(sid).name);
  });
  const a = excel.seatCell('student', layout, seats[0][0], seats[0][1]);
  const b = excel.seatCell('student', layout, seats[1][0], seats[1][1]);
  const name = ws.getCell(a.row, a.col).value;
  ws.getCell(a.row, a.col).value = ws.getCell(b.row, b.col).value;
  ws.getCell(b.row, b.col).value = name;
  const result = await excel.readExportEdits(await wb.xlsx.writeBuffer(), 'mixed.xlsx', ctxFor(cls, students));
  assert.equal(result.ops.length, 1);
  assert.deepEqual(result.ops[0].layout.rowAisles, layout.rowAisles);
});

test('课堂表现导出：删除、修改、新增、改姓名都能读回；旧表不会把网页里的新改动改回去', async () => {
  const { cls, students } = makeClass(4);
  const records = [
    { recId: 'r1', studentId: 's001', name: students[0].name, item: '迟到', value: -3, date: '2026-09-28', time: '2026-09-28 08:01:00' },
    { recId: 'r2', studentId: 's002', name: students[1].name, item: '优秀作业', value: 1, date: '2026-09-28', time: '2026-09-28 08:02:00' },
    { recId: 'r3', studentId: 's003', name: students[2].name, item: '说话', value: -3, date: '2026-09-29', time: '2026-09-29 08:03:00' }
  ];
  const wb = await roundTrip(await excel.createScoresWorkbook(cls, records, students, { exportId: 'e4', defaultDate: '2026-09-28' }));
  const sum = wb.getWorksheet('汇总统计');
  assert.match(sum.getCell('D3').value.formula, /SUMIFS/);
  const ws = wb.getWorksheet(excel.SCORE_SHEET);
  const hd = { name: 4, value: 7, id: 9 };
  const rowOf = id => { for (let r = 4; r <= ws.rowCount; r++) if (excel.cellText(ws.getCell(r, hd.id).value) === id) return r; return 0; };
  ws.getCell(rowOf('r2'), hd.value).value = 2; // 修改
  ws.getCell(rowOf('r3'), hd.name).value = students[3].name; // 改到另一名学生名下
  ws.spliceRows(rowOf('r1'), 1); // 删除
  ws.addRow(['', '2026-09-30', '', students[0].name, '', '优秀课堂表现', '', '', '']); // 新增，分值按表现项默认
  const res = await excel.readExportEdits(await wb.xlsx.writeBuffer(), 'd.xlsx', ctxFor(cls, students));
  const kinds = res.ops.map(o => o.k).sort();
  assert.deepEqual(kinds, ['score.add', 'score.add', 'score.delete', 'score.delete', 'score.edit']);
  const edit = res.ops.find(o => o.k === 'score.edit');
  assert.deepEqual(edit.patch, { value: 2 });
  const newRec = res.ops.find(o => o.k === 'score.add' && o.rec.item === '优秀课堂表现').rec;
  assert.equal(newRec.value, 1);
  assert.equal(newRec.date, '2026-09-30');

  // 把这些修改写进记录；之后网页里又把 r2 改成 5；再读一次同一份旧表，r2 不会被改回 2
  const base = records.map((r, i) => ({ k: 'score.add', id: `a${i}`, t: store.tsMake(1, i, 'a'), rec: Object.assign({ classId: 'c1' }, r) }));
  const fromExcel = res.ops.map((o, i) => Object.assign({ t: store.tsMake(2, i, 'b') }, o));
  const later = { k: 'score.edit', id: 'z1', t: store.tsMake(3, 0, 'a'), recId: 'r2', patch: { value: 5 } };
  const again = await excel.readExportEdits(await wb.xlsx.writeBuffer(), 'd.xlsx', ctxFor(cls, students));
  const reread = again.ops.map((o, i) => Object.assign({ t: store.tsMake(4, i, 'c') }, o));
  const st = store.reduce(base.concat(fromExcel, [later], reread));
  assert.equal(st.scores.get('r2').value, 5, '同一处修改编号相同，只算一次，不会盖掉之后的改动');
  assert.equal(st.scores.get('r1').deleted, true);
});

test('非本系统的表格、班级已不存在的导出表都能安全处理', async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('随便');
  assert.equal(await excel.readExportEdits(await wb.xlsx.writeBuffer(), 'x.xlsx', () => null), null);
  const { cls, students } = makeClass(2);
  const buf = await (await excel.createRosterWorkbook(cls, students, { exportId: 'e5' })).xlsx.writeBuffer();
  const res = await excel.readExportEdits(buf, 'y.xlsx', () => null);
  assert.match(res.warnings.join(''), /班级已不存在/);
});

test('文件名清理：去掉 Windows 与 U 盘不允许的字符和末尾空格', () => {
  assert.equal(excel.safeFileName('高二:1班 '), '高二1班');
  assert.equal(excel.safeFileName('a/b\\c*?'), 'abc');
  assert.equal(excel.safeFileName('...'), '未命名');
});

test('座次表：同一张导出表先后改两次，第二次读回不会被当成冲突', async () => {
  const { cls, students } = makeClass(10);
  const layout = { rows: 2, cols: 6, aisles: [2, 4] };
  const seats = core.randomSeating(students, layout, [], {}, core.mulberry32(5)).seats;
  const rec = { finalId: 'f1', date: '2026-09-28', seats, layout, mode: 'random' };
  const wb = await roundTrip(await excel.createSeatingWorkbook(cls, rec, students, { exportId: 'e9' }));
  const ws = wb.getWorksheet('学生版');
  const swap = (i, j) => {
    const a = excel.seatCell('student', layout, seats[i][0], seats[i][1]);
    const b = excel.seatCell('student', layout, seats[j][0], seats[j][1]);
    const v = ws.getCell(a.row, a.col).value;
    ws.getCell(a.row, a.col).value = ws.getCell(b.row, b.col).value;
    ws.getCell(b.row, b.col).value = v;
  };
  const clock = store.createClock('aaaa0001');
  const ops = [
    { k: 'class.create', id: 'o1', t: clock.now(), classId: cls.id, name: cls.name },
    { k: 'final.set', id: 'o2', t: clock.now(), classId: cls.id, finalId: 'f1', date: rec.date, seats, layout }
  ];
  swap(0, 1);
  (await excel.readExportEdits(await wb.xlsx.writeBuffer(), 'a.xlsx', ctxFor(cls, students))).ops.forEach(o => ops.push(Object.assign({}, o, { t: clock.now() })));
  swap(2, 3);
  const second = await excel.readExportEdits(await wb.xlsx.writeBuffer(), 'a.xlsx', ctxFor(cls, students));
  second.ops.forEach(o => ops.push(Object.assign({}, o, { t: clock.now() })));
  const c = store.reduce(ops).classes.get(cls.id);
  assert.equal(c.conflicts.length, 0);
  const now = core.seatsToMap(c.finals.get('2026-09-28').seats);
  assert.equal(now.get(core.seatKey(seats[0][0], seats[0][1])), seats[1][2]);
  assert.equal(now.get(core.seatKey(seats[2][0], seats[2][1])), seats[3][2]);
});

test('名单导入：同一班级里的同名学生按同一人合并，后面的行补充前面的字段', () => {
  const rows = [
    ['姓名', '性别', '成绩', '状态'],
    ['赵测一', '男', '', '离班'],
    ['钱测一', '女', '90', ''],
    ['赵测一', '', '88', ''],
    ['孙测一', '男', '70', '离班'],
    ['孙测一', '', '', '离班']
  ];
  const r = excel.parseRosterRows(rows, { className: '测试1班' });
  const list = r.groups[0].students;
  assert.deepEqual(list.map(s => s.name), ['赵测一', '钱测一', '孙测一']);
  assert.equal(list[0].gender, 'M');
  assert.equal(list[0].score, 88);
  assert.notEqual(list[0].active, false, '只要有一行在班就算在班');
  assert.equal(list[2].active, false);
  assert.match(r.groups[0].warnings.join(''), /2 组同名学生，已按同一名学生合并/);
});
