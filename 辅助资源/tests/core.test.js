// 规则测试：布局、分区、四种排座、抽取、排名（全部为虚构数据）
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { core, fakeStudents } = require('./helpers.js');

const L58 = { rows: 5, cols: 8, aisles: [2, 4, 6], disabled: [] };
const L68 = { rows: 6, cols: 8, aisles: [2, 4, 6], disabled: [] };

function assertValid(seats, layout, students, requireAll = true) {
  const v = core.validateSeats(seats, layout, students, { requireAll });
  assert.ok(v.ok, v.errors.join('；'));
}

test('日期按本地时区：周日归前一周，周跨度为周一～周五', () => {
  assert.equal(core.mondayOf('2026-09-27'), '2026-09-21');
  assert.equal(core.mondayOf('2026-09-28'), '2026-09-28');
  assert.equal(core.weekSpan('2026-10-01'), '9.28～10.2');
  assert.equal(core.localISODate(new Date(2026, 8, 28, 7, 30)), '2026-09-28');
  assert.equal(core.isISODate('2026-02-30'), false);
});

test('布局规范化：限制范围、去掉越界走廊与无效座位', () => {
  const n = core.normalizeLayout({ rows: 99, cols: 0, aisles: [0, 3, 3, 50], disabled: ['1,1', '9,9', 'x'] });
  assert.equal(n.rows, 20);
  assert.equal(n.cols, 1);
  assert.deepEqual(n.aisles, []);
  const m = core.normalizeLayout({ rows: 5, cols: 8, aisles: [6, 2, 4, 8], disabled: ['5,8', '5,1', '5,1'] });
  assert.deepEqual(m.aisles, [2, 4, 6]);
  assert.deepEqual(m.disabled, ['5,1', '5,8']);
  assert.equal(core.capacity(m), 38);
  assert.deepEqual(core.groupsOf(m), [[1, 2], [3, 4], [5, 6], [7, 8]]);
});

test('不同排可分别设置座位组合，前排 2+2+2、末排 3+2+3', () => {
  let layout = { rows: 4, cols: 8, aisles: [2, 4, 6], disabled: [] };
  layout = core.applyRowPattern(layout, 1, 3, [2, 2, 2]);
  layout = core.applyRowPattern(layout, 4, 4, [3, 2, 3]);
  assert.equal(core.capacity(layout), 26);
  assert.deepEqual(core.aislesForRow(layout, 1), [3, 5]);
  assert.deepEqual(core.aislesForRow(layout, 4), [3, 5]);
  assert.deepEqual(core.allSeats(layout).filter(s => s.r === 1).map(s => s.c), [2, 3, 4, 5, 6, 7]);
  assert.deepEqual(core.allSeats(layout).filter(s => s.r === 4).map(s => s.c), [1, 2, 3, 4, 5, 6, 7, 8]);
  assert.deepEqual(core.groupsOf(layout, 1), [[1, 2, 3], [4, 5], [6, 7, 8]]);
  assert.equal(core.colZone(layout, 4, 4), '中');
  assert.equal(core.deskmatePairs(layout).filter(p => p[0].r === 1).length, 3);
  assert.equal(core.deskmatePairs(layout).filter(p => p[0].r === 4).length, 3);
  assert.equal(core.displayColumns(layout, 'student').filter(x => x.type === 'aisle').length, 2);
  const students = fakeStudents(26);
  assertValid(core.randomSeating(students, layout, [], {}, core.mulberry32(4)).seats, layout, students);
  assertValid(core.scoreSeating(students, layout, []).seats, layout, students);
});

test('单排走廊可独立改变，旧布局仍保持原样', () => {
  const legacy = { rows: 2, cols: 6, aisles: [2, 4] };
  const changed = core.normalizeLayout({ rows: 2, cols: 6, aisles: [2, 4], rowAisles: { 2: [3] } });
  assert.deepEqual(core.aislesForRow(changed, 1), [2, 4]);
  assert.deepEqual(core.aislesForRow(changed, 2), [3]);
  assert.notEqual(core.layoutKey(changed), core.layoutKey(legacy));
  assert.equal(core.colZone(changed, 3, 1), '中');
  assert.equal(core.colZone(changed, 3, 2), '左');
  assert.deepEqual(core.displayColumns(changed, 'teacher').filter(x => x.type === 'aisle').map(x => x.after), [4, 3, 2]);
  assert.deepEqual(core.groupsOf(legacy), [[1, 2], [3, 4], [5, 6]]);
});

test('同桌：同一组内从左起两两成对，不隔着不可用座位；按排设置过的排从该排实际的第一个座位算起', () => {
  const text = layout => core.deskmatePairs(layout).map(p => `${p[0].r}:${p[0].c}-${p[1].c}`);
  // 旧布局保持原来的配对：4 人一组，第 2 列或第 1 列不可用时，第 3、4 列仍是同桌
  assert.deepEqual(text({ rows: 1, cols: 8, aisles: [4], disabled: ['1,2'] }), ['1:3-4', '1:5-6', '1:7-8']);
  assert.deepEqual(text({ rows: 1, cols: 8, aisles: [4], disabled: ['1,1'] }), ['1:3-4', '1:5-6', '1:7-8']);
  // 套用组合：两侧空出的位置不算座位
  let layout = core.applyRowPattern({ rows: 2, cols: 8, aisles: [2, 4, 6] }, 1, 1, [2, 2, 2]);
  layout = core.applyRowPattern(layout, 2, 2, [3, 2, 3]);
  assert.deepEqual(text(layout), ['1:2-3', '1:4-5', '1:6-7', '2:1-2', '2:4-5', '2:6-7']);
  // 组合的走廊恰好与默认走廊相同时也一样
  assert.deepEqual(text(core.applyRowPattern({ rows: 1, cols: 8, aisles: [3, 5] }, 1, 1, [2, 2, 2])), ['1:2-3', '1:4-5', '1:6-7']);
});

test('布局写法不同但每排座位与走廊都相同，视为同一布局，可以继续轮换', () => {
  const base = core.normalizeLayout({ rows: 6, cols: 8, aisles: [2, 4, 6] });
  const same = core.applyRowPattern(base, 1, 6, [2, 2, 2, 2]);
  assert.equal(core.layoutKey(same), core.layoutKey(base));
  // 编辑器里把某排走廊点开再点掉
  assert.equal(core.layoutKey({ rows: 6, cols: 8, aisles: [2, 4, 6], rowAisles: { 3: [6, 4, 2] } }), core.layoutKey(base));
  // 每排都单独设成同样的走廊，与默认走廊就是这样的布局相同
  assert.equal(core.layoutKey({ rows: 2, cols: 8, aisles: [2, 4, 6], rowAisles: { 1: [3, 5], 2: [3, 5] } }), core.layoutKey({ rows: 2, cols: 8, aisles: [3, 5] }));
  assert.notEqual(core.layoutKey(core.applyRowPattern(base, 6, 6, [3, 2, 3])), core.layoutKey(base));
  const students = fakeStudents(44);
  const src = { seats: core.randomSeating(students, base, [], {}, core.mulberry32(3)).seats, layout: base };
  assertValid(core.rotateSeating(src, students, same, []).seats, same, students);
});

test('默认左右轮换：旧布局按第一组的列数；第 1 排套用组合时不算两侧空位', () => {
  assert.equal(core.defaultRotation({ rows: 6, cols: 8, aisles: [2, 4, 6], disabled: ['1,1'] }).colShift, 2);
  assert.equal(core.defaultRotation({ rows: 6, cols: 8, aisles: [4], disabled: ['1,1'] }).colShift, 4);
  assert.equal(core.defaultRotation(core.applyRowPattern({ rows: 6, cols: 8, aisles: [2, 4, 6] }, 1, 6, [2, 2, 2])).colShift, 2);
  assert.equal(core.defaultRotation(core.applyRowPattern({ rows: 6, cols: 8, aisles: [2, 4, 6] }, 1, 6, [3, 3])).colShift, 3);
});

test('前中后区按排数三等分，左中右按大组', () => {
  const zones = rows => Array.from({ length: rows }, (_, i) => core.rowZone({ rows, cols: 4 }, i + 1)).join('');
  assert.equal(zones(6), '前前中中后后');
  assert.equal(zones(7), '前前前中中后后');
  assert.equal(zones(5), '前前中中后');
  assert.equal(zones(2), '前后');
  assert.equal(zones(1), '前');
  const cz = l => Array.from({ length: l.cols }, (_, i) => core.colZone(l, i + 1)).join('');
  assert.equal(cz(L68), '左左中中中中右右');
  assert.equal(cz({ rows: 1, cols: 6, aisles: [3] }), '左左左右右右');
  assert.equal(cz({ rows: 1, cols: 8, aisles: [] }), '左左左中中右右右');
});

test('教师版是学生版旋转 180°：列与走廊镜像、排倒序', () => {
  const s = core.displayColumns(L58, 'student').map(x => (x.type === 'seat' ? x.c : '|')).join('');
  const t = core.displayColumns(L58, 'teacher').map(x => (x.type === 'seat' ? x.c : '|')).join('');
  assert.equal(s, '12|34|56|78');
  assert.equal(t, '87|65|43|21');
  assert.deepEqual(core.displayRows(L58, 'teacher'), [5, 4, 3, 2, 1]);
  const odd = { rows: 2, cols: 5, aisles: [1, 3] };
  assert.equal(core.displayColumns(odd, 'student').map(x => (x.type === 'seat' ? x.c : '|')).join(''), '1|23|45');
  assert.equal(core.displayColumns(odd, 'teacher').map(x => (x.type === 'seat' ? x.c : '|')).join(''), '54|32|1');
});

test('填座顺序：前排优先、同排由中间向两侧，空座留在最后一排两侧', () => {
  const students = fakeStudents(34);
  const res = core.randomSeating(students, L58, [], {}, core.mulberry32(1));
  assertValid(res.seats, L58, students);
  const lastRow = res.seats.filter(t => t[0] === 5).map(t => t[1]).sort();
  assert.deepEqual(lastRow, [4, 5]);
});

test('随机排座：每人恰好一个座位；固定学生不动；离班学生不排；座位不够时报错', () => {
  const students = fakeStudents(30);
  students[29].active = false;
  const pins = [{ sid: 's001', r: 5, c: 8 }, { sid: 's002', r: 1, c: 1 }, { sid: 's030', r: 2, c: 2 }];
  for (let seed = 1; seed <= 20; seed++) {
    const res = core.randomSeating(students, L58, pins, {}, core.mulberry32(seed));
    assertValid(res.seats, L58, students);
    const m = core.seatsToMap(res.seats);
    assert.equal(m.get('5,8'), 's001');
    assert.equal(m.get('1,1'), 's002');
    assert.ok(!res.seats.some(t => t[2] === 's030'), '离班学生不应有座位');
  }
  const a = core.randomSeating(students, L58, [], {}, core.mulberry32(7)).seats;
  const b = core.randomSeating(students, L58, [], {}, core.mulberry32(7)).seats;
  assert.deepEqual(a, b, '同一随机种子结果应一致');
  assert.throws(() => core.randomSeating(fakeStudents(41), L58, [], {}), /座位不够/);
});

test('随机排座·男女搭配：男女人数足够时同桌都是一男一女', () => {
  const students = fakeStudents(32); // 男女各 16
  const res = core.randomSeating(students, { rows: 4, cols: 8, aisles: [2, 4, 6] }, [], { genderPair: true }, core.mulberry32(3));
  const byId = new Map(students.map(s => [s.id, s]));
  const m = core.seatsToMap(res.seats);
  for (const [a, b] of core.deskmatePairs({ rows: 4, cols: 8, aisles: [2, 4, 6] })) {
    const ga = byId.get(m.get(core.seatKey(a.r, a.c))).gender;
    const gb = byId.get(m.get(core.seatKey(b.r, b.c))).gender;
    assert.notEqual(ga, gb, `第${a.r}排第${a.c}列与第${b.c}列应为一男一女`);
  }
  assert.match(res.notes.join(''), /16 对/);
});

test('按成绩自动排：第一名坐前排正中，没有成绩的排在最后', () => {
  const students = fakeStudents(20);
  students[0].score = null;
  students[5].score = 200;
  const res = core.scoreSeating(students, L58, []);
  const m = core.seatsToMap(res.seats);
  assert.equal(m.get('1,4'), 's006');
  const order = core.fillOrder(L58);
  const last = order[19];
  assert.equal(m.get(core.seatKey(last.r, last.c)), 's001');
  assert.match(res.notes.join(''), /1 人没有成绩/);
});

test('排名：按成绩从高到低，并列按序号，没有成绩在后；全班无成绩时按名次', () => {
  const list = [
    { id: 'a', name: '甲', score: 90, seq: 3 },
    { id: 'b', name: '乙', score: 95, seq: 9 },
    { id: 'c', name: '丙', score: 90, seq: 1 },
    { id: 'd', name: '丁', score: null, seq: 2 }
  ];
  assert.deepEqual(core.rankStudents(list).map(s => s.id), ['b', 'c', 'a', 'd']);
  const byRank = [{ id: 'a', name: '甲', rank: 3 }, { id: 'b', name: '乙', rank: 1 }, { id: 'c', name: '丙' }];
  assert.deepEqual(core.rankStudents(byRank).map(s => s.id), ['b', 'a', 'c']);
});

function fullSource(students, layout, seed) {
  return { seats: core.randomSeating(students, layout, [], {}, core.mulberry32(seed)).seats, layout };
}

test('轮换：前后每列后移、左右每排右移；是一一对应；空座与固定学生不动', () => {
  const students = fakeStudents(46); // 48 座，2 个空座
  const src = fullSource(students, L68, 11);
  const empties = core.allSeats(L68).map(s => core.seatKey(s.r, s.c)).filter(k => !core.seatsToMap(src.seats).has(k));
  const pinned = src.seats[10];
  const pins = [{ sid: pinned[2], r: pinned[0], c: pinned[1] }];
  const res = core.rotateSeating(src, students, L68, pins, { rowShift: 2, colShift: 2 });
  assertValid(res.seats, L68, students);
  const after = core.seatsToMap(res.seats);
  empties.forEach(k => assert.ok(!after.has(k), `空座 ${k} 应保持为空`));
  assert.equal(after.get(core.seatKey(pinned[0], pinned[1])), pinned[2], '固定学生应原地不动');

  // 满座时的精确位移：第 r 排第 c 列 → 第 r+2 排、第 c+2 列（循环）
  const full = fakeStudents(48);
  const src2 = fullSource(full, L68, 5);
  const moved = core.seatsToMap(core.rotateSeating(src2, full, L68, [], { rowShift: 2, colShift: 2 }).seats);
  src2.seats.forEach(([r, c, sid]) => {
    const r2 = ((r - 1 + 2) % 6) + 1;
    const c2 = ((c - 1 + 2) % 8) + 1;
    assert.equal(moved.get(core.seatKey(r2, c2)), sid);
  });
});

test('轮换：连续轮换若干次后回到原位；每次都换前中后区（满座、按区宽移动）', () => {
  const full = fakeStudents(48);
  let cur = fullSource(full, L68, 9);
  const start = core.seatsToMap(cur.seats);
  const zone = new Map(cur.seats.map(t => [t[2], core.rowZone(L68, t[0])]));
  for (let i = 1; i <= 12; i++) {
    const next = core.rotateSeating(cur, full, L68, [], core.defaultRotation(L68));
    next.seats.forEach(t => {
      const z = core.rowZone(L68, t[0]);
      assert.notEqual(z, zone.get(t[2]), '每次轮换都应换区');
      zone.set(t[2], z);
    });
    cur = { seats: next.seats, layout: L68 };
  }
  // 前后周期 3、左右周期 4 → 12 次后回到原位
  assert.deepEqual(core.seatsToMap(cur.seats), start);
});

test('轮换：新加入与离班学生；布局改过则拒绝轮换', () => {
  const students = fakeStudents(30);
  const src = fullSource(students.slice(0, 29), L58, 2);
  students[3].active = false;
  const res = core.rotateSeating(src, students, L58, [], {});
  assertValid(res.seats, L58, students);
  assert.ok(res.seats.some(t => t[2] === 's030'), '新同学应有座位');
  assert.ok(!res.seats.some(t => t[2] === 's004'), '离班同学不应有座位');
  assert.match(res.notes.join(''), /新加入/);
  assert.throws(() => core.rotateSeating(src, students, { rows: 6, cols: 8, aisles: [2, 4, 6] }, [], {}), /布局/);
});

test('沿用：座位不变；布局缩小后原座位不存在的学生补到空座', () => {
  const students = fakeStudents(30);
  const src = fullSource(students, L58, 4);
  const same = core.holdSeating(src, students, L58, []);
  assert.deepEqual(same.seats, src.seats);
  const smaller = { rows: 4, cols: 8, aisles: [2, 4, 6] };
  const res = core.holdSeating(src, students, smaller, []);
  assertValid(res.seats, smaller, students);
});

test('校验：座位重复、不可用座位、学生重复、在班学生没座位都能发现', () => {
  const students = fakeStudents(3);
  const layout = { rows: 1, cols: 3, disabled: ['1,3'] };
  const v = core.validateSeats([[1, 1, 's001'], [1, 1, 's002'], [1, 3, 's001']], layout, students, { requireAll: true });
  assert.equal(v.ok, false);
  const text = v.errors.join('；');
  assert.match(text, /重复安排/);
  assert.match(text, /不是可用座位/);
  assert.match(text, /两个座位/);
  assert.match(text, /没有座位/);
});

test('抽取：不重复、人数有上下限、可排除已抽过的；多次抽取大致均匀', () => {
  const pool = Array.from({ length: 10 }, (_, i) => `p${i}`);
  const rng = core.mulberry32(42);
  const three = core.drawStudents(pool, 3, rng);
  assert.equal(new Set(three).size, 3);
  assert.equal(core.drawStudents(pool, 99, rng).length, 10);
  assert.equal(core.drawStudents(pool, 0, rng).length, 0);
  assert.deepEqual(core.drawStudents(pool, 5, rng, pool.slice(0, 8)).sort(), ['p8', 'p9']);
  const count = new Map(pool.map(p => [p, 0]));
  for (let i = 0; i < 20000; i++) {
    const [x] = core.drawStudents(pool, 1, rng);
    count.set(x, count.get(x) + 1);
  }
  count.forEach(n => assert.ok(n > 1800 && n < 2200, `分布不均匀：${n}`));
});

test('姓名与性别规范化', () => {
  assert.equal(core.normalizeName(' 赵　测一 '), '赵测一');
  assert.equal(core.normalizeName('阿卜杜•测'), '阿卜杜·测');
  assert.equal(core.normalizeName('Anna  Lee'), 'Anna Lee');
  assert.equal(core.normalizeGender('男生'), 'M');
  assert.equal(core.normalizeGender('F'), 'F');
  assert.equal(core.normalizeGender(''), '');
  assert.equal(core.normalizeGender('未知'), null);
  const names = core.displayNames([{ id: 'a1', name: '赵测一', seq: 3 }, { id: 'a2', name: '赵测一', seq: 8 }, { id: 'a3', name: '钱测一' }]);
  assert.equal(names.get('a1'), '赵测一(3号)');
  assert.equal(names.get('a3'), '钱测一');
});

test('标注性别：画笔标注，再点同一性别取消；剩余反标只改未标注的在班学生', () => {
  assert.equal(core.applyGenderBrush('', 'F'), 'F');
  assert.equal(core.applyGenderBrush('M', 'F'), 'F', '标错了直接改');
  assert.equal(core.applyGenderBrush('F', 'F'), '', '再点一次取消');
  assert.equal(core.applyGenderBrush('M', 'M'), '');
  const students = [
    { id: 'a', name: '赵测一', gender: 'F' },
    { id: 'b', name: '钱测一', gender: '' },
    { id: 'c', name: '孙测一', gender: 'M' },
    { id: 'd', name: '李测一', gender: '' },
    { id: 'e', name: '周测一', gender: '', active: false }
  ];
  assert.deepEqual(core.fillRemainingGender(students, 'F'), [{ sid: 'b', gender: 'M' }, { sid: 'd', gender: 'M' }]);
  assert.deepEqual(core.fillRemainingGender(students, 'M'), [{ sid: 'b', gender: 'F' }, { sid: 'd', gender: 'F' }]);
  assert.deepEqual(core.fillRemainingGender(students, ''), []);
});

test('课堂表现汇总与默认表现项', () => {
  const items = core.normalizeItems(core.DEFAULT_SCORE_ITEMS);
  assert.equal(items.length, 14);
  assert.equal(items.filter(i => i.value > 0).length, 2);
  const sum = core.summarizeScores([
    { studentId: 'a', value: 1 }, { studentId: 'a', value: -3 }, { studentId: 'a', value: 1, deleted: true }, { studentId: 'b', value: -10 }
  ]);
  assert.deepEqual(sum.get('a'), { bonusCount: 1, bonusSum: 1, penaltyCount: 1, penaltySum: -3, net: -2 });
  assert.equal(sum.get('b').net, -10);
});

test('同名学生：同班按一人对应（优先在班、序号靠前）；不同班级给出提醒', () => {
  const list = [
    { id: 'a', name: '赵测一', seq: 5, active: true },
    { id: 'b', name: '赵测一', seq: 2, active: false },
    { id: 'c', name: '赵测一', seq: 9, active: true },
    { id: 'd', name: '钱测一', seq: 1, active: true }
  ];
  const m = core.rosterByName(list);
  assert.equal(m.get('赵测一').id, 'a');
  assert.equal(m.get('钱测一').id, 'd');

  const clash = core.crossClassNames([
    { key: 'c1', className: '测试1班', names: ['赵测一', '钱测一'], incoming: false },
    { key: 'c2', className: '测试2班', names: ['孙测一', '李测一'], incoming: false },
    { key: 'c3', className: '测试3班', names: ['李测一'], incoming: false },
    { key: 'c1', className: '测试1班', names: ['钱测一', '周测一'], incoming: true },
    { key: 'new:1', className: '测试4班', names: ['赵测一', '周测一', '周测一'], incoming: true }
  ]);
  assert.deepEqual(clash, [
    { name: '赵测一', classes: ['测试1班', '测试4班'] },
    { name: '周测一', classes: ['测试1班', '测试4班'] }
  ], '合并到同一班的同名不算；两个已有班级之间原有的同名不在这次提醒范围');
});
