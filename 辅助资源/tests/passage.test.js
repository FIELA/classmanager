'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { core, store, fakeStudents } = require('./helpers.js');

test('分组过关：固定样本 5～10 人，所有非课代表都有机会向老师背诵；其余学生均衡分组', () => {
  for (const count of [15, 16, 20, 29, 30, 32, 35, 40]) {
    for (const sampleCount of [5, 7, 10]) {
      const students = fakeStudents(count);
      const representativeId = students[0].id;
      const sampleIds = students.slice(1, sampleCount + 1).map(s => s.id);
      const teachersSeen = new Set();
      let sawDualRole = false;
      for (let seed = 1; seed <= 100; seed++) {
        const result = core.drawPassageGroups(students, { representativeId, sampleIds }, core.mulberry32(seed));
        assert.equal(result.teacher.length, 5);
        assert.equal(result.teacher[0], representativeId);
        assert.equal(new Set(result.teacher).size, 5);
        assert.equal(result.representative.length, 5);
        assert.equal(new Set(result.representative).size, 5);
        assert.equal(result.groups.length, 5);
        assert.deepEqual(new Set(result.groups.map(g => g.leader)), new Set(result.representative));
        assert.ok(result.representative.every(id => sampleIds.includes(id)));
        const members = result.groups.flatMap(g => g.members);
        const roles = new Set(result.teacher.concat(result.representative));
        assert.equal(new Set(members).size, members.length);
        assert.ok(members.every(id => !roles.has(id)));
        assert.deepEqual(new Set(result.teacher.concat(result.representative, members)), new Set(students.map(s => s.id)));
        const sizes = result.groups.map(g => g.members.length);
        assert.ok(Math.max(...sizes) - Math.min(...sizes) <= 1);
        assert.ok(Math.min(...sizes) >= 1);
        result.teacher.slice(1).forEach(id => teachersSeen.add(id));
        if (result.teacher.some(id => result.representative.includes(id))) sawDualRole = true;
      }
      assert.deepEqual(teachersSeen, new Set(students.slice(1).map(s => s.id)));
      assert.ok(sawDualRole);
    }
  }
});

test('分组过关：无效配置、缺席固定成员及固定角色人数不足时拒绝抽取', () => {
  const students = fakeStudents(30);
  const sampleIds = students.slice(1, 11).map(s => s.id);
  assert.throws(() => core.drawPassageGroups(students, { representativeId: students[0].id, sampleIds: sampleIds.slice(0, 4) }), /5.*10/);
  assert.throws(() => core.drawPassageGroups(students, { representativeId: students[0].id, sampleIds: sampleIds.concat(students[11].id) }), /5.*10/);
  assert.throws(() => core.drawPassageGroups(students, { representativeId: students[0].id, sampleIds: [students[0].id].concat(sampleIds.slice(1)) }), /课代表/);
  students[1].active = false;
  assert.throws(() => core.drawPassageGroups(students, { representativeId: students[0].id, sampleIds }), /在班/);
  students[1].active = true;
  assert.throws(() => core.drawPassageGroups(students.slice(0, 14), { representativeId: students[0].id, sampleIds }), /至少 15/);
});

test('不足 30 人或组员超过 5 人的抽取结果都可以按周定版并读回', () => {
  for (const count of [20, 40]) {
    const students = fakeStudents(count);
    const result = core.drawPassageGroups(students, { representativeId: students[0].id, sampleIds: students.slice(1, 6).map(s => s.id) }, core.mulberry32(4));
    if (count === 40) assert.ok(result.groups.some(g => g.members.length > 5));
    const cls = store.reduce([
      { k: 'class.create', id: 'c', t: store.tsMake(1, 0, 'a'), classId: 'c1', name: '测试班' },
      { k: 'passage.final.set', id: 'p', t: store.tsMake(2, 0, 'a'), classId: 'c1', finalId: 'pf1', week: '2026-09-28', result }
    ]).classes.get('c1');
    assert.equal(cls.passageFinals.get('2026-09-28').finalId, 'pf1');
    assert.deepEqual(cls.passageFinals.get('2026-09-28').result, result);
  }
});

test('分组定版：允许老师名单与样本组长兼任，拒绝组员重复', () => {
  const students = fakeStudents(20);
  const sampleIds = students.slice(1, 6).map(s => s.id);
  const base = { k: 'class.create', id: 'c', t: store.tsMake(1, 0, 'a'), classId: 'c1', name: '测试班' };
  const config = { k: 'passage.config.set', id: 'pc', t: store.tsMake(2, 0, 'a'), classId: 'c1', configId: 'pc1', config: { representativeId: students[0].id, sampleIds } };
  const result = core.drawPassageGroups(students, config.config, core.mulberry32(4));
  assert.ok(result.teacher.some(id => result.representative.includes(id)));
  const op = { k: 'passage.final.set', id: 'pf', t: store.tsMake(3, 0, 'a'), classId: 'c1', finalId: 'pf1', week: '2026-09-28', result };
  const cls = store.reduce([base, config, op]).classes.get('c1');
  assert.deepEqual(cls.passageConfig.sampleIds, sampleIds);
  assert.deepEqual(cls.passageFinals.get(op.week).result, result);
  const duplicate = JSON.parse(JSON.stringify(result));
  duplicate.groups[0].members[0] = duplicate.teacher[0];
  assert.equal(store.reduce([base, { ...op, result: duplicate }]).classes.get('c1').passageFinals.size, 0);
});

test('分组定版：按周生效、同周顺序改版保留历史，异地并行定版提示冲突', () => {
  const base = { k: 'class.create', id: 'c', t: store.tsMake(1, 0, 'a'), classId: 'c1', name: '测试班' };
  const result = {
    teacher: ['s001', 's002', 's003', 's004', 's005'],
    representative: ['s006', 's007', 's008', 's009', 's010'],
    groups: Array.from({ length: 5 }, (_, i) => ({ leader: `s${String(i + 6).padStart(3, '0')}`, members: Array.from({ length: 4 }, (_, j) => `m${i}${j}`) }))
  };
  const op = (id, t, finalId, baseId) => ({
    k: 'passage.final.set', id, t: store.tsMake(t, 0, 'a'), classId: 'c1', finalId,
    week: '2026-09-28', result, base: baseId
  });
  let cls = store.reduce([base, op('a', 2, 'f1', null), op('b', 3, 'f2', 'f1')]).classes.get('c1');
  assert.equal(cls.passageFinals.get('2026-09-28').finalId, 'f2');
  assert.deepEqual(cls.passageHistory.map(x => x.finalId), ['f1', 'f2']);
  assert.equal(cls.conflicts.length, 0);
  cls = store.reduce([base, op('a', 2, 'f1', null), op('b', 3, 'f2', null)]).classes.get('c1');
  assert.equal(cls.conflicts[0].kind, 'passage');
  assert.equal(cls.conflicts[0].loser.finalId, 'f1');
});

test('固定人员设置：异地同时修改时提示冲突，顺序修改不提示', () => {
  const base = { k: 'class.create', id: 'c', t: store.tsMake(1, 0, 'a'), classId: 'c1', name: '测试班' };
  const config = { representativeId: 's001', sampleIds: Array.from({ length: 10 }, (_, i) => `s${String(i + 2).padStart(3, '0')}`) };
  const op = (id, t, configId, baseId) => ({ k: 'passage.config.set', id, t: store.tsMake(t, 0, 'a'), classId: 'c1', configId, config, base: baseId });
  let cls = store.reduce([base, op('a', 2, 'pc1', null), op('b', 3, 'pc2', 'pc1')]).classes.get('c1');
  assert.equal(cls.conflicts.length, 0);
  cls = store.reduce([base, op('a', 2, 'pc1', null), op('b', 3, 'pc2', null)]).classes.get('c1');
  assert.equal(cls.conflicts[0].kind, 'passage-config');
  assert.equal(cls.conflicts[0].loser.configId, 'pc1');
});
