#!/usr/bin/env node
/**
 * 提交 / 推送前的真实数据检查（仓库是公开的）。
 *   --staged  检查暂存区（pre-commit 使用）
 *   --push    检查即将推送、远端还没有的全部提交（pre-push 使用，从标准输入读取引用）
 *   --all     检查当前已跟踪的全部文件
 *   --message <文件>  检查提交信息（commit-msg 使用）
 * 拦截：数据/、导出/ 下的文件；表格与数据文件类型；以及出现在本机 数据/ 里的任何学生姓名。
 * 输出中的姓名一律打码。
 */
'use strict';

const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const git = (args, opts = {}) => execFileSync('git', args, Object.assign({ maxBuffer: 256 * 1024 * 1024 }, opts));
const ROOT = git(['rev-parse', '--show-toplevel'], { encoding: 'utf8' }).trim();

const BLOCKED_DIRS = ['数据/', '导出/'];
const BLOCKED_EXT = /\.(xlsx|xls|xlsm|et|csv|jsonl|numbers)$/i;

function mask(name) {
  const chars = Array.from(name);
  return chars[0] + '*'.repeat(Math.max(1, chars.length - 1));
}

// 从本机 数据/ 读取真实姓名（会话记录 *.jsonl 与快照 *.json）
function collectRealNames() {
  const names = new Set();
  const dir = path.join(ROOT, '数据');
  const addFromOp = op => {
    if (op && op.k === 'student.upsert' && op.fields && typeof op.fields.name === 'string') {
      const n = op.fields.name.trim();
      if (Array.from(n).length >= 2) names.add(n);
    }
  };
  const walk = d => {
    let entries = [];
    try { entries = fs.readdirSync(d, { withFileTypes: true }); } catch (e) { return; }
    for (const e of entries) {
      const p = path.join(d, e.name);
      if (e.isDirectory()) { walk(p); continue; }
      if (e.name.startsWith('.')) continue;
      let text = '';
      try { text = fs.readFileSync(p, 'utf8'); } catch (err) { continue; }
      if (e.name.endsWith('.jsonl')) {
        for (const line of text.split('\n')) {
          try { addFromOp(JSON.parse(line)); } catch (err) { /* 半行或损坏行忽略 */ }
        }
      } else if (e.name.endsWith('.json')) {
        try {
          const snap = JSON.parse(text);
          (snap.ops || []).forEach(addFromOp);
        } catch (err) { /* 忽略 */ }
      }
    }
  };
  walk(dir);
  return names;
}

function isBinary(buf) {
  const n = Math.min(buf.length, 8000);
  for (let i = 0; i < n; i++) if (buf[i] === 0) return true;
  return false;
}

// 需要检查的 (路径, 读取内容) 列表
function targets() {
  const argv = process.argv;
  const mode = argv.includes('--message') ? 'message' : argv.includes('--push') ? 'push' : argv.includes('--all') ? 'all' : 'staged';
  const list = [];
  if (mode === 'message') {
    const file = argv[argv.indexOf('--message') + 1];
    // 以 # 开头的是 git 的注释行，不会进入提交信息
    const text = fs.readFileSync(file, 'utf8').split('\n').filter(l => !l.startsWith('#')).join('\n');
    list.push({ file: '（提交信息）', where: '提交信息', message: true, read: () => Buffer.from(text, 'utf8') });
  } else if (mode === 'staged') {
    const files = git(['diff', '--cached', '--name-only', '--diff-filter=ACMR', '-z'], { encoding: 'utf8' }).split('\0').filter(Boolean);
    for (const f of files) list.push({ file: f, where: '暂存区', read: () => git(['show', ':' + f]) });
  } else if (mode === 'all') {
    const files = git(['ls-files', '-z'], { encoding: 'utf8' }).split('\0').filter(Boolean);
    for (const f of files) list.push({ file: f, where: '已跟踪', read: () => fs.readFileSync(path.join(ROOT, f)) });
  } else {
    const input = fs.readFileSync(0, 'utf8');
    const seenBlobs = new Set();
    for (const line of input.split('\n')) {
      const parts = line.trim().split(/\s+/);
      if (parts.length < 4) continue;
      const localSha = parts[1];
      if (/^0+$/.test(localSha)) continue; // 删除远端分支
      const commits = git(['rev-list', localSha, '--not', '--remotes=origin'], { encoding: 'utf8' }).split('\n').filter(Boolean);
      for (const c of commits) {
        const tree = git(['ls-tree', '-r', '-z', c], { encoding: 'utf8' }).split('\0').filter(Boolean);
        for (const entry of tree) {
          const m = entry.match(/^\d+ blob ([0-9a-f]+)\t(.+)$/);
          if (!m) continue;
          const key = m[1] + '\0' + m[2];
          if (seenBlobs.has(key)) continue;
          seenBlobs.add(key);
          const blob = m[1];
          list.push({ file: m[2], where: `提交 ${c.slice(0, 7)}`, read: () => git(['cat-file', 'blob', blob]) });
        }
      }
    }
  }
  return list;
}

function main() {
  const names = [...collectRealNames()];
  const problems = [];
  for (const t of targets()) {
    if (!t.message && BLOCKED_DIRS.some(d => t.file.startsWith(d))) {
      problems.push(`${t.where}：${t.file} 位于真实数据目录`);
      continue;
    }
    if (!t.message && BLOCKED_EXT.test(t.file)) {
      problems.push(`${t.where}：${t.file} 是表格/数据文件类型`);
      continue;
    }
    if (!names.length) continue;
    const buf = t.read();
    if (isBinary(buf)) continue;
    const text = buf.toString('utf8');
    const hits = names.filter(n => text.includes(n));
    if (hits.length) {
      problems.push(`${t.where}：${t.message ? '' : t.file + ' '}含有真实学生姓名（${hits.slice(0, 5).map(mask).join('、')}${hits.length > 5 ? ` 等 ${hits.length} 个` : ''}）`);
    }
  }
  if (problems.length) {
    console.error('⛔ 已阻止：以下内容可能包含真实数据，公开仓库不允许上传');
    problems.forEach(p => console.error('  - ' + p));
    console.error('请移除后再试（真实数据只放在 数据/、导出/，这两个目录已被忽略）。');
    process.exit(1);
  }
  if (process.argv.includes('--verbose')) console.log(`✓ 未发现真实数据（已比对 ${names.length} 个本机学生姓名）`);
}

main();
