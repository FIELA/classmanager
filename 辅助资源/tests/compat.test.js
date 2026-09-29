// 兼容性与静态检查：Win7 上最高只有 Chrome/Edge 109，网页只能用 Chrome 86 已支持的写法
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const ROOT = path.join(__dirname, '..', '..');
const ASSETS = path.join(ROOT, '辅助资源', 'assets');
const read = p => fs.readFileSync(p, 'utf8');
const OWN_JS = ['core.js', 'store.js', 'excel.js', 'app.js'];

// 去掉注释和字符串里的内容，避免误报
function codeOnly(src) {
  return src
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/(^|[^:\\])\/\/.*$/gm, '$1')
    .replace(/`(?:\\[\s\S]|[^`\\])*`/g, '``')
    .replace(/'(?:\\.|[^'\\\n])*'/g, "''")
    .replace(/"(?:\\.|[^"\\\n])*"/g, '""');
}

test('脚本不使用 Chrome 86 之后才有的语法和接口', () => {
  const banned = [
    [/\.at\(/, 'Array.prototype.at（Chrome 92）'],
    [/\bstructuredClone\b/, 'structuredClone（Chrome 98）'],
    [/\brandomUUID\b/, 'crypto.randomUUID（Chrome 92）'],
    [/\bObject\.hasOwn\b/, 'Object.hasOwn（Chrome 93）'],
    [/\.findLast(Index)?\(/, 'findLast（Chrome 97）'],
    [/\?\?=|\|\|=|&&=/, '逻辑赋值（Chrome 85，保守起见不用）'],
    [/\bPromise\.any\b|\bAggregateError\b/, 'Promise.any（Chrome 85）'],
    [/\bWeakRef\b|\bFinalizationRegistry\b/, 'WeakRef（Chrome 84）'],
    [/#[a-zA-Z_]\w*\s*[(=;]/, '私有字段/方法'],
    [/^\s*import\s|^\s*export\s/m, 'ES 模块（file:// 下无法加载）'],
    [/\?\.\s*[\w([]/, '可选链（保守起见不用）'],
    [/\.replaceAll\(/, 'replaceAll（保守起见不用）'],
    [/\bstatic\s+\w+\s*=/, '静态类字段（保守起见不用）']
  ];
  for (const f of OWN_JS) {
    const code = codeOnly(read(path.join(ASSETS, f)));
    for (const [re, what] of banned) assert.ok(!re.test(code), `${f} 使用了 ${what}`);
  }
});

test('样式不使用 Chrome 86 不支持的特性', () => {
  const css = read(path.join(ASSETS, 'app.css')).replace(/\/\*[\s\S]*?\*\//g, '');
  const banned = [
    [/aspect-ratio\s*:/, 'aspect-ratio（Chrome 88）'],
    [/(^|[;{\s])inset\s*:/, 'inset（Chrome 87）'],
    [/:is\(|:where\(|:has\(/, ':is/:where/:has'],
    [/@container|container-type/, '容器查询'],
    [/color-mix\(/, 'color-mix'],
    [/\b\d+(dvh|svh|lvh)\b/, '动态视口单位'],
    [/accent-color/, 'accent-color（Chrome 93）'],
    [/&\s*[.:#[]/, 'CSS 嵌套']
  ];
  for (const [re, what] of banned) assert.ok(!re.test(css), `app.css 使用了 ${what}`);
});

test('页面引用的脚本与样式都存在；app.js 用到的元素编号都在页面里', () => {
  const html = read(path.join(ROOT, '班级助理.html'));
  const refs = [...html.matchAll(/(?:src|href)="(辅助资源\/[^"]+)"/g)].map(m => m[1]);
  assert.ok(refs.length >= 5);
  refs.forEach(r => assert.ok(fs.existsSync(path.join(ROOT, r)), `缺少文件：${r}`));
  const ids = new Set([...html.matchAll(/\sid="([^"]+)"/g)].map(m => m[1]));
  const app = read(path.join(ASSETS, 'app.js'));
  const used = new Set([...app.matchAll(/\$\('([^']+)'\)/g)].map(m => m[1]));
  const dynamic = new Set(['opt-gender-pair', 'opt-row-shift', 'opt-row-dir', 'opt-col-shift', 'opt-col-dir']);
  const missing = [...used].filter(id => !ids.has(id) && !dynamic.has(id) && !id.startsWith('panel-'));
  assert.deepEqual(missing, [], '页面里缺少这些元素');
  ['seating', 'passage', 'scores', 'student', 'settings'].forEach(t => assert.ok(ids.has(`panel-${t}`)));
});

test('浏览器存储名与旧项目不同，不会互相覆盖', () => {
  const app = read(path.join(ASSETS, 'app.js'));
  assert.ok(!/classroom_|ClassroomManagementDB/.test(app));
  assert.match(app, /classmanager\.v1\./);
});
