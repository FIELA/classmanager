// 测试工具：内存伪目录（模拟浏览器 File System Access）、SyncTime 双向同步模拟器、虚构名单
// 测试一律使用虚构姓名，不读取任何真实数据。
'use strict';

process.env.TZ = 'Asia/Shanghai';

const path = require('node:path');
const ASSETS = path.join(__dirname, '..', 'assets');
global.ExcelJS = require(path.join(ASSETS, 'exceljs.min.js'));
const core = require(path.join(ASSETS, 'core.js'));
const store = require(path.join(ASSETS, 'store.js'));
const excel = require(path.join(ASSETS, 'excel.js'));

const enc = new TextEncoder();
const dec = new TextDecoder();

let globalClock = 1000000;

class FakeDir {
  constructor(name) {
    this.kind = 'directory';
    this.name = name || 'root';
    this.files = new Map(); // name -> { data: Uint8Array, lastModified }
    this.dirs = new Map();
    this.failWrites = false;
    this.writeLog = [];
  }

  async getFileHandle(name, opts = {}) {
    if (this.dirs.has(name)) throw Object.assign(new Error('是目录'), { name: 'TypeMismatchError' });
    if (!this.files.has(name)) {
      if (!opts.create) throw Object.assign(new Error('not found'), { name: 'NotFoundError' });
      if (this.failWrites) throw Object.assign(new Error('U 盘已拔出'), { name: 'NotFoundError' });
      this.files.set(name, { data: new Uint8Array(0), lastModified: ++globalClock });
    }
    const dir = this;
    return {
      kind: 'file',
      name,
      async getFile() {
        const f = dir.files.get(name);
        if (!f) throw Object.assign(new Error('not found'), { name: 'NotFoundError' });
        const bytes = f.data;
        return {
          name,
          size: bytes.length,
          lastModified: f.lastModified,
          async text() { return dec.decode(bytes); },
          async arrayBuffer() { return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength); }
        };
      },
      async createWritable() {
        let buf = new Uint8Array(0);
        return {
          async write(data) {
            if (typeof data === 'string') buf = enc.encode(data);
            else if (data instanceof Uint8Array) buf = data;
            else if (data instanceof ArrayBuffer) buf = new Uint8Array(data);
            else if (data && data.buffer) buf = new Uint8Array(data.buffer, data.byteOffset || 0, data.byteLength);
            else throw new Error('未知数据类型');
          },
          async close() {
            if (dir.failWrites) throw Object.assign(new Error('U 盘已拔出'), { name: 'NotFoundError' });
            dir.files.set(name, { data: buf, lastModified: ++globalClock });
            dir.writeLog.push(name);
          }
        };
      }
    };
  }

  async getDirectoryHandle(name, opts = {}) {
    if (!this.dirs.has(name)) {
      if (!opts.create) throw Object.assign(new Error('not found'), { name: 'NotFoundError' });
      const d = new FakeDir(name);
      d.failWrites = this.failWrites;
      this.dirs.set(name, d);
    }
    return this.dirs.get(name);
  }

  async removeEntry(name) {
    if (!this.files.delete(name) && !this.dirs.delete(name)) throw Object.assign(new Error('not found'), { name: 'NotFoundError' });
  }

  async *values() {
    for (const name of Array.from(this.files.keys())) yield await this.getFileHandle(name);
    for (const d of Array.from(this.dirs.values())) yield d;
  }

  setFailWrites(v) {
    this.failWrites = v;
    this.dirs.forEach(d => d.setFailWrites(v));
  }

  // 递归列出 { 路径 -> 文件 }
  flatten(prefix = '') {
    const out = new Map();
    this.files.forEach((f, n) => out.set(prefix + n, f));
    this.dirs.forEach((d, n) => d.flatten(prefix + n + '/').forEach((f, p) => out.set(p, f)));
    return out;
  }

  // 按路径写入（同步模拟器用：保留修改时间，如同复制）
  putPath(p, file) {
    const parts = p.split('/');
    let d = this;
    for (const seg of parts.slice(0, -1)) {
      if (!d.dirs.has(seg)) d.dirs.set(seg, new FakeDir(seg));
      d = d.dirs.get(seg);
    }
    d.files.set(parts[parts.length - 1], { data: file.data, lastModified: file.lastModified });
  }

  deletePath(p) {
    const parts = p.split('/');
    let d = this;
    for (const seg of parts.slice(0, -1)) {
      d = d.dirs.get(seg);
      if (!d) return;
    }
    d.files.delete(parts[parts.length - 1]);
  }

  async writeText(p, text) {
    const parts = p.split('/');
    let d = this;
    for (const seg of parts.slice(0, -1)) d = await d.getDirectoryHandle(seg, { create: true });
    const fh = await d.getFileHandle(parts[parts.length - 1], { create: true });
    const w = await fh.createWritable();
    await w.write(text);
    await w.close();
  }

  async readText(p) {
    const f = this.flatten().get(p);
    return f ? dec.decode(f.data) : null;
  }
}

/**
 * SyncTime“双向同步”的模拟（按其帮助文档）：
 *  - 每次同步后记录两边的快照（修改时间 + 大小）；
 *  - 只在一边新增/修改 → 复制到另一边；只在一边删除 → 另一边也删除；
 *  - 两边都改过 → 冲突：记一条错误，两边都不动。
 */
class SyncTimeSim {
  constructor(a, b) {
    this.a = a;
    this.b = b;
    this.snap = null;
    this.log = [];
  }

  sync() {
    const A = this.a.flatten();
    const B = this.b.flatten();
    const conflicts = [];
    const sig = f => (f ? `${f.lastModified}|${f.data.length}` : null);
    const paths = new Set([...A.keys(), ...B.keys()]);
    const first = !this.snap;
    for (const p of paths) {
      const fa = A.get(p);
      const fb = B.get(p);
      const s = this.snap ? this.snap.get(p) : null;
      if (first) {
        if (fa && fb) {
          if (sig(fa) !== sig(fb)) conflicts.push(p);
        } else if (fa) this.b.putPath(p, fa);
        else this.a.putPath(p, fb);
        continue;
      }
      const changedA = fa ? (!s || s.a !== sig(fa)) : !!(s && s.a);
      const changedB = fb ? (!s || s.b !== sig(fb)) : !!(s && s.b);
      if (changedA && changedB) {
        if (fa && fb && Buffer.compare(Buffer.from(fa.data), Buffer.from(fb.data)) === 0) continue;
        conflicts.push(p);
        continue;
      }
      if (changedA) {
        if (fa) this.b.putPath(p, fa);
        else this.b.deletePath(p);
      } else if (changedB) {
        if (fb) this.a.putPath(p, fb);
        else this.a.deletePath(p);
      }
    }
    const A2 = this.a.flatten();
    const B2 = this.b.flatten();
    this.snap = new Map();
    new Set([...A2.keys(), ...B2.keys()]).forEach(p => this.snap.set(p, { a: sig(A2.get(p)), b: sig(B2.get(p)) }));
    this.log.push({ conflicts });
    return { conflicts };
  }
}

// 虚构姓名：“姓 + 测 + 编号字”，不会与真实学生重名
const SURNAMES = ['赵', '钱', '孙', '李', '周', '吴', '郑', '冯', '陈', '褚', '卫', '蒋', '沈', '韩', '杨'];
const DIGITS = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];
function fakeName(i) {
  return `${SURNAMES[i % SURNAMES.length]}测${DIGITS[Math.floor(i / SURNAMES.length) % DIGITS.length]}${i >= 150 ? DIGITS[Math.floor(i / 150) % 10] : ''}`;
}

function fakeStudents(n, opts = {}) {
  const out = [];
  for (let i = 0; i < n; i++) {
    out.push({
      id: `s${String(i + 1).padStart(3, '0')}`,
      name: fakeName(i),
      gender: opts.noGender ? '' : (i % 2 ? 'F' : 'M'),
      score: opts.noScore ? null : 100 - i,
      rank: null,
      seq: i + 1,
      no: '',
      active: true
    });
  }
  return out;
}

module.exports = { core, store, excel, FakeDir, SyncTimeSim, fakeName, fakeStudents, enc, dec };
