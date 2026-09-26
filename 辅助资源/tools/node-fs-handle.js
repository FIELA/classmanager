/**
 * 让 Node 工具用与网页相同的数据层读写磁盘：把真实目录包装成 File System Access 风格的句柄。
 * 写入先写到同目录的隐藏临时文件（以 . 开头，读取时会被忽略），再改名替换，避免写到一半的文件被同步走。
 */
'use strict';

const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');

function notFound(p) {
  return Object.assign(new Error(`not found: ${p}`), { name: 'NotFoundError' });
}

class NodeFileHandle {
  constructor(filePath) {
    this.kind = 'file';
    this.path = filePath;
    this.name = path.basename(filePath);
  }

  async getFile() {
    let st;
    try {
      st = await fsp.stat(this.path);
    } catch (e) {
      throw notFound(this.path);
    }
    const p = this.path;
    return {
      name: this.name,
      size: st.size,
      lastModified: Math.round(st.mtimeMs),
      async text() { return fsp.readFile(p, 'utf8'); },
      async arrayBuffer() {
        const b = await fsp.readFile(p);
        return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
      }
    };
  }

  async createWritable() {
    const chunks = [];
    const target = this.path;
    return {
      async write(data) {
        if (typeof data === 'string') chunks.push(Buffer.from(data, 'utf8'));
        else if (data instanceof ArrayBuffer) chunks.push(Buffer.from(data));
        else if (ArrayBuffer.isView(data)) chunks.push(Buffer.from(data.buffer, data.byteOffset, data.byteLength));
        else throw new Error('未知数据类型');
      },
      async close() {
        const tmp = path.join(path.dirname(target), `.${path.basename(target)}.${process.pid}.tmp`);
        await fsp.writeFile(tmp, Buffer.concat(chunks));
        await fsp.rename(tmp, target);
      }
    };
  }
}

class NodeDirHandle {
  constructor(dirPath) {
    this.kind = 'directory';
    this.path = dirPath;
    this.name = path.basename(dirPath);
  }

  async getDirectoryHandle(name, opts = {}) {
    const p = path.join(this.path, name);
    if (!fs.existsSync(p)) {
      if (!opts.create) throw notFound(p);
      await fsp.mkdir(p);
    } else if (!fs.statSync(p).isDirectory()) {
      throw Object.assign(new Error(`不是目录：${p}`), { name: 'TypeMismatchError' });
    }
    return new NodeDirHandle(p);
  }

  async getFileHandle(name, opts = {}) {
    const p = path.join(this.path, name);
    if (!fs.existsSync(p)) {
      if (!opts.create) throw notFound(p);
      await fsp.writeFile(p, '');
    }
    return new NodeFileHandle(p);
  }

  async *values() {
    const entries = await fsp.readdir(this.path, { withFileTypes: true });
    for (const e of entries) {
      const p = path.join(this.path, e.name);
      if (e.isDirectory()) yield new NodeDirHandle(p);
      else if (e.isFile()) yield new NodeFileHandle(p);
    }
  }
}

module.exports = { NodeDirHandle, NodeFileHandle };
