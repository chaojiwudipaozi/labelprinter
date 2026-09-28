/**
 * 路径工具：兼容「网页版（源码运行）」与「Electron 桌面版（asar 打包）」
 *  - PowerShell 脚本必须用真实文件路径（asar 内不可被 child_process 执行），
 *    因此在打包环境自动指向 app.asar.unpacked
 *  - 可写数据目录在打包环境指向用户数据目录（app.asar 内是只读的）
 */
const path = require('path');
const fs = require('fs');

const ROOT = path.join(__dirname, '..');

/** 获取 scripts 目录下脚本的真实路径 */
function scriptPath(name) {
  const p = path.join(ROOT, 'scripts', name);
  const marker = `app.asar${path.sep}`;
  return p.includes(marker) ? p.replace(marker, `app.asar.unpacked${path.sep}`) : p;
}

/** 可写数据目录（默认 <项目>/data，Electron 下为 userData/data） */
function dataDir(...sub) {
  const base = process.env.LABEL_DATA_DIR
    ? path.resolve(process.env.LABEL_DATA_DIR)
    : path.join(ROOT, 'data');
  const p = path.join(base, ...sub);
  try { fs.mkdirSync(p, { recursive: true }); } catch (e) { /* ignore */ }
  return p;
}

/** 是否运行在 Electron 打包环境 */
function isPackaged() {
  return !!process.versions.electron && !!process.resourcesPath && __dirname.includes('app.asar');
}

module.exports = { ROOT, scriptPath, dataDir, isPackaged };
