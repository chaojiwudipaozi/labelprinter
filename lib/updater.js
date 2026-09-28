/**
 * 网页界面层「增量更新」
 *
 * 思路：
 *   1) 网页层 = public/**（界面/样式/前端逻辑）与 docs/**（手册、版本日志），这些文件由内置服务端托管，
 *      因此可以在不重装安装包的前提下替换；
 *   2) 覆盖目录：<数据目录>/web（默认 <安装目录同级数据目录>/web），服务端优先从这里取文件，
 *      缺失的再回落到安装包内的 public/ —— 既实现热更新，又不动安装文件（无需管理员权限、可一键回滚）；
 *   3) 客户端把自己的文件清单（相对路径 → SHA-256）报给「更新服务端」，
 *      服务端只返回 SHA 不同的文件内容（增量），客户端校验后写入覆盖目录；
 *   4) 主程序层（Electron 主进程、内置服务、PowerShell 脚本、依赖库）无法这样替换，
 *      只能通过完整安装包升级 —— 接口会在 needInstaller 字段中明确告知。
 *
 * 【重要：热更新是「补丁超前」，不是「版本回退」】
 *   覆盖层只在「覆盖层版本 ≥ 安装包版本」时生效（active）；
 *   一旦安装了更新的安装包，旧的覆盖层就变成「过期」（stale）——必须忽略并在启动时清理，
 *   否则旧网页层会一直遮住安装包里更新的文件，界面上会出现
 *   「版本停在旧号（如 1.0.4），且网页层版本比 exe 还低、无法更新」的脏状态。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { dataDir, ROOT } = require('./paths');

/* 参与热更新的文件白名单：根级 index.html + 这些子目录下的文件（含中文文件名，如 docs/使用手册.md） */
const ROOT_FILES = ['index.html'];
const ALLOWED_DIRS = ['js', 'css', 'lib', 'docs'];
const SAFE_REL = /^(index\.html|(js|css|lib|docs)\/[\w.\-\u4e00-\u9fa5]+)$/;

const bundledDir = path.join(ROOT, 'public');           // 安装包内自带
const overrideDir = () => dataDir('web');               // 热更新写入处（dataDir 会自动创建）

const sha256 = buf => crypto.createHash('sha256').update(buf).digest('hex');

/** 递归列出目录下的文件（相对路径，POSIX 分隔符） */
function walk(dir, base = '') {
  const out = [];
  let list = [];
  try { list = fs.readdirSync(dir, { withFileTypes: true }); } catch (e) { return out; }
  list.forEach(d => {
    const rel = base ? base + '/' + d.name : d.name;
    if (d.isDirectory()) {
      if (base === '' && !ALLOWED_DIRS.includes(d.name)) return;   // 只遍历白名单子目录
      out.push(...walk(path.join(dir, d.name), rel));
    } else if (SAFE_REL.test(rel)) {
      out.push(rel);
    }
  });
  return out;
}

/** 版本号比较：逐段数值比较（"1.10.0" > "1.9.0"；缺位按 0）。a>b→1，a<b→-1，相等→0 */
function cmpVersion(a, b) {
  const pa = String(a == null ? '' : a).split(/[.\-+_]/).map(x => parseInt(x, 10) || 0);
  const pb = String(b == null ? '' : b).split(/[.\-+_]/).map(x => parseInt(x, 10) || 0);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const d = (pa[i] || 0) - (pb[i] || 0);
    if (d) return d > 0 ? 1 : -1;
  }
  return 0;
}

/**
 * 覆盖层状态：
 *   none   —— 没有覆盖层（或只剩空目录）
 *   active —— 覆盖层版本 ≥ 安装包版本：正常的「热更新超前」，生效
 *   stale  —— 覆盖层版本 < 安装包版本（或没有版本号）：装包已更新，旧覆盖层必须忽略
 */
function overrideState(appVersion) {
  const dir = overrideDir();
  let meta = null;
  try { meta = JSON.parse(fs.readFileSync(path.join(dir, 'update.json'), 'utf-8')); } catch (e) { meta = null; }
  const files = walk(dir);
  const has = files.length > 0 && fs.existsSync(path.join(dir, 'index.html'));
  if (!has) return { state: 'none', dir, meta, files, version: '' };
  const version = meta && meta.version ? String(meta.version) : '';
  const stale = !version || cmpVersion(version, appVersion) < 0;
  return { state: stale ? 'stale' : 'active', dir, meta, files, version };
}

/** 清理过期覆盖层：优先整目录删除；删不掉（占用/权限）则改名到 web.stale-*，确保不再被加载 */
function purgeStale(appVersion) {
  const st = overrideState(appVersion);
  if (st.state !== 'stale') return { purged: false, version: st.version, files: st.files.length };
  try {
    fs.rmSync(st.dir, { recursive: true, force: true });
    return { purged: true, mode: 'removed', version: st.version, files: st.files.length };
  } catch (e) { /* 删除失败 → 尝试改名 */ }
  try {
    const backup = path.join(path.dirname(st.dir), 'web.stale-' + (st.version || 'unknown') + '-' + Date.now());
    fs.renameSync(st.dir, backup);
    return { purged: true, mode: 'renamed', version: st.version, files: st.files.length, backup };
  } catch (e) {
    return { purged: false, version: st.version, files: st.files.length, error: e.message };
  }
}

/** 生效的网页层根目录（覆盖层仅在 active 时参与，且优先于安装包内 public/） */
function webRoots(appVersion) {
  return overrideState(appVersion).state === 'active' ? [overrideDir(), bundledDir] : [bundledDir];
}

/** 生效的网页层版本 = max(安装包版本, 生效的覆盖层版本) */
function webVersion(appVersion) {
  const ov = overrideState(appVersion);
  return ov.state === 'active' ? ov.version : String(appVersion == null ? '' : appVersion);
}

/** 网页层清单：{ version, appVersion, files: { rel: sha } } */
function manifest(appVersion) {
  const files = {};
  webRoots(appVersion).forEach(root => {
    walk(root).forEach(rel => {
      if (files[rel]) return;                            // 覆盖目录优先，已存在则不再覆盖
      try { files[rel] = sha256(fs.readFileSync(path.join(root, rel))); } catch (e) { /* 忽略读失败 */ }
    });
  });
  return { version: webVersion(appVersion), appVersion: String(appVersion == null ? '' : appVersion), files };
}

/** 读取某个相对路径的有效内容（生效的覆盖层优先） */
function readWebFile(rel, appVersion) {
  if (!SAFE_REL.test(rel)) return null;
  for (const root of webRoots(appVersion)) {
    const p = path.join(root, ...rel.split('/'));
    try { if (fs.existsSync(p)) return fs.readFileSync(p); } catch (e) { /* 继续下一个 */ }
  }
  return null;
}

/** 与客户端清单对比，返回需要下发/删除的文件（增量） */
function diff(appVersion, clientFiles = {}) {
  const mine = manifest(appVersion);
  const changed = [];
  Object.keys(mine.files).forEach(rel => {
    if (clientFiles[rel] === mine.files[rel]) return;             // 一致，跳过
    const buf = readWebFile(rel, appVersion);
    if (!buf) return;
    changed.push({ path: rel, sha: mine.files[rel], size: buf.length, data: buf.toString('base64') });
  });
  const removed = Object.keys(clientFiles).filter(rel => SAFE_REL.test(rel) && !mine.files[rel]);
  return { version: mine.version, appVersion: mine.appVersion, total: Object.keys(mine.files).length, changed, removed };
}

/** 覆盖目录当前记录的版本号（没有 update.json 时返回空串） */
function currentOverrideVersion() {
  try {
    const meta = JSON.parse(fs.readFileSync(path.join(overrideDir(), 'update.json'), 'utf-8'));
    return String((meta || {}).version || '');
  } catch (e) { return ''; }
}

/**
 * 校验并写入覆盖目录；返回写入的文件数与版本。
 * 若覆盖目录里已有「另一个版本」的残留文件，先整体清空再写 ——
 * 否则未随本次更新下发的旧文件会继续生效，造成界面新旧混杂。
 */
function applyUpdate(version, files = []) {
  if (!Array.isArray(files) || !files.length) throw new Error('更新包为空');
  const dir = overrideDir();
  const curVer = currentOverrideVersion();
  const ver = String(version || '');
  if (walk(dir).length > 0 && curVer !== ver) {
    try { fs.rmSync(dir, { recursive: true, force: true }); } catch (e) { /* 忽略：下面按文件覆写 */ }
  }
  let written = 0;
  files.forEach(f => {
    const rel = String(f.path || '');
    if (!SAFE_REL.test(rel)) throw new Error('非法文件路径：' + rel);
    const buf = Buffer.from(String(f.data || ''), 'base64');
    const sha = sha256(buf);
    if (f.sha && f.sha !== sha) throw new Error(`文件校验失败（SHA-256 不一致）：${rel}`);
    const dest = path.join(dir, ...rel.split('/'));
    fs.mkdirSync(path.dirname(dest), { recursive: true });
    fs.writeFileSync(dest, buf);
    written++;
  });
  const meta = {
    version: ver, files: files.map(f => f.path), count: files.length,
    appliedAt: new Date().toISOString(),
  };
  fs.writeFileSync(path.join(dir, 'update.json'), JSON.stringify(meta, null, 2), 'utf-8');
  return { written, version: meta.version };
}

/** 回滚：清空覆盖目录，回到安装包自带版本 */
function rollback() {
  const dir = overrideDir();
  let removed = 0;
  ['js', 'css', 'lib', 'docs', 'index.html', 'update.json'].forEach(name => {
    const p = path.join(dir, name);
    try {
      if (!fs.existsSync(p)) return;
      const st = fs.statSync(p);
      if (st.isDirectory()) { fs.rmSync(p, { recursive: true, force: true }); removed++; }
      else { fs.rmSync(p, { force: true }); removed++; }
    } catch (e) { /* 忽略 */ }
  });
  return { removed };
}

/**
 * 当前热更新状态：
 *   appVersion     安装包（主程序）版本
 *   webVersion     网页层实际生效版本（= max(安装包版本, 生效的覆盖层版本)）
 *   hotUpdated     是否已应用「超前」的热更新
 *   staleOverride  是否存在被忽略的过期热更新（版本低于安装包）
 */
function status(appVersion) {
  const ov = overrideState(appVersion);
  const active = ov.state === 'active';
  const stale = ov.state === 'stale';
  return {
    appVersion: String(appVersion == null ? '' : appVersion),
    webVersion: active ? ov.version : String(appVersion == null ? '' : appVersion),
    hotUpdated: active,                                  // 是否已应用过热更新（且仍然有效）
    overrideDir: ov.dir,
    overrideFiles: active ? ov.files.length : 0,
    appliedAt: active && ov.meta && ov.meta.appliedAt ? ov.meta.appliedAt : '',
    appliedVersion: active ? ov.version : '',
    staleOverride: stale,                                // 覆盖层过期（已被忽略）
    staleVersion: stale ? ov.version : '',
    staleFiles: stale ? ov.files.length : 0,
  };
}

module.exports = {
  manifest, diff, applyUpdate, rollback, status,
  readWebFile, overrideDir, bundledDir, SAFE_REL,
  cmpVersion, overrideState, purgeStale, webVersion,
};
