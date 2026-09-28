/**
 * 本地自检脚本（CI 与开发自测共用，不依赖 Electron）
 *   1) 语法检查：server.js / electron/main.js / lib/*.js / public/js/app.js / scripts/*.js
 *   2) 启动内置服务（使用临时数据目录，避免污染本机模板）
 *   3) 探活关键接口：/api/version、/api/labels、/api/printers、/api/update/manifest
 *   4) 校验内置文档（public/docs）已同步
 *
 * 用法：node scripts/verify.js   （或 npm run verify）
 * 退出码：0 全部通过；1 有失败项（便于 CI 直接使用）
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn, execFileSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = Number(process.env.VERIFY_PORT || 12399);
const ok = [], bad = [];
const chk = (cond, msg) => (cond ? ok : bad).push(msg);

/* ---------- 1. 语法检查 ---------- */
function listJs(dir, out = []) {
  fs.readdirSync(dir, { withFileTypes: true }).forEach(e => {
    if (e.name === 'node_modules' || e.name === 'dist' || e.name === 'data' || e.name.startsWith('.')) return;
    const p = path.join(dir, e.name);
    if (e.isDirectory()) listJs(p, out);
    else if (e.name.endsWith('.js') && !e.name.endsWith('.min.js')) out.push(p);
  });
  return out;
}
const jsFiles = listJs(ROOT).filter(p => !p.includes(path.join('public', 'lib')));
jsFiles.forEach(p => {
  try {
    execFileSync(process.execPath, ['--check', p], { stdio: 'pipe' });
    chk(true, '语法检查通过：' + path.relative(ROOT, p));
  } catch (e) {
    chk(false, '语法检查失败：' + path.relative(ROOT, p) + ' → ' + String(e.stderr || e.message).split('\n')[0]);
  }
});

/* ---------- 2. 内置文档是否已同步 ---------- */
[['docs/使用手册.md', 'public/docs/使用手册.md'],
  ['docs/使用说明.md', 'public/docs/使用说明.md'],
  ['CHANGELOG.md', 'public/docs/CHANGELOG.md']].forEach(([src, out]) => {
  const a = path.join(ROOT, src), b = path.join(ROOT, out);
  const exists = fs.existsSync(a) && fs.existsSync(b);
  chk(exists && fs.readFileSync(a, 'utf8') === fs.readFileSync(b, 'utf8'),
    exists ? `${out} 与 ${src} 一致` : `缺少文档：${!fs.existsSync(a) ? src : out}（请执行 npm run sync-manual）`);
});

/* ---------- 3. 启动服务并探活接口 ---------- */
(async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'label-print-verify-'));
  const child = spawn(process.execPath, [path.join(ROOT, 'server.js'), '--port', String(PORT)], {
    cwd: ROOT,
    env: { ...process.env, LABEL_DATA_DIR: dataDir, PORT: String(PORT) },
    windowsHide: true,
  });
  let out = '';
  child.stdout.on('data', d => { out += d; });
  child.stderr.on('data', d => { out += d; });

  const base = `http://127.0.0.1:${PORT}`;
  const get = async p => {
    const r = await fetch(base + p);
    if (!r.ok) throw new Error(`${p} → HTTP ${r.status}`);
    return r.json();
  };
  let ready = false;
  for (let i = 0; i < 40 && !ready; i++) {
    try { await get('/api/version'); ready = true; } catch (e) { await new Promise(r => setTimeout(r, 250)); }
  }
  chk(ready, ready ? '服务启动成功：' + base : '服务启动失败（输出片段：' + out.slice(-200) + '）');

  if (ready) {
    try {
      const v = await get('/api/version');
      const pkg = require(path.join(ROOT, 'package.json'));
      chk(v.ok && v.version === pkg.version, `/api/version 版本一致（V${v.version}）`);
      const m = await get('/api/update/manifest');
      chk(m.ok && m.fileCount > 0 && !!m.files['index.html'] && !!m.files['js/app.js'],
        `/api/update/manifest 清单完整（${m.fileCount} 个文件）`);
      const l = await get('/api/labels');
      chk(l.ok && Array.isArray(l.list), '/api/labels 可读取模板列表');
      const pr = await get('/api/printers');
      chk(pr.ok && Array.isArray(pr.printers), `/api/printers 可读取打印机（${(pr.printers || []).length} 台）`);
      const home = await (await fetch(base + '/')).text();
      chk(/<title>/.test(home) && /js\/app\.js/.test(home), '首页可访问且引用了前端脚本');
      const manual = await (await fetch(base + '/docs/' + encodeURIComponent('使用手册.md'))).text();
      chk(/^#\s/.test(manual.trim()), '应用内使用手册可读取');
    } catch (e) {
      chk(false, '接口探活异常：' + e.message);
    }
  }

  child.kill();
  try { fs.rmSync(dataDir, { recursive: true, force: true }); } catch (e) { /* 忽略 */ }

  console.log('===== 通过 (' + ok.length + ') =====');
  ok.forEach(s => console.log('  √ ' + s));
  if (bad.length) {
    console.log('===== 失败 (' + bad.length + ') =====');
    bad.forEach(s => console.log('  ✗ ' + s));
  }
  console.log(bad.length ? '\n自检未通过' : '\n自检全部通过 ✓');
  process.exit(bad.length ? 1 : 0);
})();
