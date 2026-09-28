/**
 * 网页版标签打印工具站 - 后端服务
 * 功能：
 *   1) 静态站点托管
 *   2) 标签模板的本地文件持久化（保存 / 列表 / 读取 / 删除），存于 data/labels/*.json
 *   3) 本机打印机识别与状态监测（GET /api/printers）
 * 部署：0.0.0.0:11235（端口可用 PORT 环境变量或 --port=xxxx 覆盖）
 */
const express = require('express');
const path = require('path');
const fs = require('fs');
const { collectPrinters } = require('./lib/printers');
const { silentPrint, validatePrintRequest, cleanTmp } = require('./lib/silent-print');
const { collectDiag, buildAddPrinterHint } = require('./lib/printer-diag');
const { dataDir } = require('./lib/paths');
const updater = require('./lib/updater');

/**
 * 端口与监听地址可通过环境变量或启动参数覆盖（默认 11235）：
 *   node server.js
 *   PORT=11235 node server.js
 *   node server.js --port=9000 --host=127.0.0.1
 */
function readArg(name) {
  const hit = process.argv.find(a => a.startsWith(`--${name}=`));
  return hit ? hit.split('=').slice(1).join('=') : '';
}
const PORT = Number(readArg('port') || process.env.PORT || 11235);
const HOST = readArg('host') || process.env.HOST || '0.0.0.0';
/* 程序版本：读打包内/项目内的 package.json 的 version（安装后即为安装包版本号） */
const APP_VERSION = (() => {
  try { return require('./package.json').version || '0.0.0'; }
  catch (e) { return process.env.npm_package_version || '0.0.0'; }
})();
const startedAt = Date.now();
/* 数据目录：默认 <项目>/data/labels；Electron 打包版通过 LABEL_DATA_DIR 指向用户数据目录 */
let DATA_DIR = dataDir('labels');
let ACTIVE_PORT = PORT;   // 实际监听端口（桌面版可能是系统分配的空闲端口）

// 确保标签存储目录存在
fs.mkdirSync(DATA_DIR, { recursive: true });

const app = express();
app.use(express.json({ limit: '20mb' }));

/**
 * 跨设备打印代理所需的 CORS：浏览器无法直接读取本机打印机，
 * 因此允许局域网内其它页面把打印任务/打印机查询发到本服务
 * （例如 A 机器打开网页 → 打印到 B 机器上运行的打印服务）
 */
app.use('/api', (req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,DELETE,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Max-Age', '86400');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

/* 静态托管：若存在「有效的热更新覆盖目录」（<数据目录>/web）则优先生效，再回落到安装包内 public/。
   注意两点：
   1) 覆盖层版本低于安装包版本时视为「过期」——启动时直接清理，否则旧网页层会一直遮住安装包里更新的文件，
      界面上就会出现「网页层版本比 exe 还低、无法更新」的脏状态；
   2) 覆盖目录必须「常挂载 + 运行时门控」：热更新是程序运行中应用的，若只在启动时按状态决定是否挂载，
      则刚下载的更新要重启程序才生效。 */
const stalePurge = updater.purgeStale(APP_VERSION);
if (stalePurge.purged) {
  console.log(`[更新] 已清理过期热更新：V${stalePurge.version}（${stalePurge.files} 个文件，${stalePurge.mode}）`);
}
const overrideStatic = express.static(updater.overrideDir());
let overrideActive = updater.overrideState(APP_VERSION).state === 'active';
/** 应用更新/回滚后重新判定覆盖层是否生效（避免每个请求都去遍历目录） */
const refreshOverride = () => { overrideActive = updater.overrideState(APP_VERSION).state === 'active'; };
app.use((req, res, next) => (overrideActive ? overrideStatic(req, res, next) : next()));
app.use(express.static(path.join(__dirname, 'public')));

/* ============================================================
 * 标签模板持久化
 * ============================================================ */

/** 校验并清洗文件名，防止路径穿越 */
function safeName(name) {
  const n = String(name || '').trim().replace(/\.json$/i, '');
  if (!n || /[\\/:*?"<>|]/.test(n) || n.length > 100) return null;
  return n;
}

/** 列出所有已保存的标签模板 */
app.get('/api/labels', (req, res) => {
  try {
    const files = fs.readdirSync(DATA_DIR).filter(f => f.endsWith('.json'));
    const list = files.map(f => {
      try {
        const full = path.join(DATA_DIR, f);
        const stat = fs.statSync(full);
        const json = JSON.parse(fs.readFileSync(full, 'utf-8'));
        return {
          name: f.replace(/\.json$/i, ''),
          updatedAt: stat.mtime,
          width: json.label && json.label.width,
          height: json.label && json.label.height,
          elementCount: (json.elements || []).length,
          printer: json.printer || '',
        };
      } catch (e) {
        return { name: f.replace(/\.json$/i, ''), updatedAt: null, error: '文件损坏' };
      }
    });
    list.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
    res.json({ ok: true, list });
  } catch (e) {
    res.status(500).json({ ok: false, message: e.message });
  }
});

/** 读取指定模板 */
app.get('/api/labels/:name', (req, res) => {
  const name = safeName(req.params.name);
  if (!name) return res.status(400).json({ ok: false, message: '非法的模板名' });
  const file = path.join(DATA_DIR, name + '.json');
  if (!fs.existsSync(file)) return res.status(404).json({ ok: false, message: '模板不存在' });
  try {
    res.json({ ok: true, data: JSON.parse(fs.readFileSync(file, 'utf-8')) });
  } catch (e) {
    res.status(500).json({ ok: false, message: '模板文件损坏: ' + e.message });
  }
});

/** 保存模板（body: { name, label, elements, printer? }） */
app.post('/api/labels', (req, res) => {
  const name = safeName(req.body.name);
  if (!name) return res.status(400).json({ ok: false, message: '模板名为空或含非法字符' });
  const { label, elements, printer } = req.body;
  if (!label || !Array.isArray(elements)) {
    return res.status(400).json({ ok: false, message: '缺少 label 或 elements 数据' });
  }
  try {
    const payload = {
      name, label, elements,
      printer: typeof printer === 'string' ? printer : '',
      savedAt: new Date().toISOString(),
    };
    fs.writeFileSync(path.join(DATA_DIR, name + '.json'), JSON.stringify(payload, null, 2), 'utf-8');
    res.json({ ok: true, name });
  } catch (e) {
    res.status(500).json({ ok: false, message: e.message });
  }
});

/** 删除模板 */
app.delete('/api/labels/:name', (req, res) => {
  const name = safeName(req.params.name);
  if (!name) return res.status(400).json({ ok: false, message: '非法的模板名' });
  const file = path.join(DATA_DIR, name + '.json');
  try {
    if (fs.existsSync(file)) fs.unlinkSync(file);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ ok: false, message: e.message });
  }
});

/* ============================================================
 * 打印机识别与状态监测
 * 采集与状态归一化实现见 lib/printers.js（Windows 走 PowerShell 批量查询，
 * Linux/macOS 走 lpstat 兜底），此处仅做 HTTP 暴露与 5 秒缓存。
 * ============================================================ */
let printerCache = { at: 0, data: null };

app.get('/api/printers', async (req, res) => {
  const force = req.query.refresh === '1';
  if (!force && printerCache.data && Date.now() - printerCache.at < 5000) return res.json(printerCache.data);
  try {
    const payload = await collectPrinters();
    printerCache = { at: Date.now(), data: payload };
    res.json(payload);
  } catch (e) {
    res.status(500).json({
      ok: false, platform: process.platform, message: e.message,
      printers: [], summary: { total: 0, idle: 0, busy: 0, offline: 0, error: 0, jobs: 0 },
    });
  }
});

/* ============================================================
 * 静默打印（不经浏览器，直接投递到指定打印机队列）
 * 实现见 lib/silent-print.js + scripts/print-images.ps1
 * ============================================================ */
let printing = false;   // 打印队列串行化，避免多个任务并发争抢打印机

app.post('/api/print', async (req, res) => {
  const v = validatePrintRequest(req.body);
  if (v.error) return res.status(400).json({ ok: false, message: v.error });
  if (printing) return res.status(429).json({ ok: false, message: '已有打印任务正在执行，请稍候重试' });

  printing = true;
  const t0 = Date.now();
  try {
    cleanTmp();
    const r = await silentPrint({ printer: v.printer, widthMm: v.widthMm, heightMm: v.heightMm, pages: v.pages });
    res.json({
      ok: true, printer: r.printer, pages: r.pages, failed: r.failed || 0,
      elapsedMs: Date.now() - t0,
      message: `已提交 ${r.pages} 份标签到「${r.printer}」`,
    });
  } catch (e) {
    res.status(500).json({ ok: false, message: '打印失败：' + e.message });
  } finally {
    printing = false;
  }
});

/* ============================================================
 * 打印机诊断：解释「某台打印机为什么没出现在列表里」
 * ============================================================ */
let diagCache = { at: 0, data: null };

app.get('/api/printers/diag', async (req, res) => {
  const force = req.query.refresh === '1';
  if (!force && diagCache.data && Date.now() - diagCache.at < 30000) return res.json(diagCache.data);
  try {
    const d = await collectDiag();
    const unused = d.unusedTcpPorts || [];
    const orphans = d.orphanDevices || [];
    const payload = {
      ok: true,
      updatedAt: new Date().toISOString(),
      queues: d.queues || [],
      drivers: d.drivers || [],
      zebraDrivers: d.zebraDrivers || [],
      unusedTcpPorts: unused,
      orphanDevices: orphans,
      unusedDrivers: d.unusedDrivers || [],
      // 面向用户的诊断结论
      findings: [
        ...((d.unusedDrivers || []).some(x => x.zebra) ? [{
          level: 'info',
          title: '已安装斑马驱动但没有对应的打印机',
          detail: (d.unusedDrivers || []).filter(x => x.zebra).map(x => x.name).join('、')
            + ' —— 说明驱动已就绪，只差把打印机（USB/网络）添加进系统，命令见下。',
        }] : []),
        ...(unused.length ? [{
          level: 'warn',
          title: `发现 ${unused.length} 个已创建但没有任何打印机使用的 TCP/IP 端口`,
          detail: unused.map(p => `${p.name}${p.host ? '（' + p.host + '）' : ''}`).join('、')
            + ' —— 通常表示打印机队列被删除或从未创建；可据此用 scripts\\add-printer.ps1 把设备加进来。',
        }] : []),
        ...(orphans.length ? [{
          level: 'warn',
          title: `发现 ${orphans.length} 个已连接但未安装为打印机的设备`,
          detail: orphans.map(x => x.friendlyName + (x.vid === '0A5F' ? '（Zebra 设备）' : '')).join('、')
            + ' —— 硬件已识别但缺少驱动/队列，安装厂商驱动后即可在本工具中选择。',
        }] : []),
        ...(!(d.zebraDrivers || []).length ? [{
          level: 'info',
          title: '本机未安装任何 Zebra(ZDesigner) 驱动',
          detail: '斑马打印机需先安装 ZDesigner 驱动并添加打印机，Windows 才会把它列出来。',
        }] : [{
          level: 'ok',
          title: `已安装 ${(d.zebraDrivers || []).length} 个 Zebra 驱动`,
          detail: (d.zebraDrivers || []).join('、'),
        }]),
      ],
      addHint: buildAddPrinterHint({ ip: (unused[0] || {}).host || '' }),
    };
    diagCache = { at: Date.now(), data: payload };
    res.json(payload);
  } catch (e) {
    res.status(500).json({ ok: false, message: e.message, findings: [], drivers: [], unusedTcpPorts: [], orphanDevices: [] });
  }
});

/* ============================================================
 * 局域网访问信息（用于其它主机访问本工具站）
 * ============================================================ */
/** 枚举本机所有可用的局域网 IPv4 地址（过滤 169.254/虚拟网卡，物理网卡优先） */
function lanAddresses() {
  const all = [];
  const ifaces = require('os').networkInterfaces();
  const isVirtual = name => /vEthernet|VirtualBox|VMware|Hyper-V|Loopback|Bluetooth|蓝牙|本地连接\s*\*/i.test(name);
  Object.keys(ifaces).forEach(name => {
    (ifaces[name] || []).forEach(net => {
      if (net.family !== 'IPv4' || net.internal) return;
      if (net.address.startsWith('169.254.')) return;   // APIPA 自动专用地址，局域网不可达
      all.push({ iface: name, address: net.address, virtual: isVirtual(name) });
    });
  });
  const physical = all.filter(x => !x.virtual);
  const list = physical.length ? physical : all;      // 没有物理网卡时才退回虚拟网卡
  const score = x => (/WLAN|无线|Wi-?Fi/i.test(x.iface) ? 2 : /以太网|Ethernet/i.test(x.iface) ? 1 : 0);
  list.sort((a, b) => score(b) - score(a));
  return list;
}

/* 版本信息：供界面显示与「安装包是否已更新到最新版」核对 */
app.get('/api/version', (req, res) => {
  const st = updater.status(APP_VERSION);
  res.json({
    ok: true,
    version: APP_VERSION,               // 安装包（主程序）版本
    appVersion: APP_VERSION,
    webVersion: st.webVersion,          // 网页层实际生效版本（热更新后高于安装包版本；不会低于安装包版本）
    hotUpdated: st.hotUpdated,
    staleOverride: st.staleOverride,    // 存在被忽略的过期热更新
    staleVersion: st.staleVersion,
    stalePurged: stalePurge.purged,     // 启动时是否清理了过期热更新
    stalePurgeInfo: stalePurge,
    productName: '标签打印工具站',
    packaged: !!process.versions.electron,
    electron: process.versions.electron || '',
    node: process.versions.node,
    dataDir: dataDir(),
    startedAt: new Date(startedAt).toISOString(),
  });
});

/* ============================================================
 * 网页界面层增量更新（详见 lib/updater.js）
 * ============================================================ */
/** 本机状态：安装包版本 / 网页层版本 / 是否已热更新 */
app.get('/api/update/status', (req, res) => {
  res.json({ ok: true, ...updater.status(APP_VERSION), node: process.versions.node });
});

/** 本机网页层清单（供客户端比对，也便于人工核查） */
app.get('/api/update/manifest', (req, res) => {
  const m = updater.manifest(APP_VERSION);
  res.json({ ok: true, ...m, fileCount: Object.keys(m.files).length });
});

/** 差异计算：客户端上报自己的清单 → 只返回有变化的文件（增量包） */
app.post('/api/update/diff', (req, res) => {
  try {
    const clientFiles = (req.body && req.body.files) || {};
    const version = APP_VERSION;
    const d = updater.diff(version, clientFiles);
    res.json({
      ok: true,
      from: String((req.body && req.body.version) || ''),
      to: d.version,
      total: d.total,
      changedCount: d.changed.length,
      changed: d.changed,
      removed: d.removed,
    });
  } catch (e) {
    res.status(400).json({ ok: false, message: e.message });
  }
});

/** 应用更新：把（已校验的）文件写入本机覆盖目录 */
app.post('/api/update/apply', (req, res) => {
  try {
    const body = req.body || {};
    const ver = String(body.version || APP_VERSION);
    // 拒绝「降级」写入：热更新只用于补丁超前，旧版本网页层会让界面倒退（版本显示变低、功能缺失）
    if (updater.cmpVersion(ver, APP_VERSION) < 0) {
      return res.status(400).json({
        ok: false,
        downgrade: true,
        message: `更新包版本 V${ver} 低于本机安装包 V${APP_VERSION}，已拒绝写入（如需降级请安装对应版本的完整安装包）`,
      });
    }
    const r = updater.applyUpdate(ver, body.files || []);
    refreshOverride();                       // 立即生效，无需重启程序
    res.json({ ok: true, ...r, needReload: true });
  } catch (e) {
    res.status(400).json({ ok: false, message: e.message });
  }
});

/** 回滚：清除覆盖目录，回到安装包自带版本 */
app.post('/api/update/rollback', (req, res) => {
  try {
    const r = updater.rollback();
    refreshOverride();
    res.json({ ok: true, ...r, needReload: true });
  } catch (e) { res.status(400).json({ ok: false, message: e.message }); }
});

/* ============================================================
 * 安装包下载（网页右上角「⬇ 下载安装包」按钮）
 * ------------------------------------------------------------
 * 用途：把最新安装包放到服务器上，局域网内任何电脑打开网页即可一键下载安装
 *      （也用于「检查更新」提示"需要完整安装包"时，直接从更新源下载）。
 * 搜索目录（取版本号最高者；同名多份时取修改时间最新的）：
 *   1) 环境变量 LABEL_INSTALLER_DIR 指定的目录
 *   2) <数据目录>/installer      ← 推荐：运维/桌面版用户把安装包放这里
 *   3) 程序目录及其各级上级目录下的 dist / installer 目录，以及这些目录本身
 *      （覆盖「源码版 <项目>/dist」「免安装版 …\dist\win-unpacked\…」等布局）
 * 只识别 LabelPrint-Setup-<版本>.exe，不接受任何来自请求的文件名（无路径穿越风险）。
 * ============================================================ */
const INSTALLER_RE = /^LabelPrint-Setup-([0-9A-Za-z.\-_]+)\.exe$/;
function installerDirs() {
  const list = [];
  if (process.env.LABEL_INSTALLER_DIR) list.push(path.resolve(process.env.LABEL_INSTALLER_DIR));
  list.push(dataDir('installer'));
  let p = __dirname;
  for (let i = 0; i < 4; i++) {                              // 逐级向上找常见放置位置
    list.push(p, path.join(p, 'dist'), path.join(p, 'installer'));
    const up = path.dirname(p);
    if (up === p) break;
    p = up;
  }
  return list.filter((d, i) => list.indexOf(d) === i);       // 去重
}
/** 实际存在的搜索目录（用于"没找到安装包"时的提示，最多列 6 个） */
function installerSearchDirs() {
  return installerDirs().filter(d => {
    try { return fs.statSync(d).isDirectory(); } catch (e) { return false; }
  }).slice(0, 6);
}
/** 扫描本机所有可用安装包，按版本号倒序（同版本取最新修改的） */
function findInstallers() {
  const found = [];
  installerDirs().forEach(dir => {
    let names = [];
    try { names = fs.readdirSync(dir); } catch (e) { return; }
    names.forEach(name => {
      const m = INSTALLER_RE.exec(name);
      if (!m) return;
      const p = path.join(dir, name);
      let st = null;
      try { st = fs.statSync(p); } catch (e) { return; }
      if (!st.isFile() || st.size < 100 * 1024) return;      // 忽略残留的空文件/半截文件
      found.push({ name, path: p, dir, version: m[1], size: st.size, mtime: st.mtimeMs });
    });
  });
  found.sort((a, b) => updater.cmpVersion(b.version, a.version) || (b.mtime - a.mtime));
  return found;
}
/** 安装包信息（供界面显示按钮状态与版本号） */
app.get('/api/installer', (req, res) => {
  const list = findInstallers();
  const best = list[0] || null;
  const dirs = installerSearchDirs();
  res.json({
    ok: true,
    available: !!best,
    appVersion: APP_VERSION,               // 当前运行版本（用于提示"服务器上是不是更新的版本"）
    version: best ? best.version : '',
    name: best ? best.name : '',
    size: best ? best.size : 0,
    sizeText: best ? (best.size / 1024 / 1024).toFixed(1) + ' MB' : '',
    mtime: best ? new Date(best.mtime).toISOString() : '',
    count: list.length,
    others: list.slice(1, 6).map(x => x.name),
    downloadUrl: best ? '/api/installer/download' : '',
    searched: dirs,
    hint: `把 LabelPrint-Setup-*.exe 放到「${dirs[1]}」或程序目录 dist\\ 下，刷新页面即可下载`,
  });
});
/** 下载最新安装包（附件方式，浏览器另存为） */
app.get('/api/installer/download', (req, res) => {
  const best = findInstallers()[0];
  if (!best) return res.status(404).json({ ok: false, message: '服务器上未找到安装包文件（LabelPrint-Setup-*.exe）' });
  res.setHeader('X-Installer-Version', best.version);
  res.setHeader('X-Installer-Size', String(best.size));
  res.download(best.path, best.name, err => {
    if (err && !res.headersSent) res.status(500).end();
  });
});

app.get('/api/network', (req, res) => {
  const list = lanAddresses();
  const port = ACTIVE_PORT;
  res.json({
    ok: true,
    hostname: require('os').hostname(),
    port,
    localUrl: `http://localhost:${port}`,
    urls: list.map(x => `http://${x.address}:${port}`),
    interfaces: list,
    firewallHint: '若局域网其他主机无法访问，请以管理员身份运行 scripts/fix-lan-firewall.ps1 放行入站端口 '
      + port + '（Windows 防火墙默认拦截入站；且 node.exe 可能存在阻止规则，阻止优先于允许）。',
  });
});

/* ============================================================ */
/**
 * 启动服务（可被网页版直接调用，也可被 Electron 主进程内嵌调用）
 * @param {{port?:number, host?:string, dataDir?:string, quiet?:boolean}} opts
 * @returns {Promise<import('http').Server>}
 */
function startServer(opts = {}) {
  if (opts.dataDir) {
    DATA_DIR = path.resolve(opts.dataDir);
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
  // 注意：端口 0 表示“由系统分配空闲端口”，不能被 || 吞掉
  const port = (opts.port === undefined || opts.port === null) ? PORT : Number(opts.port);
  const host = opts.host || HOST;
  const quiet = !!opts.quiet;

  return new Promise((resolve, reject) => {
    const server = app.listen(port, host, () => {
      if (!quiet) {
        console.log(`标签打印工具站已启动: 监听 ${host}:${port}`);
        console.log(`  本机访问:     http://localhost:${port}`);
        const lan = lanAddresses();
        if (lan.length) {
          console.log('  局域网访问（其他主机请用这些地址）:');
          lan.forEach(x => console.log(`    http://${x.address}:${port}    [${x.iface}]`));
        } else {
          console.log('  未检测到可用的局域网 IPv4 地址');
        }
        console.log(`  提示: 若局域网无法访问，请以管理员身份运行: `
          + `scripts\\fix-lan-firewall.ps1 -UnblockNode -Port ${port}`);
        console.log(`标签模板存储目录: ${DATA_DIR}`);
        if (process.platform === 'win32') console.log('打印机识别: 已启用 (Get-Printer / Win32_Printer)');
      }
      ACTIVE_PORT = server.address() ? server.address().port : port;
      server.__port = ACTIVE_PORT;
      resolve(server);
    });

    server.on('error', e => {
      if (e.code === 'EADDRINUSE') {
        e.friendly = `端口 ${port} 已被占用`;
        if (!quiet) {
          console.error(`[启动失败] 端口 ${port} 已被占用。`
            + `\n  可换端口启动: node server.js --port=其他端口`
            + `\n  或查看占用:   netstat -ano | findstr :${port}`);
        }
      } else if (e.code === 'EACCES') {
        e.friendly = `端口 ${port} 无权限监听（可能被系统保留或被安全软件拦截）`;
        if (!quiet) console.error(`[启动失败] ${e.friendly}，请换端口或使用管理员身份运行。`);
      } else {
        e.friendly = e.message;
        if (!quiet) console.error('[启动失败] ' + e.message);
      }
      reject(e);
    });
  });
}

module.exports = { app, startServer, lanAddresses, DATA_DIR };

/* 直接运行 node server.js 时自动启动（被 Electron 引入时不会自动启动） */
if (require.main === module) {
  startServer().catch(() => process.exit(1));
}
