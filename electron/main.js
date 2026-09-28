/**
 * 标签打印工具站 · 桌面版（Electron）
 * 设计要点：
 *  1. 主进程内嵌启动本项目自带的 Express 服务（127.0.0.1），窗口加载该地址，
 *     因此「安装到哪台主机，就识别哪台主机的打印机」——打印机识别/静默打印都跑在本机。
 *  2. 端口默认 11235，被占用时自动改用随机空闲端口。
 *  3. 数据（标签模板）写入用户数据目录，卸载重装不丢；打包后 asar 内只读。
 *  4. 支持 --selftest 自检模式：不开窗口，启动服务→查询打印机→打印结果并退出（用于自动化验证）。
 */
const { app, BrowserWindow, Menu, dialog, shell } = require('electron');
const path = require('path');
const net = require('net');
const updater = require('../lib/updater');   // 仅用于自检时比较「网页层版本 vs 安装包版本」

const DEFAULT_PORT = 11235;
let mainWindow = null;
let httpServer = null;
let actualPort = DEFAULT_PORT;
let localUrls = { local: '', lan: [] };

/* ---------- 单实例：第二次启动时聚焦已有窗口 ---------- */
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) { app.quit(); }

/** 启动内嵌服务：优先用 11235，被占用则自动改用系统分配的空闲端口 */
async function bootServer() {
  // 打包后把数据目录放到用户数据目录（asar 内不可写）
  process.env.LABEL_DATA_DIR = path.join(app.getPath('userData'), 'data');
  const { startServer, lanAddresses } = require(path.join(__dirname, '..', 'server.js'));

  const tryStart = async port => startServer({ port, host: '0.0.0.0', quiet: true });
  try {
    httpServer = await tryStart(DEFAULT_PORT);
  } catch (e) {
    if (e.code !== 'EADDRINUSE') throw e;
    console.log(`[桌面版] 端口 ${DEFAULT_PORT} 已被占用，改用系统分配的空闲端口`);
    httpServer = await tryStart(0);           // 0 = 让系统分配
  }
  actualPort = httpServer.address().port;

  const lan = lanAddresses();
  localUrls = {
    local: `http://localhost:${actualPort}`,
    lan: lan.map(x => `http://${x.address}:${actualPort}`),
    ifaces: lan,
  };
  console.log(`[桌面版] 服务已启动：${localUrls.local}`);
  localUrls.lan.forEach(u => console.log(`[桌面版] 局域网访问：${u}`));
  console.log(`[桌面版] 数据目录：${process.env.LABEL_DATA_DIR}`);
  return actualPort;
}

/** 主窗口 */
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 940,
    minWidth: 1100,
    minHeight: 700,
    title: '标签打印工具站',
    backgroundColor: '#f0f2f5',
    autoHideMenuBar: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      spellcheck: false,
    },
  });

  mainWindow.loadURL(`http://127.0.0.1:${actualPort}/`);

  // 打印预览等 window.open 允许弹出（用系统默认浏览器打开外部链接）
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:\/\/(127\.0\.0\.1|localhost)/.test(url)) return { action: 'allow' };
    shell.openExternal(url);
    return { action: 'deny' };
  });

  mainWindow.on('closed', () => { mainWindow = null; });
}

/** 菜单（中文） */
function buildMenu() {
  const template = [
    {
      label: '文件',
      submenu: [
        {
          label: '打开数据目录（标签模板）',
          click: () => shell.openPath(path.join(app.getPath('userData'), 'data', 'labels')),
        },
        {
          label: '显示访问地址（局域网共享）',
          click: () => {
            dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: '访问地址',
              message: `本机：${localUrls.local}\n`
                + (localUrls.lan.length ? '局域网（其他主机可用）：\n' + localUrls.lan.join('\n') : '未检测到局域网地址')
                + '\n\n提示：其他主机访问前，需在本机放行该端口的入站连接（Windows 防火墙）。',
              buttons: ['确定'],
            });
          },
        },
        { type: 'separator' },
        { label: '退出', role: 'quit' },
      ],
    },
    {
      label: '视图',
      submenu: [
        { label: '重新加载', role: 'reload' },
        { label: '强制重新加载', role: 'forceReload' },
        { label: '放大', role: 'zoomIn' },
        { label: '缩小', role: 'zoomOut' },
        { label: '实际大小', role: 'resetZoom' },
        { type: 'separator' },
        { label: '全屏', role: 'togglefullscreen' },
        { label: '开发者工具', role: 'toggleDevTools' },
      ],
    },
    {
      label: '帮助',
      submenu: [
        {
          label: '使用手册（F1）',
          accelerator: 'F1',
          click: () => {
            if (mainWindow) mainWindow.webContents.executeJavaScript('window.openManual && window.openManual(false)');
          },
        },
        {
          label: '版本日志（更新记录）',
          click: () => {
            if (mainWindow) mainWindow.webContents.executeJavaScript('window.openChangelog && window.openChangelog()');
          },
        },
        {
          label: '检查更新（增量更新）',
          click: () => {
            if (mainWindow) mainWindow.webContents.executeJavaScript(
              `(() => { const t = document.querySelector('.rtab[data-tab="update"]'); if (t) t.click();
                        const b = document.getElementById('btn-check-update'); if (b) b.click(); })()`);
          },
        },
        { type: 'separator' },
        {
          label: '关于本工具',
          click: () => dialog.showMessageBox(mainWindow, {
            type: 'info',
            title: '关于',
            message: '标签打印工具站（桌面版）',
            detail: `版本 ${app.getVersion()}\n`
              + `内嵌服务端口：${actualPort}\n`
              + `数据目录：${path.join(app.getPath('userData'), 'data')}\n\n`
              + '打印机识别与静默打印均在本机执行：安装到哪台电脑，就使用哪台电脑的打印机。\n'
              + 'Excel 数据规范与使用说明见安装目录下的「使用说明.md」。',
            buttons: ['确定'],
          }),
        },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

/* ---------- 自检模式（无窗口，便于自动化验证） ---------- */
async function selfTest() {
  try {
    await bootServer();
    const get = async p => (await fetch(`http://127.0.0.1:${actualPort}${p}`)).json();
    const net_ = await get('/api/network');
    const pr = await get('/api/printers?refresh=1');
    const labels = await get('/api/labels');

    // 关键：验证 PowerShell 脚本在打包环境下的真实路径可执行（静默打印/诊断依赖它）
    const fs = require('fs');
    const { scriptPath } = require(path.join(__dirname, '..', 'lib', 'paths.js'));
    const scripts = ['list-printers.ps1', 'print-images.ps1', 'printer-diag.ps1', 'fix-lan-firewall.ps1', 'add-printer.ps1'];
    const missing = scripts.filter(s => !fs.existsSync(scriptPath(s)));
    let diagOk = false, diagMsg = '';
    try {
      const d = await get('/api/printers/diag?refresh=1');
      diagOk = !!d.ok;
      diagMsg = (d.findings || []).length + ' 条诊断结论';
    } catch (e) { diagMsg = e.message; }

    // 安装包下载（网页右上角「⬇ 下载安装包」依赖它；服务器上没放安装包时 available=false 属正常）
    let installerAvailable = false, installerVersion = '', installerSize = 0;
    try {
      const ins = await get('/api/installer');
      installerAvailable = !!ins.available;
      installerVersion = ins.version || '';
      installerSize = ins.size || 0;
    } catch (e) { /* 忽略 */ }

    // 内置使用手册是否可读（软件内「❓ 使用手册」与菜单 F1 依赖它）
    let manualOk = false, manualLen = 0;
    try {
      const r = await fetch(`http://127.0.0.1:${actualPort}/docs/` + encodeURIComponent('使用手册.md'));
      const t = await r.text();
      manualOk = r.ok && /^#\s/.test(t.trim());
      manualLen = t.length;
    } catch (e) { /* 忽略 */ }

    // 前端是否为最新版：直接向本程序内嵌服务请求前端资源并查找各轮功能标记
    const ver = await get('/api/version');
    let frontOk = false, frontMissing = [];
    try {
      const js = await (await fetch(`http://127.0.0.1:${actualPort}/js/app.js`)).text();
      const html = await (await fetch(`http://127.0.0.1:${actualPort}/`)).text();
      const features = {
        zoom150: /DEFAULT_ZOOM = 1\.5/,
        zoomPersist: /zoomExplicit/,
        anyExcel: /detectHeaderRow/,
        importPanel: /ir-reparse/,
        tableCells: /tblRemapCells/,
        tableFix: /cloneElement/,
        dateISO: /isDateFormat/,
        printMm: /elementContent\(el, row, 'mm'\)/,
        manualBtn: /btn-manual/,
        versionChip: /app-version/,
        dataPanel: /data-toggle/,           // 底部 Excel 数据区可上拉展开/向下收起
        panelCollapse: /btn-data-collapse/,
        changelogTab: /data-tab="update"/,  // 「版本日志」页签
        incrementalUpdate: /api\/update\/diff/,  // 增量更新（清单/差异/应用）
        updateUi: /btn-check-update/,
        cellAutoFit: /qe-autofit/,               // 单元格自适应字号
        versionSplit: /staleVersion/,            // 安装包版本/网页层版本分开显示 + 过期热更新标识
        installerBtn: /btn-installer/,           // 右上角「⬇ 下载安装包」
        multiSheet: /mergeSheets/,               // 多工作表导入（按列名合并）
        undoRedo: /btn-undo/,                    // 上一步 / 下一步
        autoFitAll: /fitFontSizePt/,             // 文本框/条码也支持自适应字号
      };
      frontMissing = Object.keys(features).filter(k => !features[k].test(js) && !features[k].test(html));
      frontOk = frontMissing.length === 0;
    } catch (e) { frontMissing = ['请求失败: ' + e.message]; }

    // 增量更新：网页层清单 + 覆盖目录状态（安装版必须可用）
    let updateOk = false, updateFiles = 0, hotUpdated = false, overrideFiles = 0;
    let staleOverride = false, staleVersion = '', versionOk = true;
    try {
      const u = await get('/api/update/manifest');
      const us = await get('/api/update/status');
      updateOk = !!u.ok && !!us.ok;
      updateFiles = u.fileCount || 0;
      hotUpdated = !!us.hotUpdated;
      overrideFiles = us.overrideFiles || 0;
      staleOverride = !!us.staleOverride;
      staleVersion = us.staleVersion || '';
      // 网页层生效版本不得低于安装包版本（低于即为过期热更新未清理的脏状态）
      versionOk = updater.cmpVersion(us.webVersion, ver.version) >= 0;
      if (!versionOk) console.log(`[自检] 警告：网页层版本 V${us.webVersion} 低于安装包 V${ver.version}（存在过期热更新）`);
    } catch (e) { updateOk = false; }

    console.log('SELFTEST_RESULT ' + JSON.stringify({
      ok: true,
      version: ver.version,
      webVersion: ver.webVersion,
      packaged: __dirname.includes('app.asar'),
      updateOk, updateFiles, hotUpdated, overrideFiles,
      versionOk, staleOverride, staleVersion, stalePurged: !!ver.stalePurged,
      installerAvailable, installerVersion, installerSize,
      port: actualPort,
      hostname: net_.hostname,
      localUrl: net_.localUrl,
      lanUrls: net_.urls,
      frontOk, frontMissing,
      printers: (pr.printers || []).length,
      printerNames: (pr.printers || []).slice(0, 5).map(p => p.name),
      zebra: (pr.printers || []).filter(p => /Zebra/i.test(p.vendor)).length,
      labels: (labels.list || []).length,
      manualOk, manualLen,
      scriptsOk: missing.length === 0,
      scriptsMissing: missing,
      diagOk, diagMsg,
      dataDir: process.env.LABEL_DATA_DIR,
    }));
    httpServer && httpServer.close();
    app.exit(0);
  } catch (e) {
    console.log('SELFTEST_RESULT ' + JSON.stringify({ ok: false, message: e.message }));
    app.exit(1);
  }
}

/* ---------- 窗口冒烟测试（验证界面能真正加载） ---------- */
async function smokeTest() {
  const timer = setTimeout(() => {
    console.log('SMOKETEST_RESULT ' + JSON.stringify({ ok: false, message: '窗口加载超时' }));
    app.exit(1);
  }, 30000);
  try {
    await bootServer();
    buildMenu();
    createWindow();
    mainWindow.webContents.once('did-finish-load', async () => {
      clearTimeout(timer);
      const title = mainWindow.webContents.getTitle();
      const url = mainWindow.webContents.getURL();
      const hasApp = await mainWindow.webContents.executeJavaScript(
        'document.querySelectorAll(".pal[data-type]").length + "|" + (document.querySelector("#printer-select") ? "printer-select-ok" : "no-select")');
      // 界面上的版本号（app-version 由 /api/version 填充），用于核对安装包版本
      let uiVersion = '';
      for (let i = 0; i < 20; i++) {
        uiVersion = await mainWindow.webContents.executeJavaScript('(document.querySelector("#app-version")||{}).textContent || ""');
        if (uiVersion && uiVersion !== 'V—') break;
        await new Promise(r => setTimeout(r, 200));
      }
      console.log('SMOKETEST_RESULT ' + JSON.stringify({
        ok: true, title, url, dom: hasApp, version: app.getVersion(), uiVersion,
      }));
      app.exit(0);
    });
    mainWindow.webContents.once('did-fail-load', (e, code, desc) => {
      clearTimeout(timer);
      console.log('SMOKETEST_RESULT ' + JSON.stringify({ ok: false, code, desc }));
      app.exit(1);
    });
  } catch (e) {
    clearTimeout(timer);
    console.log('SMOKETEST_RESULT ' + JSON.stringify({ ok: false, message: e.message }));
    app.exit(1);
  }
}

/* ---------- 生命周期 ---------- */
app.on('second-instance', () => {
  if (mainWindow) {
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.focus();
  }
});

app.whenReady().then(async () => {
  if (process.argv.includes('--selftest')) return selfTest();
  if (process.argv.includes('--smoketest')) return smokeTest();
  try {
    await bootServer();
  } catch (e) {
    dialog.showErrorBox('启动失败', `${e.friendly || e.message}\n\n请检查端口占用或安全软件拦截。`);
    app.exit(1);
    return;
  }
  buildMenu();
  createWindow();
  app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
});

app.on('window-all-closed', () => {
  if (httpServer) try { httpServer.close(); } catch (e) { }
  app.quit();
});
