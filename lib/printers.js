/**
 * 本机打印机识别与状态归一化
 * - Windows：调用 scripts/list-printers.ps1（Get-Printer / Win32_Printer / Win32_PrintJob 批量查询）
 * - Linux / macOS：解析 lpstat -p -d
 */
const path = require('path');
const os = require('os');
const { execFile } = require('child_process');

const { scriptPath } = require('./paths');
const PS_SCRIPT = scriptPath('list-printers.ps1');

/** Win32_Printer.DetectedErrorState 映射 */
const ERROR_TEXT = {
  0: '', 1: '其他错误', 2: '', 3: '纸张不足', 4: '缺纸', 5: '碳粉不足',
  6: '碳粉耗尽', 7: '仓门打开', 8: '卡纸', 9: '脱机', 10: '需维修', 11: '出纸盒已满', 12: '缺墨',
  13: '墨盒不足', 14: '需要更换硒鼓', 15: '需要更换墨盒', 16: '需要用户干预',
  17: '内存不足', 18: '输出纸张不可用', 19: '服务器未知', 20: '驱动不支持',
};

/** Get-Printer.PrinterStatus 枚举字符串 → [状态码, 中文] */
const STATUS_TEXT = {
  normal: ['idle', '空闲'], ready: ['idle', '空闲'], idle: ['idle', '空闲'],
  printing: ['printing', '打印中'], warmup: ['printing', '预热中'], warmingup: ['printing', '预热中'],
  processing: ['printing', '处理中'], initializing: ['printing', '正在初始化'],
  paused: ['paused', '已暂停'], stopped: ['paused', '已停止'], stoppedprinting: ['paused', '已停止'],
  offline: ['offline', '脱机'],
  error: ['error', '打印机错误'], userintervention: ['error', '需用户干预'],
  paperout: ['error', '缺纸'], paperjam: ['error', '卡纸'], paperproblem: ['error', '纸张问题'],
  tonerlow: ['error', '碳粉不足'], notoner: ['error', '碳粉耗尽'], noretoner: ['error', '碳粉耗尽'],
  dooropen: ['error', '仓门打开'], outputbinfull: ['error', '出纸盒已满'],
  serverunknown: ['error', '服务未知'], notavailable: ['error', '不可用'],
  powersave: ['idle', '节能待机'],
};

/** PrinterStatus 数字枚举（Win32_Printer）：3 空闲 / 4 打印中 / 5 预热 / 6 停止 / 7 脱机 */
const STATUS_NUM = { 3: ['idle', '空闲'], 4: ['printing', '打印中'], 5: ['printing', '预热中'], 6: ['paused', '已停止'], 7: ['offline', '脱机'] };

/** 依据名称/驱动识别品牌（惠普、斑马等） */
function detectVendor(p) {
  const s = (String(p.name || '') + ' ' + String(p.driver || '')).toLowerCase();
  // 注意匹配顺序：先判定各品牌专有名称/机型，最后才用 'zpl' 兜底
  //（TSC/Godex 等也支持 ZPL，若先匹配 zpl 会被误判为斑马）
  const brands = [
    ['zebra', 'Zebra 斑马'], ['zdesigner', 'Zebra 斑马'], ['zebra technologies', 'Zebra 斑马'],
    // Zebra 常见系列/机型（名称里可能不含 Zebra 字样）
    ['gx430', 'Zebra 斑马'], ['gx420', 'Zebra 斑马'], ['gk420', 'Zebra 斑马'], ['gk430', 'Zebra 斑马'],
    ['zd4', 'Zebra 斑马'], ['zd5', 'Zebra 斑马'], ['zd6', 'Zebra 斑马'],
    ['zt2', 'Zebra 斑马'], ['zt4', 'Zebra 斑马'], ['zt6', 'Zebra 斑马'],
    ['zq3', 'Zebra 斑马'], ['zq5', 'Zebra 斑马'], ['zq6', 'Zebra 斑马'],
    ['ql2', 'Zebra 斑马'], ['ql3', 'Zebra 斑马'], ['ql4', 'Zebra 斑马'], ['ql5', 'Zebra 斑马'],
    ['qln', 'Zebra 斑马'], ['zp4', 'Zebra 斑马'], ['zp5', 'Zebra 斑马'],
    ['lp28', 'Zebra 斑马'], ['tlp28', 'Zebra 斑马'], ['rw4', 'Zebra 斑马'], ['rz4', 'Zebra 斑马'],
    ['s4m', 'Zebra 斑马'], ['z4m', 'Zebra 斑马'], ['z6m', 'Zebra 斑马'],
    ['105sl', 'Zebra 斑马'], ['110xi', 'Zebra 斑马'], ['140xi', 'Zebra 斑马'],
    ['170xi', 'Zebra 斑马'], ['220xi', 'Zebra 斑马'], ['mz220', 'Zebra 斑马'],
    // 其他标签机品牌（须排在 zpl 之前）
    ['godex', 'Godex 科诚'], ['toshiba', 'Toshiba 东芝'], ['tsc', 'TSC 台半'], ['sato', 'SATO'],
    ['postek', 'Postek 博思得'], ['citizen', 'Citizen 西铁城'], ['zicox', 'Zicox 芝柯'],
    ['北洋', '北洋'], ['佳博', '佳博'], ['映美', '映美'], ['得力', '得力'], ['汉印', '汉印'],
    ['hprt', 'HPRT 汉印'], ['gprinter', '佳博'], ['argox', 'Argox 立象'], ['cipherlab', 'CipherLab'],
    // 办公打印品牌
    ['hp ', 'HP 惠普'], ['hp-', 'HP 惠普'], ['hewlett', 'HP 惠普'], ['canon', 'Canon 佳能'],
    ['epson', 'Epson 爱普生'], ['brother', 'Brother 兄弟'], ['xerox', 'Xerox 施乐'],
    ['ricoh', 'Ricoh 理光'], ['kyocera', 'Kyocera 京瓷'], ['lexmark', 'Lexmark 利盟'],
    ['samsung', 'Samsung 三星'], ['sharp', 'Sharp 夏普'], ['konica', 'Konica 柯尼卡'],
    // 兜底：驱动名带 ZPL 且未被上述命中，才认为可能是斑马
    ['zpl', 'Zebra 斑马（ZPL 驱动）'],
  ];
  for (const [k, v] of brands) if (s.includes(k)) return v;
  return '';
}

/** 将原始状态归一化：状态码 / 中文描述 / 队列长度 */
function normalizePrinter(p) {
  const rawStatus = String(p.status || '').trim();
  const num = parseInt(rawStatus, 10);
  const jobs = Number(p.jobs) || 0;
  const errCode = Number(p.errorCode) || 0;
  const errText = ERROR_TEXT[errCode] || '';

  let code = 'unknown', text = '未知状态';
  const key = rawStatus.toLowerCase().replace(/[\s_-]/g, '');
  const byName = STATUS_TEXT[key];
  const byNum = Number.isFinite(num) ? STATUS_NUM[num] : null;

  if (p.offline) { code = 'offline'; text = '脱机'; }
  else if (errCode > 2 && errCode !== 9) { code = 'error'; text = errText || '异常'; }
  else if (byName) { code = byName[0]; text = byName[1]; }
  else if (byNum) { code = byNum[0]; text = byNum[1]; }
  else if (jobs > 0) { code = 'printing'; text = '打印中'; }
  else if (errText) { code = 'error'; text = errText; }

  // 有排队任务时统一按“打印中”体现，并附带队列长度
  if (jobs > 0 && code !== 'offline') {
    if (code === 'idle') { code = 'printing'; text = '打印中'; }
    text += `（队列 ${jobs}）`;
  }
  return {
    name: String(p.name || '').trim(),
    driver: String(p.driver || '').trim(),
    port: String(p.port || '').trim(),
    vendor: detectVendor(p),
    rawStatus,
    statusCode: code,
    statusText: text,
    jobs,
    isDefault: !!p.default,
    offline: !!p.offline,
    shared: !!p.shared,
    errorText: errText,
    source: p.source || '',
  };
}

/** Windows：执行 PowerShell 脚本 */
function listPrintersWin() {
  return new Promise((resolve, reject) => {
    execFile('powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', PS_SCRIPT],
      { timeout: 20000, windowsHide: true, maxBuffer: 8 * 1024 * 1024, encoding: 'utf8' },
      (err, stdout, stderr) => {
        if (err && !stdout) return reject(new Error((stderr || err.message || '').trim() || 'PowerShell 执行失败'));
        // PS 5.1 在输出重定向时可能写入 UTF-8 BOM，需先剥离再解析
        const txt = String(stdout || '').replace(/^\uFEFF+/, '').trim();
        if (!txt) return resolve([]);
        try {
          const parsed = JSON.parse(txt);
          resolve(Array.isArray(parsed) ? parsed : [parsed]);
        } catch (e) {
          reject(new Error('解析打印机数据失败: ' + e.message));
        }
      });
  });
}

/** Linux / macOS：解析 lpstat */
function listPrintersUnix() {
  return new Promise(resolve => {
    execFile('lpstat', ['-p', '-d'], { timeout: 8000, encoding: 'utf8' }, (err, stdout) => {
      if (err && !stdout) return resolve([]);
      const defMatch = /system default destination:\s*(\S+)/i.exec(stdout || '');
      const def = defMatch ? defMatch[1] : '';
      const list = [];
      (stdout || '').split('\n').forEach(line => {
        const m = /^printer\s+(\S+)\s+is\s+(\w+)/i.exec(line.trim());
        if (!m) return;
        list.push({
          name: m[1], driver: '', port: '', status: m[2],
          jobs: 0, offline: /disabled/i.test(m[2]), default: m[1] === def, shared: false, errorCode: 2,
        });
      });
      resolve(list);
    });
  });
}

/** 采集本机打印机（含状态汇总），带 5 秒缓存 */
async function collectPrinters(force) {
  const raw = process.platform === 'win32' ? await listPrintersWin() : await listPrintersUnix();
  const printers = raw.map(normalizePrinter).filter(p => p.name);
  return {
    ok: true,
    platform: process.platform,
    hostname: os.hostname(),
    updatedAt: new Date().toISOString(),
    printers,
    summary: {
      total: printers.length,
      idle: printers.filter(p => p.statusCode === 'idle').length,
      busy: printers.filter(p => p.statusCode === 'printing').length,
      offline: printers.filter(p => p.statusCode === 'offline').length,
      error: printers.filter(p => p.statusCode === 'error').length,
      jobs: printers.reduce((s, p) => s + p.jobs, 0),
    },
  };
}

module.exports = { collectPrinters, normalizePrinter, detectVendor, ERROR_TEXT, STATUS_TEXT };
