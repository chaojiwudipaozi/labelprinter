/**
 * 打印机诊断：排查「某台打印机在工具里看不到」的原因
 * 采集已安装驱动、端口使用情况、即插即用设备，并找出：
 *   - 已创建但没有任何队列使用的 TCP/IP 端口（通常是队列被删或从未创建）
 *   - 已连接硬件但没有对应打印队列的设备（驱动未装/打印机未添加）
 * 实现：调用 scripts/printer-diag.ps1（只读查询）
 */
const path = require('path');
const { execFile } = require('child_process');

const { scriptPath } = require('./paths');
const DIAG_PS = scriptPath('printer-diag.ps1');

/** 去掉 PowerShell 输出可能带有的 UTF-8 BOM（PS 5.1 重定向时会写入） */
function stripBom(s) {
  return String(s || '').replace(/^\uFEFF+/, '').trim();
}

function collectDiag() {
  return new Promise((resolve, reject) => {
    execFile('powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', DIAG_PS],
      { timeout: 40000, windowsHide: true, maxBuffer: 8 * 1024 * 1024, encoding: 'utf8' },
      (err, stdout, stderr) => {
        const txt = stripBom(stdout);
        if (!txt) {
          return reject(new Error(stripBom(stderr) || (err && err.message) || '诊断脚本无输出'));
        }
        try {
          const j = JSON.parse(txt);
          resolve(j);
        } catch (e) {
          reject(new Error('诊断数据解析失败: ' + e.message));
        }
      });
  });
}

/** 生成可执行的添加网络打印机命令（供界面直接展示给用户） */
function buildAddPrinterHint({ name, ip, port = 9100 }) {
  const printerName = name || (ip ? `网络打印机 ${ip}` : '新打印机');
  const lines = [
    `# 以管理员身份运行；-ListDrivers 可先查看本机可用驱动名`,
    `powershell -ExecutionPolicy Bypass -File scripts\\add-printer.ps1 -ListDrivers`,
    `powershell -ExecutionPolicy Bypass -File scripts\\add-printer.ps1 -Name "${printerName}" -Ip "${ip || '打印机IP'}" -Driver "ZDesigner GX430t (ZPL)" -Port ${port}`,
  ];
  return lines.join('\n');
}

module.exports = { collectDiag, stripBom, buildAddPrinterHint };
