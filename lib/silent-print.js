/**
 * 服务端静默打印：把前端渲染好的标签位图直接提交到指定打印机的后台队列
 * （使用 .NET PrintDocument + StandardPrintController，不经浏览器、不弹任何对话框）
 */
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFile } = require('child_process');

const { scriptPath, dataDir } = require('./paths');
const PRINT_PS = scriptPath('print-images.ps1');
const TMP_DIR = dataDir('print-tmp');
const MAX_PAGES = 2000;          // 单次任务最大页数
const MAX_UNIQUE_IMAGES = 400;   // 单次任务最大唯一图片数

fs.mkdirSync(TMP_DIR, { recursive: true });

/** 解析 dataURL -> Buffer（仅允许 PNG） */
function decodePng(dataUrl) {
  const m = /^data:image\/png;base64,([A-Za-z0-9+/=]+)$/.exec(String(dataUrl || '').trim());
  if (!m) return null;
  const buf = Buffer.from(m[1], 'base64');
  // PNG 魔数校验
  if (buf.length < 8 || buf[0] !== 0x89 || buf[1] !== 0x50 || buf[2] !== 0x4e || buf[3] !== 0x47) return null;
  return buf;
}

/** 清理超过 1 小时的临时文件 */
function cleanTmp() {
  try {
    const now = Date.now();
    fs.readdirSync(TMP_DIR).forEach(f => {
      const p = path.join(TMP_DIR, f);
      try { if (now - fs.statSync(p).mtimeMs > 3600e3) fs.unlinkSync(p); } catch (_) { }
    });
  } catch (_) { }
}

/** 执行静默打印：pages = [{ image: dataURL, copies: n }] */
function silentPrint({ printer, widthMm, heightMm, pages }) {
  return new Promise((resolve, reject) => {
    const stamp = Date.now();
    const written = [];
    const listLines = [];

    try {
      pages.forEach((pg, i) => {
        const buf = decodePng(pg.image);
        if (!buf) throw new Error(`第 ${i + 1} 张标签图片数据无效（仅支持 PNG dataURL）`);
        const file = path.join(TMP_DIR, `job${stamp}-${i}.png`);
        fs.writeFileSync(file, buf);
        written.push(file);
        const copies = Math.max(1, Math.min(999, parseInt(pg.copies, 10) || 1));
        for (let c = 0; c < copies; c++) listLines.push(file);
      });

      const listFile = path.join(TMP_DIR, `job${stamp}.txt`);
      fs.writeFileSync(listFile, listLines.join('\r\n'), 'utf-8');
      written.push(listFile);

      execFile('powershell.exe', [
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', PRINT_PS,
        '-Printer', String(printer),
        '-List', listFile,
        '-WidthMm', String(widthMm || 100),
        '-HeightMm', String(heightMm || 60),
      ], { timeout: 120000, windowsHide: true, maxBuffer: 4 * 1024 * 1024, encoding: 'utf8' },
        (err, stdout, stderr) => {
          written.forEach(f => { try { fs.unlinkSync(f); } catch (_) { } });
          const txt = String(stdout || '').replace(/^\uFEFF+/, '').trim();
          let parsed = null;
          try { parsed = JSON.parse(txt.split('\n').pop().trim()); } catch (_) { }
          if (parsed && parsed.ok) return resolve({ ...parsed, pages: listLines.length });
          const msg = (parsed && parsed.message)
            || (stderr || '').trim() || (err && err.message) || '打印失败（打印机无响应或驱动异常）';
          reject(new Error(msg));
        });
    } catch (e) {
      written.forEach(f => { try { fs.unlinkSync(f); } catch (_) { } });
      reject(e);
    }
  });
}

/** 校验打印请求体，返回 { ok, error, printer, widthMm, heightMm, pages } */
function validatePrintRequest(body) {
  const printer = String((body && body.printer) || '').trim();
  if (!printer) return { error: '未选择打印机，请先在顶部选择目标打印机' };
  const pages = (body && body.pages) || [];
  if (!Array.isArray(pages) || !pages.length) return { error: '没有可打印的标签内容' };
  if (pages.length > MAX_UNIQUE_IMAGES) return { error: `单次任务标签种类过多（上限 ${MAX_UNIQUE_IMAGES}）` };

  let total = 0;
  for (const pg of pages) {
    if (!pg || !decodePng(pg.image)) return { error: '标签图片数据无效（仅支持 PNG dataURL）' };
    total += Math.max(1, Math.min(999, parseInt(pg.copies, 10) || 1));
  }
  if (total > MAX_PAGES) return { error: `单次任务总份数过多（${total} > ${MAX_PAGES}），请分批打印` };

  const widthMm = Math.max(10, Math.min(500, Number(body.widthMm) || 100));
  const heightMm = Math.max(10, Math.min(500, Number(body.heightMm) || 60));
  return { ok: true, printer, widthMm, heightMm, pages, total };
}

module.exports = { silentPrint, validatePrintRequest, decodePng, cleanTmp, TMP_DIR, MAX_PAGES };
