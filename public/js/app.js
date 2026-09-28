/* ============================================================
 * 网页版标签打印工具站 - 前端主逻辑
 * 1) 可视化标签设计器（拖动 / 八向缩放 / 属性编辑 / 层级 / 画布拖拽创建文本区域）
 * 2) 双击元素快捷编辑：直接改文字、绑定下方 Excel 字段、调字体字号
 * 3) 标准化 Excel：下载模板 → 填写 → 导入校验 → 自动识别每行标签数据
 * 4) 元素绑定数据列 → 预览 / 逐行批量打印
 * 5) 模板保存 / 读取 / 删除（存于服务端 data/labels 目录）
 * ============================================================ */
'use strict';

/* ---------- 常量与状态 ---------- */
const MM = 96 / 25.4; // 1mm 对应屏幕像素(96dpi)
const DEFAULT_ZOOM = 1.5;   // 标签区默认显示比例：150%（等比，仅影响屏幕显示）
const FONT_LIST = ['宋体', '黑体', '微软雅黑', '楷体', '仿宋', '等线', 'Arial', 'Times New Roman', 'Courier New'];
const DEFAULT_FONT = '微软雅黑';
/* 标准 Excel 模板列（导出模板与导入校验共用） */
const STANDARD_COLUMNS = [
  '物料编码', '物料名称', '规格', '型号', '数量', '单位', '批号', '序列号',
  '供应商编码', '供应商名称', '生产日期', '有效期至', '采购单号', '任务单号',
  '二维码内容', '条码内容', '打印数量',
];
/* 缺失时可按规则自动生成的列 */
const DERIVED = {
  '二维码内容': r => [r['物料编码'], r['批号'], r['生产日期']].filter(Boolean).join(';'),
  '条码内容': r => r['物料编码'] || '',
};

const state = {
  name: '',                                   // 当前模板名
  label: { width: 100, height: 60 },          // 标签尺寸 mm
  elements: [],                               // 画布元素
  selectedId: null,
  zoom: DEFAULT_ZOOM,                         // 画布显示缩放（默认 150%，等比缩放，不影响打印尺寸）
  zoomExplicit: false,                        // 用户是否手动选过缩放（用于草稿恢复时保留用户选择）
  headers: [],                                // Excel 列名
  rows: [],                                   // Excel 数据行
  qty: [],                                    // 每行打印数量（与 rows 平行）
  currentRow: -1,                             // 预览/打印选中的行号
  tableEdit: null,                            // 正在编辑的表格元素 id
  tableSel: null,                             // 表格单元格框选范围 {r1,c1,r2,c2}
  printer: '',                                // 目标打印机（随模板保存）
  printerAgent: '',                           // 打印服务地址（空=当前服务端；可指向局域网其他设备的打印服务）
  printers: [],                               // 本机打印机列表
  printersSummary: null,                      // 打印机状态汇总
  printersError: '',                          // 打印机检测错误
  printersUpdatedAt: '',                      // 打印机状态更新时间
};

const $ = s => document.querySelector(s);
const canvas = $('#canvas');
const mm2px = v => v * MM * state.zoom;
let uid = 1;

/* ---------- 元素默认值 ---------- */
const DEFAULTS = {
  text:    { type: 'text',    x: 10, y: 10, w: 50, h: 7, text: '文本内容', binding: '', prefix: '', fontSize: 10, bold: false, align: 'left', fontFamily: DEFAULT_FONT, dynamic: '', dateFormat: 'YYYY/M/DD HH:MM', autoFit: true },
  barcode: { type: 'barcode', x: 10, y: 10, w: 45, h: 12, binding: '', text: '1234567890', barcodeFormat: 'CODE128', showText: true, barWidth: 2, fontSize: 10, fontFamily: DEFAULT_FONT, autoFit: true },
  qrcode:  { type: 'qrcode',  x: 10, y: 10, w: 20, h: 20, binding: '', text: '二维码内容', correctLevel: 'M' },
  rect:    { type: 'rect',    x: 10, y: 10, w: 60, h: 30, borderWidth: 0.4 },
  line:    { type: 'line',    x: 10, y: 10, w: 60, h: 0.4 },
  table:   { type: 'table',   x: 5, y: 5, w: 90, h: 24, rows: 3, cols: 4, colWidths: [], rowHeights: [], merges: [], cells: {}, borderWidth: 0.3 },
};

/* ============================================================
 * 动态内容：当前日期时间
 * ============================================================ */
/** 可选日期时间格式（默认与需求一致：月份不补零） */
const DATETIME_FORMATS = [
  'YYYY/M/DD HH:MM',
  'YYYY/M/DD  HH:MM',
  'YYYY/MM/DD HH:MM',
  'YYYY-MM-DD HH:MM',
  'YYYY年M月D日 HH:MM',
  'YYYY/M/DD',
  'HH:MM',
  'YYYY/M/DD HH:MM:SS',
];

/** 按格式串生成时间文本：YYYY 年 / MM 月(补零) / M 月 / DD 日 / D 日 / HH 时 / H 时 / MM 分 / SS 秒 */
function formatDateTime(fmt, date) {
  const d = date || new Date();
  const p2 = n => String(n).padStart(2, '0');
  // 时间部分用小写 mm/ss 会被上面的 MM 冲突，这里按占位符顺序替换，先长后短
  const map = {
    YYYY: String(d.getFullYear()),
    MM: p2(d.getMonth() + 1),
    M: String(d.getMonth() + 1),
    DD: p2(d.getDate()),
    D: String(d.getDate()),
    HH: p2(d.getHours()),
    H: String(d.getHours()),
    mm: p2(d.getMinutes()),
    ss: p2(d.getSeconds()),
  };
  // 歧义消解：冒号后的 MM/SS 表示"分/秒"，其余位置的 MM 表示"月份"
  const s = String(fmt || 'YYYY/M/DD HH:MM')
    .replace(/:MM/g, ':mm')
    .replace(/:SS/g, ':ss');
  return s.replace(/YYYY|MM|DD|HH|mm|ss|M|D|H/g, t => (map[t] !== undefined ? map[t] : t));
}

/* ============================================================
 * 数据取值：静态文字 + 可选绑定列
 * ============================================================ */
function cellValue(el, row) {
  if (el.dynamic === 'datetime') return formatDateTime(el.dateFormat);   // 动态：当前日期时间
  if (!el.binding) return el.text || '';
  if (!row || row[el.binding] === undefined) return null; // 无数据 → 显示占位
  const v = row[el.binding];
  return v === '' || v === undefined ? null : String(v);
}

/* ============================================================
 * 表格模型（行高/列宽自由调节、合并单元格）
 * 坐标与尺寸单位统一为毫米，表格左上角为 (el.x, el.y)
 * ============================================================ */
function makeTable(rows, cols, w, h, borderWidth) {
  // 重要：行高/列宽/合并块/单元格数据必须「每个表格各自独立」。
  // 若沿用 DEFAULTS.table 的同名对象（展开运算符是浅拷贝），会让所有表格共享同一份单元格数据，
  // 表现为「绑定不 1:1、文本跑到别的表格/别的行」。
  const t = {
    ...DEFAULTS.table, id: 'e' + (uid++), rows, cols, w, h, borderWidth,
    colWidths: [], rowHeights: [], merges: [], cells: {},
  };
  t.colWidths = new Array(cols).fill(Math.round((w / cols) * 100) / 100);
  t.rowHeights = new Array(rows).fill(Math.round((h / rows) * 100) / 100);
  return t;
}

/** 深拷贝元素：表格需拷贝行高/列宽/合并块/单元格数据，避免复制体与原件共享（串位） */
function cloneElement(el) {
  const c = { ...el, id: 'e' + (uid++) };
  if (el.type === 'table') {
    c.colWidths = [...(el.colWidths || [])];
    c.rowHeights = [...(el.rowHeights || [])];
    c.merges = (el.merges || []).map(m => ({ ...m }));
    c.cells = JSON.parse(JSON.stringify(el.cells || {}));
  }
  return c;
}

/** 规整合并块：丢弃越界/不完整的合并块（行列增删后若不清理，会出现引用不存在行列的块 → 渲染出 NaN 坐标） */
function tblClampMerges(el) {
  el.merges = (el.merges || []).filter(m => m.rs >= 1 && m.cs >= 1
    && m.r >= 0 && m.c >= 0 && m.r + m.rs <= el.rows && m.c + m.cs <= el.cols);
}

/** 载入历史数据（草稿/模板）时规整表格：补齐行列尺寸、修正 cells/merges，避免脏数据导致错位 */
function tblSanitize(el) {
  if (!el || el.type !== 'table') return el;
  el.rows = Math.max(1, parseInt(el.rows, 10) || 1);
  el.cols = Math.max(1, parseInt(el.cols, 10) || 1);
  el.colWidths = Array.isArray(el.colWidths) ? el.colWidths.slice(0, el.cols) : [];
  el.rowHeights = Array.isArray(el.rowHeights) ? el.rowHeights.slice(0, el.rows) : [];
  const defW = Math.max(2, Math.round(((el.w || 90) / el.cols) * 100) / 100);
  const defH = Math.max(1, Math.round(((el.h || 24) / el.rows) * 100) / 100);
  while (el.colWidths.length < el.cols) el.colWidths.push(defW);
  while (el.rowHeights.length < el.rows) el.rowHeights.push(defH);
  el.cells = (el.cells && typeof el.cells === 'object') ? el.cells : {};
  el.merges = Array.isArray(el.merges) ? el.merges : [];
  tblClampMerges(el);
  tblSyncSize(el);
  return el;
}

/** 行列增删后重映射单元格数据（键形如 "r,c"）；fn 返回 null 表示该单元格被删除 */
function tblRemapCells(el, fn) {
  const src = el.cells || {};
  const out = {};
  Object.keys(src).forEach(k => {
    const p = k.split(',');
    const pos = fn(+p[0], +p[1]);
    if (!pos) return;
    out[pos[0] + ',' + pos[1]] = src[k];
  });
  el.cells = out;
}

/** 列左边界相对 X（mm） */
function tblColX(el, c) { let x = 0; for (let i = 0; i < c; i++) x += el.colWidths[i]; return x; }
/** 行上边界相对 Y（mm） */
function tblRowY(el, r) { let y = 0; for (let i = 0; i < r; i++) y += el.rowHeights[i]; return y; }

/** 包含 (r,c) 的合并块（无则返回 1×1 块） */
function tblBlockAt(el, r, c) {
  for (const m of (el.merges || [])) {
    if (r >= m.r && r < m.r + m.rs && c >= m.c && c < m.c + m.cs) return m;
  }
  return { r, c, rs: 1, cs: 1 };
}

/** 全部可见单元格块（合并块只返回一次，按行优先） */
function tblBlocks(el) {
  const covered = new Set();
  (el.merges || []).forEach(m => {
    for (let r = m.r; r < m.r + m.rs; r++) {
      for (let c = m.c; c < m.c + m.cs; c++) {
        if (r !== m.r || c !== m.c) covered.add(r + ',' + c);
      }
    }
  });
  const blocks = [];
  for (let r = 0; r < el.rows; r++) {
    for (let c = 0; c < el.cols; c++) {
      if (covered.has(r + ',' + c)) continue;
      blocks.push(tblBlockAt(el, r, c));
    }
  }
  return blocks;
}

/** 块的范围（相对表格左上角，mm） */
function tblBlockRect(el, b) {
  const x = tblColX(el, b.c);
  const y = tblRowY(el, b.r);
  let bw = 0, bh = 0;
  for (let i = b.c; i < b.c + b.cs; i++) bw += el.colWidths[i];
  for (let i = b.r; i < b.r + b.rs; i++) bh += el.rowHeights[i];
  return { x, y, w: bw, h: bh };
}

/** 单元格数据（按块的左上角键存储，保证合并后内容不丢） */
function tblCellData(el, b) {
  const key = b.r + ',' + b.c;
  if (!el.cells[key]) {
    el.cells[key] = {
      dynamic: '', text: '', binding: '', fontSize: 9, bold: false,
      align: 'center', valign: 'middle',
      autoFit: true,          // 自适应字号：默认开启——文字放不进一行时自动缩小字号
    };
  }
  return el.cells[key];
}

/**
 * 单元格显示文本（静态文字或 Excel 字段）
 * mode==='mm' 为打印路径：绑定字段为空值时**输出空白**（屏幕上仍显示 {字段名} 便于设计时辨认）
 */
function tblCellText(cd, row, mode) {
  if (cd.dynamic === 'datetime') return formatDateTime(cd.dateFormat);
  if (cd.binding) {
    const v = row ? row[cd.binding] : undefined;
    if (v === undefined || v === '') return mode === 'mm' ? '' : `{${cd.binding}}`;
    return String(v);
  }
  return cd.text || '';
}

/** 表格总尺寸按行高列宽重新计算 */
function tblSyncSize(el) {
  el.w = Math.round(el.colWidths.reduce((s, v) => s + v, 0) * 100) / 100;
  el.h = Math.round(el.rowHeights.reduce((s, v) => s + v, 0) * 100) / 100;
}

/** 选中区域归一化 {r1,c1,r2,c2}（含端点） */
function tblSelRect(sel) {
  if (!sel) return null;
  return {
    r1: Math.min(sel.r1, sel.r2), r2: Math.max(sel.r1, sel.r2),
    c1: Math.min(sel.c1, sel.c2), c2: Math.max(sel.c1, sel.c2),
  };
}
function tblInSel(sel, r, c) {
  if (!sel) return false;
  const s = tblSelRect(sel);
  return r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2;
}

/* ============================================================
 * 元素内容渲染（屏幕与打印共用，单位 mm，绝对定位）
 * ============================================================ */
function barcodeSVG(el, value) {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  try {
    // 条码下方的可读文字同样支持自适应字号：放不进元素宽度就按比例缩小
    const txt = String(value);
    const pt = el.autoFit === false
      ? (el.fontSize || 10)
      : fitFontSizePt(txt, Math.max(0.5, (el.w || 0) - 0.4), el.fontSize || 10, el.fontFamily, el.bold);
    JsBarcode(svg, value, {
      format: el.barcodeFormat || 'CODE128',
      width: el.barWidth || 2,
      height: Math.round(el.h * 4),          // 按高度换算分辨率
      displayValue: !!el.showText,
      fontSize: pt * 1.4,
      font: el.fontFamily || DEFAULT_FONT,
      margin: 0,
    });
  } catch (e) {
    return `<div style="width:100%;height:100%;display:flex;align-items:center;justify-content:center;
      border:1px dashed #f00;color:#f00;font-size:9px;overflow:hidden">条码错误:${e.message}</div>`;
  }
  svg.setAttribute('preserveAspectRatio', 'none');
  svg.style.width = '100%';
  svg.style.height = '100%';
  return svg.outerHTML;
}

/* 二维码 canvas 缓存：屏幕展示取 dataURL，服务端打印直接绘制 canvas */
const qrCanvasCache = new Map();
function qrCanvasFor(el, value, px) {
  const key = value + '|' + px + '|' + (el.correctLevel || 'M');
  if (!qrCanvasCache.has(key)) {
    const holder = document.createElement('div');
    holder.style.cssText = 'position:absolute;left:-9999px;top:0;';
    document.body.appendChild(holder);
    let cv = null;
    try {
      new QRCode(holder, {
        text: value, width: px, height: px,
        correctLevel: QRCode.CorrectLevel[el.correctLevel || 'M'] || QRCode.CorrectLevel.M,
      });
      cv = holder.querySelector('canvas');
    } catch (e) { cv = null; }
    document.body.removeChild(holder);
    qrCanvasCache.set(key, cv);
  }
  return qrCanvasCache.get(key);
}

/** 保留 3 位小数（缩放计算会产生 0.6000000000000001 这类浮点尾数） */
const r3 = v => Math.round(v * 1000) / 1000;

function elementContent(el, row, mode) {
  const value = cellValue(el, row);
  /* mode==='mm' 是打印路径（浏览器打印预览）；此时「绑定字段为空值」的元素输出空白，
     不再把 {字段名} 占位一起打出来（屏幕上仍保留占位，方便设计时确认绑定关系）。 */
  const printing = mode === 'mm';
  /* zf：画布放大倍数。字号(pt)、以 mm 表示的线宽在屏幕上按缩放等比放大，
     使“放大显示”是真正的等比放大（所见比例与打印一致）；打印路径（mode==='mm'）始终用真实尺寸。 */
  const zf = mode === 'mm' ? 1 : state.zoom;
  switch (el.type) {
    case 'text': {
      const ph = value === null;
      if (ph && printing) return '';                 // 打印：该字段为空 → 整块不打印（连前缀一起去掉）
      const full = ph ? `{${el.binding}}` : (el.prefix || '') + value;
      // 自适应字号：一行放不下就自动缩小字号（与表格单元格同一套算法）
      const fit = el.autoFit !== false;
      const pt = fit
        ? fitFontSizePt(full, Math.max(0.5, (el.w || 0) - 0.4), el.fontSize || 10, el.fontFamily, el.bold)
        : (el.fontSize || 10);
      const body = ph ? `<span class="bind-ph">{${el.binding}}</span>` : escHtml(full);
      /* 布局样式全部行内：打印窗口不加载本站样式表，
         否则 display:flex 与 span 的 text-align 都会失效 → 打印预览里居中/对齐不生效。 */
      return `<div class="el-text" style="display:flex;
        justify-content:${el.align === 'center' ? 'center' : el.align === 'right' ? 'flex-end' : 'flex-start'};
        align-items:center;width:100%;height:100%;overflow:hidden;line-height:1.2;word-break:break-all;
        ${fit ? 'white-space:nowrap;' : 'white-space:pre-wrap;'}
        font-size:${pt * zf}pt;font-weight:${el.bold ? 'bold' : 'normal'};
        font-family:'${el.fontFamily || DEFAULT_FONT}'"
        ><span style="flex:1 1 auto;min-width:0;text-align:${el.align || 'left'}">${body}</span></div>`;
    }
    case 'barcode': {
      if (value === null) {
        if (printing) return '';                     // 打印：字段为空 → 不打印（连占位框也去掉）
        return `<div style="width:100%;height:100%;display:flex;align-items:center;justify-content:center;
          border:1px dashed #b085f5;color:#b085f5;font-size:${9 * zf}px;overflow:hidden">{${el.binding || '条码'}}</div>`;
      }
      return barcodeSVG(el, value);
    }
    case 'qrcode': {
      if (value === null) {
        if (printing) return '';
        return `<div style="width:100%;height:100%;display:flex;align-items:center;justify-content:center;
          border:1px dashed #b085f5;color:#b085f5;font-size:${9 * zf}px;overflow:hidden">{${el.binding || '二维码'}}</div>`;
      }
      // 分辨率同时考虑打印精度与屏幕缩放，避免放大后二维码模糊
      const px = Math.max(96, Math.round(el.w * MM * 3 * zf));
      const url = qrDataUrlFor(el, value, px);
      return url ? `<img src="${url}" style="width:100%;height:100%">`
                 : `<div style="width:100%;height:100%;border:1px dashed #f00;color:#f00;font-size:${9 * zf}px;display:flex;align-items:center;justify-content:center">二维码错误</div>`;
    }
    case 'rect':
      return `<div style="width:100%;height:100%;border:${r3((el.borderWidth || 0.4) * zf)}mm solid #000"></div>`;
    case 'line': {
      const vertical = el.h > el.w;
      const bw = r3((el.borderWidth || 0.4) * zf);
      return `<div style="width:100%;height:100%;${vertical
        ? `border-left:${bw}mm solid #000;` : `border-top:${bw}mm solid #000;`}"></div>`;
    }
    case 'table': return tableHTML(el, row, mode);
  }
  return '';
}

/* 生成二维码 dataURL（带缓存） */
function qrDataUrlFor(el, value, px) {
  const cv = qrCanvasFor(el, value, px);
  return cv ? cv.toDataURL('image/png') : '';
}

/* ============================================================
 * 单元格「自适应字号」
 * 需求：单元格字数多到一行放不下时，自动缩小字号，使内容在一行内完整显示；
 *      可在单元格属性里开关（默认开启）。
 * 实现：用离屏 canvas 按「pt 字号 + 字体」量出文字宽度（mm），
 *      再按单元格可用宽度线性缩放字号 —— 与屏幕缩放无关，
 *      因此画布、打印预览、直连打印位图三种路径结果一致。
 * ============================================================ */
const PT2MM = 25.4 / 72;                       // 1pt = 0.3528mm
let __measureCtx = null;
/** 量出文字在指定字号(pt)/字体下的宽度（mm） */
function textWidthMm(text, pt, family, bold) {
  const s = String(text == null ? '' : text).replace(/\r?\n/g, ' ');
  if (!s) return 0;
  if (!__measureCtx) __measureCtx = document.createElement('canvas').getContext('2d');
  __measureCtx.font = `${bold ? 'bold ' : ''}${pt}pt "${family || DEFAULT_FONT}"`;
  return __measureCtx.measureText(s).width / MM;   // canvas 量出的是 px（96dpi），换算为 mm
}

/**
 * 计算单元格实际使用的字号(pt)：
 *  - 未开启自适应 → 直接用设定字号（文字可换行）
 *  - 开启自适应 → 放得下就用设定字号；放不下则按比例缩小到刚好放得下（最小 3pt）
 * 说明：宽度留 1% 余量并向下取整到 0.1pt，确保「缩小后一定不会超宽」
 *      （canvas 测量与实际排版有极小差异，宁可再小一点点也不让文字被裁掉）。
 */
function cellFontSizePt(cd, text, availMm) {
  const base = cd.fontSize || 9;
  if (cd.autoFit === false) return base;
  return fitFontSizePt(text, availMm, base, cd.fontFamily, cd.bold);
}

/**
 * 通用「自适应字号」（表格单元格 / 文本框 / 条码文字共用）：
 * 放得下就用设定字号（不放大）；放不下则按比例缩小到刚好一行放得下，最小 3pt。
 */
function fitFontSizePt(text, availMm, basePt, family, bold) {
  const base = basePt || 10;
  if (!text || !(availMm > 0)) return base;
  const w = textWidthMm(text, base, family, bold);
  if (w <= availMm) return base;               // 本来就放得下：保持原字号（不放大）
  return Math.max(3, Math.floor((availMm * 0.99 / w) * base * 10) / 10);
}

/** 单元格可用文字宽度（mm）：块宽 − 两侧边框 − 两侧内边距(约 0.6mm) */
function tblCellAvailMm(el, b) {
  const r = tblBlockRect(el, b);
  return Math.max(1, r.w - 2 * (el.borderWidth || 0.3) - 0.6);
}

/**
 * 表格元素 → HTML（边框采用 上/左 + 末行末列，避免相邻边框叠加变粗）
 * mode='mm' 时用毫米输出（打印路径），否则用屏幕像素（画布路径，跟随缩放）；
 * 之前打印路径也走像素，导致画布缩放≠100% 时表格打印尺寸被放大。
 */
function tableHTML(el, row, mode) {
  const editing = state.tableEdit === el.id;
  const mmMode = mode === 'mm';
  const K = mmMode ? 1 : MM * state.zoom;                     // 1mm 对应的输出单位
  const U = mmMode ? 'mm' : 'px';
  const px = v => `${Math.round(v * 10000) / 10000}${U}`;
  const bw = mmMode
    ? px(el.borderWidth || 0.3)
    : `${Math.max(1, (el.borderWidth || 0.3) * K)}px`;         // 屏幕上保证至少 1px 可见
  const parts = [];
  const blocks = tblBlocks(el);

  blocks.forEach(b => {
    const r = tblBlockRect(el, b);
    const cd = tblCellData(el, b);
    const inSel = editing && tblInSel(state.tableSel, b.r, b.c);
    const sel = tblSelRect(state.tableSel);
    const single = sel && sel.r1 === sel.r2 && sel.c1 === sel.c2;
    // 自适应字号：按单元格可用宽度算出实际字号（放不下则缩小），保证一行显示完整
    const txt = tblCellText(cd, row, mode);
    const fit = cd.autoFit !== false;
    const pt = cellFontSizePt(cd, txt, tblCellAvailMm(el, b));
    const style = [
      // 布局样式全部行内给出：打印窗口不加载本站样式表，只有行内样式才能保证打印与画布一致
      'position:absolute', 'display:flex', 'box-sizing:border-box', 'overflow:hidden', 'line-height:1.15',
      'word-break:break-all',
      `left:${px(r.x * K)}`, `top:${px(r.y * K)}`, `width:${px(r.w * K)}`, `height:${px(r.h * K)}`,
      `border-top:${bw} solid #000`, `border-left:${bw} solid #000`,
      b.c + b.cs >= el.cols ? `border-right:${bw} solid #000` : '',
      b.r + b.rs >= el.rows ? `border-bottom:${bw} solid #000` : '',
      `font-size:${pt * (mmMode ? 1 : state.zoom)}pt`,   // 画布按缩放等比放大字号
      `font-weight:${cd.bold ? 'bold' : 'normal'}`,
      `font-family:'${cd.fontFamily || DEFAULT_FONT}'`,
      `justify-content:${cd.align === 'left' ? 'flex-start' : cd.align === 'right' ? 'flex-end' : 'center'}`,
      `align-items:${cd.valign === 'top' ? 'flex-start' : cd.valign === 'bottom' ? 'flex-end' : 'center'}`,
      fit ? 'white-space:nowrap' : '',                  // 自适应：强制不换行（放不下就缩小字号）
    ].filter(Boolean).join(';');
    // 注意：span 是整宽（width:100%），此时容器上的 justify-content 对文字水平位置不再起作用，
    // 真正的水平对齐必须由 span 的 text-align 决定 —— 之前缺这一条，导致单元格「居中」失效。
    parts.push(`<div class="tbl-cell${inSel ? (single ? ' sel-single' : ' sel') : ''}${editing ? ' editing' : ''}"
      data-r="${b.r}" data-c="${b.c}" style="${style}">
      <span class="cell-text" style="width:100%;padding:0 1px;text-align:${cd.align === 'right' ? 'right' : cd.align === 'left' ? 'left' : 'center'}">${escHtml(txt)}</span></div>`);
  });

  // 编辑模式（仅画布）：显示可拖拽的行/列分隔线（按块边缘生成，合并单元格内部不出现分隔线）
  if (editing && !mmMode) {
    blocks.forEach(b => {
      const r = tblBlockRect(el, b);
      if (b.c > 0) {
        parts.push(`<div class="tbl-sep v" data-kind="col" data-i="${b.c}"
          style="left:${r.x * K - 3}px;top:${r.y * K}px;height:${r.h * K}px"></div>`);
      }
      if (b.r > 0) {
        parts.push(`<div class="tbl-sep h" data-kind="row" data-i="${b.r}"
          style="top:${r.y * K - 3}px;left:${r.x * K}px;width:${r.w * K}px"></div>`);
      }
    });
  }
  return parts.join('');
}

function escHtml(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

/** 当前站点的访问地址（换端口 / 换主机后提示文案自动跟随，无需改代码） */
function siteOrigin() {
  return location.protocol.startsWith('http') ? location.origin : 'http://服务器IP:端口';
}

/* ============================================================
 * 通知与对话框
 * 说明：部分嵌入式预览/沙箱 iframe 会静默拦截原生 alert/confirm/prompt，
 *      导致操作“无任何反馈”，因此全部改用页面内自绘 UI。
 * ============================================================ */
function notify(msg, type = 'info', ms = 3600) {
  const box = $('#toasts');
  const el = document.createElement('div');
  el.className = 'toast-item ' + type;
  el.textContent = msg;
  box.appendChild(el);
  if (ms > 0) setTimeout(() => el.remove(), ms);   // ms=0 表示常驻（需手动 remove）
  return el;
}

/** 通用对话框：input 为字符串时是输入框，为 null 时是确认框；返回 输入值 / true / null(取消) */
function dialog({ title = '提示', message = '', input = null, okText = '确定', cancelText = '取消', danger = false }) {
  return new Promise(resolve => {
    const mask = $('#dialog-mask');
    const inputEl = $('#dlg-input');
    $('#dlg-title').textContent = title;
    $('#dlg-msg').textContent = message;
    $('#dlg-ok').textContent = okText;
    $('#dlg-cancel').textContent = cancelText;
    $('#dlg-ok').className = 'primary' + (danger ? ' danger' : '');
    if (input === null) { inputEl.classList.add('hidden'); inputEl.value = ''; }
    else { inputEl.classList.remove('hidden'); inputEl.value = input; }
    mask.classList.remove('hidden');
    setTimeout(() => { input === null ? $('#dlg-ok').focus() : inputEl.select(); }, 30);

    const done = val => {
      mask.classList.add('hidden');
      $('#dlg-ok').removeEventListener('click', onOk);
      $('#dlg-cancel').removeEventListener('click', onCancel);
      mask.removeEventListener('mousedown', onMask);
      document.removeEventListener('keydown', onKey);
      resolve(val);
    };
    const onOk = () => done(input === null ? true : (inputEl.value.trim() || null));
    const onCancel = () => done(null);
    const onMask = e => { if (e.target === mask) onCancel(); };
    const onKey = e => {
      if (e.key === 'Escape') { e.preventDefault(); onCancel(); }
      else if (e.key === 'Enter') { e.preventDefault(); onOk(); }
    };
    $('#dlg-ok').addEventListener('click', onOk);
    $('#dlg-cancel').addEventListener('click', onCancel);
    mask.addEventListener('mousedown', onMask);
    document.addEventListener('keydown', onKey);
  });
}
const askText = (title, def = '') => dialog({ title, input: def });
const askConfirm = (title, message, okText = '确定') => dialog({ title, message, danger: true, okText });

/** 模板状态徽标 */
function setTplStatus(saved) {
  const el = $('#tpl-status');
  const name = state.name || '未命名';
  el.textContent = (saved ? '✔ 已保存：' : '● 未保存：') + name;
  el.className = 'tpl-status ' + (saved ? 'saved' : 'dirty');
}

/** 服务端连通性检测：区分「通过 localhost 访问」与「直接打开文件」等异常场景 */
async function checkServer() {
  const el = $('#server-status');
  try {
    const res = await fetch('/api/labels', { cache: 'no-store' });
    const j = await res.json();
    if (!j.ok) throw new Error(j.message || '接口返回异常');
    el.textContent = '● 服务端已连接';
    el.className = 'srv-status ok';
    return true;   // 模板列表由初始化时统一加载
  } catch (e) {
    el.textContent = '● 服务端未连接';
    el.className = 'srv-status bad';
    el.title = '无法访问 /api/labels：' + e.message
      + `\n请确认后端已启动，并通过 ${siteOrigin()} 打开本页面（不要直接双击 index.html）`;
    notify('无法连接服务端接口 /api/labels（' + e.message + '），模板将无法保存。'
      + `请确认后端已启动，并通过 ${siteOrigin()} 访问本页面。`, 'error', 12000);
    return false;
  }
}

/* ============================================================
 * 画布渲染
 * ============================================================ */
const HANDLE_POS = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']; // 八向缩放手柄

function renderCanvas() {
  canvas.style.width = mm2px(state.label.width) + 'px';
  canvas.style.height = mm2px(state.label.height) + 'px';
  canvas.innerHTML = '';
  const previewRow = state.rows[state.currentRow] || null;

  state.elements.forEach(el => {
    const div = document.createElement('div');
    div.className = 'el el-' + el.type + (el.id === state.selectedId ? ' selected' : '')
      + (el.type === 'table' && state.tableEdit === el.id ? ' editing' : '');
    div.dataset.id = el.id;
    div.style.cssText += `left:${mm2px(el.x)}px;top:${mm2px(el.y)}px;width:${mm2px(el.w)}px;height:${mm2px(el.h)}px;`;
    div.innerHTML = elementContent(el, previewRow);
    if (el.id === state.selectedId) HANDLE_POS.forEach(pos => {
      const h = document.createElement('div');
      h.className = 'handle ' + pos;
      h.dataset.pos = pos;
      div.appendChild(h);
    });
    canvas.appendChild(div);
  });
}

/* ============================================================
 * 选中 / 拖动 / 缩放
 * ============================================================ */
function findEl(id) { return state.elements.find(e => e.id === id); }

canvas.addEventListener('mousedown', e => {
  const handle = e.target.closest('.handle');
  const elDiv = e.target.closest('.el');
  if (!elDiv) { startDrawTextArea(e); return; }  // 空白处拖拽 → 自由绘制文本区域
  const el = findEl(elDiv.dataset.id);
  if (!el) return;

  // 表格编辑模式：优先处理“拖拽行/列分隔线”与“框选单元格”
  if (el.type === 'table' && state.tableEdit === el.id) {
    const sep = e.target.closest('.tbl-sep');
    if (sep) {
      // 必须在 select()（会重建 DOM）之前取出边界信息与鼠标起点
      const kind = sep.dataset.kind, idx = +sep.dataset.i, mx = e.clientX, my = e.clientY;
      select(el.id); e.preventDefault();
      startTableSeparatorDrag(el, kind, idx, mx, my);
      return;
    }
    const cell = e.target.closest('.tbl-cell');
    if (cell) {
      select(el.id); e.preventDefault();
      startTableCellSelect(el, +cell.dataset.r, +cell.dataset.c);
      // 快捷编辑器开着时，点选其他单元格即切换编辑目标：
      // 保证「蓝色高亮的单元格」＝「正在编辑/绑定的单元格」，避免绑到上一个单元格上
      if (!$('#quick-edit').classList.contains('hidden') && qeEl && qeEl.id === el.id) {
        const k = cell.dataset.r + ',' + cell.dataset.c;
        if (qeCellKey !== k) openQuickEdit(el, undefined, undefined, k, true);
      }
      return;
    }
  }

  select(el.id);
  e.preventDefault();
  // 快捷编辑器已打开时，切换元素则同步切换编辑目标
  if (!$('#quick-edit').classList.contains('hidden') && qeEl && qeEl.id !== el.id) {
    const box = elDiv.getBoundingClientRect();
    openQuickEdit(el, box.right, box.top);
  }

  const start = {
    mx: e.clientX, my: e.clientY, el: { ...el },
    sizes: el.type === 'table' ? { colWidths: [...el.colWidths], rowHeights: [...el.rowHeights], w: el.w, h: el.h } : null,
  };
  const mode = handle ? 'resize:' + handle.dataset.pos : 'move';

  function onMove(ev) {
    const dmx = (ev.clientX - start.mx) / (MM * state.zoom);
    const dmy = (ev.clientY - start.my) / (MM * state.zoom);
    const s = start.el;
    if (mode === 'move') {
      el.x = snap(s.x + dmx); el.y = snap(s.y + dmy);
    } else {
      const dir = mode.split(':')[1];
      if (dir.includes('e')) el.w = Math.max(1, snap(s.w + dmx));
      if (dir.includes('s')) el.h = Math.max(0.4, snap(s.h + dmy));
      if (dir.includes('w')) { el.x = snap(s.x + dmx); el.w = Math.max(1, s.w - dmx); }
      if (dir.includes('n')) { el.y = snap(s.y + dmy); el.h = Math.max(0.4, s.h - dmy); }
      // 表格：整体缩放时按比例同步行高/列宽
      if (el.type === 'table' && start.sizes) {
        const sx = el.w / start.sizes.w, sy = el.h / start.sizes.h;
        el.colWidths = start.sizes.colWidths.map(v => Math.max(1, Math.round(v * sx * 100) / 100));
        el.rowHeights = start.sizes.rowHeights.map(v => Math.max(0.5, Math.round(v * sy * 100) / 100));
        tblSyncSize(el);
      }
    }
    renderCanvas(); renderProps();
  }
  function onUp() {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    saveLocal();
  }
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
});

const snap = v => Math.round(v * 10) / 10;

/* ---------- 在画布空白处拖拽，自由“画”出一个文本区域 ---------- */
function startDrawTextArea(e) {
  if (e.button !== 0) return;
  const rect = canvas.getBoundingClientRect();
  const ox = e.clientX - rect.left, oy = e.clientY - rect.top;
  const box = document.createElement('div');
  box.id = 'draw-box';
  canvas.appendChild(box);

  function onMove(ev) {
    const cx = Math.min(Math.max(ev.clientX - rect.left, 0), rect.width);
    const cy = Math.min(Math.max(ev.clientY - rect.top, 0), rect.height);
    box.style.cssText += `left:${Math.min(ox, cx)}px;top:${Math.min(oy, cy)}px;
      width:${Math.abs(cx - ox)}px;height:${Math.abs(cy - oy)}px;`;
  }
  function onUp(ev) {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    box.remove();
    const px2mm = p => p / (MM * state.zoom);
    const w = px2mm(Math.abs(ev.clientX - rect.left - ox));
    const h = px2mm(Math.abs(ev.clientY - rect.top - oy));
    if (w < 3 || h < 2) { select(null); return; }        // 视为普通点击 → 取消选中
    const el = {
      ...DEFAULTS.text, id: 'e' + (uid++),
      x: snap(px2mm(Math.min(ox, ev.clientX - rect.left))),
      y: snap(px2mm(Math.min(oy, ev.clientY - rect.top))),
      w: snap(w), h: snap(h),
      text: '文本内容', binding: '', prefix: '',
      fontSize: Math.max(6, Math.min(24, Math.round(h * 2.6))),  // 按框高自适应字号
    };
    state.elements.push(el);
    select(el.id); saveLocal();
    openQuickEdit(el);                                   // 画完直接进入编辑/绑定
  }
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

/* ---------- 表格：编辑模式 / 框选 / 分隔线拖拽 ---------- */
function enterTableEdit(id) {
  state.tableEdit = id;
  state.tableSel = { r1: 0, c1: 0, r2: 0, c2: 0 };
  state.selectedId = id;
  renderCanvas(); renderProps(); updateTableBar();
  infoBar('已进入表格编辑：拖拽框选单元格 → 点「合并单元格」；拖动蓝色分隔线调整行高列宽；Esc 或「✓ 完成」退出');
}

function exitTableEdit() {
  if (!state.tableEdit) return;
  state.tableEdit = null; state.tableSel = null;
  renderCanvas(); renderProps(); updateTableBar();
}

function updateTableBar() {
  const bar = $('#table-bar');
  const editing = !!state.tableEdit;
  bar.classList.toggle('hidden', !editing);
  if (!editing) return;
  const el = findEl(state.tableEdit);
  const sel = tblSelRect(state.tableSel);
  const n = sel ? (sel.r2 - sel.r1 + 1) * (sel.c2 - sel.c1 + 1) : 0;
  $('#tbl-sel-info').textContent = el
    ? `表格 ${el.rows}×${el.cols} · 选中 ${n} 个单元格${sel ? `（${sel.r1 + 1},${sel.c1 + 1} → ${sel.r2 + 1},${sel.c2 + 1}）` : ''}`
    : '已退出表格编辑';
  const canMerge = n > 1;
  $('#tbl-merge').disabled = !canMerge;
}

/** 拖拽框选单元格 */
function startTableCellSelect(el, r, c) {
  state.tableSel = { r1: r, c1: c, r2: r, c2: c };
  renderCanvas(); updateTableBar();

  function onMove(ev) {
    const rect = canvas.getBoundingClientRect();
    const mmx = (ev.clientX - rect.left) / (MM * state.zoom) - el.x;
    const mmy = (ev.clientY - rect.top) / (MM * state.zoom) - el.y;
    // 由坐标反推行列（支持拖出表格范围时夹紧）
    let cc = 0, acc = 0;
    for (let i = 0; i < el.cols; i++) { acc += el.colWidths[i]; if (mmx <= acc) { cc = i; break; } cc = i; }
    let rr = 0, acr = 0;
    for (let i = 0; i < el.rows; i++) { acr += el.rowHeights[i]; if (mmy <= acr) { rr = i; break; } rr = i; }
    state.tableSel.r2 = rr; state.tableSel.c2 = cc;
    renderCanvas(); updateTableBar();
  }
  function onUp() {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    saveLocal();
  }
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

/**
 * 拖拽行/列分隔线，自由调整列宽/行高（与相邻行列此消彼长，保持总尺寸）
 * 关键点：以「按下的鼠标坐标」+「按下时的行列尺寸快照」为基准计算绝对位移。
 * 不能用被拖拽元素的 getBoundingClientRect()——按下时会先 renderCanvas() 重建 DOM，
 * 原节点已脱离文档，rect 全为 0，会导致分隔线瞬间跳到鼠标绝对坐标处（即“跳动”）。
 */
function startTableSeparatorDrag(el, kind, i, mx0, my0) {
  const key = kind === 'col' ? 'colWidths' : 'rowHeights';
  const arr0 = (el[key] || []).slice();      // 尺寸快照，全程以此为准，避免误差累积
  const min = kind === 'col' ? 2 : 1;        // 最小列宽 2mm / 行高 1mm
  const K = MM * state.zoom;
  const r2 = v => Math.round(v * 100) / 100;

  function onMove(ev) {
    const d = (kind === 'col' ? ev.clientX - mx0 : ev.clientY - my0) / K;
    const a = arr0[i - 1], b = arr0[i];
    if (a === undefined) return;
    if (b === undefined) {                   // 末列/末行：拖动直接改变表格总尺寸
      el[key][i - 1] = r2(Math.max(min, a + d));
    } else {                                 // 有相邻行列：此消彼长，两者之和保持不变
      const total = a + b;
      const na = Math.min(Math.max(min, a + d), total - min);
      el[key][i - 1] = r2(na);
      el[key][i] = r2(total - na);
    }
    tblSyncSize(el);
    renderCanvas(); updateTableBar();
  }
  function onUp() {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    renderProps();                           // 拖拽结束再刷新属性面板，拖拽过程更跟手
    saveLocal();
  }
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

/* ---------- 表格：合并 / 拆分 / 行列增删 ---------- */
function mergeSelectedCells() {
  const el = findEl(state.tableEdit);
  const sel = tblSelRect(state.tableSel);
  if (!el || el.type !== 'table' || !sel) return;
  if (sel.r1 === sel.r2 && sel.c1 === sel.c2) { notify('请先拖拽框选多个单元格再合并', 'warn'); return; }

  // 记录左上角块的原有内容（锚点归一到合并块左上角），丢弃被合并区域内的其它内容
  const anchor = tblBlockAt(el, sel.r1, sel.c1);
  const keep = { ...tblCellData(el, { r: anchor.r, c: anchor.c, rs: 1, cs: 1 }) };
  el.merges = (el.merges || []).filter(m => !(m.r >= sel.r1 && m.c >= sel.c1
    && m.r + m.rs - 1 <= sel.r2 && m.c + m.cs - 1 <= sel.c2));
  el.merges.push({ r: sel.r1, c: sel.c1, rs: sel.r2 - sel.r1 + 1, cs: sel.c2 - sel.c1 + 1 });
  tblClampMerges(el);

  // 合并后默认上下左右居中（按需求）
  el.cells[sel.r1 + ',' + sel.c1] = { ...keep, align: 'center', valign: 'middle' };
  state.tableSel = { r1: sel.r1, c1: sel.c1, r2: sel.r1, c2: sel.c1 };
  renderCanvas(); renderProps(); updateTableBar(); saveLocal();
  notify(`已合并 ${sel.r2 - sel.r1 + 1}×${sel.c2 - sel.c1 + 1} 个单元格（内容默认上下左右居中）`, 'success', 3000);
}

function splitSelectedCells() {
  const el = findEl(state.tableEdit);
  const sel = tblSelRect(state.tableSel);
  if (!el || !sel) return;
  const before = (el.merges || []).length;
  el.merges = (el.merges || []).filter(m => !(m.r <= sel.r2 && m.r + m.rs - 1 >= sel.r1
    && m.c <= sel.c2 && m.c + m.cs - 1 >= sel.c1));
  renderCanvas(); renderProps(); updateTableBar(); saveLocal();
  notify(before === el.merges.length ? '选区内没有合并单元格' : '已拆分合并单元格',
    before === el.merges.length ? 'warn' : 'success', 2600);
}

function tableInsertRow(after) {
  const el = findEl(state.tableEdit);
  const sel = tblSelRect(state.tableSel);
  if (!el || !sel) return;
  const at = after ? sel.r2 + 1 : sel.r1;
  const h = Math.round((el.rowHeights[sel.r1] || 6) * 100) / 100;
  el.rowHeights.splice(at, 0, h);
  // 行列数变化后：既要把插入点之后的合并块整体下移，也要同步迁移单元格数据
  (el.merges || []).forEach(m => { if (m.r >= at) m.r += 1; });
  tblRemapCells(el, (r, c) => [r >= at ? r + 1 : r, c]);
  el.rows += 1;
  tblSyncSize(el); tblClampMerges(el);
  state.tableSel = { r1: at, c1: sel.c1, r2: at, c2: sel.c2 };
  renderCanvas(); renderProps(); updateTableBar(); saveLocal();
}

function tableInsertCol(after) {
  const el = findEl(state.tableEdit);
  const sel = tblSelRect(state.tableSel);
  if (!el || !sel) return;
  const at = after ? sel.c2 + 1 : sel.c1;
  const w = Math.round((el.colWidths[sel.c1] || 20) * 100) / 100;
  el.colWidths.splice(at, 0, w);
  (el.merges || []).forEach(m => { if (m.c >= at) m.c += 1; });
  tblRemapCells(el, (r, c) => [r, c >= at ? c + 1 : c]);
  el.cols += 1;
  tblSyncSize(el); tblClampMerges(el);
  state.tableSel = { r1: sel.r1, c1: at, r2: sel.r2, c2: at };
  renderCanvas(); renderProps(); updateTableBar(); saveLocal();
}

function tableDeleteRow() {
  const el = findEl(state.tableEdit);
  const sel = tblSelRect(state.tableSel);
  if (!el || !sel) return;
  if (el.rows <= 1) { notify('至少保留 1 行', 'warn'); return; }
  const n = sel.r2 - sel.r1 + 1;
  el.rowHeights.splice(sel.r1, n);
  el.merges = (el.merges || []).filter(m => !(m.r >= sel.r1 && m.r + m.rs - 1 <= sel.r2));
  (el.merges || []).forEach(m => { if (m.r > sel.r2) m.r -= n; });
  tblRemapCells(el, (r, c) => (r >= sel.r1 && r <= sel.r2) ? null : [r > sel.r2 ? r - n : r, c]);
  el.rows -= n;
  tblSyncSize(el); tblClampMerges(el);
  state.tableSel = { r1: Math.min(sel.r1, el.rows - 1), c1: sel.c1, r2: Math.min(sel.r1, el.rows - 1), c2: sel.c2 };
  renderCanvas(); renderProps(); updateTableBar(); saveLocal();
}

function tableDeleteCol() {
  const el = findEl(state.tableEdit);
  const sel = tblSelRect(state.tableSel);
  if (!el || !sel) return;
  if (el.cols <= 1) { notify('至少保留 1 列', 'warn'); return; }
  const n = sel.c2 - sel.c1 + 1;
  el.colWidths.splice(sel.c1, n);
  el.merges = (el.merges || []).filter(m => !(m.c >= sel.c1 && m.c + m.cs - 1 <= sel.c2));
  (el.merges || []).forEach(m => { if (m.c > sel.c2) m.c -= n; });
  tblRemapCells(el, (r, c) => (c >= sel.c1 && c <= sel.c2) ? null : [r, c > sel.c2 ? c - n : c]);
  el.cols -= n;
  tblSyncSize(el); tblClampMerges(el);
  state.tableSel = { r1: sel.r1, c1: Math.min(sel.c1, el.cols - 1), r2: sel.r2, c2: Math.min(sel.c1, el.cols - 1) };
  renderCanvas(); renderProps(); updateTableBar(); saveLocal();
}

$('#tbl-merge').addEventListener('click', mergeSelectedCells);
$('#tbl-split').addEventListener('click', splitSelectedCells);
$('#tbl-addrow').addEventListener('click', () => tableInsertRow(true));
$('#tbl-addcol').addEventListener('click', () => tableInsertCol(true));
$('#tbl-delrow').addEventListener('click', tableDeleteRow);
$('#tbl-delcol').addEventListener('click', tableDeleteCol);
$('#tbl-done').addEventListener('click', exitTableEdit);

function select(id) {
  state.selectedId = id;
  // 注意：切换到其他元素时**不退出**表格编辑模式，
  // 便于点击别处后再回到表格继续框选/调整（退出请按 Esc 或浮动条的「✓ 完成」）
  renderCanvas(); renderProps();
  if (id === null) closeQuickEdit();
}

/* ============================================================
 * 双击快捷编辑器：改文字 / 绑定 Excel 字段 / 调字体字号
 * ============================================================ */
const TYPE_NAME = { text: '文本', barcode: '条形码', qrcode: '二维码', rect: '矩形框', line: '直线', table: '表格' };
let qeEl = null;          // 当前编辑的元素
let qeCellKey = null;     // 非空表示正在编辑表格中的某个单元格（键为 "r,c"）

/** 当前编辑目标（元素本身，或表格单元格数据） */
function qeTarget() {
  if (!qeEl) return null;
  if (qeCellKey) return tblCellData(qeEl, { r: +qeCellKey.split(',')[0], c: +qeCellKey.split(',')[1], rs: 1, cs: 1 });
  return qeEl;
}
/** 按当前目标重绘（单元格改动需重绘所属表格） */
function qeApply() { renderCanvas(); renderProps(); saveLocal(); }

canvas.addEventListener('dblclick', e => {
  const elDiv = e.target.closest('.el');
  if (!elDiv) return;
  const el = findEl(elDiv.dataset.id);
  if (!el) return;
  if (el.type === 'table') {
    if (state.tableEdit !== el.id) { enterTableEdit(el.id); return; }
    const cell = e.target.closest('.tbl-cell');
    if (cell) {
      const b = tblBlockAt(el, +cell.dataset.r, +cell.dataset.c);
      state.tableSel = { r1: b.r, c1: b.c, r2: b.r, c2: b.c };
      renderCanvas(); updateTableBar();
      openQuickEdit(el, e.clientX, e.clientY, b.r + ',' + b.c);
    } else {
      openQuickEdit(el, e.clientX, e.clientY);   // 双击表格边框 → 编辑表格属性
    }
    return;
  }
  openQuickEdit(el, e.clientX, e.clientY);
});

function openQuickEdit(el, mx, my, cellKey, keepPos) {
  qeEl = el;
  qeCellKey = cellKey || null;
  const t = qeTarget();
  if (!t) return;
  if (state.selectedId !== el.id) { state.selectedId = el.id; renderCanvas(); renderProps(); }

  const isCell = !!qeCellKey;
  const tableOnly = el.type === 'table' && !isCell;   // 表格整体：不能绑定字段/设字体，引导按单元格绑定
  $('#quick-edit').classList.toggle('table-only', tableOnly);
  if (tableOnly) {
    const b = state.tableSel ? tblBlockAt(el, tblSelRect(state.tableSel).r1, tblSelRect(state.tableSel).c1) : null;
    $('#qe-table-hint').innerHTML = `表格是不能整体绑定字段的，<b>绑定与字体都属于单元格</b>。<br>`
      + `请<b>双击表格里的具体单元格</b>（第 N 行 / 第 M 列），在弹出的窗口里点字段名完成绑定。`
      + (b ? `<br>当前选中：第 ${b.r + 1} 行 / 第 ${b.c + 1} 列` : '');
  }
  $('#qe-title').textContent = isCell
    ? `编辑单元格（第 ${+cellKey.split(',')[0] + 1} 行 / 第 ${+cellKey.split(',')[1] + 1} 列）`
    : '编辑元素 · ' + (TYPE_NAME[el.type] || el.type);

  // 内容来源（静态文字 / 当前日期时间）
  const dynSel = $('#qe-dynamic');
  dynSel.value = t.dynamic || '';
  $('#qe-datetime-row').classList.toggle('hidden', t.dynamic !== 'datetime');
  $('#qe-prefix-row').classList.toggle('hidden', isCell);
  $('#qe-valign-row').classList.toggle('hidden', !isCell);
  if (!dynSel.dataset.inited) {
    dynSel.innerHTML = '<option value="">静态文字</option><option value="datetime">当前日期时间</option>';
    const df = $('#qe-dateformat');
    df.innerHTML = DATETIME_FORMATS.map(f => `<option>${escHtml(f)}</option>`).join('');
    dynSel.dataset.inited = '1';
  }
  $('#qe-dateformat').value = t.dateFormat || DATETIME_FORMATS[0];

  $('#qe-text').value = t.text || '';
  $('#qe-prefix').value = t.prefix || '';
  $('#qe-binding').value = t.binding || '未绑定';
  $('#qe-text').placeholder = (el.type === 'qrcode' || el.type === 'barcode')
    ? '未绑定字段时使用的固定内容' : '直接输入固定文字内容';

  // 字段快捷绑定（来自下方 Excel 区域识别到的列名）
  const fw = $('#qe-fields');
  if (!state.headers.length) {
    fw.innerHTML = '<span class="muted" style="font-size:12px">暂无字段：请先「⬇ 下载Excel模板」填写后导入，'
      + '或在属性面板选择绑定列</span>';
  } else {
    fw.innerHTML = state.headers
      .map(h => `<span class="chip${t.binding === h ? ' active' : ''}" data-f="${escHtml(h)}">${escHtml(h)}</span>`)
      .join('');
    fw.querySelectorAll('.chip').forEach(c => c.addEventListener('click', () => {
      const tg = qeTarget();
      tg.binding = c.dataset.f;
      tg.dynamic = '';
      $('#qe-dynamic').value = '';
      $('#qe-datetime-row').classList.add('hidden');
      $('#qe-binding').value = c.dataset.f;
      fw.querySelectorAll('.chip').forEach(x => x.classList.toggle('active', x === c));
      qeApply();
    }));
  }

  // 字体与字号
  const fontSel = $('#qe-font');
  if (!fontSel.options.length) fontSel.innerHTML = FONT_LIST.map(f => `<option>${f}</option>`).join('');
  const fam = t.fontFamily || DEFAULT_FONT;
  if (![...fontSel.options].some(o => o.value === fam)) fontSel.insertAdjacentHTML('beforeend', `<option>${escHtml(fam)}</option>`);
  fontSel.value = fam;
  $('#qe-size').value = t.fontSize || (isCell ? 9 : 10);
  // 自适应字号：表格单元格、文本框、条码（条码下方文字）都有
  const canFit = isCell || el.type === 'text' || el.type === 'barcode';
  $('#qe-autofit-row').classList.toggle('hidden', !canFit);
  $('#qe-autofit').checked = t.autoFit !== false;
  $('#qe-bold').checked = !!t.bold;
  $('#qe-align').value = t.align || (isCell ? 'center' : 'left');
  $('#qe-valign').value = t.valign || 'middle';

  const qe = $('#quick-edit');
  qe.classList.remove('hidden');
  if (!keepPos) positionQuickEdit(mx, my);   // keepPos：只切换编辑目标，弹窗位置保持不变
}

function positionQuickEdit(mx, my) {
  const qe = $('#quick-edit');
  const w = qe.offsetWidth || 300, h = qe.offsetHeight || 380;
  let left = (mx ?? window.innerWidth / 2) + 16;
  let top = (my ?? 120) - 20;
  if (left + w > window.innerWidth - 10) left = window.innerWidth - w - 10;
  if (top + h > window.innerHeight - 10) top = Math.max(10, window.innerHeight - h - 10);
  qe.style.left = Math.max(10, left) + 'px';
  qe.style.top = Math.max(10, top) + 'px';
}

function closeQuickEdit() { qeEl = null; qeCellKey = null; $('#quick-edit').classList.add('hidden'); }

/* 快捷编辑器事件（只绑定一次，作用于 qeTarget()） */
$('#qe-close').addEventListener('click', closeQuickEdit);
$('#qe-ok').addEventListener('click', closeQuickEdit);
$('#qe-del').addEventListener('click', () => {
  if (qeCellKey) {                       // 单元格：清空内容而非删除元素
    const t = qeTarget();
    t.text = ''; t.binding = ''; t.dynamic = '';
    $('#qe-text').value = ''; $('#qe-binding').value = '未绑定'; $('#qe-dynamic').value = '';
    qeApply();
    return;
  }
  deleteSelected(); closeQuickEdit();
});
$('#qe-del').textContent = '删除元素';
$('#qe-unbind').addEventListener('click', () => {
  const t = qeTarget();
  if (!t) return;
  t.binding = ''; t.dynamic = '';
  $('#qe-binding').value = '未绑定';
  $('#qe-dynamic').value = '';
  $('#qe-datetime-row').classList.add('hidden');
  $('#qe-fields').querySelectorAll('.chip').forEach(x => x.classList.remove('active'));
  qeApply();
});
$('#qe-dynamic').addEventListener('change', e => {
  const t = qeTarget();
  if (!t) return;
  t.dynamic = e.target.value;
  $('#qe-datetime-row').classList.toggle('hidden', t.dynamic !== 'datetime');
  if (t.dynamic === 'datetime') infoBar('该文本框将自动显示当前日期时间：' + formatDateTime(t.dateFormat));
  qeApply();
});
$('#qe-dateformat').addEventListener('change', e => {
  const t = qeTarget();
  if (!t) return;
  t.dateFormat = e.target.value;
  if (t.dynamic !== 'datetime') { t.dynamic = 'datetime'; $('#qe-dynamic').value = 'datetime'; $('#qe-datetime-row').classList.remove('hidden'); }
  infoBar('时间格式已设为：' + e.target.value + '（当前显示 ' + formatDateTime(e.target.value) + '）');
  qeApply();
});
$('#qe-text').addEventListener('input', e => { const t = qeTarget(); if (t) { t.text = e.target.value; qeApply(); } });
$('#qe-prefix').addEventListener('input', e => { const t = qeTarget(); if (t) { t.prefix = e.target.value; qeApply(); } });
$('#qe-font').addEventListener('change', e => { const t = qeTarget(); if (t) { t.fontFamily = e.target.value; qeApply(); } });
$('#qe-size').addEventListener('input', e => { const t = qeTarget(); if (t) { t.fontSize = parseFloat(e.target.value) || 10; qeApply(); } });
$('#qe-autofit').addEventListener('change', e => {
  const t = qeTarget(); if (!t) return;
  t.autoFit = e.target.checked; qeApply();
});
$('#qe-bold').addEventListener('change', e => { const t = qeTarget(); if (t) { t.bold = e.target.checked; qeApply(); } });
$('#qe-align').addEventListener('change', e => { const t = qeTarget(); if (t) { t.align = e.target.value; qeApply(); } });
$('#qe-valign').addEventListener('change', e => { const t = qeTarget(); if (t) { t.valign = e.target.value; qeApply(); } });

/* ============================================================
 * 属性面板
 * ============================================================ */
function renderProps() {
  const box = $('#props-body');
  const el = findEl(state.selectedId);
  if (!el) { box.innerHTML = '<p class="muted">未选中元素</p>'; return; }

  const num = (key, label, step = 1) =>
    `<div class="prop-row"><label>${label}</label><input type="number" step="${step}" data-k="${key}" value="${el[key]}"></div>`;
  const text = (key, label, list = '') =>
    `<div class="prop-row"><label>${label}</label><input type="text" data-k="${key}" value="${escHtml(el[key] || '')}" ${list ? `list="${list}"` : ''}></div>`;
  /* 绑定列：自动列出导入 Excel 的表头字段（下拉选择，避免手工输入错别字） */
  const bindRow = (e2) => {
    const cur = e2.binding || '';
    const hs = state.headers || [];
    if (!hs.length) {
      return `<div class="prop-row"><label>绑定列</label>
        <select data-k="binding" disabled><option>（请先导入 Excel）</option></select></div>`;
    }
    const opts = ['<option value="">（不绑定）</option>']
      .concat(hs.map(h => `<option value="${escHtml(h)}"${h === cur ? ' selected' : ''}>${escHtml(h)}</option>`));
    if (cur && !hs.includes(cur)) {
      opts.push(`<option value="${escHtml(cur)}" selected>${escHtml(cur)}（不在当前数据中）</option>`);
    }
    return `<div class="prop-row"><label>绑定列</label><select data-k="binding" title="选择该元素显示 Excel 的哪一列数据">${opts.join('')}</select></div>`;
  };
  const fontRow = (e2) => {
    const fam = e2.fontFamily || DEFAULT_FONT;
    const opts = FONT_LIST.includes(fam) ? FONT_LIST : [...FONT_LIST, fam];
    return `<div class="prop-row"><label>字体</label><select data-k="fontFamily">${
      opts.map(f => `<option ${f === fam ? 'selected' : ''}>${escHtml(f)}</option>`).join('')}</select></div>`;
  };
  /* 内容来源 / 时间格式（动态内容） */
  const dynRow = (o, attr = 'data-k') => {
    const cur = o.dynamic || '';
    return `<div class="prop-row"><label>内容来源</label><select ${attr}="dynamic">
      <option value=""${cur === '' ? ' selected' : ''}>静态文字 / Excel字段</option>
      <option value="datetime"${cur === 'datetime' ? ' selected' : ''}>当前日期时间</option></select></div>`;
  };
  const dateFmtRow = (o, attr = 'data-k') => {
    if (o.dynamic !== 'datetime') return '';
    const cur = o.dateFormat || DATETIME_FORMATS[0];
    const opts = DATETIME_FORMATS.includes(cur) ? DATETIME_FORMATS : [...DATETIME_FORMATS, cur];
    return `<div class="prop-row"><label>时间格式</label><select ${attr}="dateFormat">${
      opts.map(f => `<option value="${escHtml(f)}"${f === cur ? ' selected' : ''}>${escHtml(f)}（${escHtml(formatDateTime(f))}）</option>`).join('')}</select></div>`;
  };

  let html = `<div class="prop-section">位置与大小 (mm)</div>`
    + num('x', 'X', 0.1) + num('y', 'Y', 0.1) + num('w', '宽', 0.1) + num('h', '高', 0.1);

  if (el.type === 'text') {
    html += `<div class="prop-section">文本</div>`
      + dynRow(el) + dateFmtRow(el)
      + text('prefix', '前缀') + bindRow(el) + text('text', '静态内容')
      + fontRow(el) + num('fontSize', '字号(pt)', 0.5)
      + `<div class="prop-row"><label>自适应字号</label><input type="checkbox" data-k="autoFit"
          ${el.autoFit !== false ? 'checked' : ''}
          title="勾选后文字不换行：内容超出元素宽度时自动缩小字号，保证一行显示完整"></div>`
      + `<p class="muted" style="font-size:11px;margin:-2px 0 6px">勾选「自适应字号」：内容始终一行，超出宽度时自动缩小字号（最小 3pt）；取消勾选则按原字号换行显示</p>`
      + `<div class="prop-row"><label>加粗</label><input type="checkbox" data-k="bold" ${el.bold ? 'checked' : ''}></div>`
      + `<div class="prop-row"><label>对齐</label><select data-k="align">
          <option ${el.align === 'left' ? 'selected' : ''} value="left">左</option>
          <option ${el.align === 'center' ? 'selected' : ''} value="center">中</option>
          <option ${el.align === 'right' ? 'selected' : ''} value="right">右</option></select></div>`
      + `<p class="muted" style="font-size:11px;margin-top:6px">提示：在画布上双击该元素可快速编辑并绑定 Excel 字段；`
      + `「当前日期时间」会在预览/打印时自动取当时时间</p>`;
  } else if (el.type === 'barcode') {
    html += `<div class="prop-section">条形码</div>`
      + bindRow(el) + text('text', '静态内容')
      + `<div class="prop-row"><label>码制</label><select data-k="barcodeFormat">
          ${['CODE128', 'CODE39', 'EAN13', 'EAN8', 'UPC', 'ITF14', 'ITF', 'MSI', 'pharmacode', 'codabar']
            .map(f => `<option ${el.barcodeFormat === f ? 'selected' : ''}>${f}</option>`).join('')}</select></div>`
      + num('barWidth', '线宽(px)', 1)
      + `<div class="prop-row"><label>显示文字</label><input type="checkbox" data-k="showText" ${el.showText ? 'checked' : ''}></div>`
      + fontRow(el) + num('fontSize', '字号(pt)', 0.5)
      + `<div class="prop-row"><label>自适应字号</label><input type="checkbox" data-k="autoFit"
          ${el.autoFit !== false ? 'checked' : ''}
          title="勾选后条码下方的文字在放不下时自动缩小字号"></div>`;
  } else if (el.type === 'qrcode') {
    html += `<div class="prop-section">二维码</div>`
      + bindRow(el) + text('text', '静态内容')
      + `<div class="prop-row"><label>纠错</label><select data-k="correctLevel">
          ${['L', 'M', 'Q', 'H'].map(l => `<option ${el.correctLevel === l ? 'selected' : ''}>${l}</option>`).join('')}</select></div>`;
  } else if (el.type === 'rect') {
    html += `<div class="prop-section">矩形</div>` + num('borderWidth', '边框粗(mm)', 0.1);
  } else if (el.type === 'table') {
    const sel = tblSelRect(state.tableSel);
    const editing = state.tableEdit === el.id;
    // 选区锚点归一到「真实单元格」：若左上角落在合并块内部，则归到该块左上角，
    // 否则会出现“读取一个不存在的单元格 → 显示未绑定”的假象。
    const probe = sel ? tblBlockAt(el, sel.r1, sel.c1) : null;
    const cellPos = probe ? { r: probe.r, c: probe.c } : (sel ? { r: sel.r1, c: sel.c1 } : null);
    const oneCell = sel && sel.r1 === sel.r2 && sel.c1 === sel.c2;
    const oneBlock = probe && sel && probe.r === sel.r1 && probe.c === sel.c1
      && probe.rs === sel.r2 - sel.r1 + 1 && probe.cs === sel.c2 - sel.c1 + 1;
    const single = !!(editing && cellPos && (oneCell || oneBlock));
    html += `<div class="prop-section">表格</div>`
      + `<div class="prop-row"><label>行列数</label><span>${el.rows} 行 × ${el.cols} 列</span></div>`
      + num('borderWidth', '线宽(mm)', 0.1)
      + `<div class="prop-row"><label>合并块</label><span>${(el.merges || []).length} 处</span></div>`
      + `<p class="muted" style="font-size:11px;margin-top:6px">双击表格进入编辑模式：`
      + `拖拽框选单元格 → 合并；拖动蓝色分隔线调整行高/列宽（拖动时与相邻行列此消彼长）；Esc 退出</p>`;

    if (single) {
      const cd = tblCellData(el, { r: cellPos.r, c: cellPos.c, rs: 1, cs: 1 });
      const bindOpts = ['<option value="">（不绑定）</option>']
        .concat((state.headers || []).map(h => `<option value="${escHtml(h)}"${cd.binding === h ? ' selected' : ''}>${escHtml(h)}</option>`))
        .concat(cd.binding && !(state.headers || []).includes(cd.binding)
          ? [`<option value="${escHtml(cd.binding)}" selected>${escHtml(cd.binding)}（不在当前数据中）</option>`] : []);
      const merged = probe && (probe.rs > 1 || probe.cs > 1);
      html += `<div class="prop-section">选中单元格（第 ${cellPos.r + 1} 行 / 第 ${cellPos.c + 1} 列${merged ? `，合并 ${probe.rs}×${probe.cs}` : ''}）</div>`
        + dynRow(cd, 'data-c') + dateFmtRow(cd, 'data-c')
        + `<div class="prop-row"><label>静态内容</label><input type="text" data-c="text" value="${escHtml(cd.text || '')}"></div>`
        + `<div class="prop-row"><label>绑定列</label><select data-c="binding">${bindOpts.join('')}</select></div>`
        + `<div class="prop-row"><label>字号(pt)</label><input type="number" step="0.5" data-c="fontSize" value="${cd.fontSize || 9}"></div>`
        + `<div class="prop-row"><label>自适应字号</label><input type="checkbox" data-c="autoFit"
            ${cd.autoFit !== false ? 'checked' : ''}
            title="勾选后文字不换行：字数多到一行放不下时自动缩小字号，保证内容一行显示完整"></div>`
        + `<p class="muted" style="font-size:11px;margin:-2px 0 6px">勾选「自适应字号」：内容始终一行；超出单元格宽度时自动缩小字号（最小 3pt），取消勾选则按原字号换行显示</p>`
        + `<div class="prop-row"><label>加粗</label><input type="checkbox" data-c="bold" ${cd.bold ? 'checked' : ''}></div>`
        + `<div class="prop-row"><label>水平对齐</label><select data-c="align">
            <option ${cd.align === 'left' ? 'selected' : ''} value="left">左</option>
            <option ${cd.align === 'center' ? 'selected' : ''} value="center">中</option>
            <option ${cd.align === 'right' ? 'selected' : ''} value="right">右</option></select></div>`
        + `<div class="prop-row"><label>垂直对齐</label><select data-c="valign">
            <option ${cd.valign === 'top' ? 'selected' : ''} value="top">上</option>
            <option ${cd.valign === 'middle' ? 'selected' : ''} value="middle">中</option>
            <option ${cd.valign === 'bottom' ? 'selected' : ''} value="bottom">下</option></select></div>`;
    } else if (editing) {
      html += `<p class="muted" style="font-size:11px">已选中 ${sel ? (sel.r2 - sel.r1 + 1) * (sel.c2 - sel.c1 + 1) : 0} 个单元格：`
        + `${sel && !(sel.r1 === sel.r2 && sel.c1 === sel.c2) ? '点上方浮动条的「合并单元格」；拖动框选范围可改变选区' : '拖拽框选多个单元格后可合并'}</p>`;
    } else {
      html += `<button id="props-table-edit" class="primary" style="width:100%">✎ 编辑表格（框选合并 / 调整行高列宽）</button>`
        + `<p class="muted" style="font-size:11px;margin-top:6px">也可直接双击表格进入编辑；编辑中点击其他地方不会退出，按 Esc 或浮动条「✓ 完成」结束</p>`;
    }
    // 处于编辑模式时始终提供“完成”按钮（含单选单元格的情况）
    if (editing) {
      html += `<button id="props-table-exit" style="width:100%;margin-top:6px">✓ 完成表格编辑</button>`;
    }
  }

  box.innerHTML = html + '<datalist id="headers-dl">' +
    state.headers.map(h => `<option value="${escHtml(h)}">`).join('') + '</datalist>';

  /** 通用赋值：容器为元素本身，或表格中被选中的单元格 */
  const assign = (obj, k, inp) => {
    if (inp.type === 'checkbox') obj[k] = inp.checked;
    else if (inp.type === 'number') obj[k] = parseFloat(inp.value) || 0;
    else obj[k] = inp.value;
  };

  // 表格编辑入口/出口按钮
  if ($('#props-table-edit')) $('#props-table-edit').addEventListener('click', () => enterTableEdit(el.id));
  if ($('#props-table-exit')) $('#props-table-exit').addEventListener('click', () => { exitTableEdit(); renderProps(); });

  box.querySelectorAll('[data-k]').forEach(inp => {
    const handler = () => {
      const k = inp.dataset.k;
      assign(el, k, inp);
      // 选择时间格式时自动切换为“当前日期时间”来源
      if (k === 'dateFormat') el.dynamic = 'datetime';
      renderCanvas(); saveLocal();
      if (k === 'dynamic' || k === 'dateFormat') renderProps();   // 格式行显示/隐藏需要重绘面板
      // 绑定列变化时，同步刷新快捷编辑器
      if (k === 'binding' && qeEl && qeEl.id === el.id) $('#qe-binding').value = el.binding || '未绑定';
    };
    inp.addEventListener('input', handler);
    inp.addEventListener('change', handler);
  });

  // 表格中被选中单元格的属性（锚点同上：落在合并块内时归到块左上角）
  const selRect = el.type === 'table' ? tblSelRect(state.tableSel) : null;
  const cellAnchor = selRect ? tblBlockAt(el, selRect.r1, selRect.c1) : null;
  if (el.type === 'table' && cellAnchor
    && (selRect.r1 === selRect.r2 && selRect.c1 === selRect.c2
      || (cellAnchor.r === selRect.r1 && cellAnchor.c === selRect.c1
        && cellAnchor.rs === selRect.r2 - selRect.r1 + 1 && cellAnchor.cs === selRect.c2 - selRect.c1 + 1))) {
    box.querySelectorAll('[data-c]').forEach(inp => {
      const handler = () => {
        const k = inp.dataset.c;
        const cd = tblCellData(el, { r: cellAnchor.r, c: cellAnchor.c, rs: 1, cs: 1 });
        assign(cd, k, inp);
        if (k === 'dateFormat') cd.dynamic = 'datetime';
        renderCanvas(); saveLocal();
        if (k === 'dynamic' || k === 'dateFormat') renderProps();
        // 若快捷编辑器正开着同一个单元格：只同步显示值，绝不切换编辑目标
        //（之前这里调用 openQuickEdit(el) 会丢掉单元格，导致编辑器变成“表格元素”、绑定框显示未绑定）
        if (qeEl && qeEl.id === el.id && qeCellKey === cellAnchor.r + ',' + cellAnchor.c) {
          $('#qe-binding').value = cd.binding || '未绑定';
          $('#qe-text').value = cd.text || '';
        }
      };
      inp.addEventListener('input', handler);
      inp.addEventListener('change', handler);
    });
  }
}

/* ============================================================
 * 工具栏事件
 * ============================================================ */
document.querySelectorAll('.pal[data-type]').forEach(btn => {
  btn.addEventListener('click', () => {
    // 新元素继承工具栏当前字体与字号
    const el = { ...DEFAULTS[btn.dataset.type], id: 'e' + (uid++), ...currentFont() };
    el.x = Math.max(2, (state.label.width - el.w) / 2);
    el.y = Math.max(2, (state.label.height - el.h) / 2);
    state.elements.push(el);
    select(el.id); saveLocal();
    openQuickEdit(el, window.innerWidth / 2, 120);   // 创建后直接进入编辑/绑定
  });
});

/* ---------- 插入：当前日期时间 ---------- */
$('#btn-insert-datetime').addEventListener('click', () => {
  const el = {
    ...DEFAULTS.text, id: 'e' + (uid++),
    dynamic: 'datetime', text: '', binding: '', prefix: '',
    fontSize: 12, w: 48, h: 8,
    fontFamily: $('#font-family').value || DEFAULT_FONT,
    dateFormat: DATETIME_FORMATS[0],
  };
  el.x = Math.max(2, Math.round((state.label.width - el.w) / 2));
  el.y = Math.max(2, Math.round((state.label.height - el.h) / 2));
  state.elements.push(el);
  select(el.id); saveLocal();
  notify(`已插入「当前日期时间」文本框，当前显示：${formatDateTime(el.dateFormat)}`
    + '（可双击修改格式，打印/预览时会取当时的时间）', 'success', 5500);
});

/* ---------- 插入：表格（可预选行列数） ---------- */
const TABLE_PRESETS = [[2, 2], [3, 3], [3, 4], [4, 4], [4, 6], [5, 4], [5, 8], [6, 4]];

function buildTablePresets() {
  $('#tbl-presets').innerHTML = TABLE_PRESETS
    .map(([r, c]) => `<button data-rc="${r},${c}">${r} × ${c}</button>`).join('');
  $('#tbl-presets').querySelectorAll('[data-rc]').forEach(btn => btn.addEventListener('click', () => {
    const [r, c] = btn.dataset.rc.split(',').map(Number);
    $('#tbl-rows').value = r; $('#tbl-cols').value = c;
    syncTableDialog();
  }));
}

function syncTableDialog() {
  const r = Math.min(30, Math.max(1, parseInt($('#tbl-rows').value, 10) || 1));
  const c = Math.min(20, Math.max(1, parseInt($('#tbl-cols').value, 10) || 1));
  $('#tbl-presets').querySelectorAll('[data-rc]').forEach(b => {
    b.classList.toggle('active', b.dataset.rc === `${r},${c}`);
  });
  // 预览：按实际行列等比绘制网格
  const maxW = 200, maxH = 110;
  const boxW = 160, boxH = Math.max(30, Math.min(maxH, Math.round(boxW * r / Math.max(c, 1) / 3)));
  const pv = $('#tbl-preview');
  pv.style.width = boxW + 'px';
  pv.style.height = boxH + 'px';
  let html = '';
  for (let i = 1; i < r; i++) html += `<div class="pv-h" style="top:${(boxH * i / r).toFixed(1)}px"></div>`;
  for (let i = 1; i < c; i++) html += `<div class="pv-v" style="left:${(boxW * i / c).toFixed(1)}px"></div>`;
  pv.innerHTML = html;
  const info = $('#tbl-rows').value + ' 行 × ' + $('#tbl-cols').value + ' 列';
  pv.title = info;
}

function openTableDialog() {
  $('#table-mask').classList.remove('hidden');
  $('#tbl-width').value = Math.min(state.label.width - 4, 90);
  syncTableDialog();
  setTimeout(() => $('#tbl-rows').select(), 30);
}

function insertTable() {
  const rows = Math.min(30, Math.max(1, parseInt($('#tbl-rows').value, 10) || 1));
  const cols = Math.min(20, Math.max(1, parseInt($('#tbl-cols').value, 10) || 1));
  const w = Math.max(10, Math.min(500, parseFloat($('#tbl-width').value) || 90));
  const h = Math.max(5, Math.min(500, parseFloat($('#tbl-height').value) || 24));
  const bw = Math.max(0.1, Math.min(3, parseFloat($('#tbl-border').value) || 0.3));
  const t = makeTable(rows, cols, w, h, bw);
  t.x = Math.max(0, Math.round((state.label.width - w) / 2 * 10) / 10);
  t.y = Math.max(0, Math.round((state.label.height - h) / 2 * 10) / 10);
  state.elements.push(t);
  $('#table-mask').classList.add('hidden');
  select(t.id); saveLocal();
  enterTableEdit(t.id);          // 插入后直接进入编辑模式，方便框选合并
  notify(`已插入 ${rows} × ${cols} 表格：可拖拽框选单元格合并、拖动蓝色分隔线调整行高列宽`, 'success', 6000);
}

$('#btn-insert-table').addEventListener('click', openTableDialog);
$('#table-ok').addEventListener('click', insertTable);
$('#table-cancel').addEventListener('click', () => $('#table-mask').classList.add('hidden'));
$('#table-close').addEventListener('click', () => $('#table-mask').classList.add('hidden'));
$('#table-mask').addEventListener('click', e => { if (e.target.id === 'table-mask') $('#table-mask').classList.add('hidden'); });
['#tbl-rows', '#tbl-cols'].forEach(s => $(s).addEventListener('input', syncTableDialog));
buildTablePresets();

$('#btn-del').addEventListener('click', deleteSelected);
$('#btn-copy').addEventListener('click', copySelected);
$('#btn-layer-up').addEventListener('click', () => layer(1));
$('#btn-layer-down').addEventListener('click', () => layer(-1));
/** 「上一步 / 下一步」（撤销、重做）：见文件后部的「撤销 / 重做」实现 */
$('#btn-undo').addEventListener('click', () => undo());
$('#btn-redo').addEventListener('click', () => redo());

function deleteSelected() {
  const i = state.elements.findIndex(e => e.id === state.selectedId);
  if (i < 0) return;
  const removed = state.elements[i];
  state.elements.splice(i, 1);
  if (removed && removed.id === state.tableEdit) {   // 删除表格时同步退出编辑模式
    state.tableEdit = null; state.tableSel = null; updateTableBar();
  }
  select(null); saveLocal();
}
function copySelected() {
  const el = findEl(state.selectedId);
  if (!el) return;
  const c = cloneElement(el);              // 深拷贝：复制表格不会与原件共享行列/合并/单元格数据
  c.x = snap(el.x + 3); c.y = snap(el.y + 3);
  state.elements.push(c); select(c.id); saveLocal();
  if (el.type === 'table') notify('已复制表格：副本的单元格数据与原件相互独立', 'success', 3200);
}
function layer(dir) {
  const i = state.elements.findIndex(e => e.id === state.selectedId);
  const j = i + dir;
  if (i < 0 || j < 0 || j >= state.elements.length) return;
  [state.elements[i], state.elements[j]] = [state.elements[j], state.elements[i]];
  renderCanvas(); saveLocal();
}

/* 键盘：撤销/重做、Del 删除、方向键微调 */
document.addEventListener('keydown', e => {
  if (['INPUT', 'SELECT', 'TEXTAREA'].includes(e.target.tagName)) return;   // 输入框内保留原生撤销
  const key = String(e.key || '').toLowerCase();
  if ((e.ctrlKey || e.metaKey) && key === 'z') {          // Ctrl+Z 撤销 / Ctrl+Shift+Z 重做
    e.preventDefault();
    e.shiftKey ? redo() : undo();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && key === 'y') {          // Ctrl+Y 重做
    e.preventDefault(); redo(); return;
  }
  if (e.key === 'Escape') {          // Esc 退出表格编辑模式
    if (state.tableEdit) { exitTableEdit(); e.preventDefault(); }
    return;
  }
  const el = findEl(state.selectedId);
  if (!el) return;
  // 表格编辑模式下，Delete 用于删除选中行（避免误删整个表格）
  if (state.tableEdit === el.id && (e.key === 'Delete' || e.key === 'Backspace')) {
    tableDeleteRow(); e.preventDefault(); return;
  }
  if (e.key === 'Delete' || e.key === 'Backspace') { deleteSelected(); return; }
  const step = e.shiftKey ? 0.2 : 1;
  const mv = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key];
  if (mv) { el.x = snap(el.x + mv[0]); el.y = snap(el.y + mv[1]); renderCanvas(); renderProps(); saveLocal(); e.preventDefault(); }
});

/* ============================================================
 * 标签纸尺寸设置：标准尺寸预设 + 自定义（mm / cm / 英寸）+ 预览
 * 内部统一以毫米（mm）存储，界面按所选单位换算显示
 * ============================================================ */
const MM_MIN = 10, MM_MAX = 500;
const UNIT_INFO = { mm: { label: 'mm', rate: 1 }, cm: { label: 'cm', rate: 10 }, in: { label: 'inch', rate: 25.4 } };

/** 标准标签纸尺寸预设（尺寸单位 mm） */
const PAPER_PRESETS = [
  {
    group: '常用条码 / 热敏标签',
    items: [[100, 150, '快递面单'], [100, 100, '方形大标签'], [100, 80, '物料标签'], [100, 70, '物料标签'],
      [100, 60, '物料标签'], [90, 60, ''], [80, 60, ''], [70, 50, ''], [60, 40, ''], [50, 30, '']],
  },
  {
    group: '仓储 / 物流面单',
    items: [[100, 180, '大面单'], [76, 130, '小面单'], [101.6, 152.4, '4×6 英寸'], [76.2, 101.6, '3×4 英寸'],
      [50.8, 76.2, '2×3 英寸'], [102, 76, '横向面单']],
  },
  {
    group: '线缆 / 资产 / 小标签',
    items: [[40, 30, ''], [30, 20, '设备标签'], [25, 15, ''], [20, 10, ''], [24, 12, '线缆'],[38, 12, '线缆']],
  },
  {
    group: '办公打印纸（整页打印）',
    items: [[210, 297, 'A4'], [148, 210, 'A5'], [176, 250, 'B5'], [216, 279, 'Letter'], [105, 148, 'A6']],
  },
];

let lastPaperUnit = 'mm';                  // 记录上一次使用的单位，用于切换时换算
const paperUnit = () => $('#paper-unit').value || 'mm';
const toMm = v => Math.round((Number(v) * UNIT_INFO[paperUnit()].rate) * 10) / 10;
const fromMm = mm => {
  const r = UNIT_INFO[paperUnit()].rate;
  return Math.round((mm / r) * 100) / 100;
};
const fmtMm = v => (Math.round(v * 10) / 10) + ' mm';

/** 绘制尺寸预览 + 预设高亮（不触碰输入框，供输入时实时预览使用） */
function renderPaperPreview(w, h) {
  const box = $('#paper-preview');
  const maxW = 210, maxH = 120;
  const scale = Math.min(maxW / w, maxH / h);
  box.style.width = Math.max(6, Math.round(w * scale)) + 'px';
  box.style.height = Math.max(6, Math.round(h * scale)) + 'px';
  const inchW = (w / 25.4).toFixed(2), inchH = (h / 25.4).toFixed(2);
  $('#paper-preview-text').textContent = `${w} × ${h} mm　|　${inchW} × ${inchH} inch　|　`
    + `${w >= h ? '横向' : '纵向'}　|　比例 ${(w / h).toFixed(2)}:1`;
  document.querySelectorAll('#paper-presets .pp-item').forEach(btn => {
    const [pw, ph] = btn.dataset.size.split(',').map(Number);
    btn.classList.toggle('active', Math.abs(pw - w) < 0.05 && Math.abs(ph - h) < 0.05);
  });
}

/** 完整同步：功能区徽标 + 输入框 + 预览（用于载入模板 / 尺寸变更后） */
function syncPaperUI() {
  const { width: w, height: h } = state.label;
  $('#paper-chip-text').textContent = `${Math.round(w * 10) / 10} × ${Math.round(h * 10) / 10} mm`;
  $('#lbl-w').value = fromMm(w);
  $('#lbl-h').value = fromMm(h);
  renderPaperPreview(w, h);
}

/* 对话框内的草稿尺寸：始终以毫米精确保存。
   输入框显示值会被四舍五入（如英寸 3.94），因此仅在用户真正编辑过输入框
   （paperInputDirty）时才回读输入框，其它情况一律以毫米草稿为准，
   避免反复切换单位造成精度累积误差。 */
let paperDraft = { w: 100, h: 60 };
let paperInputDirty = false;

/** 按草稿值刷新对话框输入框与预览（切换单位 / 选预设 / 互换时使用） */
function showPaperDraft() {
  $('#lbl-w').value = fromMm(paperDraft.w);
  $('#lbl-h').value = fromMm(paperDraft.h);
  paperInputDirty = false;
  renderPaperPreview(paperDraft.w, paperDraft.h);
  checkPaperRange();
}

/** 输入框 → 草稿（mm）；unit 可指定按哪个单位折算（默认当前单位） */
function draftFromInputs(unit) {
  const rate = UNIT_INFO[unit || paperUnit()].rate;
  const w = parseFloat($('#lbl-w').value) * rate;
  const h = parseFloat($('#lbl-h').value) * rate;
  if (Number.isFinite(w) && w > 0) paperDraft.w = Math.round(w * 100) / 100;
  if (Number.isFinite(h) && h > 0) paperDraft.h = Math.round(h * 100) / 100;
}

/** 范围校验：错误信息写回对话框；返回是否通过 */
function checkPaperRange(silent) {
  const err = $('#paper-error');
  const { w, h } = paperDraft;
  let msg = '';
  if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) msg = '请输入有效的宽和高';
  else if (w < MM_MIN || w > MM_MAX || h < MM_MIN || h > MM_MAX) {
    msg = `尺寸超出范围：宽高需在 ${MM_MIN} ~ ${MM_MAX} mm 之间（当前 ${Math.round(w * 10) / 10} × ${Math.round(h * 10) / 10} mm）`;
  }
  err.textContent = msg;
  err.classList.toggle('hidden', !msg || !!silent);
  return !msg;
}

/** 输入时的实时预览（不覆盖正在输入的内容） */
function previewFromInputs() {
  paperInputDirty = true;
  draftFromInputs();
  renderPaperPreview(paperDraft.w, paperDraft.h);
  checkPaperRange();
}

/** 应用标签纸尺寸 */
function applyPaperSize() {
  if (paperInputDirty) draftFromInputs();      // 仅在用户改动过输入框时回读
  if (!checkPaperRange()) return;              // 非法尺寸：提示并阻止关闭
  const w = Math.round(paperDraft.w * 10) / 10;
  const h = Math.round(paperDraft.h * 10) / 10;
  const out = state.elements.filter(el => el.x + el.w > w + 0.05 || el.y + el.h > h + 0.05);
  state.label.width = Math.min(MM_MAX, Math.max(MM_MIN, w));
  state.label.height = Math.min(MM_MAX, Math.max(MM_MIN, h));
  renderCanvas(); renderProps(); saveLocal(); syncPaperUI();
  const label = `${fmtMm(state.label.width)} × ${fmtMm(state.label.height)}`;
  if (out.length) {
    notify(`标签纸已设为 ${label}，其中 ${out.length} 个元素超出新范围（可拖动或缩小）`, 'warn', 6000);
  } else {
    notify(`标签纸尺寸已更新为 ${label}`, 'success', 3500);
  }
  $('#paper-mask').classList.add('hidden');
}

function buildPaperPresets() {
  $('#paper-presets').innerHTML = PAPER_PRESETS.map(g => `
    <div class="pp-group">
      <div class="pp-group-name">${escHtml(g.group)}</div>
      <div class="pp-items">${g.items.map(([w, h, note]) => `
        <button class="pp-item" data-size="${w},${h}" title="${w} × ${h} mm${note ? '（' + note + '）' : ''}">
          ${w} × ${h}<small>${note || h <= 150 ? 'mm' : ''}</small>
        </button>`).join('')}</div>
    </div>`).join('');

  document.querySelectorAll('#paper-presets .pp-item').forEach(btn => btn.addEventListener('click', () => {
    const [w, h] = btn.dataset.size.split(',').map(Number);
    paperDraft = { w, h };          // 草稿以 mm 精确保存预设值
    $('#paper-error').classList.add('hidden');
    showPaperDraft();
    infoBar(`已选中标准尺寸 ${w} × ${h} mm，点「应用到标签」生效`);
  }));
}

function openPaperDialog() {
  $('#paper-mask').classList.remove('hidden');
  $('#paper-error').classList.add('hidden');
  lastPaperUnit = paperUnit();                       // 单位保持上次选择（默认 mm）
  paperDraft = { w: state.label.width, h: state.label.height };
  showPaperDraft();
}

$('#btn-paper').addEventListener('click', openPaperDialog);
$('#paper-close').addEventListener('click', () => $('#paper-mask').classList.add('hidden'));
$('#paper-cancel').addEventListener('click', () => $('#paper-mask').classList.add('hidden'));
$('#paper-mask').addEventListener('click', e => { if (e.target.id === 'paper-mask') $('#paper-mask').classList.add('hidden'); });
$('#paper-apply').addEventListener('click', applyPaperSize);
$('#paper-mask').addEventListener('keydown', e => {
  if (e.key === 'Enter') { e.preventDefault(); applyPaperSize(); }
  else if (e.key === 'Escape') { e.preventDefault(); $('#paper-mask').classList.add('hidden'); }
});
$('#paper-swap').addEventListener('click', () => {
  draftFromInputs();
  const w = paperDraft.w;
  paperDraft.w = paperDraft.h;
  paperDraft.h = w;
  $('#paper-error').classList.add('hidden');
  showPaperDraft();
  infoBar('已交换宽高（纵向 / 横向），点「应用到标签」生效');
});
$('#paper-standard').addEventListener('click', () => {
  $('#paper-unit').value = 'mm';
  lastPaperUnit = 'mm';
  paperDraft = { w: 100, h: 60 };
  $('#paper-error').classList.add('hidden');
  showPaperDraft();
  infoBar('已复位为 100 × 60 mm，点「应用到标签」生效');
});
$('#paper-unit').addEventListener('change', () => {
  // 用户改过输入框 → 按旧单位折算；否则直接用毫米草稿精确换算显示
  if (paperInputDirty) draftFromInputs(lastPaperUnit);
  lastPaperUnit = paperUnit();
  $('#paper-error').classList.add('hidden');
  showPaperDraft();
});
[$('#lbl-w'), $('#lbl-h')].forEach(inp => {
  inp.addEventListener('input', previewFromInputs);
  inp.addEventListener('change', previewFromInputs);
});

/* 缩放 */
$('#zoom').addEventListener('change', e => {
  state.zoom = +e.target.value;
  state.zoomExplicit = true;            // 用户手动选择 → 草稿恢复时保留
  renderCanvas(); saveLocal();
  infoBar(`标签区显示比例：${Math.round(state.zoom * 100)}%（等比缩放，不影响打印尺寸）`);
});

/* 字体 / 字号：选中元素生效；未选中则应用到全部文本与条码元素 */
function currentFont() {
  return { fontFamily: $('#font-family').value, fontSize: parseFloat($('#font-size').value) || 10 };
}
function applyFont(key, value) {
  const sel = findEl(state.selectedId);
  const targets = sel ? [sel] : state.elements.filter(e => e.type === 'text' || e.type === 'barcode');
  if (!targets.length) { notify('画布上还没有文本/条码元素，请先创建', 'warn'); return; }
  targets.forEach(t => { t[key] = value; });
  renderCanvas(); renderProps(); saveLocal();
  if (qeEl) { $('#qe-font').value = qeEl.fontFamily || DEFAULT_FONT; $('#qe-size').value = qeEl.fontSize || 10; }
  infoBar(sel ? '已修改选中元素的' + (key === 'fontFamily' ? '字体' : '字号')
              : `已应用到全部 ${targets.length} 个文本/条码元素`);
}
$('#font-family').addEventListener('change', e => applyFont('fontFamily', e.target.value));
$('#font-size').addEventListener('input', e => {
  const v = parseFloat(e.target.value); if (v) applyFont('fontSize', v);
});
$('#font-inc').addEventListener('click', () => {
  const v = Math.min(72, (parseFloat($('#font-size').value) || 10) + 1);
  $('#font-size').value = v; applyFont('fontSize', v);
});
$('#font-dec').addEventListener('click', () => {
  const v = Math.max(4, (parseFloat($('#font-size').value) || 10) - 1);
  $('#font-size').value = v; applyFont('fontSize', v);
});

let infoTimer = null;
function infoBar(msg) {
  $('#import-info').textContent = msg;
  clearTimeout(infoTimer);
  infoTimer = setTimeout(() => { $('#import-info').textContent = ''; }, 3600);
}

/* ============================================================
 * 标准化 Excel：模板导出 / 导入 / 校验
 * ============================================================ */
const SAMPLE_ROWS = [
  ['W0000123456', '高耐磨涂层纸', '787×1092mm', 'A-2024', '500', '张', 'L2408-001', 'SN000001',
   'SUP001', '广东邦固化纤科技有限公司', '2026-08-24', '2027-08-23', 'CG20260824-01', 'RW20260824-001', '', '', '2'],
  ['W0000123457', '激光全息防伪膜', '650×900mm', 'B-1130', '120', '卷', 'L2408-002', 'SN000002',
   'SUP002', '湖北华工图像技术开发有限公司', '2026-08-25', '2027-02-24', 'CG20260824-02', 'RW20260824-002', '', '', '1'],
];

$('#btn-tpl').addEventListener('click', () => {
  const ws = XLSX.utils.aoa_to_sheet([STANDARD_COLUMNS, ...SAMPLE_ROWS]);
  ws['!cols'] = STANDARD_COLUMNS.map(c => ({ wch: Math.max(12, c.length * 2 + 4) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '标签数据');

  const notes = [
    ['标签打印数据模板 · 填写说明'],
    ['0. 本模板只是推荐格式，并非必须：软件支持导入任意 Excel/CSV —— 第一行（第一行有内容的行）的每一列会自动识别为字段名，往下每行数据对应一张标签。'],
    ['1. 请勿修改第 1 行（列名）的内容与顺序。若你的表格前面有标题行/说明行，导入时软件会自动跳过并给出提示，也可手动指定「表头行」。'],
    ['2. 从第 2 行开始，一行数据 = 一张标签；整行为空的行会被自动忽略。'],
    ['3. 日期请使用 YYYY-MM-DD 文本格式（如 2026-08-24），避免 Excel 日期序列号。'],
    ['4. 数量、序列号等请使用普通文本或数值，勿使用公式与合并单元格。'],
    ['5. 「二维码内容」「条码内容」可留空：导入时会自动生成二维码内容（物料编码;批号;生产日期）与条码内容（物料编码）。'],
    ['6. 需要额外字段时，可在标准列右侧自行追加列，导入后可同样绑定到标签元素。'],
    ['7. 「打印数量」列：填写该行需要连续打印的份数（留空或填 1 表示打印 1 份）；导入后也可在网页数据表格最后一列直接修改。'],
    ['8. 填写完成后保存为 .xlsx，回到网页点击「📥 导入Excel」选择该文件，系统会给出校验报告。'],
    ['9. 标准列清单（共 ' + STANDARD_COLUMNS.length + ' 列）：' + STANDARD_COLUMNS.join('、')],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(notes);
  ws2['!cols'] = [{ wch: 100 }];
  XLSX.utils.book_append_sheet(wb, ws2, '填写说明');

  XLSX.writeFile(wb, '标签打印数据模板.xlsx');
  notify('标准模板「标签打印数据模板.xlsx」已下载：含 16 列标准字段与填写说明，填写后点击「📥 导入 Excel」', 'success', 6000);
});

$('#btn-import').addEventListener('click', () => $('#excel-file').click());
$('#excel-file').addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try { startImport(new Uint8Array(ev.target.result), file.name); }
    catch (err) { notify('解析文件失败：' + err.message, 'error', 7000); }
    e.target.value = '';
  };
  reader.onerror = () => { notify('读取文件失败，请重新选择', 'error', 6000); e.target.value = ''; };
  reader.readAsArrayBuffer(file);
});

/* ============================================================
 * 通用 Excel / CSV 导入（兼容任意表格文件）
 * 规则：默认把「第一行（第一行有内容的行）」的每一列识别为字段名，
 *      其下每一行作为该字段的一组取值（一行 = 一张标签）。
 * 兼容：任意工作表（自动挑选数据表，可手动切换）、前置标题/说明/空行、
 *      合并单元格、重复列名、空白列名、整列全空、疑似合计行、
 *      GBK 编码 CSV，以及 .xlsx/.xlsm/.xlsb/.xls/.csv/.tsv/.txt
 * ============================================================ */
let importCtx = null;   // { wb, fileName, sheets[], sheet, headerRow, res, encoding, norm }

/** 文本字节解码：先按严格 UTF-8，失败再依次尝试 GB18030/GBK/Big5（用浏览器原生 TextDecoder，不依赖 SheetJS 码表） */
function decodeTextBytes(bytes) {
  try {
    return { text: new TextDecoder('utf-8', { fatal: true }).decode(bytes), encoding: 'UTF-8' };
  } catch (e) {
    for (const enc of ['gb18030', 'gbk', 'big5']) {
      try { return { text: new TextDecoder(enc).decode(bytes), encoding: enc.toUpperCase() }; } catch (e2) { /* 换下一种 */ }
    }
    return { text: new TextDecoder('utf-8').decode(bytes), encoding: 'UTF-8（容错）' };
  }
}

/** 读取文件为工作簿：文本类先做编码探测（UTF-8 / GB18030 / Big5），二进制直接解析 */
function readWorkbook(bytes, fileName) {
  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4B;      // xlsx/xlsm/xlsb（zip 容器）
  const isCfb = bytes[0] === 0xD0 && bytes[1] === 0xCF;      // xls（复合文档）
  if (isZip || isCfb) {
    return { wb: XLSX.read(bytes, { type: 'array', cellDates: false }), encoding: 'Excel 二进制' };
  }
  const bom = bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF;
  const { text, encoding } = decodeTextBytes(bom ? bytes.slice(3) : bytes);
  return { wb: XLSX.read(text, { type: 'string' }), encoding: bom ? 'UTF-8 (BOM)' : encoding };
}

/** 工作表“含数据量”评分：非空单元格越多、列越宽越可能是数据表 */
function sheetScore(ws) {
  if (!ws || !ws['!ref']) return -1;
  const range = XLSX.utils.decode_range(ws['!ref']);
  const rowN = Math.min(range.e.r - range.s.r + 1, 300);
  const colN = Math.min(range.e.c - range.s.c + 1, 60);
  let filled = 0;
  for (let r = 0; r < rowN; r++) {
    for (let c = 0; c < colN; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: range.s.r + r, c: range.s.c + c })];
      if (cell && String(cell.w !== undefined ? cell.w : cell.v).trim() !== '') filled++;
    }
  }
  return filled + Math.min(colN, 20) * 2;
}

/** 挑选最可能存放数据的工作表（同分取靠前者） */
function pickBestSheet(wb, names) {
  let best = names[0], bestScore = -Infinity;
  names.forEach(n => {
    const s = sheetScore(wb.Sheets[n]);
    if (s > bestScore) { bestScore = s; best = n; }
  });
  return best;
}

/**
 * 判断 Excel 单元格格式是否为日期/时间格式。
 * 先去掉引号文本段（"USD" 等）、方括号段（[Red]/[h]）与转义字符，再看剩余记号里是否含日期记号；
 * 比 SheetJS 的 SSF.is_date 更保守（不会把 "USD"#,##0 这类货币格式误判为日期）。
 */
function isDateFormat(z) {
  const s = String(z || '')
    .replace(/"[^"]*"/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/\\./g, '·')
    .replace(/_.|@/g, '');
  const tokens = s.toLowerCase().replace(/[^ymdhs]/g, '');
  return /y|d|h|s/.test(tokens) || /m{2,}/.test(tokens);
}

/** 单元格 → 显示文本：日期统一成 YYYY-MM-DD（不带时区问题），其它沿用 Excel 显示文本（保留 %、千分位、前导零） */
function cellToText(cell) {
  if (!cell) return '';
  // 注意：SheetJS 对数字单元格常省略 t 字段，故按「无 t 或 t==='n'」判断
  const numLike = (cell.t === 'n' || cell.t === undefined) && typeof cell.v === 'number';
  // 显示值本身是日期样式（如 8/24/26、2026-08-24、2026/8/24 09:05）也算日期；
  // 该正则不会把 100.5、46258.0005 这类普通小数误判为日期。
  const w = typeof cell.w === 'string' ? cell.w.trim() : '';
  const dateStyled = /^(?:\d{4}[-\/.]\d{1,2}[-\/.]\d{1,2}|\d{1,2}[-\/.]\d{1,2}[-\/.]\d{2,4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/.test(w);
  if (numLike && (dateStyled || (cell.z && isDateFormat(cell.z)))
    && typeof XLSX.SSF.parse_date_code === 'function') {
    const d = XLSX.SSF.parse_date_code(cell.v);
    if (d && d.y) {
      const p2 = n => String(n || 0).padStart(2, '0');
      const date = `${d.y}-${p2(d.m)}-${p2(d.d)}`;
      return (d.H || d.M) ? `${date} ${p2(d.H)}:${p2(d.M)}` : date;   // 秒级浮点残差不当作时间
    }
  }
  if (cell.t === 'd' && cell.v instanceof Date) {
    const p2 = n => String(n).padStart(2, '0');
    return `${cell.v.getFullYear()}-${p2(cell.v.getMonth() + 1)}-${p2(cell.v.getDate())}`;
  }
  if (cell.w !== undefined && cell.w !== null) return String(cell.w);
  return cell.v === undefined || cell.v === null ? '' : String(cell.v);
}

/** 工作表 → 二维数组（全部转字符串、保留空行位置、去掉尾部空行），并给出首行偏移 */
function sheetToAoa(ws) {
  if (!ws || !ws['!ref']) return { aoa: [], rowOffset: 0 };
  const range = XLSX.utils.decode_range(ws['!ref']);
  const aoa = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      row.push(cellToText(ws[XLSX.utils.encode_cell({ r, c })]));
    }
    aoa.push(row);
  }
  let last = aoa.length - 1;
  while (last >= 0 && !aoa[last].some(v => String(v).trim() !== '')) last--;
  return { aoa: aoa.slice(0, last + 1), rowOffset: range.s.r };
}

/** 字段名清洗：去首尾空白；中文表头去掉内部空白（“物料 编码”→“物料编码”）；超长截断 */
function cleanHeaderName(raw, colIdx) {
  let s = String(raw == null ? '' : raw).replace(/\u00a0/g, ' ').trim();
  if (!s) return '';
  if (/[\u4e00-\u9fff]/.test(s)) s = s.replace(/\s+/g, '');
  else s = s.replace(/\s+/g, ' ');
  return s.slice(0, 60) || `列${colIdx + 1}`;
}

/**
 * 识别表头行：默认取第一行；
 * 仅当后续某行“明显更像表头”（列更多、名称唯一、多为短文本，如前面是标题行/说明行）时才跳过。
 */
function detectHeaderRow(aoa) {
  const limit = Math.min(10, aoa.length);
  const widthOf = row => row.filter(v => String(v).trim() !== '').length;
  /** 该行的“占用列跨度”：最后一个有内容的列序号 + 1（空列名不影响跨度判定） */
  const spanOf = row => { let last = 0; (row || []).forEach((v, i) => { if (String(v).trim() !== '') last = i + 1; }); return last; };
  let maxW = 0, maxSpan = 0;
  for (let i = 0; i < limit; i++) {
    maxW = Math.max(maxW, widthOf(aoa[i] || []));
    maxSpan = Math.max(maxSpan, spanOf(aoa[i]));
  }
  const score = i => {
    const row = aoa[i] || [];
    const cells = row.map(v => String(v).trim()).filter(Boolean);
    if (cells.length < 2) return -99;                 // 标题行/单格行不当表头
    const uniq = new Set(cells).size / cells.length;
    const textRatio = cells.filter(v => !/^-?[\d.,%¥$ ]+$/.test(v)).length / cells.length;
    const avgLen = cells.reduce((s, v) => s + v.length, 0) / cells.length;
    // “像取值”的比例：纯数字、百分比、日期（如 2026/8/24）—— 表头里出现这些的可能很低
    const valueish = cells.filter(v => /^-?[\d.,%¥$ ]+$/.test(v)
      || /^\d{1,4}[\/\-.]\d{1,2}([\/\-.]\d{1,4})?$/.test(v)).length / cells.length;
    let below = 0;
    for (let j = i + 1; j < Math.min(i + 4, aoa.length); j++) below = Math.max(below, widthOf(aoa[j] || []));
    return cells.length * 2 + uniq * 4 + textRatio * 3
      + (spanOf(row) >= maxSpan ? 3 : 0)              // 表头通常横向铺满（含空列名）
      + (avgLen <= 10 ? 2 : avgLen >= 20 ? -2 : 0)
      + (below >= 1 ? 4 : 0) + (below >= cells.length * 0.6 ? 2 : 0)
      - valueish * 6                                  // 取值行明显扣分
      - i * 0.5;
  };
  let best = 0, bestScore = score(0);
  for (let i = 1; i < limit; i++) {
    const sc = score(i);
    if (sc > bestScore + 3) { best = i; bestScore = sc; }   // 需明显领先才跳过前面的行
  }
  return best;
}

/** 解析工作表：表头行 → 字段名（清洗/去重/命名）→ 数据行（跳过整行空行） */
function analyzeSheet(ws, headerRow) {
  const { aoa, rowOffset } = sheetToAoa(ws);
  if (!aoa.length) return { headers: [], rows: [], headerRow: 0, rowOffset: 0, emptySkipped: 0, notes: { droppedCols: [], summaryRows: 0, leadingSkipped: 0 } };
  let hr = (headerRow === null || headerRow === undefined) ? detectHeaderRow(aoa) : headerRow;
  hr = Math.max(0, Math.min(hr, aoa.length - 1));
  const head = aoa[hr] || [];
  const body = aoa.slice(hr + 1);

  let width = head.length;
  body.forEach(r => { if (r.length > width) width = r.length; });

  const used = new Map();
  const names = [];          // 与列一一对应；null 表示该列被丢弃
  const droppedCols = [];
  for (let c = 0; c < width; c++) {
    const rawName = String(head[c] === undefined ? '' : head[c]).trim();
    const hasData = body.some(r => String(r[c] === undefined ? '' : r[c]).trim() !== '');
    if (!rawName && !hasData) { names.push(null); droppedCols.push(c + 1); continue; }   // 无名且整列全空 → 丢弃
    let n = cleanHeaderName(rawName, c) || `列${c + 1}`;
    if (used.has(n)) {                                     // 重复字段名 → 追加序号，避免互相覆盖
      let k = used.get(n) + 1;
      while (used.has(`${n}(${k})`)) k++;
      used.set(n, k);
      n = `${n}(${k})`;
    }
    used.set(n, 1);
    names.push(n);
  }

  const rows = [];
  let emptySkipped = 0, summaryRows = 0;
  body.forEach(r => {
    const o = {}; let has = false;
    names.forEach((n, c) => {
      if (!n) return;
      const v = String(r[c] === undefined || r[c] === null ? '' : r[c]).trim();
      o[n] = v;
      if (v !== '') has = true;
    });
    if (!has) { emptySkipped++; return; }                  // 整行空行忽略
    const firstNonEmpty = (r.find(v => String(v == null ? '' : v).trim() !== '') || '').trim();
    if (/^(合计|总计|小计|汇总|共计|总价)\s*[:：]?/.test(firstNonEmpty)) summaryRows++;
    rows.push(o);
  });

  return {
    headers: names.filter(Boolean), rows, headerRow: hr, rowOffset, emptySkipped,
    notes: { droppedCols, summaryRows, leadingSkipped: hr },
  };
}

/** 开始一次导入：读取工作簿 → 挑选工作表 → 解析（默认导入全部工作表） */
function startImport(bytes, fileName) {
  const { wb, encoding } = readWorkbook(bytes, fileName);
  const names = (wb.SheetNames || []).filter(n => wb.Sheets[n] && wb.Sheets[n]['!ref']);
  if (!names.length) { notify('文件中没有可读取的工作表（可能是空文件）', 'error', 7000); return; }
  importCtx = {
    wb, fileName, sheets: names, sheet: pickBestSheet(wb, names),
    headerRow: null, encoding, res: null, norm: null,
    useAll: names.length > 1,      // 多工作表时默认「全部导入」
    cache: new Map(),
  };
  applyImport();
}

/** 解析某张工作表（同一「工作表+表头行」结果缓存，避免重复解析） */
function parseSheetCached(name, headerRowOverride) {
  const ctx = importCtx;
  const auto = headerRowOverride === undefined;
  const key = name + '|' + (auto ? String(ctx.headerRow) : 'auto');
  if (ctx.cache.has(key)) return ctx.cache.get(key);
  const res = analyzeSheet(ctx.wb.Sheets[name], auto ? ctx.headerRow : headerRowOverride);
  ctx.cache.set(key, res);
  return res;
}

/** 归一化列名（忽略大小写与空格），用于跨表按列名合并 */
const normColKey = s => String(s || '').replace(/\s+/g, '').toLowerCase();

/**
 * 合并全部工作表：
 *  - 主表 = 最可能是数据表的那张（其列名作为基础列）
 *  - 其余工作表按「列名相同」合并（列名大小写/空格不同也算同一列），多出的新列追加到最后
 *  - 与主表没有任何同名列的表（如「填写说明」）会被跳过，并在报告里说明原因
 */
function mergeSheets() {
  const ctx = importCtx;
  const primary = pickBestSheet(ctx.wb, ctx.sheets);
  const base = parseSheetCached(primary, null);
  const headers = [...base.headers];
  const canonical = new Map(base.headers.map(h => [normColKey(h), h]));   // 归一化名 → 规范列名
  const rows = base.rows.map(r => ({ ...r }));
  const sheetInfo = [{
    name: primary, rows: base.rows.length, headerRow: base.headerRow,
    used: true, reason: '主表', headers: base.headers,
  }];
  ctx.sheets.forEach(name => {
    if (name === primary) return;
    const res = parseSheetCached(name, null);
    if (!res.rows.length) { sheetInfo.push({ name, rows: 0, used: false, reason: '没有数据行' }); return; }
    const inter = res.headers.filter(h => canonical.has(normColKey(h)));
    if (!inter.length) {
      sheetInfo.push({ name, rows: res.rows.length, used: false, reason: '列名与主表无相同项', headers: res.headers });
      return;
    }
    res.headers.forEach(h => {
      if (!canonical.has(normColKey(h))) { canonical.set(normColKey(h), h); headers.push(h); }
    });
    res.rows.forEach(r => {                       // 列名不一致但同义的列合并到同一列
      const o = {};
      Object.keys(r).forEach(k => { o[canonical.get(normColKey(k)) || k] = r[k]; });
      rows.push(o);
    });
    sheetInfo.push({ name, rows: res.rows.length, headerRow: res.headerRow, used: true,
      reason: `与主表共有 ${inter.length} 列`, headers: res.headers });
  });
  return {
    res: {
      headers, rows, headerRow: base.headerRow, rowOffset: base.rowOffset,
      notes: base.notes, sheetInfo, multi: true,
    },
    sheetInfo,
  };
}

/** 按当前设置解析并应用（重新解析、切换工作表/表头行都走这里） */
function applyImport() {
  const ctx = importCtx;
  if (!ctx) return;
  ctx.cache = new Map();                                   // 重新解析时清缓存
  let res;
  if (ctx.useAll && ctx.sheets.length > 1) {
    res = mergeSheets().res;
    ctx.res = res;
    if (!res.rows.length) {
      renderReport();
      notify('全部工作表中都没有识别到数据行：可取消「导入全部工作表」后单独指定工作表与表头行', 'warn', 8000);
      return;
    }
    applyImportResult(res);
    return;
  }
  res = parseSheetCached(ctx.sheet, undefined);
  ctx.headerRow = res.headerRow;
  ctx.res = res;
  if (!res.rows.length) {
    renderReport();                                        // 报告里给出可操作的提示
    notify(`工作表「${ctx.sheet}」未识别到数据行：请更换「工作表」或调整「表头行」后点「重新解析」`, 'warn', 8000);
    return;
  }
  applyImportResult(res);
}

/** 应用导入结果：写入状态、初始化打印数量、刷新界面 */
function applyImportResult(res) {
  const ctx = importCtx;
  const fileName = ctx ? ctx.fileName : (res.fileName || '');
  // 后处理：表头规范化 + 派生列（二维码内容/条码内容缺失或整列为空时自动生成）
  const norm = normalizeImport(res.rows, res.headers, fileName);
  if (ctx) ctx.norm = norm;
  state.headers = norm.headers;
  state.rows = norm.rows;
  // 打印数量：读取表格中的「打印数量」列（留空或非法值按 1 份处理）
  state.qty = norm.rows.map(r => {
    const v = parseInt(r['打印数量'], 10);
    return Number.isFinite(v) && v >= 1 ? Math.min(999, v) : 1;
  });
  state.currentRow = 0;
  renderTable(); renderCanvas(); renderProps(); renderReport();
  saveLocal();                     // 导入结果（含打印数量）写入草稿，刷新后可恢复
  if (qeEl) openQuickEdit(qeEl);   // 刷新快捷编辑器里的字段列表
  const multi = state.qty.filter(q => q > 1).length;
  const usedSheets = (res.sheetInfo || []).filter(s => s.used);
  const fromTxt = usedSheets.length > 1 ? `（来自 ${usedSheets.length} 个工作表）` : '';
  $('#import-info').textContent = `已导入 ${norm.rows.length} 行 × ${norm.headers.length} 列${fromTxt}`;
  if (multi) notify(`已按「打印数量」列设置 ${multi} 行的打印份数`, 'info', 4000);
  if (usedSheets.length > 1) {
    notify(`已合并导入 ${usedSheets.length} 个工作表：` + usedSheets.map(s => `${s.name} ${s.rows} 行`).join('、')
      + `，共 ${norm.rows.length} 行`, 'success', 6000);
  }
}

/** 数据标准化：表头去空格、剔除空行、生成派生列 */
function normalizeImport(rawRows, rawHeaders, fileName) {
  const headers = rawHeaders.map(h => String(h).trim()).filter(h => h !== '');
  const rows = [];
  let emptySkipped = 0;
  rawRows.forEach(r => {
    const o = {};
    headers.forEach(h => { const v = r[h]; o[h] = v === undefined || v === null ? '' : String(v).trim(); });
    if (headers.every(h => o[h] === '')) { emptySkipped++; return; }
    rows.push(o);
  });
  const derived = [];
  Object.keys(DERIVED).forEach(col => {
    const exists = headers.includes(col);
    // 列不存在，或列存在但整列为空 → 按规则自动生成
    if (exists && rows.some(r => r[col] !== '')) return;
    if (!rows.some(r => DERIVED[col](r) !== '')) return;
    if (!exists) headers.push(col);
    rows.forEach(r => { r[col] = DERIVED[col](r); });
    derived.push(col);
  });
  return { headers, rows, emptySkipped, derived, fileName };
}

/**
 * 导入结果条（精简版）：解析正常时只保留「工作表 / 表头行 / 重新解析」这几个必要控件与行数列数，
 * 不再输出字段清单与各类提示文案，导入后直接看下方表格；仅当没解析出数据行时才给出原因提示。
 */
function renderReport() {
  const box = $('#import-report');
  const ctx = importCtx;
  if (!ctx) { box.classList.add('hidden'); box.innerHTML = ''; return; }
  const res = ctx.res || { headers: [], rows: [], headerRow: 0, emptySkipped: 0, notes: {} };
  const notes = res.notes || {};
  const parsedOk = (res.rows || []).length > 0;                 // 本次解析是否拿到数据行
  const rowN = parsedOk ? state.rows.length : 0;
  const colN = parsedOk ? state.headers.length : (res.headers || []).length;
  const hrNo = res.headerRow + (res.rowOffset || 0) + 1;     // 表头行号（按 Excel 实际行号显示）

  const multi = ctx.sheets.length > 1;
  const all = multi && ctx.useAll;                            // 是否「导入全部工作表」
  const info = res.sheetInfo || [];
  const usedN = info.filter(s => s.used).length;
  const sheetList = info.length
    ? '工作表：' + info.map(s => s.used
      ? `<b>${escHtml(s.name)}</b> ${s.rows} 行${s.reason === '主表' ? '（主表）' : ''}`
      : `<span class="muted">${escHtml(s.name)}</span> <span class="muted">已跳过（${escHtml(s.reason)}）</span>`
    ).join('　|　')
    : '';

  box.classList.remove('hidden');
  box.className = parsedOk ? '' : 'warn';                    // 成功时不加醒目底色，保持安静
  box.innerHTML = `
    <div class="ir-head">
      ${multi ? `<label class="ir-ctl" title="勾选后把所有工作表的数据合并导入（按列名对齐，多出的列自动追加）；取消勾选可单独指定某一张工作表与表头行">
        <input type="checkbox" id="ir-all"${all ? ' checked' : ''}> 导入全部 ${ctx.sheets.length} 个工作表</label>` : ''}
      <label class="ir-ctl"${all ? ' style="opacity:.55"' : ''}>工作表 <select id="ir-sheet"${all ? ' disabled' : ''}>${
        ctx.sheets.map(n => `<option${n === ctx.sheet ? ' selected' : ''}>${escHtml(n)}</option>`).join('')}</select></label>
      <label class="ir-ctl"${all ? ' style="opacity:.55"' : ''} title="${escHtml([
        `默认取第一行（第一行有内容的行）作为字段名`,
        all ? '「导入全部工作表」时每张表各自自动识别表头行' : '',
        notes.leadingSkipped ? `本次已自动跳过前 ${notes.leadingSkipped} 行（标题/说明/空行），表头取自第 ${hrNo} 行` : '',
        ctx.encoding ? `文件编码 ${ctx.encoding}` : '',
      ].filter(Boolean).join('；'))}">表头行
        <input type="number" id="ir-header" min="${(res.rowOffset || 0) + 1}" max="9999" value="${hrNo}"${all ? ' disabled' : ''}></label>
      <button id="ir-reparse" title="按当前设置重新解析该文件">↻ 重新解析</button>
      <button id="ir-hide" title="收起本行，不影响已导入的数据">收起</button>
      <span class="ir-stat">已导入 <b>${rowN}</b> 行 × <b>${colN}</b> 列${all && usedN ? `（来自 <b>${usedN}</b> 个工作表）` : ''}</span>
    </div>
    ${all && sheetList ? `<div class="ir-line">${sheetList}</div>` : ''}
    ${parsedOk ? '' : `<div class="ir-line">未识别到数据行：请确认「工作表」是否正确，或把「表头行」改为字段名所在行后点「↻ 重新解析」。`
      + (state.rows.length ? `（下方仍显示上一次导入的 ${state.rows.length} 行数据）` : '') + '</div>'}`;

  if ($('#ir-all')) $('#ir-all').addEventListener('change', e => {
    ctx.useAll = e.target.checked;
    applyImport();
  });
  $('#ir-sheet').addEventListener('change', e => {
    ctx.sheet = e.target.value; ctx.headerRow = null; applyImport();
  });
  $('#ir-header').addEventListener('change', e => {
    const v = parseInt(e.target.value, 10);
    ctx.headerRow = Math.max(0, (Number.isFinite(v) ? v : 1) - 1 - (res.rowOffset || 0));
    applyImport();
  });
  $('#ir-header').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); e.target.blur(); } });
  $('#ir-reparse').addEventListener('click', () => applyImport());
  $('#ir-hide').addEventListener('click', () => box.classList.add('hidden'));
}

function renderTable() {
  const wrap = $('#data-table-wrap');
  if (!state.rows.length) {
    wrap.innerHTML = '<p class="muted" style="padding:12px">尚未导入数据。</p>';
    $('#data-count').textContent = '';
    return;
  }
  const th = state.headers.map(h => `<th>${escHtml(h)}</th>`).join('');
  const trs = state.rows.map((r, i) => {
    const tds = state.headers.map(h => `<td title="${escHtml(r[h])}">${escHtml(r[h])}</td>`).join('');
    const cls = i === state.currentRow ? ' class="active"' : '';
    return `<tr${cls} data-i="${i}"><td class="rowno">${i + 1}</td>${tds}`
      + `<td class="qty-cell"><input type="number" class="qty-input" min="0" max="999" step="1" `
      + `data-qty="${i}" value="${qtyOf(i)}" title="该行连续打印的份数"></td>`
      + `<td class="op" data-print="${i}" title="双击直接打印该行">打印</td></tr>`;
  }).join('');
  wrap.innerHTML = `<table id="data-table"><thead><tr><th class="rowno">#</th>${th}`
    + `<th class="qty-head">打印数量</th><th>操作</th></tr></thead><tbody>${trs}</tbody></table>`;
  renderDataCount();

  wrap.querySelectorAll('tr[data-i]').forEach(tr => tr.addEventListener('click', () => {
    state.currentRow = +tr.dataset.i;
    renderTable(); renderCanvas();
  }));
  // 打印数量可编辑：输入即生效并记忆
  wrap.querySelectorAll('[data-qty]').forEach(inp => {
    inp.addEventListener('click', ev => ev.stopPropagation());
    inp.addEventListener('change', ev => {
      const i = +ev.target.dataset.qty;
      let v = parseInt(ev.target.value, 10);
      if (!Number.isFinite(v) || v < 0) v = 0;
      state.qty[i] = Math.min(999, v);
      ev.target.value = state.qty[i];
      renderDataCount(); saveLocal();
      const row = state.rows[i];
      notify(`第 ${i + 1} 行「${row && row['物料编码'] ? row['物料编码'] : '数据'}」打印数量已设为 ${state.qty[i]} 份`, 'info', 2200);
    });
  });
  wrap.querySelectorAll('[data-print]').forEach(td => td.addEventListener('dblclick', async ev => {
    ev.stopPropagation();
    const i = +td.dataset.print;
    printDirect([{ row: state.rows[i], qty: qtyOf(i) }]);
  }));
}

/** 数据区统计：行数 + 合计打印份数（收起时也显示在标题条上） */
function renderDataCount() {
  if (!state.rows.length) { $('#data-count').textContent = ''; return; }
  const total = state.rows.reduce((s, _, i) => s + qtyOf(i), 0);
  const multi = state.rows.filter((_, i) => qtyOf(i) > 1).length;
  const cur = (state.currentRow >= 0 && state.currentRow < state.rows.length)
    ? ` · 预览第 ${state.currentRow + 1} 行` : '';
  $('#data-count').textContent = `共 ${state.rows.length} 行 · 合计 ${total} 份`
    + (multi ? `（${multi} 行设置了多份）` : '') + cur;
}

/* ============================================================
 * 底部 Excel 数据区：上拉展开 / 点击向下收起（默认收起，避免遮挡标签区）
 * ============================================================ */
const DATA_PANEL_KEY = 'label-print-data-panel';
const DATA_H_EXPANDED = 232, DATA_H_COLLAPSED = 34;

function setDataPanel(collapsed, opt = {}) {
  const box = $('#data-area');
  if (!box) return;
  box.classList.toggle('collapsed', !!collapsed);
  document.documentElement.style.setProperty('--data-h', (collapsed ? DATA_H_COLLAPSED : DATA_H_EXPANDED) + 'px');
  const hint = $('#data-toggle-hint');
  if (hint) hint.textContent = collapsed ? '点击展开' : '点击收起';
  const t = $('#data-toggle');
  if (t) t.title = collapsed ? '点击展开 Excel 数据区（上拉显示）' : '点击收起 Excel 数据区（向下收起，腾出标签空间）';
  if (!opt.silent) {
    try { localStorage.setItem(DATA_PANEL_KEY, collapsed ? 'collapsed' : 'expanded'); } catch (e) { /* 忽略 */ }
  }
}

$('#data-toggle').addEventListener('click', e => {
  if (e.target.closest('button')) return;            // 标题条上的快捷按钮不触发折叠
  setDataPanel(!$('#data-area').classList.contains('collapsed'));
});
$('#btn-data-collapse').addEventListener('click', () => setDataPanel(true));
$('#btn-import-mini').addEventListener('click', () => $('#excel-file').click());

/* ============================================================
 * 打印（生成独立打印窗口，@page 按标签尺寸）
 * ============================================================ */
/* ============================================================
 * 标签位图渲染（用于服务端静默打印，不依赖任何截图库）
 * 按元素类型用 Canvas 2D 直接绘制，1mm = PRINT_K 像素（≈305dpi）
 * ============================================================ */
const PRINT_K = 12;

function wrapText(ctx, text, maxW) {
  const out = [];
  String(text).split('\n').forEach(seg => {
    if (!seg) { out.push(''); return; }
    let line = '';
    for (const ch of seg) {
      if (line && ctx.measureText(line + ch).width > maxW) { out.push(line); line = ch; }
      else line += ch;
    }
    out.push(line);
  });
  return out;
}

function drawPlaceholder(ctx, x, y, w, h, label) {
  ctx.save();
  ctx.strokeStyle = '#bdbdbd';
  ctx.lineWidth = Math.max(1, PRINT_K * 0.2);
  ctx.setLineDash([PRINT_K, PRINT_K]);
  ctx.strokeRect(x, y, w, h);
  ctx.setLineDash([]);
  ctx.fillStyle = '#9e9e9e';
  ctx.font = `${Math.max(6, PRINT_K * 1.1)}px sans-serif`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(label, x + w / 2, y + h / 2);
  ctx.restore();
}

/** 绘制单个元素（单位 mm → 像素） */
function drawElement(ctx, el, row, K) {
  const x = el.x * K, y = el.y * K, w = el.w * K, h = el.h * K;
  const value = cellValue(el, row);
  ctx.save();
  ctx.fillStyle = '#000';
  ctx.strokeStyle = '#000';
  switch (el.type) {
    case 'text': {
      // 直连打印（位图）：绑定字段为空 → 整个元素不画（不再把 {字段名} 打出来）
      if (value === null) break;
      const txt = (el.prefix || '') + value;
      // 自适应字号：一行放不下就缩小字号（与画布、打印预览同一套算法）
      const autoFit = el.autoFit !== false;
      const pt = autoFit
        ? fitFontSizePt(txt, Math.max(0.5, el.w - 0.4), el.fontSize || 10, el.fontFamily, el.bold)
        : (el.fontSize || 10);
      const fontPx = Math.max(1, pt * 25.4 / 72 * K);
      ctx.font = `${el.bold ? 'bold ' : ''}${fontPx}px "${el.fontFamily || DEFAULT_FONT}"`;
      ctx.textBaseline = 'middle';
      ctx.textAlign = 'left';
      const lines = autoFit ? [String(txt).replace(/\r?\n/g, ' ')] : wrapText(ctx, txt, w);
      const lh = fontPx * 1.2;
      let ty = y + h / 2 - (lines.length * lh) / 2 + lh / 2;
      ctx.save();
      ctx.beginPath();
      ctx.rect(x, y, w, h);                     // 限制在元素框内，避免溢出到标签其它位置
      ctx.clip();
      lines.forEach(line => {
        const tw = ctx.measureText(line).width;
        const tx = el.align === 'center' ? x + (w - tw) / 2 : el.align === 'right' ? x + w - tw : x;
        ctx.fillText(line, tx, ty);
        ty += lh;
      });
      ctx.restore();
      break;
    }
    case 'barcode': {
      if (value === null) break;                     // 打印：空值不画占位框
      const bcPt = el.autoFit === false
        ? (el.fontSize || 10)
        : fitFontSizePt(String(value), Math.max(0.5, el.w - 0.4), el.fontSize || 10, el.fontFamily, el.bold);
      const bc = document.createElement('canvas');
      try {
        JsBarcode(bc, value, {
          format: el.barcodeFormat || 'CODE128',
          width: Math.max(1, Math.round((el.barWidth || 2) * K / 3)),
          height: Math.round(h * (el.showText ? 0.75 : 1)),
          displayValue: !!el.showText,
          fontSize: Math.round(bcPt * 25.4 / 72 * K),
          font: el.fontFamily || DEFAULT_FONT,
          margin: 0,
        });
        ctx.drawImage(bc, x, y, w, h);
      } catch (e) { drawPlaceholder(ctx, x, y, w, h, '条码错误'); }
      break;
    }
    case 'qrcode': {
      if (value === null) break;                     // 打印：空值不画占位框
      const cv = qrCanvasFor(el, value, Math.max(64, Math.round(el.w * PRINT_K)));
      if (cv) ctx.drawImage(cv, x, y, w, h); else drawPlaceholder(ctx, x, y, w, h, '二维码错误');
      break;
    }
    case 'rect': {
      const lw = Math.max(1, (el.borderWidth || 0.4) * K);
      ctx.lineWidth = lw;
      ctx.strokeRect(x + lw / 2, y + lw / 2, Math.max(0, w - lw), Math.max(0, h - lw));
      break;
    }
    case 'line': {
      ctx.lineWidth = Math.max(1, 0.4 * K);
      ctx.beginPath();
      if (el.h > el.w) { ctx.moveTo(x + w / 2, y); ctx.lineTo(x + w / 2, y + h); }
      else { ctx.moveTo(x, y + h / 2); ctx.lineTo(x + w, y + h / 2); }
      ctx.stroke();
      break;
    }
    case 'table': {
      const bw = Math.max(1, (el.borderWidth || 0.3) * K);
      ctx.textBaseline = 'middle';
      tblBlocks(el).forEach(b => {
        const r = tblBlockRect(el, b);
        const x0 = x + r.x * K, y0 = y + r.y * K, w0 = r.w * K, h0 = r.h * K;
        // 边框：每条边只画一次（上/左 + 末行末列），避免相邻单元格边框重叠变粗
        ctx.fillStyle = '#000';
        ctx.fillRect(x0, y0, w0, bw);
        ctx.fillRect(x0, y0, bw, h0);
        if (b.c + b.cs >= el.cols) ctx.fillRect(x0 + w0 - bw, y0, bw, h0);
        if (b.r + b.rs >= el.rows) ctx.fillRect(x0, y0 + h0 - bw, w0, bw);

        const cd = tblCellData(el, b);
        const txt = tblCellText(cd, row, 'mm');     // 直连打印：绑定字段为空 → 该单元格留空
        if (!txt) return;
        // 自适应字号：可用宽度(mm) = 块宽 − 边框 − 内边距；放不下则缩小字号，且只画一行
        const autoFit = cd.autoFit !== false;
        const availMm = (w0 - bw * 2) / K - 0.6;
        const pt = cellFontSizePt(cd, txt, availMm);
        const fontPx = Math.max(1, pt * 25.4 / 72 * K);
        ctx.save();
        ctx.beginPath();
        ctx.rect(x0, y0, w0, h0);                    // 限制在单元格内，避免溢出到相邻格子
        ctx.clip();
        ctx.font = `${cd.bold ? 'bold ' : ''}${fontPx}px "${cd.fontFamily || DEFAULT_FONT}"`;
        ctx.fillStyle = '#000';
        ctx.textAlign = cd.align === 'left' ? 'left' : cd.align === 'right' ? 'right' : 'center';
        const oneLine = autoFit ? [String(txt).replace(/\r?\n/g, ' ')] : wrapText(ctx, txt, Math.max(2, w0 - bw * 2));
        const lh = fontPx * 1.2;
        const totalH = oneLine.length * lh;
        const tx = cd.align === 'left' ? x0 + bw : cd.align === 'right' ? x0 + w0 - bw : x0 + w0 / 2;
        let firstCenter = y0 + h0 / 2 - totalH / 2 + lh / 2;
        if (cd.valign === 'top') firstCenter = y0 + bw + lh / 2;
        else if (cd.valign === 'bottom') firstCenter = y0 + h0 - bw - totalH + lh / 2;
        oneLine.forEach((line, i) => ctx.fillText(line, tx, firstCenter + i * lh));
        ctx.restore();
      });
      break;
    }
  }
  ctx.restore();
}

/** 把一行数据的标签渲染为 PNG dataURL（可复用同一 canvas） */
function renderLabelPng(row, reuseCanvas) {
  const K = PRINT_K;
  const c = reuseCanvas || document.createElement('canvas');
  c.width = Math.max(1, Math.round(state.label.width * K));
  c.height = Math.max(1, Math.round(state.label.height * K));
  const ctx = c.getContext('2d');
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, c.width, c.height);
  state.elements.forEach(el => drawElement(ctx, el, row, K));
  return c.toDataURL('image/png');
}

/** 直接打印：前端渲染位图 → 服务端静默投递到已选打印机（无浏览器打印对话框） */
async function printDirect(jobs) {
  if (!state.elements.length) { notify('标签模板为空，请先设计标签', 'warn'); return; }
  if (!jobs.length) { notify('没有可打印的数据行，请先导入 Excel 或选中行', 'warn'); return; }
  if (!state.printer) { notify('请先在顶部选择目标打印机', 'warn'); await loadPrinters(true); return; }

  const total = jobs.reduce((s, j) => s + (j.qty || 1), 0);
  const p = (state.printers || []).find(x => x.name === state.printer);
  if (p && ['offline', 'error', 'paused'].includes(p.statusCode)) {
    const go = await askConfirm('打印机当前不可用',
      `「${state.printer}」当前状态：${p.statusText}。作业可能会排队等待或无法输出，是否仍然发送？`, '仍然发送');
    if (!go) return;
  }

  const busy = notify(`正在渲染并投递 ${jobs.length} 行 / ${total} 份标签到「${state.printer}」…`, 'info', 0);
  const t0 = Date.now();
  try {
    const canvas = document.createElement('canvas');
    const pages = [];
    const seen = new Map();       // 相同内容的标签合并为一份图片 + 份数
    jobs.forEach(j => {
      const dataUrl = renderLabelPng(j.row, canvas);
      if (seen.has(dataUrl)) seen.get(dataUrl).copies += (j.qty || 1);
      else { const o = { image: dataUrl, copies: j.qty || 1 }; seen.set(dataUrl, o); pages.push(o); }
    });
    const res = await fetch(apiUrl('/api/print'), {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        printer: state.printer, widthMm: state.label.width, heightMm: state.label.height, pages,
      }),
    });
    const j = await res.json();
    if (!j.ok) throw new Error(j.message || '服务端打印失败');
    notify(`✔ 已直接打印 ${j.pages} 份到「${j.printer}」`
      + (isRemoteAgent() ? `（打印服务：${agentBase()}）` : '')
      + `${j.failed ? `（${j.failed} 张渲染失败）` : ''}，耗时 ${((Date.now() - t0) / 1000).toFixed(1)} 秒`, 'success', 7000);
    setTimeout(() => loadPrinters(true), 800);   // 打印后刷新队列状态
  } catch (e) {
    notify('打印失败：' + e.message, 'error', 9000);
  } finally {
    busy.remove();
  }
}

function labelOuter(el, row) {
  return `<div style="position:absolute;left:${el.x}mm;top:${el.y}mm;width:${el.w}mm;height:${el.h}mm;">`
    + elementContent(el, row, 'mm') + `</div>`;    // 打印：表格内部也用 mm，避免画布缩放影响打印尺寸
}

function printLabels(rows) {
  if (!rows.length) { notify('没有可打印的数据行，请先导入 Excel 或选中行', 'warn'); return; }
  if (!state.elements.length) { notify('标签模板为空，请先设计标签', 'warn'); return; }
  const { width, height } = state.label;
  // 打印时不带表格编辑态的辅助线/选框
  const savedTableEdit = state.tableEdit;
  state.tableEdit = null;
  const labels = rows.map(r => `<div class="lbl">${state.elements.map(el => labelOuter(el, r)).join('')}</div>`).join('');
  state.tableEdit = savedTableEdit;

  // 打印机提示：附带当前选择的本机打印机与其状态
  const cur = (state.printers || []).find(p => p.name === state.printer);
  const printerInfo = state.printer
    ? `<span class="pr-tip">目标打印机：<b>${escHtml(state.printer)}</b>`
      + (cur ? `（${escHtml(cur.statusText)}${cur.isDefault ? '，系统默认' : ''}）` : '')
      + (cur && ['offline', 'error', 'paused'].includes(cur.statusCode) ? ' ⚠ 该打印机当前不可用' : '')
      + '，请在打印对话框中选择该打印机</span>'
    : '<span class="pr-tip">未指定打印机，请在打印对话框中选择</span>';

  const w = window.open('', '_blank', 'width=900,height=700');
  if (!w) { notify('打印窗口被浏览器拦截，请允许本站弹出窗口后重试', 'error', 7000); return; }
  w.document.write(`<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8"><title>标签打印 (${rows.length} 份)</title>
  <style>
    @page { size: ${width}mm ${height}mm; margin: 0; }
    html, body { margin: 0; padding: 0; background: #888; }
    .lbl {
      width: ${width}mm; height: ${height}mm; position: relative; overflow: hidden;
      background: #fff; page-break-after: always; margin: 6px auto; outline: 1px dashed #aaa;
    }
    #pt-bar { position: fixed; top: 0; left: 0; right: 0; background: #1668dc; color: #fff;
      padding: 8px 14px; z-index: 9; display: flex; gap: 10px; align-items: center; }
    #pt-bar button { padding: 4px 16px; }
    #pt-bar .pr-tip { font-size: 12px; opacity: .95; }
    @media print { #pt-bar { display: none; } .lbl { margin: 0; outline: none; } body { background: #fff; } }
  </style></head><body>
  <div id="pt-bar"><span>共 ${rows.length} 份标签 · ${width}×${height}mm</span>
    ${printerInfo}
    <button onclick="window.print()">打印</button>
    <button onclick="window.close()">关闭</button></div>
  ${labels}
  <script>setTimeout(function(){window.print()},400)<\/script>
  </body></html>`);
  w.document.close();
}

/** 取某行的打印数量：未设置默认 1；设置为 0 表示该行不打印 */
function qtyOf(i) {
  const raw = state.qty[i];
  if (raw === undefined || raw === null || raw === '') return 1;
  const v = parseInt(raw, 10);
  if (!Number.isFinite(v)) return 1;
  return Math.max(0, Math.min(999, v));
}
/** 当前所有数据行 → 打印任务 [{row, qty}] */
function allJobs() { return state.rows.map((row, i) => ({ row, qty: qtyOf(i) })); }
function totalCopies(jobs) { return jobs.reduce((s, j) => s + j.qty, 0); }

async function printCurrentRow() {
  if (state.currentRow < 0 || !state.rows[state.currentRow]) {
    notify('请先导入 Excel 并在底部表格中选中一行', 'warn'); return;
  }
  const qty = qtyOf(state.currentRow);
  if (qty < 1) {
    notify(`第 ${state.currentRow + 1} 行的打印数量为 0（不打印），请先修改该行份数`, 'warn', 4500);
    return;
  }
  const yes = await askConfirm('打印当前行',
    `第 ${state.currentRow + 1} 行将连续打印 ${qty} 份（纸张 ${state.label.width}×${state.label.height}mm）。\n`
    + `目标打印机：${state.printer || '未选择'}（直连打印机，不弹浏览器打印框）`, '开始打印');
  if (!yes) return;
  printDirect([{ row: state.rows[state.currentRow], qty }]);
}

async function printAllRows() {
  if (!state.rows.length) { notify('请先导入 Excel 数据', 'warn'); return; }
  const jobs = allJobs();
  const total = totalCopies(jobs);
  const zero = jobs.filter(j => j.qty < 1).length;
  const yes = await askConfirm('批量打印',
    `共 ${state.rows.length} 行数据，按各行「打印数量」合计 ${total} 份标签`
    + `（纸张 ${state.label.width}×${state.label.height}mm）${zero ? '，其中 ' + zero + ' 行份数为 0 将被跳过' : ''}。\n`
    + `目标打印机：${state.printer || '未选择'}（直连打印机，不弹浏览器打印框）`, '开始打印');
  if (!yes) return;
  printDirect(jobs.filter(j => j.qty > 0));
}

/* 浏览器预览打印（备用方式，会弹出浏览器打印对话框） */
async function printPreviewAll() {
  if (!state.rows.length) { notify('请先导入 Excel 数据', 'warn'); return; }
  const rows = [];
  allJobs().forEach(j => { for (let i = 0; i < j.qty; i++) rows.push(j.row); });
  if (!rows.length) { notify('各行打印数量均为 0，没有可打印内容', 'warn'); return; }
  printLabels(rows);
}

/* 功能区、数据区、数据区标题条三处入口都绑定（直接打印到已选打印机） */
['#btn-print-row', '#btn-print-row2'].forEach(s => $(s) && $(s).addEventListener('click', printCurrentRow));
['#btn-print-all', '#btn-print-all2', '#btn-print-all-mini'].forEach(s => $(s) && $(s).addEventListener('click', printAllRows));
$('#btn-print-preview') && $('#btn-print-preview').addEventListener('click', printPreviewAll);

/* ============================================================
 * 模板：保存 / 管理（服务端 data/labels 目录）
 * ============================================================ */
/** 保存当前设计到服务端（供工具栏与模板弹窗复用） */
async function saveTemplate() {
  if (!state.elements.length) { notify('画布为空，请先设计标签内容再保存', 'warn'); return; }
  const name = await askText('保存标签模板', state.name || '我的标签');
  if (name === null) return;                       // 用户取消
  if (!name) { notify('模板名不能为空', 'warn'); return; }

  const btn = $('#btn-save');
  btn.disabled = true; btn.textContent = '保存中…';
  try {
    const res = await fetch('/api/labels', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, label: state.label, elements: state.elements, printer: state.printer || '' }),
    });
    const j = await res.json();
    if (!j.ok) throw new Error(j.message || '服务端返回失败');
    state.name = name;
    saveLocal(); setTplStatus(true);
    notify(`模板「${name}」已保存：服务端 data/labels/${name}.json（${state.elements.length} 个元素，`
      + `${state.label.width}×${state.label.height}mm${state.printer ? '，打印机 ' + state.printer : ''}）`, 'success', 6000);
    loadTplList();
  } catch (e) {
    notify('保存失败：' + e.message
      + (String(e.message).includes('fetch') ? `（无法连接服务端，请通过 ${siteOrigin()} 访问本页面）` : ''), 'error', 8000);
    setTplStatus(false);
  } finally {
    btn.disabled = false; btn.textContent = '保存模板';
  }
}
$('#btn-save').addEventListener('click', saveTemplate);

$('#btn-open').addEventListener('click', openModal);
$('#modal-close').addEventListener('click', () => $('#modal-mask').classList.add('hidden'));
$('#modal-mask').addEventListener('click', e => { if (e.target.id === 'modal-mask') $('#modal-mask').classList.add('hidden'); });

async function openModal() {
  $('#modal-mask').classList.remove('hidden');
  const list = $('#modal-list');
  list.innerHTML = '<p class="muted" style="padding:12px">正在读取服务端模板列表…</p>';
  try {
    const res = await fetch('/api/labels', { cache: 'no-store' });
    const j = await res.json();
    if (!j.ok) throw new Error(j.message);
    if (!j.list.length) {
      list.innerHTML = '<p class="muted" style="padding:12px">暂无已保存的模板。<br>'
        + '请先设计标签，然后点击工具栏「保存模板」或本窗口的「＋ 保存当前设计」。<br>'
        + '模板会保存为服务端 data/labels/模板名.json。</p>';
      return;
    }
    list.innerHTML = j.list.map(t => `<div class="tpl-item">
      <div><div><b>${escHtml(t.name)}</b>${t.name === state.name ? ' <span class="muted">(当前)</span>' : ''}</div>
      <div class="meta">${t.elementCount ?? '-'} 个元素 · ${t.width || '?'}×${t.height || '?'}mm
        ${t.updatedAt ? ' · ' + new Date(t.updatedAt).toLocaleString() : ''}</div></div>
      <div class="ops"><button data-load="${escHtml(t.name)}">载入</button>
      <button class="danger" data-del="${escHtml(t.name)}">删除</button></div></div>`).join('');

    list.querySelectorAll('[data-load]').forEach(b => b.addEventListener('click', () => loadTemplate(b.dataset.load)));
    list.querySelectorAll('[data-del]').forEach(b => b.addEventListener('click', async () => {
      const name = b.dataset.del;
      const info = await loadTemplate(name, true);   // 取元素信息用于确认提示
      const desc = info ? `${(info.elements || []).length} 个元素 · ${info.label ? info.label.width + '×' + info.label.height + 'mm' : ''}` : '';
      const yes = await askConfirm('删除标签模板',
        `确定删除模板「${name}」吗？${desc ? '（' + desc + '）' : ''}该操作会删除服务端 data/labels/${name}.json，不可恢复。`, '删除');
      if (!yes) return;
      try {
        const res = await fetch('/api/labels/' + encodeURIComponent(name), { method: 'DELETE' });
        const j = await res.json();
        if (!j.ok) throw new Error(j.message);
        if (state.name === name) { state.name = ''; saveLocal(); setTplStatus(false); }
        notify(`模板「${name}」已删除`, 'success');
        openModal(); loadTplList();
      } catch (e) { notify('删除失败：' + e.message, 'error', 7000); }
    }));
  } catch (e) {
    list.innerHTML = `<p class="muted" style="padding:12px">读取模板列表失败：${escHtml(e.message)}<br>`
      + `请确认后端已启动，并通过 ${siteOrigin()} 访问本页面。</p>`;
    notify('读取模板列表失败：' + e.message, 'error', 8000);
  }
}

async function loadTemplate(name, onlyInfo) {
  try {
    const res = await fetch('/api/labels/' + encodeURIComponent(name), { cache: 'no-store' });
    const j = await res.json();
    if (!j.ok) throw new Error(j.message);
    if (onlyInfo) return j.data;
    state.name = j.data.name || name;
    state.label = j.data.label;
    state.elements = (j.data.elements || []).map(tblSanitize);
    if (j.data.printer) setPrinter(j.data.printer, { silent: true });
    // 恢复自增 id
    state.elements.forEach(e => { const n = parseInt(String(e.id).slice(1)); if (n >= uid) uid = n + 1; });
    syncPaperUI();
    select(null); saveLocal(); setTplStatus(true);
    $('#modal-mask').classList.add('hidden');
    renderTplList();
    notify(`已载入模板「${name}」：${state.elements.length} 个元素，${state.label.width}×${state.label.height}mm`
      + (state.printer ? `，打印机 ${state.printer}` : ''), 'success', 4500);
  } catch (e) {
    if (onlyInfo) return null;
    notify('载入失败：' + e.message, 'error', 7000);
    return null;
  }
}

/* 模板弹窗内的按钮 */
$('#tpl-refresh').addEventListener('click', openModal);
$('#tpl-save-cur').addEventListener('click', async () => {
  $('#modal-mask').classList.add('hidden');
  await saveTemplate();
  openModal();
});

/* ============================================================
 * 示例模板（参照样例软件截图的物料标签示例，字段与标准模板一致）
 * ============================================================ */
$('#btn-sample').addEventListener('click', async () => {
  if (state.elements.length && !(await askConfirm('载入示例模板',
      '载入示例模板将覆盖当前画布设计（未保存的改动会丢失），是否继续？', '载入'))) return;
  state.name = '物料标签示例';
  state.label = { width: 100, height: 70 };
  syncPaperUI();

  const T = (o) => ({ ...DEFAULTS.text, id: 'e' + (uid++), fontSize: 8, ...o });
  const QR = (o) => ({ ...DEFAULTS.qrcode, id: 'e' + (uid++), ...o });
  const BC = (o) => ({ ...DEFAULTS.barcode, id: 'e' + (uid++), fontSize: 7, ...o });

  state.elements = [
    T({ text: '湖北华工图像技术开发有限公司', x: 5, y: 2, w: 90, h: 7, fontSize: 13, bold: true, align: 'center' }),
    QR({ x: 71, y: 11, w: 25, h: 25, binding: '二维码内容' }),
    T({ prefix: '供应商：', binding: '供应商名称', x: 5, y: 11, w: 64, h: 6 }),
    T({ prefix: '采购单号：', binding: '采购单号', x: 5, y: 17, w: 64, h: 6 }),
    T({ prefix: '物料编码：', binding: '物料编码', x: 5, y: 23, w: 64, h: 6 }),
    T({ prefix: '物料名称：', binding: '物料名称', x: 5, y: 29, w: 64, h: 6 }),
    T({ prefix: '规格：', binding: '规格', x: 5, y: 35, w: 31, h: 6 }),
    T({ prefix: '型号：', binding: '型号', x: 38, y: 35, w: 31, h: 6 }),
    T({ prefix: '批号：', binding: '批号', x: 5, y: 41, w: 31, h: 6 }),
    T({ prefix: '数量：', binding: '数量', x: 38, y: 41, w: 31, h: 6 }),
    T({ prefix: '生产日期：', binding: '生产日期', x: 5, y: 47, w: 31, h: 6 }),
    T({ prefix: '有效期至：', binding: '有效期至', x: 38, y: 47, w: 31, h: 6 }),
    T({ prefix: '序列号：', binding: '序列号', x: 5, y: 53, w: 31, h: 6 }),
    T({ prefix: '任务单号：', binding: '任务单号', x: 38, y: 53, w: 31, h: 6 }),
    BC({ x: 5, y: 60, w: 91, h: 9, binding: '条码内容' }),
    { ...DEFAULTS.rect, id: 'e' + (uid++), x: 4, y: 9.5, w: 92, h: 59.5 },
  ];
  select(null); saveLocal(); setTplStatus(false);
  notify('示例模板已载入：双击任一文字可改字/绑定字段；点击「⬇ 下载Excel模板」开始填数据', 'success', 5000);
});

/* 新建 */
$('#btn-new').addEventListener('click', async () => {
  if (state.elements.length && !(await askConfirm('新建标签',
      '将清空当前画布与草稿（未保存的改动会丢失），是否继续？', '新建'))) return;
  state.name = ''; state.elements = []; select(null); saveLocal(); setTplStatus(false);
  renderTplList();
  notify('已新建空白标签画布', 'success');
});
['#btn-new2'].forEach(s => $(s) && $(s).addEventListener('click', () => $('#btn-new').click()));

/* ============================================================
 * 本地草稿（localStorage），防止误关页面丢失
 * ============================================================ */
/* ============================================================
 * 撤销 / 重做（设计历史）
 * ------------------------------------------------------------
 * 快照内容 = 模板名 + 标签尺寸 + 画布元素（表格的行列/合并/单元格数据都在元素里）。
 * 所有改动最终都会走到 saveLocal()（拖拽结束、属性修改、插入删除、合并拆分…），
 * 因此在这里统一记录历史，不需要改动每个操作入口。
 * 连续操作（拖拽、输入框连续修改）在 HISTORY_COALESCE_MS 内合并为一步，避免碎步。
 * ============================================================ */
const HISTORY_MAX = 60;
const HISTORY_COALESCE_MS = 500;
let history = { stack: [], idx: -1, at: 0, suspend: false };

function snapshotDesign() {
  return JSON.stringify({ name: state.name, label: state.label, elements: state.elements });
}

/** 应用快照（撤销/重做共用） */
function applySnapshot(json) {
  let d = null;
  try { d = JSON.parse(json); } catch (e) { return false; }
  const keepSel = state.selectedId;
  history.suspend = true;                     // 应用过程中不再记历史
  try {
    state.name = d.name || '';
    if (d.label) state.label = d.label;
    state.elements = (d.elements || []).map(tblSanitize);
    // id 自增游标不能回退，否则新元素会与历史元素撞 id
    state.elements.forEach(e => { const n = parseInt(String(e.id).slice(1)); if (n >= uid) uid = n + 1; });
    state.tableEdit = null; state.tableSel = null;
    const wantSel = d.sel === undefined ? keepSel : d.sel;
    state.selectedId = state.elements.some(e => e.id === wantSel) ? wantSel : null;
    // 快捷编辑器若指向已被撤销掉的元素，直接关闭，避免后续操作作用在“幽灵元素”上
    if (qeEl && !state.elements.some(e => e.id === qeEl.id)) {
      qeEl = null; qeCellKey = null;
      $('#quick-edit').classList.add('hidden');
    }
    syncPaperUI(); renderCanvas(); renderProps(); updateTableBar();
    saveLocal();                              // 同步草稿（挂起中，不会再记历史）
  } finally { history.suspend = false; }
  return true;
}

/** 记录一步（由 saveLocal 自动调用）；force=true 时不做时间合并 */
function historyPush(force) {
  if (history.suspend) return;
  const snap = snapshotDesign();
  if (history.stack[history.idx] === snap) return;         // 内容没有变化
  const now = Date.now();
  if (!force && history.idx >= 1 && now - history.at < HISTORY_COALESCE_MS) {
    history.stack[history.idx] = snap;                     // 合并到上一步（连续拖拽/输入）
  } else {
    history.stack = history.stack.slice(0, history.idx + 1);
    history.stack.push(snap);
    if (history.stack.length > HISTORY_MAX) history.stack.shift();
    history.idx = history.stack.length - 1;
  }
  history.at = now;
  updateHistoryUI();
}

/** 以当前状态为第一步重置历史（载入草稿/模板后调用） */
function historyReset() {
  history = { stack: [snapshotDesign()], idx: 0, at: Date.now(), suspend: false };
  updateHistoryUI();
}

function undo() {
  if (history.idx <= 0) { notify('已经是第一步，没有可撤销的操作', 'info', 2600); return; }
  history.idx -= 1;
  applySnapshot(history.stack[history.idx]);
  history.at = 0;                            // 撤销后紧接着的修改单独成步
  updateHistoryUI();
  notify(`已撤销一步（当前第 ${history.idx + 1}/${history.stack.length} 步）`, 'info', 2200);
}

function redo() {
  if (history.idx >= history.stack.length - 1) { notify('没有可重做的操作', 'info', 2600); return; }
  history.idx += 1;
  applySnapshot(history.stack[history.idx]);
  history.at = 0;
  updateHistoryUI();
  notify(`已重做一步（当前第 ${history.idx + 1}/${history.stack.length} 步）`, 'info', 2200);
}

/** 刷新「撤销/重做」按钮的可用状态与提示 */
function updateHistoryUI() {
  const u = $('#btn-undo'), r = $('#btn-redo');
  if (u) {
    u.disabled = history.idx <= 0;
    u.title = `撤销上一步（Ctrl+Z）　可回退 ${Math.max(0, history.idx)} 步`;
  }
  if (r) {
    const n = Math.max(0, history.stack.length - 1 - history.idx);
    r.disabled = n <= 0;
    r.title = `重做下一步（Ctrl+Y / Ctrl+Shift+Z）　可前进 ${n} 步`;
  }
}

function saveLocal() {
  historyPush(false);                        // 记录设计历史（撤销/重做）
  try {
    localStorage.setItem('label-print-draft', JSON.stringify({
      name: state.name, label: state.label, elements: state.elements,
      zoom: state.zoom, zoomExplicit: state.zoomExplicit,
      headers: state.headers, rows: state.rows, qty: state.qty,
      currentRow: state.currentRow, printer: state.printer,
    }));
  } catch (e) { /* 数据过大时忽略 */ }
  if (state.name) setTplStatus(false);   // 有改动 → 标记为“未保存”
}
function restoreLocal() {
  try {
    const d = JSON.parse(localStorage.getItem('label-print-draft') || 'null');
    if (!d) return;
    Object.assign(state, {
      name: d.name || '', label: d.label || state.label,
      elements: (d.elements || []).map(tblSanitize),
      // 缩放：用户手动选过就沿用；否则一律用默认 150%（老草稿里存的旧默认 100% 不再沿用）
      zoom: d.zoomExplicit ? (+d.zoom || DEFAULT_ZOOM) : DEFAULT_ZOOM,
      zoomExplicit: !!d.zoomExplicit,
      headers: d.headers || [], rows: d.rows || [],
      qty: Array.isArray(d.qty) ? d.qty : [],
      currentRow: d.currentRow ?? -1,
      printer: d.printer || '',
    });
    syncPaperUI();
    $('#zoom').value = String(state.zoom);
    if (state.rows.length) {
      renderTable();                       // 草稿恢复后直接展示表格，不再输出整片提示文字
      $('#import-report').classList.add('hidden');
      $('#import-info').textContent = `已恢复草稿数据 ${state.rows.length} 行`;
    }
  } catch (e) { /* ignore */ }
}

/* ============================================================
 * 功能区页签切换（Excel 风格）
 * ============================================================ */
document.querySelectorAll('.rtab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.rtab').forEach(t => t.classList.toggle('active', t === tab));
    document.querySelectorAll('.rpanel').forEach(p => p.classList.toggle('active', p.dataset.tab === tab.dataset.tab));
    if (tab.dataset.tab === 'update') loadUpdateStatus();      // 切到版本日志页签时刷新版本/更新状态
  });
});

/* ============================================================
 * 左侧标签模板列表（点击即载入）
 * ============================================================ */
let tplListCache = [];

async function loadTplList(silent) {
  const box = $('#tpl-list');
  if (!silent && !tplListCache.length) box.innerHTML = '<p class="muted small" style="padding:10px">正在读取模板…</p>';
  try {
    const res = await fetch('/api/labels', { cache: 'no-store' });
    const j = await res.json();
    if (!j.ok) throw new Error(j.message);
    tplListCache = j.list;
    renderTplList();
    return true;
  } catch (e) {
    box.innerHTML = `<p class="muted small" style="padding:10px">读取模板失败：${escHtml(e.message)}<br>`
      + `请确认后端已启动，并通过 ${siteOrigin()} 访问本页面。</p>`;
    return false;
  }
}

function renderTplList() {
  const box = $('#tpl-list');
  if (!tplListCache.length) {
    box.innerHTML = '<p class="muted small" style="padding:10px">暂无已保存的模板。<br>'
      + '设计好标签后点击「💾 保存模板」即可保存到这里。</p>';
    return;
  }
  box.innerHTML = tplListCache.map(t => `
    <div class="tpl-card${t.name === state.name ? ' active' : ''}" data-name="${escHtml(t.name)}">
      <button class="tc-del" data-del="${escHtml(t.name)}" title="删除该模板">🗑</button>
      <div class="tc-name">${escHtml(t.name)}</div>
      <div class="tc-meta">${t.elementCount ?? '-'} 个元素 · ${t.width || '?'}×${t.height || '?'}mm</div>
      <div class="tc-meta">${t.printer ? '🖨 ' + escHtml(t.printer) + '<br>' : ''}${t.updatedAt ? new Date(t.updatedAt).toLocaleString() : ''}</div>
    </div>`).join('');

  box.querySelectorAll('.tpl-card').forEach(card => card.addEventListener('click', async e => {
    if (e.target.closest('[data-del]')) return;
    await loadTemplate(card.dataset.name);
  }));
  box.querySelectorAll('[data-del]').forEach(btn => btn.addEventListener('click', async e => {
    e.stopPropagation();
    const name = btn.dataset.del;
    const info = await loadTemplate(name, true);
    const desc = info ? `${(info.elements || []).length} 个元素 · ${info.label ? info.label.width + '×' + info.label.height + 'mm' : ''}` : '';
    const yes = await askConfirm('删除标签模板',
      `确定删除模板「${name}」吗？${desc ? '（' + desc + '）' : ''}该操作会删除服务端 data/labels/${name}.json，不可恢复。`, '删除');
    if (!yes) return;
    try {
      const res = await fetch('/api/labels/' + encodeURIComponent(name), { method: 'DELETE' });
      const j = await res.json();
      if (!j.ok) throw new Error(j.message);
      if (state.name === name) { state.name = ''; saveLocal(); setTplStatus(false); }
      notify(`模板「${name}」已删除`, 'success');
      loadTplList();
    } catch (err) { notify('删除失败：' + err.message, 'error', 7000); }
  }));
}
$('#btn-tpl-list-refresh').addEventListener('click', () => { loadTplList(); loadPrinters(true); });

/* ============================================================
 * 打印机识别与状态监测
 * ============================================================ */
const PR_STATUS_LABEL = { idle: '空闲', printing: '打印中', offline: '脱机', paused: '已暂停', error: '异常', unknown: '未知' };

/* ---------- 打印服务（打印代理）：支持打印到局域网内其他设备的打印机 ----------
 * 浏览器出于安全限制无法枚举“打开网页这台设备”的打印机，
 * 因此每台要打印的设备需运行本程序作为打印服务，页面把任务发到对应服务。
 * 空字符串 = 当前打开页面的服务端（默认） */
const AGENT_KEY = 'print-agent';
const AGENTS_KEY = 'print-agents';

function agentBase() {
  const a = state.printerAgent || '';
  return a ? a.replace(/\/+$/, '') : '';      // '' → 使用相对路径（当前服务端）
}
function apiUrl(path) { return agentBase() + path; }
function isRemoteAgent() { return !!agentBase(); }

function savedAgents() {
  try { return JSON.parse(localStorage.getItem(AGENTS_KEY) || '[]'); } catch (e) { return []; }
}
function saveAgentList(list) {
  try { localStorage.setItem(AGENTS_KEY, JSON.stringify([...new Set(list)].slice(0, 8))); } catch (e) { }
}
function renderAgentSelect() {
  const sel = $('#agent-select');
  const list = savedAgents();
  sel.innerHTML = `<option value="">当前服务端（本机打印机）</option>`
    + list.map(a => `<option value="${escHtml(a)}"${a === agentBase() ? ' selected' : ''}>${escHtml(a)}</option>`).join('');
  sel.value = agentBase();
  sel.classList.toggle('remote', !!agentBase());
  sel.title = agentBase()
    ? `当前打印由 ${agentBase()} 执行（该设备上的打印机）`
    : '当前打印由本页面的服务端执行（其所在电脑的打印机）';
}

/** 切换打印服务（会重新加载该服务上的打印机列表） */
async function switchAgent(addr, silent) {
  const base = (addr || '').replace(/\/+$/, '');
  state.printerAgent = base;
  try { localStorage.setItem(AGENT_KEY, base); } catch (e) { }
  state.printer = '';
  state.printers = [];
  renderAgentSelect();
  const ok = await loadPrinters(true);
  if (ok) {
    const n = (state.printers || []).length;
    if (base && !silent) notify(`已切换打印服务：${base}（该设备上有 ${n} 台打印机）`, 'success', 5000);
  }
  return ok;
}

$('#agent-select').addEventListener('change', e => switchAgent(e.target.value));
$('#btn-agent-set').addEventListener('click', async () => {
  const port = location.port || '11235';
  const addr = await askText('添加局域网打印服务', `http://其他设备IP:${port}`);
  if (!addr) return;
  const base = addr.replace(/\/+$/, '');
  if (!/^https?:\/\//i.test(base)) { notify('请填写完整地址，例如 http://192.168.1.33:11235', 'warn', 6000); return; }
  // 校验该地址确实是本程序的打印服务
  try {
    const r = await fetch(base + '/api/network', { cache: 'no-store' });
    const j = await r.json();
    if (!j.ok) throw new Error(j.message || '接口异常');
    saveAgentList([...savedAgents(), base]);
    await switchAgent(base);
    notify(`已添加打印服务 ${base}（主机 ${j.hostname}）。若要打印该设备上的打印机，请在目标设备上运行本程序。`, 'success', 7000);
  } catch (e) {
    notify(`无法连接打印服务 ${base}：${e.message}。请确认目标设备已运行本程序，且其防火墙已放行端口（npm run lan）。`, 'error', 9000);
  }
});

/** 选择/记录目标打印机（随模板保存） */
function setPrinter(name, opt = {}) {
  state.printer = name || '';
  const sel = $('#printer-select');
  if (sel && name && ![...sel.options].some(o => o.value === name)) {
    sel.insertAdjacentHTML('beforeend', `<option value="${escHtml(name)}">${escHtml(name)}</option>`);
  }
  if (sel) sel.value = state.printer;
  if (!opt.noSave) saveLocal();
  if (name && !opt.silent) { infoBar(`目标打印机：${name}`); if (state.name) setTplStatus(false); }
  renderPrinterStatus();
}

function renderPrinterStatus() {
  const el = $('#printer-status');
  const p = (state.printers || []).find(x => x.name === state.printer);
  if (!state.printer) {
    el.className = 'pr-status unknown';
    el.textContent = (state.printers || []).length ? '未选择' : '未检测到';
    return;
  }
  if (!p) {
    el.className = 'pr-status unknown';
    el.textContent = '状态未知';
    return;
  }
  el.className = 'pr-status ' + p.statusCode;
  el.textContent = `${p.statusText}${p.isDefault ? ' · 默认' : ''}`;
  el.title = `${p.name}\n驱动：${p.driver || '-'}\n端口：${p.port || '-'}\n状态：${p.statusText}`
    + `\n队列：${p.jobs} 个任务${p.vendor ? '\n品牌：' + p.vendor : ''}`
    + (p.statusCode === 'offline' ? '\n提示：设备可能未连接，或已勾选「脱机使用打印机」' : '');
}

/** 打印页签中的打印机概览卡片 */
function renderPrinterCards() {
  const box = $('#printer-cards');
  const list = state.printers || [];
  if (!list.length) {
    box.innerHTML = '<span class="muted small">未检测到本机打印机（或后端不在 Windows 环境）</span>';
    return;
  }
  box.innerHTML = list.slice(0, 6).map(p => `
    <span class="pr-card ${p.statusCode}">
      <span class="dot"></span>
      <b>${escHtml(p.name.length > 22 ? p.name.slice(0, 22) + '…' : p.name)}</b>
      <span class="muted">${escHtml(p.statusText)}</span>
      ${p.jobs ? `<span class="muted">队列 ${p.jobs}</span>` : ''}
      ${p.isDefault ? '<span class="muted">★</span>' : ''}
    </span>`).join('') + (list.length > 6 ? `<span class="muted small">等共 ${list.length} 台，点「详情」查看</span>` : '');
}

function renderPrinterTable() {
  const wrap = $('#printer-table-wrap');
  const list = state.printers || [];
  const agentLabel = isRemoteAgent() ? `打印服务：${agentBase()}　` : '';
  $('#printer-updated').textContent = state.printersUpdatedAt
    ? `${agentLabel}最近更新：${state.printersUpdatedAt}${state.printersError ? '（' + state.printersError + '）' : ''}`
    : (agentLabel + (state.printersError || '—'));
  if (!list.length) {
    wrap.innerHTML = '<p class="muted" style="padding:12px">未检测到本机打印机。'
      + (state.printersError ? '<br>检测失败：' + escHtml(state.printersError) : '')
      + '<br>提示：打印机识别依赖 Windows 的 Get-Printer / Win32_Printer，请确认已安装打印机且 Print Spooler（打印后台处理程序）服务正在运行。</p>';
    return;
  }
  const s = state.printersSummary || {};
  wrap.innerHTML = `<p class="muted" style="margin-bottom:6px">
      共 <b>${s.total || 0}</b> 台：空闲 ${s.idle || 0} · 打印中 ${s.busy || 0} · 脱机 ${s.offline || 0} · 异常 ${s.error || 0}
      · 队列合计 ${s.jobs || 0} 个任务
    </p>
    <table id="printer-table"><thead><tr>
      <th>打印机名称</th><th>品牌</th><th>驱动/型号</th><th>端口</th><th>状态</th><th>队列</th><th>操作</th>
    </tr></thead><tbody>${list.map(p => `
      <tr class="${p.isDefault ? 'is-default' : ''}">
        <td>${escHtml(p.name)}${p.isDefault ? ' <span class="muted">默认</span>' : ''}</td>
        <td class="vendor">${escHtml(p.vendor || '-')}</td>
        <td>${escHtml(p.driver || '-')}</td>
        <td>${escHtml(p.port || '-')}</td>
        <td><span class="st ${p.statusCode}">${escHtml(p.statusText)}</span></td>
        <td>${p.jobs}</td>
        <td><button data-pick="${escHtml(p.name)}">${p.name === state.printer ? '已选定' : '选为打印目标'}</button></td>
      </tr>`).join('')}</tbody></table>`;

  wrap.querySelectorAll('[data-pick]').forEach(b => b.addEventListener('click', () => {
    setPrinter(b.dataset.pick);
    renderPrinterTable();
    notify(`已选定打印目标：${b.dataset.pick}（随模板一起保存）`, 'success', 4000);
  }));
}

async function loadPrinters(force) {
  try {
    const res = await fetch('/api/printers' + (force ? '?refresh=1' : ''), { cache: 'no-store' });
    const j = await res.json();
    const list = j.printers || [];
    state.printers = list;
    state.printersSummary = j.summary || null;
    state.printersError = j.ok ? '' : (j.message || '检测失败');
    state.printersUpdatedAt = j.updatedAt ? new Date(j.updatedAt).toLocaleTimeString() : '';
    // 保留用户选择；未选择时默认取系统默认打印机
    if (!state.printer || !list.some(p => p.name === state.printer)) {
      const def = list.find(p => p.isDefault);
      if (def || list.length) setPrinter((def || list[0]).name, { silent: true, noSave: true });
    }
    renderPrinterSelect();
    renderPrinterStatus();
    renderPrinterCards();
    if (!$('#printer-mask').classList.contains('hidden')) renderPrinterTable();
    return true;
  } catch (e) {
    state.printers = [];
    state.printersError = e.message;
    renderPrinterSelect(); renderPrinterStatus(); renderPrinterCards();
    if (!$('#printer-mask').classList.contains('hidden')) renderPrinterTable();
    if (force) notify('打印机检测失败：' + e.message, 'error', 7000);
    return false;
  }
}

function renderPrinterSelect() {
  const sel = $('#printer-select');
  const list = state.printers || [];
  if (!list.length) {
    sel.innerHTML = `<option value="">${state.printersError ? '检测失败' : '未检测到打印机'}</option>`;
    return;
  }
  sel.innerHTML = list.map(p => `<option value="${escHtml(p.name)}"${p.name === state.printer ? ' selected' : ''}>`
    + `${p.isDefault ? '★ ' : ''}${escHtml(p.name)} — ${escHtml(PR_STATUS_LABEL[p.statusCode] || p.statusText)}`
    + `${p.jobs ? '（队列 ' + p.jobs + '）' : ''}</option>`).join('');
}

$('#printer-select').addEventListener('change', e => setPrinter(e.target.value));
$('#btn-printer-refresh').addEventListener('click', async () => {
  infoBar('正在重新检测本机打印机…');
  await loadPrinters(true);
  notify(`打印机状态已刷新：共 ${(state.printers || []).length} 台`, 'success', 3000);
});
$('#btn-printer-detail').addEventListener('click', () => { $('#printer-mask').classList.remove('hidden'); renderPrinterTable(); loadPrinters(true); });

/* ---------- 打印机诊断：为什么某台打印机没出现在列表里 ---------- */
async function loadPrinterDiag(force) {
  const box = $('#printer-diag');
  box.classList.remove('hidden');
  box.innerHTML = '<p class="muted">正在采集本机驱动、端口与设备信息…（首次约需数秒）</p>';
  try {
    const j = await (await fetch(apiUrl('/api/printers/diag') + (force ? '?refresh=1' : ''), { cache: 'no-store' })).json();
    if (!j.ok) throw new Error(j.message || '诊断失败');
    const drivers = (j.drivers || []);
    const zebra = j.zebraDrivers || [];
    const findings = (j.findings || []).map(f =>
      `<div class="finding ${f.level}"><b>${escHtml(f.title)}</b>${escHtml(f.detail)}</div>`).join('');
    box.innerHTML = `<h4>诊断结果</h4>${findings}
      <div class="kv" style="margin:6px 0">本机打印队列 ${(j.queues || []).length} 个 ·
        已安装驱动 ${drivers.length} 个（其中斑马 ${zebra.length} 个）</div>
      ${drivers.length ? `<details><summary>查看已安装驱动清单</summary>
        <div class="kv" style="margin-top:4px">${drivers.map(d =>
          escHtml(d.name) + (d.zebra ? ' <b style="color:#1a7f37">[斑马]</b>' : '') + '　').join('<br>')}</div></details>` : ''}
      ${(j.unusedTcpPorts || []).length ? `<details><summary>未使用的端口详情</summary>
        <div class="kv" style="margin-top:4px">${(j.unusedTcpPorts).map(p =>
          escHtml(p.name) + '（' + escHtml(p.host || p.description || '') + '）').join('<br>')}</div></details>` : ''}
      <h4 style="margin-top:10px">如果打印机没出现在上面的列表里</h4>
      <div class="finding info"><b>工具只列出「Windows 已安装的打印机」</b>
        若打印机只是插着 USB / 接了网线，但没装驱动、没在系统里添加打印机，本工具就看不到它（Windows 自身的「设备和打印机」里也没有）。</div>
      <div class="kv" style="margin:6px 0">
        1. 先确认 Windows「设置 → 蓝牙和其他设备 → 打印机和扫描仪」里能看到该打印机；看不到就是系统层面未安装。<br>
        2. 斑马打印机请先安装 <b>ZDesigner 驱动</b>（官网按型号下载，如 GX430t）。<br>
        3. 然后用下面的命令把它添加进系统（以管理员身份运行，命令已按本机情况生成）：
      </div>
      <pre>${escHtml(j.addHint || '')}</pre>
      <div class="kv">添加成功后回到本页面点「↻ 立即刷新」，即可在打印机下拉中选择它。</div>`;
  } catch (e) {
    box.innerHTML = `<p class="muted">诊断失败：${escHtml(e.message)}</p>`;
  }
}
$('#btn-printer-diag').addEventListener('click', () => loadPrinterDiag(true));
$('#printer-refresh2').addEventListener('click', () => loadPrinters(true));
$('#printer-close').addEventListener('click', () => $('#printer-mask').classList.add('hidden'));
$('#printer-mask').addEventListener('click', e => { if (e.target.id === 'printer-mask') $('#printer-mask').classList.add('hidden'); });

/* ============================================================
 * 局域网访问地址（供其他主机访问本工具站）
 * ============================================================ */
async function openLanDialog() {
  $('#lan-mask').classList.remove('hidden');
  $('#lan-list').innerHTML = '<span class="muted small">正在获取…</span>';
  try {
    const j = await (await fetch('/api/network', { cache: 'no-store' })).json();
    $('#lan-local').textContent = j.localUrl || '—';
    // 端口动态展示（换端口后提示自动跟随）
    document.querySelectorAll('#lan-mask .lan-port').forEach(el => { el.textContent = j.port || ''; });
    if ($('#lan-cmd-1')) $('#lan-cmd-1').textContent = `scripts\\fix-lan-firewall.ps1 -UnblockNode -Port ${j.port}`;
    const urls = j.urls || [];
    $('#lan-list').innerHTML = urls.length
      ? urls.map((u, i) => `<div class="lan-item">
          <code id="lan-url-${i}">${escHtml(u)}</code>
          <button data-copy="lan-url-${i}">复制</button>
          <span class="iface">${escHtml((j.interfaces[i] || {}).iface || '')}</span></div>`).join('')
      : '<span class="muted small">未检测到局域网地址（可能只有回环网卡）</span>';
    bindCopyButtons();
  } catch (e) {
    $('#lan-list').innerHTML = `<span class="muted small">获取失败：${escHtml(e.message)}</span>`;
  }
}

function bindCopyButtons() {
  document.querySelectorAll('#lan-mask [data-copy]').forEach(btn => {
    btn.onclick = async () => {
      const text = ($('#' + btn.dataset.copy) || {}).textContent || '';
      try {
        await navigator.clipboard.writeText(text);
        btn.textContent = '已复制';
        setTimeout(() => { btn.textContent = '复制'; }, 1500);
      } catch (e) {
        // 无剪贴板权限时退化为全选提示
        notify('复制失败，请手动选择文本：' + text, 'warn', 5000);
      }
    };
  });
}

$('#server-status').addEventListener('click', openLanDialog);
$('#lan-close').addEventListener('click', () => $('#lan-mask').classList.add('hidden'));
$('#lan-mask').addEventListener('click', e => { if (e.target.id === 'lan-mask') $('#lan-mask').classList.add('hidden'); });

/* 每 10 秒自动刷新打印机状态（页面不可见时暂停） */
setInterval(() => {
  if (document.visibilityState === 'visible') loadPrinters(false);
}, 10000);

/* ============================================================
 * 清空已导入数据
 * ============================================================ */
$('#btn-clear-data').addEventListener('click', async () => {
  if (!state.rows.length) { notify('当前没有已导入的数据', 'warn'); return; }
  const yes = await askConfirm('清空数据', `将清空已导入的 ${state.rows.length} 行标签数据（不影响标签模板），是否继续？`, '清空');
  if (!yes) return;
  state.rows = []; state.headers = []; state.qty = []; state.currentRow = -1;
  $('#import-report').classList.add('hidden');
  $('#import-info').textContent = '';
  renderTable(); renderCanvas(); renderProps(); saveLocal();
  notify('已清空导入的数据', 'success');
});

/* 统一设置所有行的打印数量 */
$('#qty-apply') && $('#qty-apply').addEventListener('click', () => {
  if (!state.rows.length) { notify('请先导入 Excel 数据', 'warn'); return; }
  let v = parseInt($('#qty-all').value, 10);
  if (!Number.isFinite(v) || v < 1) { notify('请输入 ≥ 1 的份数', 'warn'); return; }
  v = Math.min(999, v);
  state.qty = state.rows.map(() => v);
  renderTable(); saveLocal();
  notify(`已将全部 ${state.rows.length} 行的打印数量统一设为 ${v} 份`, 'success', 3000);
});

/* ============================================================
 * 使用手册（内置离线文档）
 * 文档来源：public/docs/使用手册.md（由 scripts/sync-manual.js 从项目根同步）
 * 自带轻量 Markdown 渲染，不依赖任何第三方库，断网亦可用
 * ============================================================ */
/** 转义 HTML */
function mdEsc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
/** 行内格式：粗体 / 行内代码 / 链接 / 复选框 */
function mdInline(s) {
  let t = mdEsc(s);
  t = t.replace(/`([^`]+)`/g, '<code>$1</code>');
  t = t.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  t = t.replace(/(https?:\/\/[^\s<）)，。]+)/g, '<a href="$1" target="_blank" rel="noreferrer">$1</a>');
  t = t.replace(/^\[ \]\s*/, '<input type="checkbox" disabled> ')
    .replace(/^\[[xX]\]\s*/, '<input type="checkbox" disabled checked> ');
  return t;
}
/** 轻量 Markdown → HTML（支持标题/表格/列表/引用/代码块/分隔线/粗体/行内代码/链接/复选框） */
function renderMarkdown(md) {
  const lines = String(md).replace(/\r\n?/g, '\n').split('\n');
  const out = [];
  const toc = [];
  let i = 0, headIdx = 0, inCode = false, codeBuf = [];

  const isTableSep = s => /^\s*\|?[\s:|-]+\|[\s:|-]*$/.test(s) && s.includes('-');
  const cells = s => s.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(x => x.trim());

  while (i < lines.length) {
    const line = lines[i];

    if (/^```/.test(line)) {                       // 代码块
      if (inCode) { out.push('<pre><code>' + mdEsc(codeBuf.join('\n')) + '</code></pre>'); codeBuf = []; inCode = false; }
      else inCode = true;
      i++; continue;
    }
    if (inCode) { codeBuf.push(line); i++; continue; }

    if (/^\s*$/.test(line)) { i++; continue; }

    if (/^\s*(---+|\*\*\*+)\s*$/.test(line)) { out.push('<hr>'); i++; continue; }

    let m = /^(#{1,4})\s+(.*)$/.exec(line);        // 标题
    if (m) {
      const lv = m[1].length;
      const id = 'mdh' + (++headIdx);
      if (lv === 2) toc.push(`<a href="#${id}">${mdEsc(m[2].trim())}</a>`);
      out.push(`<h${lv} id="${id}">${mdInline(m[2].trim())}</h${lv}>`);
      i++; continue;
    }

    if (/^>/.test(line)) {                          // 引用块
      const buf = [];
      while (i < lines.length && /^>/.test(lines[i])) { buf.push(lines[i].replace(/^>\s?/, '')); i++; }
      out.push('<blockquote><p>' + buf.map(mdInline).join('<br>') + '</p></blockquote>');
      continue;
    }

    if (line.includes('|') && i + 1 < lines.length && isTableSep(lines[i + 1])) {   // 表格
      const head = cells(line);
      i += 2;
      const rows = [];
      while (i < lines.length && lines[i].includes('|') && !/^\s*$/.test(lines[i])) { rows.push(cells(lines[i])); i++; }
      out.push('<table><thead><tr>' + head.map(h => `<th>${mdInline(h)}</th>`).join('')
        + '</tr></thead><tbody>' + rows.map(r => '<tr>'
          + head.map((_, ci) => `<td>${mdInline(r[ci] === undefined ? '' : r[ci])}</td>`).join('') + '</tr>').join('')
        + '</tbody></table>');
      continue;
    }

    if (/^\s*[-*+]\s+/.test(line)) {                // 无序列表（含任务清单）
      const items = [];
      while (i < lines.length && /^\s*[-*+]\s+/.test(lines[i])) {
        items.push(lines[i].replace(/^\s*[-*+]\s+/, ''));
        i++;
      }
      out.push('<ul>' + items.map(x => `<li>${mdInline(x.trim())}</li>`).join('') + '</ul>');
      continue;
    }
    if (/^\s*\d+[.)]\s+/.test(line)) {              // 有序列表
      const items = [];
      while (i < lines.length && /^\s*\d+[.)]\s+/.test(lines[i])) {
        items.push(lines[i].replace(/^\s*\d+[.)]\s+/, ''));
        i++;
      }
      out.push('<ol>' + items.map(x => `<li>${mdInline(x.trim())}</li>`).join('') + '</ol>');
      continue;
    }

    // 普通段落（合并连续行）
    const buf = [line];
    i++;
    while (i < lines.length && !/^\s*$/.test(lines[i]) && !/^[#>|`]/.test(lines[i])
      && !/^\s*[-*+]\s+/.test(lines[i]) && !/^\s*\d+[.)]\s+/.test(lines[i]) && !/^\s*(---+|\*\*\*+)\s*$/.test(lines[i])) {
      buf.push(lines[i]); i++;
    }
    out.push('<p>' + buf.map(mdInline).join('<br>') + '</p>');
  }
  if (inCode && codeBuf.length) out.push('<pre><code>' + mdEsc(codeBuf.join('\n')) + '</code></pre>');

  const tocHtml = toc.length
    ? `<div class="toc"><b>目录：</b>${toc.join('')}</div>` : '';
  return tocHtml + out.join('\n');
}

/* 文档查看器：使用手册与版本日志共用（Markdown 渲染 + 打印） */
const DOCS = {
  manual: { title: '使用手册（安装版）', url: '/docs/使用手册.md', cache: '' },
  changelog: { title: '版本日志（每个版本的版本号与更新内容）', url: '/docs/CHANGELOG.md', cache: '' },
};
let curDocKey = 'manual';

async function openDoc(key, force) {
  const doc = DOCS[key] || DOCS.manual;
  curDocKey = DOCS[key] ? key : 'manual';
  $('#manual-title').textContent = doc.title;
  const box = $('#manual-body');
  $('#manual-mask').classList.remove('hidden');
  if (doc.cache && !force) { box.innerHTML = renderMarkdown(doc.cache); box.scrollTop = 0; return; }
  box.innerHTML = `<p class="muted">正在加载${escHtml(doc.title)}…</p>`;
  try {
    const res = await fetch(doc.url + (force ? '?t=' + Date.now() : ''), { cache: 'no-store' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const md = await res.text();
    if (!/^#\s/.test(md.trim())) throw new Error('文档内容异常');
    doc.cache = md;
    box.innerHTML = renderMarkdown(md);
    box.scrollTop = 0;
  } catch (e) {
    box.innerHTML = `<p class="muted">${escHtml(doc.title)}加载失败：${escHtml(e.message)}</p>`
      + '<p class="muted">可查看软件安装目录下的 <code>resources\\docs\\</code>，'
      + '或源码目录中对应的 Markdown 文件。</p>';
  }
}
const openManual = force => openDoc('manual', force);
/** 打印手册（新窗口渲染后调用系统打印） */
function printManual() {
  const doc = DOCS[curDocKey] || DOCS.manual;
  if (!doc.cache) { notify('文档尚未加载完成，请稍后再试', 'warn'); return; }
  const w = window.open('', '_blank', 'width=900,height=700');
  if (!w) { notify('打印窗口被拦截，请允许本站弹出窗口', 'error'); return; }
  w.document.write(`<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">
    <title>${escHtml(doc.title)} · 标签打印工具站</title>
    <style>
      body{font-family:"微软雅黑",sans-serif;line-height:1.8;color:#24292f;padding:22mm 18mm;font-size:13px;}
      h1{font-size:22px;border-bottom:2px solid #1668dc;padding-bottom:6px;}
      h2{font-size:18px;border-left:4px solid #1668dc;padding-left:8px;margin-top:22px;}
      h3{font-size:15px;color:#0b4c9c;} table{border-collapse:collapse;width:100%;font-size:12px;margin:10px 0;}
      th,td{border:1px solid #ccc;padding:5px 8px;} th{background:#f2f2f2;}
      code{background:#f0f2f5;padding:1px 4px;border-radius:3px;} pre{background:#0d1117;color:#d1d5da;padding:10px;border-radius:5px;}
      blockquote{background:#f5faff;border-left:4px solid #1668dc;margin:10px 0;padding:8px 12px;}
      .toc{display:none;} @page{margin:14mm;}
    </style></head><body>${renderMarkdown(doc.cache).replace(/<div class="toc">[\s\S]*?<\/div>/, '')}</body></html>`);
  w.document.close();
  setTimeout(() => { try { w.focus(); w.print(); } catch (e) { } }, 350);
}

$('#btn-manual').addEventListener('click', () => openManual(false));
$('#manual-reload').addEventListener('click', () => openDoc(curDocKey, true));
$('#manual-print').addEventListener('click', printManual);
$('#manual-close').addEventListener('click', () => $('#manual-mask').classList.add('hidden'));
$('#manual-mask').addEventListener('click', e => { if (e.target.id === 'manual-mask') $('#manual-mask').classList.add('hidden'); });
$('#btn-changelog').addEventListener('click', () => openDoc('changelog', false));
window.openManual = openManual;      // 供桌面版菜单调用
window.openChangelog = () => openDoc('changelog', false);

/* ============================================================
 * 下载安装包（右上角「⬇ 下载安装包」）
 * ------------------------------------------------------------
 * 服务器上的安装包由 /api/installer 扫描（搜索目录见 server.js）：
 * 运维/桌面版用户把 LabelPrint-Setup-*.exe 放到 <数据目录>\installer\ 或程序目录 dist\，
 * 局域网内任何电脑打开网页即可一键下载安装（也用于「检查更新」提示需装完整包时）。
 * ============================================================ */
let installerInfo = null;

/** 刷新按钮悬停提示（显示服务器上可下载的版本与大小） */
async function refreshInstallerBtn() {
  const btn = $('#btn-installer');
  if (!btn) return null;
  try {
    const j = await (await fetch('/api/installer', { cache: 'no-store' })).json();
    installerInfo = j;
    btn.classList.toggle('muted-btn', !j.available);
    btn.title = j.available
      ? `下载安装包 ${j.name}\n版本 V${j.version}　大小 ${j.sizeText}\n可保存到本机安装，也可发给其他电脑安装`
      : '下载最新版安装包（当前服务器上还没有安装包文件，点击查看放置位置）';
    return j;
  } catch (e) {
    btn.title = '下载最新版安装包（无法获取安装包信息）';
    return null;
  }
}

/** 触发浏览器下载（大文件用 <a download>，不占内存） */
function triggerInstallerDownload(url, name) {
  const a = document.createElement('a');
  a.href = url + (url.includes('?') ? '&' : '?') + 't=' + Date.now();
  if (name) a.download = name;
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  a.remove();
}

/** 点击「⬇ 下载安装包」：有安装包就下载，没有就给出放置路径 */
async function downloadInstaller() {
  const btn = $('#btn-installer');
  if (btn) btn.disabled = true;
  try {
    const j = (installerInfo && installerInfo.available) ? installerInfo : await refreshInstallerBtn();
    if (!j || !j.available) {
      const dirs = (j && j.searched) || [];
      notify('服务器上还没有安装包文件，已给出放置路径', 'warn', 8000);
      await dialog({
        title: '还没有可下载的安装包',
        message: '请把 LabelPrint-Setup-*.exe 放到下面任意一个目录，然后刷新本页面：\n\n'
          + dirs.join('\n')
          + '\n\n放好后点「⬇ 下载安装包」即可下载；局域网内其他电脑打开本页面也能下载。',
        okText: '知道了', cancelText: '关闭',
      });
      return;
    }
    triggerInstallerDownload(j.downloadUrl || '/api/installer/download', j.name);
    notify(`开始下载安装包：${j.name}（V${j.version}，${j.sizeText}）`, 'success', 6000);
    const running = (updCache.local && updCache.local.appVersion) || '';
    if (running && cmpVer(j.version, running) > 0) {
      infoBar(`注意：服务器上的安装包 V${j.version} 比当前运行版本 V${running} 新，安装后即为最新版`);
    }
  } catch (e) {
    notify('下载安装包失败：' + e.message, 'error', 7000);
  } finally { if (btn) btn.disabled = false; }
}

/** 从「更新源」下载安装包（更新源那台电脑上放了安装包时可用） */
async function downloadInstallerFromSource() {
  const base = updBase();
  const who = base || '当前服务端（本机）';
  try {
    const j = await fetchJson((base || '') + '/api/installer');
    if (!j.available) {
      notify(`更新源（${who}）上没有安装包文件，请在该电脑放入 LabelPrint-Setup-*.exe 后重试`, 'warn', 9000);
      return;
    }
    triggerInstallerDownload((base || '') + '/api/installer/download', j.name);
    notify(`开始从更新源下载安装包：${j.name}（V${j.version}，${j.sizeText}）`, 'success', 6000);
  } catch (e) {
    notify('从更新源下载安装包失败：' + e.message, 'error', 8000);
  }
}

$('#btn-installer').addEventListener('click', downloadInstaller);
refreshInstallerBtn();

/* ============================================================
 * 版本号显示：核对安装包是否已更新到最新版
 * ============================================================ */
async function loadVersion() {
  const chip = $('#app-version');
  try {
    const j = await (await fetch('/api/version', { cache: 'no-store' })).json();
    if (!j.ok) throw new Error('接口异常');
    // 徽标显示「网页层实际生效版本」：正常情况下等于安装包版本；热更新后是新版本（标 ⁺）
    const web = j.webVersion || j.version;
    chip.textContent = 'V' + web + (j.hotUpdated ? '⁺' : '');
    chip.title = `标签打印工具站\n安装包（主程序）版本：V${j.version}\n网页层（界面）版本：V${web}`
      + (j.hotUpdated ? `（已应用热更新，标 ⁺）` : '')
      + (j.staleOverride ? `\n⚠ 检测到过期热更新 V${j.staleVersion}（已自动忽略）` : '')
      + (j.stalePurged ? `\n✔ 启动时已清理过期热更新 V${(j.stalePurgeInfo || {}).version || ''}`
        + `（${(j.stalePurgeInfo || {}).files || 0} 个文件），已回到安装包自带版本` : '')
      + (j.packaged ? `\n运行方式：桌面安装版 · Electron ${j.electron}` : '\n运行方式：网页版（浏览器访问）')
      + `\n数据目录：${j.dataDir}`
      + `\nNode ${j.node}`
      + '\n\n如果界面版本长期低于安装包版本，请 Ctrl+F5 强制刷新；仍不行则关闭程序重新打开。';
    return j;
  } catch (e) {
    chip.textContent = 'V?';
    chip.title = '无法获取版本信息（服务端未连接）';
    return null;
  }
}

/* ============================================================
 * 增量更新（网页界面层热更新）
 * ------------------------------------------------------------
 * 客户端把自己的网页层文件清单（相对路径 → SHA-256）报给「更新服务端」，
 * 服务端只返回有变化的文件；客户端在本机校验后写入覆盖目录（<数据目录>/web）即生效，
 * 不需要重装安装包。主程序层（Electron/内置服务/PowerShell 脚本）仍需完整安装包升级。
 * ============================================================ */
const UPD_SRC_KEY = 'label-print-update-src';
let updSrc = '';
try { updSrc = localStorage.getItem(UPD_SRC_KEY) || ''; } catch (e) { updSrc = ''; }
let updCache = { local: null, source: null, diff: null };

const updBase = () => (updSrc ? updSrc.replace(/\/+$/, '') : '');     // 空 = 当前服务端
const updUrl = p => updBase() + p;

/** 简单版本号比较：a>b 返回 1，a<b 返回 -1，相等 0 */
function cmpVer(a, b) {
  const pa = String(a || '0').split('.').map(n => parseInt(n, 10) || 0);
  const pb = String(b || '0').split('.').map(n => parseInt(n, 10) || 0);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const d = (pa[i] || 0) - (pb[i] || 0);
    if (d) return d > 0 ? 1 : -1;
  }
  return 0;
}

async function fetchJson(url, opt, timeoutMs = 15000) {
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), timeoutMs);
  try {
    const res = await fetch(url, { cache: 'no-store', signal: ctl.signal, ...(opt || {}) });
    const txt = await res.text();
    let j = null;
    try { j = JSON.parse(txt); } catch (e) { throw new Error(`返回内容不是 JSON（HTTP ${res.status}）`); }
    if (!res.ok || j.ok === false) throw new Error(j.message || `HTTP ${res.status}`);
    return j;
  } finally { clearTimeout(timer); }
}

const updVerTag = (v, cur) => cmpVer(v, cur) > 0
  ? `<span class="upd-tag upd-new">新版本 V${escHtml(v)}</span>`
  : `<span class="upd-tag">V${escHtml(v)}</span>`;

/** 渲染「当前版本 / 更新状态」区 */
function renderUpdateInfo(html) {
  const box = $('#upd-info');
  if (!box) return;
  const st = updCache.local;
  const srcLabel = updSrc ? escHtml(updSrc) : '当前服务端（本机）';
  if (html) { box.innerHTML = html; return; }
  const lines = [];
  if (st) {
    // 两个版本分开显示：安装包（主程序）版本 与 网页层（界面）实际生效版本
    lines.push(`安装包版本：<b>V${escHtml(st.appVersion || '?')}</b>　网页层版本：<b>V${escHtml(st.webVersion || '?')}</b>`
      + (st.hotUpdated ? `　<span class="upd-ok">已热更新</span>（${st.overrideFiles} 个文件`
        + `${st.appliedAt ? ' · ' + escHtml(st.appliedAt.slice(0, 19).replace('T', ' ')) : ''}）` : ''));
    if (st.staleOverride) {
      lines.push(`<span class="upd-err">检测到过期热更新 V${escHtml(st.staleVersion)}（低于安装包 V${escHtml(st.appVersion)}）</span>`
        + `　<span class="muted">旧网页层已自动忽略，避免界面停在旧版本。</span>`
        + `　<button id="upd-clean-stale" title="删除过期热更新文件，回到安装包自带版本">清除过期文件</button>`
        + `　<button id="upd-reload" class="primary">🔄 刷新界面</button>`);
    }
  } else {
    lines.push('当前版本：读取中…');
  }
  lines.push(`更新源：${srcLabel}`);
  if (updCache.diff) {
    const d = updCache.diff;
    const srcApp = escHtml((updCache.source || {}).appVersion || d.to);
    if (d.downgrade) {
      lines.push(`<span class="upd-err">更新源网页层版本较低（V${escHtml(d.to)}）</span>`
        + `　<span class="muted">本机已是 V${escHtml((updCache.local || {}).webVersion || '')}，无需更新`
        + `（热更新只用于升级；如需降级请用「回滚到安装版本」或安装对应版本安装包）。</span>`);
    } else if (d.changedCount > 0) {
      lines.push(`<span class="upd-new">⬆ 发现可用更新：V${escHtml(d.to)}</span>`
        + `（${d.changedCount} 个文件有变化${d.changedCount < d.total ? `，共 ${d.total} 个` : ''}）`
        + `　<button id="upd-view-notes" title="查看更新源上的版本日志">查看更新内容</button>`);
      if (d.needInstaller === 'newer-app') {
        lines.push(`<span class="muted">注意：更新源主程序版本更高（V${srcApp}），本次只更新网页界面；`
          + `涉及主程序（Electron/内置服务/系统脚本）的改动仍需安装完整安装包。</span>`
          + `　<button id="upd-get-installer" title="从更新源直接下载安装包">⬇ 从更新源下载安装包</button>`);
      }
    } else if (d.needInstaller) {
      lines.push(`<span class="upd-err">更新源的主程序版本更高（V${srcApp}），但网页层文件一致</span>`
        + `——该更新涉及主程序（Electron/内置服务/系统脚本），请安装完整安装包。`
        + `　<button id="upd-get-installer" class="primary" title="从更新源直接下载安装包">⬇ 从更新源下载安装包</button>`);
    } else {
      lines.push(`<span class="upd-ok">✔ 已是最新版本</span>（网页层 V${escHtml(d.to)}，${d.total} 个网页文件一致）`);
    }
    if (d.changedCount > 0 && !d.downgrade) {
      lines.push(`<button id="upd-apply" class="primary" title="只下载有变化的文件，校验后写本机覆盖目录">⬇ 下载并应用更新</button>`
        + `　<span class="muted">应用后点「立即刷新界面」即可生效，无需重装安装包</span>`);
    }
  } else {
    lines.push('<span class="muted">点上方「🔄 检查更新」向更新源核对版本与文件差异。</span>');
  }
  box.innerHTML = lines.join('<br>');
  const notesBtn = $('#upd-view-notes');
  if (notesBtn) notesBtn.addEventListener('click', () => openRemoteChangelog());
  const applyBtn = $('#upd-apply');
  if (applyBtn) applyBtn.addEventListener('click', applyIncrementalUpdate);
  const cleanBtn = $('#upd-clean-stale');
  if (cleanBtn) cleanBtn.addEventListener('click', cleanStaleOverride);
  const getInstBtn = $('#upd-get-installer');
  if (getInstBtn) getInstBtn.addEventListener('click', downloadInstallerFromSource);
  const reloadBtn = $('#upd-reload');
  if (reloadBtn) reloadBtn.addEventListener('click', () => location.reload());
}

/** 清除热更新覆盖文件（含过期残留），界面回到安装包自带网页层；模板与数据不受影响 */
async function cleanStaleOverride() {
  const yes = await askConfirm('清除热更新覆盖文件',
    '将删除本机覆盖目录（<数据目录>\\web）里的网页层文件，界面回到安装包自带版本。\n'
    + '标签模板、Excel 数据与打印机设置都不受影响。\n\n是否继续？', '清除并刷新');
  if (!yes) return;
  try {
    const r = await fetchJson('/api/update/rollback', { method: 'POST' });
    notify(`已清除热更新文件（${r.removed} 项），正在刷新界面…`, 'success', 4000);
    setTimeout(() => location.reload(), 800);
  } catch (e) { notify('清除失败：' + e.message, 'error', 7000); }
}

/** 读取本机状态（安装包版本 / 网页层版本 / 是否已热更新） */
async function loadUpdateStatus() {
  try {
    updCache.local = await fetchJson('/api/update/status');
  } catch (e) { updCache.local = null; }
  renderUpdateInfo();
}

/** 打开更新源上的版本日志（对比前先看新版本更新了什么） */
async function openRemoteChangelog() {
  const url = updBase() ? updUrl('/docs/CHANGELOG.md') : '/docs/CHANGELOG.md';
  const box = $('#manual-body');
  $('#manual-title').textContent = '版本日志（来自更新源）';
  $('#manual-mask').classList.remove('hidden');
  box.innerHTML = '<p class="muted">正在从更新源读取版本日志…</p>';
  try {
    const res = await fetch(url + '?t=' + Date.now(), { cache: 'no-store' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const md = await res.text();
    if (!/^#\s/.test(md.trim())) throw new Error('内容异常');
    DOCS.remote = { title: '版本日志（更新源）', url, cache: md };
    curDocKey = 'remote';
    box.innerHTML = renderMarkdown(md);
    box.scrollTop = 0;
  } catch (e) {
    box.innerHTML = `<p class="muted">读取更新源版本日志失败：${escHtml(e.message)}</p>`;
  }
}

/** 检查更新：本机清单 vs 更新源清单（只算差异，不下发文件） */
async function checkUpdate() {
  const btn = $('#btn-check-update');
  if (btn) btn.disabled = true;
  renderUpdateInfo('<span class="muted">正在与更新源核对版本与文件差异…</span>');
  try {
    // 本机清单：version 已经是「网页层实际生效版本」（过期热更新会被忽略，不会低于安装包版本）
    const loc = await fetchJson('/api/update/manifest');
    const status = await fetchJson('/api/update/status');
    updCache.local = { ...status, version: loc.version, files: loc.files };
    const src = await fetchJson(updUrl('/api/update/manifest'));  // 更新源清单（同样是其生效网页层版本）
    updCache.source = src;
    // 只把「本机自己的清单」上报，服务端返回有变化的文件
    const d = await fetchJson(updUrl('/api/update/diff'), {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ version: loc.version, files: loc.files }),
    });
    // 版本比较：网页层对网页层（同口径），主程序版本另行比较
    d.verCmp = cmpVer(src.version, loc.version);
    d.downgrade = d.verCmp < 0;
    d.needInstaller = false;
    const appCmp = cmpVer(src.appVersion || src.version, loc.appVersion || loc.version);
    if (appCmp > 0) {
      if (d.changedCount > 0) d.needInstaller = 'newer-app';   // 网页层可先热更，主程序需安装包
      else d.needInstaller = true;
    }
    updCache.diff = d;
    renderUpdateInfo();
    if (d.downgrade) notify(`更新源网页层版本较低（V${d.to}），本机 V${loc.version} 无需更新`, 'warn', 5000);
    else if (d.changedCount > 0) notify(`发现可用更新：V${d.to}（${d.changedCount} 个文件有变化）`, 'success', 5000);
    else if (d.needInstaller === true) notify(`更新源主程序版本更高（V${src.appVersion || d.to}），网页层一致，请安装完整安装包`, 'warn', 7000);
    else notify(`已是最新版本（V${d.to}）`, 'info', 3500);
    return d;
  } catch (e) {
    renderUpdateInfo(`<span class="upd-err">检查更新失败：${escHtml(e.message)}</span><br>`
      + `<span class="muted">请确认更新源地址可访问（当前：${escHtml(updSrc || '本机')}）；`
      + `桌面安装版需要指定局域网内运行最新版软件的电脑，例如 http://192.168.1.23:11235。</span>`);
    notify('检查更新失败：' + e.message, 'error', 7000);
    return null;
  } finally { if (btn) btn.disabled = false; }
}

/** 下载并应用增量更新（只下变化的文件 → 本机覆盖目录） */
async function applyIncrementalUpdate() {
  const d = updCache.diff;
  if (!d || !d.changedCount) { notify('请先「检查更新」', 'warn'); return; }
  const yes = await askConfirm('应用增量更新',
    `将从更新源下载 ${d.changedCount} 个有变化的文件（约 ${(d.changed.reduce((s, f) => s + (f.size || 0), 0) / 1024).toFixed(0)} KB）`
    + `更新到 V${d.to}。\n文件会先做 SHA-256 校验，写入本机覆盖目录，不修改安装文件；`
    + `如有异常可用「回滚到安装版本」还原。\n\n是否继续？`, '下载并应用');
  if (!yes) return;
  const btn = $('#upd-apply');
  if (btn) { btn.disabled = true; btn.textContent = '下载中…'; }
  try {
    const loc = await fetchJson('/api/update/manifest');
    const diff = await fetchJson(updUrl('/api/update/diff'),
      { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ version: loc.version, files: loc.files }) });
    if (!diff.changedCount) { notify('更新源与本地一致，无需下载', 'info'); return; }
    const r = await fetchJson('/api/update/apply', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ version: diff.to, files: diff.changed }),
    });
    updCache.diff = null;
    await loadUpdateStatus();
    renderUpdateInfo(`<span class="upd-ok">✔ 更新已应用到本机（${r.written} 个文件 → V${escHtml(r.version)}）</span><br>`
      + `<span class="muted">点下面按钮刷新界面即可生效（无需重装安装包）。</span><br>`
      + `<button id="upd-reload" class="primary">🔄 立即刷新界面</button>`);
    const rb = $('#upd-reload');
    if (rb) rb.addEventListener('click', () => location.reload());
    notify(`增量更新已应用：${r.written} 个文件 → V${r.version}，刷新界面后生效`, 'success', 7000);
  } catch (e) {
    notify('应用更新失败：' + e.message, 'error', 8000);
    renderUpdateInfo(`<span class="upd-err">应用更新失败：${escHtml(e.message)}</span>`);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = '⬇ 下载并应用更新'; }
  }
}

/* 更新面板按钮 */
$('#btn-check-update').addEventListener('click', () => checkUpdate());
$('#btn-update-src').addEventListener('click', async () => {
  const v = await askText('更新源地址',
    updSrc || (location.protocol.startsWith('http') ? location.origin : ''));
  if (v === null) return;
  updSrc = /^https?:\/\//i.test(v) ? v.replace(/\/+$/, '') : '';
  try { localStorage.setItem(UPD_SRC_KEY, updSrc); } catch (e) { /* 忽略 */ }
  updCache.diff = null;
  notify(updSrc ? `更新源已设为：${updSrc}` : '更新源已设为：当前服务端（本机）', 'success', 4500);
  renderUpdateInfo();
});
$('#btn-update-rollback').addEventListener('click', async () => {
  const yes = await askConfirm('回滚到安装版本',
    '将清除本机已下载的增量更新文件，回到安装包自带版本。\n（安装目录内的文件从未被修改，数据与模板不受影响）', '回滚');
  if (!yes) return;
  try {
    const r = await fetchJson('/api/update/rollback', { method: 'POST' });
    await loadUpdateStatus();
    renderUpdateInfo(`<span class="upd-ok">✔ 已回滚到安装包自带版本</span>（清理 ${r.removed} 项）<br>`
      + `<button id="upd-reload" class="primary">🔄 立即刷新界面</button>`);
    const rb = $('#upd-reload');
    if (rb) rb.addEventListener('click', () => location.reload());
    notify('已回滚到安装包自带版本，刷新界面后生效', 'success', 6000);
  } catch (e) { notify('回滚失败：' + e.message, 'error', 7000); }
});

/* ---------- 初始化 ---------- */
buildPaperPresets();
restoreLocal();
try { state.printerAgent = localStorage.getItem(AGENT_KEY) || ''; } catch (e) { state.printerAgent = ''; }
renderAgentSelect();
syncPaperUI();
renderCanvas();
renderProps();
historyReset();                       // 以草稿恢复后的状态作为「上一步/下一步」的起点
setTplStatus(!!state.name);
/* 数据区默认收起（避免遮挡标签区）；仅当用户上次手动展开过才沿用展开状态 */
try { setDataPanel(localStorage.getItem(DATA_PANEL_KEY) !== 'expanded', { silent: true }); }
catch (e) { setDataPanel(true, { silent: true }); }
renderDataCount();
loadVersion();
loadUpdateStatus();
checkServer();
loadTplList();
loadPrinters(false);
