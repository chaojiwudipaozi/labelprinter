/**
 * 同步项目文档到应用内（供软件「❓ 使用手册 / 版本日志」在线读取，并随安装包一起分发）
 *   docs/使用手册.md   →  public/docs/使用手册.md
 *   docs/使用说明.md   →  public/docs/使用说明.md
 *   CHANGELOG.md       →  public/docs/CHANGELOG.md
 *
 * 用法：node scripts/sync-manual.js（`npm run dist` 会通过 predist 自动执行）
 * 说明：public/docs/ 下的文件是**同步产物**，改文档请改 docs/ 或根目录的 CHANGELOG.md。
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const OUT_DIR = path.join(ROOT, 'public', 'docs');
const MAP = [
  [path.join(ROOT, 'docs', '使用手册.md'), '使用手册.md'],
  [path.join(ROOT, 'docs', '使用说明.md'), '使用说明.md'],
  [path.join(ROOT, 'CHANGELOG.md'), 'CHANGELOG.md'],
];

fs.mkdirSync(OUT_DIR, { recursive: true });

let copied = 0, missing = [];
MAP.forEach(([src, outName]) => {
  const out = path.join(OUT_DIR, outName);
  if (!fs.existsSync(src)) { missing.push(path.relative(ROOT, src)); return; }
  const text = fs.readFileSync(src, 'utf-8');
  fs.writeFileSync(out, text, 'utf-8');
  copied++;
  console.log(`已同步 ${path.relative(ROOT, src)} → public/docs/${outName}（${text.length} 字符）`);
});
if (missing.length) console.log('未找到（跳过）：' + missing.join('、'));
console.log(`完成：同步 ${copied} 个文档到 ${path.relative(ROOT, OUT_DIR)}`);
