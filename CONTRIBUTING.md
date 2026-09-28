# 贡献指南

感谢你愿意改进这个项目！下面是最短的路径。

## 环境准备

```bash
git clone https://github.com/chaojiwudipaozi/labelprinter.git
cd label-print-site
npm install
npm run verify     # 语法检查 + 启动服务探活关键接口
npm start          # 打开 http://localhost:11235 开始开发
```

- 前端**没有构建步骤**：直接修改 `public/**`，浏览器 `Ctrl+F5` 即可看到效果。
- 桌面版调试：`npm run electron`（加载的是同一个本地服务）。
- 改动 `docs/` 里的文档后，执行 `npm run sync-manual` 同步到 `public/docs/`（`npm run dist` 会自动执行）。

## 代码约定

| 项 | 约定 |
| --- | --- |
| 语言 | 前端为原生 JavaScript（无框架、无打包器），服务端为 Node.js + Express |
| 缩进 | 2 空格（PowerShell 脚本 4 空格），统一 UTF-8 + LF（见 `.editorconfig`） |
| 注释 | 用中文说明「为什么这样做」，尤其是易踩坑处（单位换算、打印路径差异、历史兼容） |
| 单位 | 标签坐标与尺寸**统一用毫米**；字号用 pt；屏幕像素换算集中在 `MM` / `PRINT_K` 常量 |
| 三处一致 | 画布（DOM）、浏览器打印预览（HTML）、直连打印（Canvas 位图）必须共用同一套排版算法 |
| 提交信息 | 建议 `feat: …` / `fix: …` / `docs: …` / `refactor: …`，一句话说清「改了什么、为什么」 |

## 提交前自检

```bash
npm run verify     # 必须通过
npm run sync-manual # 如果改了 docs/
```

若改动涉及打包或桌面版，建议再跑一次主程序自检：

```bash
npm run pack
dist/win-unpacked/标签打印工具站.exe --selftest    # 输出 SELFTEST_RESULT {...}
dist/win-unpacked/标签打印工具站.exe --smoketest   # 输出 SMOKETEST_RESULT {...}
```

> 提示：桌面版有**单实例锁**。跑自检/冒烟前请先关闭已运行的窗口（否则新进程会直接退出且没有输出）。

## 提 Issue 时请附带

1. 软件版本（界面左上角徽标，或用菜单「关于」）、是网页版还是桌面版；
2. 操作系统与打印机型号（打印相关问题）；
3. 复现步骤 + 期望结果 + 实际结果；
4. 截图或 `SELFTEST_RESULT` 输出（打印问题尤其有用）。

## 发版流程（维护者）

安装包体积约 77 MB，**不要把 `.exe` 提交进仓库**（`.gitignore` 已忽略）；统一通过 **GitHub Releases** 分发，
README 里的下载入口指向 `releases/latest`，因此永远指向最新版。

1. 改版本号与日志：`package.json` 的 `version`、`CHANGELOG.md`，并同步 `docs/` 手册里的版本号；
2. 本地验证：`npm run sync-manual && npm run verify`；
3. 提交并打 tag（tag 形如 `v1.2.1`）：
   ```bash
   git add -A
   git commit -m "chore(release): v1.2.1"
   git push
   git tag v1.2.1
   git push origin v1.2.1
   ```
4. GitHub Actions 会自动构建 NSIS 安装包并**创建 Release**（同时上传带版本号的和固定名 `LabelPrint-Setup.exe`）；
5. 到 Releases 页面确认附件与 README 中的直链可用：
   - `.../releases/latest` → 最新版页面
   - `.../releases/latest/download/LabelPrint-Setup.exe` → 点击即下载最新版（依赖固定文件名）

> 想不跑 CI 直接发一版：在 Releases 页面 「Draft a new release」→ 新建 `vX.Y.Z` 标签 →
> 把 `dist/LabelPrint-Setup-<version>.exe` 拖进 Assets，并**再拖一份改名为 `LabelPrint-Setup.exe`**（保证固定直链可用）。

## 提 Pull Request

1. Fork 本仓库并从 `main` 开分支（如 `feat/multi-sheet-import`）；
2. 保持改动聚焦：一个 PR 解决一件事，避免顺手大范围重构；
3. 更新相关文档（`README.md` / `docs/` / `CHANGELOG.md`）；
4. 在 PR 描述里写清：改了什么、为什么、如何验证（命令 + 结果）。

## 不适合合并的内容

- 会把用户数据（`data/`、模板、Excel 内容）或内网地址、主机名提交进仓库的改动；
- 引入打包器/框架却未说明收益的大改（前端刻意保持零构建，便于热更新与现场排障）；
- 与「本机打印」核心场景无关的通用功能堆叠。
