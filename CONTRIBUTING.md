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

## 发版与分发（维护者）

**安装包不进入版本库**：体积约 77 MB，`.gitignore` 已忽略 `*.exe`——GitHub 单文件上限 100 MB，
而且每次提交都会在 git 历史里永久累积体积（clone 越来越慢，且不可撤销）。

分发方式任选：

| 方式 | 做法 |
| --- | --- |
| **使用者自行打包**（README 采用的方式） | 按 README「桌面安装版（自行打包）」执行 `npm run pack`（免安装目录）或 `npm run dist`（安装包） |
| **线下/内网分发** | 维护者本地 `npm run dist` 出包后直接发给同事；或放进软件数据目录 `data\installer\`，让局域网内的人从网页右上角「⬇ 下载安装包」取包 |
| **GitHub Releases**（可选） | 想让项目页提供下载时，在 Actions 页面手动触发并填写 `release_tag`，CI 会构建并创建 Release；也可在 Releases 页面手动上传 exe |

发版步骤：

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
4. 推 tag 后 CI 会**构建安装包并上传为构建产物（Artifact）**，但**不会**自动创建 Release；
5. 只有你确实想公开提供下载时，再到 Actions → CI → **Run workflow** 填 `release_tag`（如 `v1.2.1`），
   这时会额外创建 Release 并挂上 `LabelPrint-Setup-<version>.exe` 与固定名 `LabelPrint-Setup.exe`。

> README 中**没有**安装包下载入口，使用者按「自行打包」章节自建；默认也不对外发布安装包。

## 提 Pull Request

1. Fork 本仓库并从 `main` 开分支（如 `feat/multi-sheet-import`）；
2. 保持改动聚焦：一个 PR 解决一件事，避免顺手大范围重构；
3. 更新相关文档（`README.md` / `docs/` / `CHANGELOG.md`）；
4. 在 PR 描述里写清：改了什么、为什么、如何验证（命令 + 结果）。

## 不适合合并的内容

- 会把用户数据（`data/`、模板、Excel 内容）或内网地址、主机名提交进仓库的改动；
- 引入打包器/框架却未说明收益的大改（前端刻意保持零构建，便于热更新与现场排障）；
- 与「本机打印」核心场景无关的通用功能堆叠。
