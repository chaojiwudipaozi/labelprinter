<div align="center">

# 标签打印工具站

**Excel 导数据 → 网页设计标签 → 直连打印机批量出纸**

一个面向企业仓储/生产场景的标签打印工具：上传 Excel（支持多工作表、任意表头），
在浏览器里拖拽设计标签（文本 / 条码 / 二维码 / 矩形 / 直线 / 表格），
点一下就把标签**静默送进打印机队列**——不弹浏览器打印对话框、不需要专业标签软件。

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Node](https://img.shields.io/badge/node-%3E%3D18-brightgreen.svg)](https://nodejs.org/)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-0078D6.svg)](#系统要求)
[![Electron](https://img.shields.io/badge/desktop-Electron%2032-47848F.svg)](https://www.electronjs.org/)

[功能特性](#功能特性) · [快速开始](#快速开始) · [工作原理](#工作原理) · [目录结构](#目录结构) · [开发与打包](#开发与打包) · [文档](#文档)

</div>

---

<!-- 截图占位：把界面截图放到 docs/images/ 后，删掉注释并改成实际文件名
<p align="center"><img src="docs/images/screenshot-main.png" alt="界面预览" width="860"></p>
-->

## 简介

本工具解决的是「用 Excel 批量打标签」这件事上最烦的几段：

1. **数据在 Excel 里**，格式五花八门（多个工作表、前面有标题行、列名带空格、日期是数字…）；
2. **标签版式要能自己画**，还得能按行把数据替换进去；
3. **打印要快**——批量几十上百张，不能被浏览器的打印对话框一张张卡住。

它把这三点做成一个纯本地运行的小服务：**网页负责界面，服务端负责读打印机与投递作业**，
所以同一个程序既是「网页版」（局域网内谁都能打开），也能打包成「Windows 桌面安装版」（双击即用）。

> 全部数据（Excel 内容、标签模板）**只在本机处理**，不上传任何云端。

## 功能特性

### 数据导入（Excel / CSV）

- 支持 `xlsx / xls / xlsm / xlsb / csv / tsv / txt`，自动识别编码（UTF-8 / GB18030 / Big5）
- **多工作表全部识别**：默认把所有工作表合并导入——主表提供列名，其余表按列名对齐（忽略大小写与空格），
  独有列自动追加；纯说明表会跳过并在结果面板说明原因，也可取消勾选改为单独指定工作表
- 自动跳过标题行/说明行/空行，自动识别表头行（可手动指定后重新解析）
- 列名清洗（「物料 编码」→「物料编码」）、重复列名加序号、无名有数据的列自动命名
- 日期/数字按 Excel 显示效果保留；`二维码内容`/`条码内容` 留空时自动生成
- 导入结果面板给出逐表行数、跳过原因与校验提示

### 标签设计

- 元素：文本、条形码（CODE128/CODE39/EAN13 等 10 种码制）、二维码、矩形、直线、**表格**
- 文本支持「前缀 + 绑定字段」「当前日期时间（8 种格式）」、字体/字号/加粗/对齐
- **自适应字号**：文本、条码文字、表格单元格默认开启——内容始终一行，超出宽度自动缩小字号（最小 3pt）
- 表格支持框选合并、拖拽行列分隔线调整行高列宽、单元格独立绑定字段与对齐
- **上一步 / 下一步**（`Ctrl+Z` / `Ctrl+Y`，最多 60 步，连续拖动自动合并为一步）
- 标签纸尺寸：27 种常用规格 + 自定义（毫米/厘米/英寸）
- 模板保存/载入/删除，草稿自动保存（关掉页面不丢）

### 打印

- 自动列出本机全部打印机与实时状态（空闲/打印中/脱机/已暂停/异常）与队列长度
- **直连打印（推荐）**：前端按约 305dpi 渲染位图 → 服务端静默投递，无任何对话框，按每行「打印数量」连续出纸
- **浏览器预览打印（备用）**：需要自己确认版式时使用
- 可选「打印服务」切换：让 A 电脑的网页使用 B 电脑上的打印机
- 打印机诊断：驱动、TCP/IP 端口、已连接未安装设备，一站式排查「看不到打印机」
- 绑定字段为空值时**打印留空**（屏幕上仍显示 `{字段名}` 占位，便于设计时辨认）

### 部署与升级

- 网页版：`npm start` 即用；局域网内其他电脑用 `http://本机IP:11235` 打开
- 桌面版：`npm run dist` 打出 NSIS 安装包，装上即用、卸载不删数据
- **增量更新**：网页界面层（`public/**`、文档）可热更新，只需从局域网里一台装有新版的电脑拉取变化文件；
  主程序层（Electron/内置服务/系统脚本）仍走完整安装包
- 网页右上角「⬇ 下载安装包」：把自己打好的安装包放进数据目录 `data\installer\`，局域网内任意电脑即可一键下载（无需 U 盘分发）

## 快速开始

### 系统要求

| 项 | 要求 |
| --- | --- |
| 操作系统 | Windows 10 / 11（64 位）**打印相关功能依赖 Windows 打印后台**；界面本身可在其它系统运行 |
| Node.js | ≥ 18（脚本使用全局 `fetch`；开发环境实测 v20 / v23） |
| 浏览器 | Chrome / Edge 等现代浏览器 |

### 网页版（3 条命令）

```bash
git clone https://github.com/chaojiwudipaozi/labelprinter.git
cd label-print-site
npm install
npm start          # 默认 http://localhost:11235
```

局域网内其他电脑访问 `http://本机IP:11235`；若被防火墙拦截，以管理员身份执行：

```bash
npm run lan        # 放行入站端口（同时解除可能存在的 node.exe 阻止规则）
```

### 桌面安装版（自行打包）

本项目**不提供预编译的安装包**，需要桌面版时按下面步骤自己打包（Windows 环境）：

```bash
npm install        # 首次：安装依赖（含 Electron 二进制，约 100 MB，需联网）
npm run pack       # 产出免安装目录 dist/win-unpacked/ → 双击「标签打印工具站.exe」即用
npm run dist       # 产出 NSIS 安装包 dist/LabelPrint-Setup-<version>.exe（可分发给同事安装）
```

打包要点：

- 请在 **Windows** 上打包（`npm run dist` 使用 Windows 打包链，产出 x64 安装包）；
- `npm run dist` 会先自动执行 `predist → npm run sync-manual`，把 `docs/` 文档同步进 `public/docs/`；
- 公司网络下载 Electron 失败时，先设镜像再安装：

  ```bat
  set ELECTRON_MIRROR=https://npmmirror.com/mirrors/electron/
  npm install
  ```

- 打完包用内置自检确认产物正常：

  ```bat
  dist\win-unpacked\标签打印工具站.exe --selftest     :: 输出 SELFTEST_RESULT {...}
  dist\win-unpacked\标签打印工具站.exe --smoketest    :: 真实打开窗口，输出 SMOKETEST_RESULT {...}
  ```

- 安装版数据目录为 `%APPDATA%\label-print-site\data`，**卸载不会删除标签模板**；
- 打好的安装包可以放进 `data\installer\`，局域网内其他人就能从网页右上角「⬇ 下载安装包」取包（见「部署与升级」）。

### 五分钟上手

1. 打开软件 →「📋 示例模板」看一个成品标签长什么样；
2. 「⬇ 下载Excel模板」或直接准备任意 Excel（第一行是字段名）；
3. 「📥 导入Excel」→ 选文件 → 看导入结果面板（可切换工作表/表头行）；
4. 双击标签上的元素 → 点字段名完成绑定；
5. 顶部选打印机 → 「🖨 批量打印」。

## 工作原理

```text
┌─────────────────────────── 本机（或服务器） ───────────────────────────┐
│                                                                        │
│   浏览器（Chrome/Edge）              Node.js + Express 内置服务         │
│   ┌───────────────────┐             ┌──────────────────────────────┐   │
│   │ public/ 界面       │  HTTP/JSON  │ server.js                    │   │
│   │  · 画布设计器      │ ◄─────────► │  · /api/labels  模板读写      │   │
│   │  · Excel 解析(xlsx)│             │  · /api/printers 打印机状态   │   │
│   │  · 标签位图渲染    │             │  · /api/print   静默打印      │   │
│   └───────────────────┘             │  · /api/update  增量更新      │   │
│                                     └───────────┬──────────────────┘   │
│                                                 │ child_process         │
│                                     ┌───────────▼──────────────────┐   │
│                                     │ scripts/*.ps1（PowerShell）   │   │
│                                     │ 读打印机 / 投递位图到队列      │   │
│                                     └───────────┬──────────────────┘   │
└─────────────────────────────────────────────────┼──────────────────────┘
                                                  ▼
                                        Windows 打印后台（Spooler）
                                                  ▼
                                              打印机
```

- **桌面版 = 同一套服务 + Electron 窗口**：主进程内嵌启动上面的 Express 服务，窗口直接加载它；
  因此「装到哪台电脑，就读哪台电脑的打印机」。
- **浏览器拿不到本机打印机**，所以打印作业由服务端通过 PowerShell 投递到打印队列（与 Word 打印同理）。
- 表格/条码/二维码/文字在**画布（DOM）、打印预览（HTML）、直连打印（Canvas 位图）**三条路径上共用同一套
  毫米坐标与自适应字号算法，保证「所见即所得」。

## 目录结构

```
label-print-site/
├─ server.js                服务端入口（Express：静态托管 + 模板/打印机/打印/更新接口）
├─ electron/main.js         桌面版主进程（内嵌服务 + 窗口 + 菜单 + 自检/冒烟）
├─ lib/                     服务端模块
│  ├─ paths.js              打包/开发两套目录规则（asar、app.asar.unpacked、数据目录）
│  ├─ printers.js           打印机枚举与状态归一化
│  ├─ printer-diag.js       打印机诊断（驱动/端口/设备）
│  ├─ silent-print.js       静默打印（位图 → 打印队列）
│  └─ updater.js            界面层增量更新（清单/差异/应用/回滚/过期清理）
├─ public/                  前端（无构建步骤，原生 HTML/CSS/JS）
│  ├─ index.html
│  ├─ css/style.css
│  ├─ js/app.js             画布设计器 / 表格模型 / 数据导入 / 打印 / 更新面板
│  ├─ lib/                  前端依赖（SheetJS、JsBarcode、qrcode）
│  └─ docs/                 应用内文档（由 npm run sync-manual 生成，随包分发）
├─ scripts/
│  ├─ verify.js             本地自检：语法检查 + 启动服务探活（CI 复用）
│  ├─ sync-manual.js        同步 docs/ 到 public/docs/
│  └─ *.ps1                 打印机枚举 / 静默打印 / 诊断 / 添加打印机 / 防火墙放行
├─ docs/                    项目文档（使用手册 / 部署说明 / 截图）
├─ build/icon.ico           应用图标
└─ .github/                 Issue 模板、PR 模板、CI（Windows）
```

> 运行期数据（`data/`）与打包产物（`dist/`）不纳入版本控制，详见 `.gitignore`。

## 开发与打包

```bash
npm install            # 安装依赖（含 Electron）
npm run verify         # 自检：语法检查 + 启动服务探活关键接口（不依赖 Electron）
npm start              # 网页版开发（改前端直接刷新页面即可）
npm run electron       # 桌面版开发（Electron 窗口）
npm run sync-manual    # 改完 docs/ 后同步到应用内（npm run dist 会自动执行）
```

- 前端**没有构建步骤**：直接改 `public/**`，浏览器 `Ctrl+F5` 即可生效（这也让界面层可以热更新）。
- 打包：`npm run pack`（免安装目录）/ `npm run dist`（NSIS 安装包）；
  主程序自检与界面冒烟可用于自动化验证：
  - `dist/win-unpacked/标签打印工具站.exe --selftest`（不开窗口，输出 `SELFTEST_RESULT {...}`）
  - `... --smoketest`（真实开窗口，输出 `SMOKETEST_RESULT {...}`）
- 提交前请至少跑一次 `npm run verify`。

## 文档

| 文档 | 内容 |
| --- | --- |
| [docs/使用手册.md](docs/使用手册.md) | **面向使用者**：安装、界面、设计标签、Excel 数据、打印、模板、FAQ（软件内按 `F1` 也能看） |
| [docs/使用说明.md](docs/使用说明.md) | **面向部署运维**：启动命令、端口、局域网放行、打包、目录结构、故障排查 |
| [CHANGELOG.md](CHANGELOG.md) | 版本日志（每个版本的新增/优化/修复） |

> 修改文档后执行 `npm run sync-manual`（或 `npm run dist`），会同步到 `public/docs/` 供软件内读取。

## 常见问题

**为什么浏览器里选不了打印机？**
浏览器出于安全限制无法读取本机打印机。本项目让**服务端**去读：所以「软件装在哪台电脑，就用哪台的打印机」；
要让 A 电脑的网页用 B 电脑的打印机，可在顶部「打印服务」里填入 B 的地址。

**导入 Excel 后字段名不对？**
在「导入结果」面板里调整「表头行」，或取消「导入全部工作表」后单独指定工作表，再点「重新解析」。

**打印出来是 `{字段名}`？**
说明界面版本较旧；新版（≥ V1.2.1）在打印时会把空值留空，屏幕上仍显示占位方便设计。

更多问题见 [docs/使用手册.md](docs/使用手册.md) 的「常见问题」与 [docs/使用说明.md](docs/使用说明.md) 的运维排查。

## 贡献

欢迎提 Issue / PR，见 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 许可证

[MIT](LICENSE)

> 本项目为独立实现的通用标签打印工具，与任何第三方标签软件厂商无关。
> 打印功能会调用系统打印后台与本机 PowerShell 脚本，请在受控环境内使用并自行评估合规性。
