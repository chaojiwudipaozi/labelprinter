# 项目文档

本目录存放**项目级文档**。软件内「❓ 使用手册 / 版本日志」读取的是 `public/docs/`，由 `npm run sync-manual`
从本目录（以及根目录的 `CHANGELOG.md`）同步生成，**改文档请改本目录下的源文件**。

| 文档 | 面向 | 内容 |
| --- | --- | --- |
| [使用手册.md](使用手册.md) | 使用者 | 安装与启动、界面说明、设计标签（元素/表格/自适应字号）、Excel 数据准备、打印与排查、模板管理、局域网共享、FAQ、快捷键 |
| [使用说明.md](使用说明.md) | 部署运维 | 启动命令、端口与局域网放行、后台运行、打包、目录结构、命令速查、故障排查 |
| [../CHANGELOG.md](../CHANGELOG.md) | 所有人 | 版本日志（按版本倒序记录新增/优化/修复） |

## 文档更新流程

```bash
# 1) 修改 docs/使用手册.md 或 docs/使用说明.md（根目录的 CHANGELOG.md 用于记录版本）
# 2) 同步到应用内（软件内手册与版本日志页签读的就是这里）
npm run sync-manual
# 3) 打包时会自动执行同步（predist），无需手动重复
npm run dist
```

## 截图

界面截图放在 `docs/images/`，README 顶部有一段被注释掉的图片引用，替换文件名并取消注释即可。

```text
docs/images/
├─ screenshot-main.png     主界面（画布 + 功能区）
├─ screenshot-import.png   导入结果面板（多工作表）
├─ screenshot-table.png    表格编辑（合并 / 行列调整）
└─ screenshot-print.png    打印与打印机状态
```

> 截图建议：使用示例模板 + 少量示例数据，避免出现真实公司名、物料编号、内网 IP 等信息。
