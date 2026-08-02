# Excel 空 Sheet 批量清理工具 (PowerShell 轻量版)

批量扫描文件夹中的 Excel 文件（.xlsx / .xls），自动删除无数据的空 Sheet，保留原文件格式和样式。

## 特性

- **零依赖** — 基于 Windows 原生 PowerShell + COM，无需 Java
- **轻量级** — 仅 ~14KB，零外部依赖
- 自动识别 Excel/WPS，用 COM 接口处理格式
- CSV 自动检测编码（UTF-8 / GBK / GB2312）
- 自动检测 .csv 扩展名实际为 XLSX 格式的文件
- 日志末尾按文件列出明细

## 快速使用

1. 将 `ExcelSheetCleaner.ps1` 和 `运行.bat` 复制到**目标文件夹**（存放 Excel 的那个文件夹）
2. **双击 `运行.bat`**
3. 自动扫描 → 清理空 Sheet → 同目录生成 `excel-cleaner-log.txt` 日志

## 删除规则

- Sheet 仅有标题行、无实际数据 → 删除
- Sheet 完全空白 → 删除
- Sheet 名称标注数量为 0（如 `Sheet1(0)`）→ 删除
- `可用性` / `可用` 开头的 Sheet → **始终保留**

## 环境要求

- Windows（Win7 及以上）
- 安装了 Office Excel 或 WPS

## 依赖

- Windows（Win7 及以上）自带的 PowerShell
- 安装了 Office Excel 或 WPS（处理 Excel 的电脑通常已安装）
