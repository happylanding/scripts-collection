# Excel 空 Sheet 批量清理工具

批量扫描文件夹中的 Excel 文件（.xlsx / .xls / .csv），自动删除无数据的空 Sheet，保留原文件格式和样式。

## 特性

- 自动识别 Excel 真实格式（不以扩展名为准）
- CSV 自动检测编码（UTF-8 BOM / UTF-16 / GBK / GB2312）
- 双击 JAR 自动处理 JAR 所在目录，无需手动指定路径

## 快速使用

1. 将 `excel-cleaner.jar` 和 `运行.bat` 复制到目标文件夹
2. **双击 `运行.bat`**
3. 自动清理完成，查看同目录 `excel-cleaner-log.txt`

## 删除规则

- Sheet 仅有标题行、无实际数据 → 删除
- Sheet 完全空白 → 删除  
- Sheet 名称标注数量为 0（如 `Sheet1(0)`）→ 删除
- `可用性` 开头的 Sheet → **始终保留**

## 环境要求

Java 8+（绝大多数电脑已预装）

## 重新编译

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

## 技术栈

Java 1.8 + Apache POI 4.1.2
