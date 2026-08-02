package com.excelcleaner;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.*;

/**
 * Excel 空 Sheet 批量清理工具
 *
 * 功能：
 * 1. 扫描指定文件夹中所有 .csv / .xls / .xlsx 文件
 * 2. 删除空 Sheet（仅有标题行或无数据）
 * 3. 删除名称中标注数量为 0 的 Sheet（"可用性" 开头的除外）
 * 4. 保持原文件格式不变
 *
 * 使用方式：
 *   - 命令行: java -jar excel-cleaner.jar [文件夹路径]
 *   - 双击运行: 放在目标文件夹中双击，自动处理当前目录
 */
public class ExcelSheetCleaner {

    /** Sheet 名称中数量标注的正则，如 "Sheet1(5)" → base="Sheet1", count=5 */
    private static final Pattern NAME_WITH_COUNT = Pattern.compile("^(.+?)\\((\\d+)\\)\\s*$");

    /** 支持的文件扩展名 */
    private static final Set<String> SUPPORTED_EXTENSIONS = new HashSet<>(
            Arrays.asList(".xls", ".xlsx", ".csv")
    );

    /** 可保留的 Sheet 名称关键字（即使标注数量为 0 也不删除） */
    private static final String KEEP_KEYWORD = "可用性";

    // ========== 统计变量 ==========
    private int totalFilesScanned = 0;
    private int totalSheetsRemoved = 0;
    private int filesModified = 0;
    private int filesWithError = 0;
    private final List<String> logBuffer = new ArrayList<>();

    // ========== 入口 ==========
    public static void main(String[] args) {
        Path targetDir = resolveTargetDir(args);
        if (targetDir == null) {
            return;
        }

        ExcelSheetCleaner cleaner = new ExcelSheetCleaner();
        cleaner.execute(targetDir);
    }

    /**
     * 解析目标文件夹：命令行参数 > 当前工作目录
     */
    private static Path resolveTargetDir(String[] args) {
        Path dir;
        if (args.length > 0) {
            dir = Paths.get(args[0]);
        } else {
            dir = Paths.get("").toAbsolutePath();
        }

        if (!Files.isDirectory(dir)) {
            printError("路径不存在或不是文件夹: " + dir.toAbsolutePath());
            printError("用法: java -jar excel-cleaner.jar [文件夹路径]");
            printError("提示: 不放任何参数时，处理当前目录（双击 JAR 即为当前目录）");
            pauseIfDoubleClicked(args);
            return null;
        }
        return dir;
    }

    /**
     * 双击运行时暂停，让用户看到信息
     */
    private static void pauseIfDoubleClicked(String[] args) {
        if (args.length == 0) {
            System.out.println();
            System.out.print("按 Enter 键退出...");
            try {
                System.in.read();
            } catch (IOException ignored) {
            }
        }
    }

    private static void printError(String msg) {
        System.err.println(msg);
        System.out.println(msg);
    }

    // ========== 主流程 ==========
    public void execute(Path targetDir) {
        printHeader(targetDir);

        try {
            Files.list(targetDir)
                    .filter(Files::isRegularFile)
                    .filter(this::isSupportedFile)
                    .sorted()
                    .forEach(this::processOneFile);
        } catch (IOException e) {
            log("【严重错误】无法读取文件夹: " + e.getMessage());
        }

        printSummary(targetDir);
    }

    private boolean isSupportedFile(Path path) {
        String name = path.getFileName().toString().toLowerCase();
        return SUPPORTED_EXTENSIONS.stream().anyMatch(name::endsWith);
    }

    // ========== 单文件处理 ==========
    private void processOneFile(Path filePath) {
        totalFilesScanned++;
        String fileName = filePath.getFileName().toString();
        String lowerName = fileName.toLowerCase();

        log(String.format("  [%d] 检查: %s", totalFilesScanned, fileName));

        try {
            if (lowerName.endsWith(".csv")) {
                handleCsv(filePath);
            } else {
                handleExcel(filePath, lowerName);
            }
        } catch (Exception e) {
            filesWithError++;
            log("       !! 出错: " + e.getMessage());
            log("       !! 提示: 文件可能正在被 Excel/WPS 打开，请关闭后重试");
        }
    }

    // ========== Excel 处理 (.xls / .xlsx) ==========
    private void handleExcel(Path filePath, String lowerName) throws IOException {
        boolean isXls = lowerName.endsWith(".xls") && !lowerName.endsWith(".xlsx");

        Workbook workbook;
        try (FileInputStream fis = new FileInputStream(filePath.toFile())) {
            workbook = isXls ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
        }

        int totalSheets = workbook.getNumberOfSheets();
        // 记录要删除的 Sheet 索引（从后往前删，避免索引偏移）
        List<Integer> toRemove = new ArrayList<>();

        for (int i = 0; i < totalSheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            String name = sheet.getSheetName();

            if (shouldDelete(sheet)) {
                toRemove.add(i);
                log(String.format("       × 删除: \"%s\" (空Sheet)", name));
            }
        }

        if (toRemove.isEmpty()) {
            log("       √ 无需处理，所有 Sheet 均有数据");
        } else {
            // 从后往前删除，保持索引有效
            for (int i = toRemove.size() - 1; i >= 0; i--) {
                workbook.removeSheetAt(toRemove.get(i));
            }

            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }

            filesModified++;
            totalSheetsRemoved += toRemove.size();
            log(String.format("       √ 已删除 %d 个空 Sheet 并保存", toRemove.size()));
        }

        workbook.close();
    }

    // ========== CSV 处理 ==========
    private void handleCsv(Path filePath) throws IOException {
        List<String> lines;
        try {
            lines = Files.readAllLines(filePath, StandardCharsets.UTF_8);
        } catch (IOException e) {
            // UTF-8 失败，尝试 GBK
            lines = Files.readAllLines(filePath, Charset.forName("GBK"));
        }

        if (lines.isEmpty()) {
            log("       !! CSV 文件完全为空");
        } else if (lines.size() == 1) {
            log("       !! CSV 仅有标题行，无实际数据");
        } else {
            // 检查第 2 行起是否有非空数据
            boolean hasData = false;
            for (int i = 1; i < lines.size(); i++) {
                String line = lines.get(i).trim();
                if (!line.isEmpty() && !line.replaceAll(",", "").trim().isEmpty()) {
                    hasData = true;
                    break;
                }
            }
            if (hasData) {
                log("       - CSV 有数据，跳过");
            } else {
                log("       !! CSV 仅有标题行，数据行全为空");
            }
        }
    }

    // ========== 核心判断逻辑 ==========

    /**
     * 判断一个 Sheet 是否应该被删除
     */
    private boolean shouldDelete(Sheet sheet) {
        String name = sheet.getSheetName();

        // ----- 规则 1: 名称标注数量为 0 -----
        Matcher matcher = NAME_WITH_COUNT.matcher(name);
        if (matcher.matches()) {
            String baseName = matcher.group(1).trim();
            int count = Integer.parseInt(matcher.group(2));

            // "可用性" 开头的 Sheet，即使数量为 0 也保留
            if (baseName.startsWith(KEEP_KEYWORD)) {
                // 继续检查是否为空（可能仍需删除空 Sheet，但名称不构成删除理由）
            } else if (count == 0) {
                return true;  // 名称中标注数量为 0，直接删除
            }
        }

        // ----- 规则 2: Sheet 数据为空 -----
        // 检查物理行数（忽略空白格式行）
        int physicalRows = sheet.getPhysicalNumberOfRows();

        if (physicalRows == 0) {
            // 完全没有行
            return true;
        }

        if (physicalRows == 1) {
            // 仅有标题行（第 0 行）
            return true;
        }

        // 第 0 行是标题，检查第 1 行起是否有非空数据
        boolean hasData = false;
        int lastRowNum = sheet.getLastRowNum();
        for (int rowIdx = 1; rowIdx <= lastRowNum; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row != null && !isRowEmpty(row)) {
                hasData = true;
                break;
            }
        }

        return !hasData;
    }

    /**
     * 判断某行是否为全空
     */
    private boolean isRowEmpty(Row row) {
        short lastCell = row.getLastCellNum();
        if (lastCell <= 0) {
            return true;
        }

        for (int colIdx = 0; colIdx < lastCell; colIdx++) {
            Cell cell = row.getCell(colIdx);
            if (cell == null) {
                continue;
            }
            CellType type = cell.getCellType();
            if (type == CellType.BLANK) {
                continue;
            }
            if (type == CellType.STRING) {
                String val = cell.getStringCellValue();
                if (val != null && !val.trim().isEmpty()) {
                    return false;  // 有非空字符串
                }
            } else if (type == CellType.FORMULA) {
                // 公式也算数据
                return false;
            } else {
                return false;  // 数字、布尔、日期等，都算数据
            }
        }

        return true;
    }

    // ========== 日志与输出 ==========
    private void log(String msg) {
        System.out.println(msg);
        logBuffer.add(msg);
    }

    private void printHeader(Path targetDir) {
        String line = repeatStr("=", 56);
        log(line);
        log("    Excel 空 Sheet 批量清理工具 v1.0");
        log(line);
        log("  目标文件夹: " + targetDir.toAbsolutePath());
        log("  支持格式:   .xlsx  .xls  .csv");
        log("  删除规则:");
        log("    1. Sheet 中仅有标题行，无实际数据 → 删除");
        log("    2. Sheet名称标注数量为0(如Sheet1(0)) -> 删除");
        log("    3. 例外: [可用性]开头的Sheet始终保留");
        log("  处理方式:   直接修改原文件，保留格式和样式");
        log(line);
        log("");
    }

    private void printSummary(Path targetDir) {
        String line = repeatStr("=", 56);
        log("");
        log(line);
        log("    处理完毕 - " + currentTime());
        log(line);
        log(String.format("  扫描文件:   %d 个", totalFilesScanned));
        log(String.format("  已修改:     %d 个（删除了空 Sheet）", filesModified));
        log(String.format("  合计删除:   %d 个空 Sheet", totalSheetsRemoved));
        log(String.format("  出错:       %d 个", filesWithError));
        log("");

        if (filesModified == 0 && totalSheetsRemoved == 0) {
            log("  所有文件均无需处理。");
        }

        // 写日志文件
        writeReport(targetDir);
    }

    private void writeReport(Path targetDir) {
        try {
            Path reportFile = targetDir.resolve("excel-cleaner-log.txt");
            Files.write(reportFile, logBuffer, StandardCharsets.UTF_8);
            log(String.format("  📄 详细日志已保存: %s", reportFile.toAbsolutePath()));
        } catch (IOException e) {
            log("  !! 无法保存日志文件: " + e.getMessage());
        } finally {
            pauseIfDoubleClicked(new String[0]);
        }
    }

    private static String repeatStr(String s, int count) {
        StringBuilder sb = new StringBuilder(s.length() * count);
        for (int i = 0; i < count; i++) {
            sb.append(s);
        }
        return sb.toString();
    }

    private static String currentTime() {
        return LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
    }
}
