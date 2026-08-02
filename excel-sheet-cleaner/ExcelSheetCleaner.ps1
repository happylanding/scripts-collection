<#
.SYNOPSIS
    Excel 空 Sheet 批量清理工具 (PowerShell 轻量版)
.DESCRIPTION
    部署在目标文件夹中，双击运行.bat 即可自动扫描并清理所有 Excel 文件中的空 Sheet。
    基于 Windows 原生 PowerShell + COM 接口，无需安装 Java 或其他运行时。
#>

param(
    [string]$TargetDir = ""
)

# ==================== 配置 ====================
$scriptDir = if ($TargetDir) { (Resolve-Path $TargetDir).Path } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$logFile   = Join-Path $scriptDir "excel-cleaner-log.txt"
$logEncoding = [System.Text.Encoding]::UTF8

# ==================== 日志函数 ====================
function Write-Log {
    param([string]$Message)
    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$time  $Message"
    Add-Content -Path $logFile -Value $line -Encoding UTF8
    Write-Host $line
}

# ==================== CSV 编码检测 ====================
function Read-Csv {
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -eq 0) { return @() }

    # 1. BOM 检测
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $text = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $text = [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $text = [System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    else {
        # 2. 无 BOM：依次尝试严格解码
        $text = $null

        # UTF-8 严格模式：非法字节直接抛异常
        try {
            $utf8Strict = [System.Text.UTF8Encoding]::new($false, $true)
            $text = $utf8Strict.GetString($bytes)
        }
        catch { }

        # GBK / GB2312（GB 系列编码覆盖所有 256 字节值，基本不会失败）
        if ($null -eq $text) {
            try { $text = [System.Text.Encoding]::GetEncoding(936).GetString($bytes) }    # GBK
            catch { }
        }

        # 3. 兜底
        if ($null -eq $text) {
            $text = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)  # ISO-8859-1
        }
    }

    # 统一换行符后按 LF 分行
    return ($text -replace "`r`n", "`n") -split "`n"
}

# ==================== Sheet 检测函数 ====================
function Test-Sheet {
    param($Sheet, [string]$SheetName)

    # 规则 1：可用性开头 → 始终保留
    if ($SheetName -like "可用性*" -or $SheetName -like "可用*") {
        return "protected"
    }

    # 规则 2：名称含 "(0)" → 删除
    if ($SheetName -match '\(0\)') {
        return "delete:zero"
    }

    # 定位最后一个已使用的单元格
    try {
        $lastCell = $Sheet.Cells.SpecialCells(11)  # 11 = xlCellTypeLastCell
        $lastRow  = $lastCell.Row
        $lastCol  = $lastCell.Column
    }
    catch {
        return "delete:empty"
    }

    if ($lastRow -le 0 -or $lastCol -le 0) {
        return "delete:empty"
    }

    # 逐格扫描是否有任何数据
    $hasAny = $false
    for ($r = 1; $r -le $lastRow; $r++) {
        for ($c = 1; $c -le $lastCol; $c++) {
            $v = $Sheet.Cells.Item($r, $c).Text
            if ($v -and $v.Trim().Length -gt 0) {
                $hasAny = $true
                break
            }
        }
        if ($hasAny) { break }
    }

    if (-not $hasAny) {
        return "delete:empty"
    }

    # 检查标题行之下是否有数据
    if ($lastRow -le 1) {
        return "keep"
    }

    $hasBelow = $false
    for ($r = 2; $r -le $lastRow; $r++) {
        for ($c = 1; $c -le $lastCol; $c++) {
            $v = $Sheet.Cells.Item($r, $c).Text
            if ($v -and $v.Trim().Length -gt 0) {
                $hasBelow = $true
                break
            }
        }
        if ($hasBelow) { break }
    }

    if (-not $hasBelow) {
        return "delete:header"
    }

    return "keep"
}

# ==================== 主流程 ====================

# 初始化日志
[System.IO.File]::WriteAllText($logFile, "", $logEncoding)

Write-Log "========================================"
Write-Log "  Excel 空 Sheet 批量清理工具 v2.0 (PS)"
Write-Log "  目标文件夹: $scriptDir"
Write-Log "========================================"

$totalDeleted  = 0
$totalFiles    = 0
$fileResults   = @()  # 记录每个文件的处理结果，用于最终汇总

# ==================== 文件收集与格式检测 ====================
# 注意：PowerShell 中 -Include 必须配合 \* 或 -Recurse 才生效
$allFiles = @(Get-ChildItem -Path $scriptDir -File)

# Excel 文件
$excelFiles = @($allFiles | Where-Object { $_.Extension -match '^\.(xlsx?|xls)$' -and $_.Name -notmatch '^~\$' })

# CSV 文件：检测是否有扩展名 .csv 但实际是 XLSX 的文件
$realCsv = @()
$fakeCsv = @()
$csvCandidate = @($allFiles | Where-Object { $_.Extension -eq '.csv' })
foreach ($f in $csvCandidate) {
    try {
        $sig = [System.IO.File]::ReadAllBytes($f.FullName)
        if ($sig.Length -ge 4 -and $sig[0] -eq 0x50 -and $sig[1] -eq 0x4B -and $sig[2] -eq 0x03 -and $sig[3] -eq 0x04) {
            # ZIP 文件头 (PK..) → 实际是 XLSX，归入 Excel 处理队列
            $fakeCsv += $f
        }
        else {
            $realCsv += $f
        }
    }
    catch {
        $realCsv += $f
    }
}
if ($fakeCsv.Count -gt 0) {
    Write-Log ""
    Write-Log "[!] 检测到 $($fakeCsv.Count) 个 .csv 文件实际为 XLSX 格式："
    foreach ($f in $fakeCsv) { Write-Log "     -> $($f.Name)" }
    $excelFiles += $fakeCsv
}
$csvFiles = $realCsv

# ==================== Excel 处理 ====================
if ($excelFiles.Count -gt 0) {
    Write-Log ""

    # 初始化 Excel COM
    try {
        $excel = New-Object -ComObject Excel.Application
    }
    catch {
        Write-Log "错误: 未检测到 Excel 或 WPS，请先安装 Office / WPS"
        Write-Log "提示: 如果仅有 CSV 文件，可忽略此错误继续"
        $excel = $null
    }

    if ($excel) {
        $excel.Visible       = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        try {
            foreach ($file in $excelFiles) {
                $totalFiles++
                Write-Log ""
                Write-Log "[$totalFiles] 检查: $($file.Name)"

                $wb = $null
                try {
                    $wb = $excel.Workbooks.Open($file.FullName)
                    $fileDeleted = 0
                    $toDelete = @()

                    for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
                        $sheet = $wb.Sheets.Item($i)
                        $name  = $sheet.Name
                        $result = Test-Sheet -Sheet $sheet -SheetName $name

                        switch -Wildcard ($result) {
                            "protected" {
                                Write-Log "  [+] 保留: $name (实时可用性监控)"
                            }
                            "delete:*" {
                                $reason = if ($result -eq "delete:zero")   { "名称标注为 0" }
                                     elseif ($result -eq "delete:empty")  { "无任何数据" }
                                     else                                  { "仅有标题行" }
                                Write-Log "  [-] 删除: $name ($reason)"
                                $toDelete += $i
                            }
                            "keep" {
                                Write-Log "  [+] 保留: $name"
                            }
                        }
                    }

                    # 倒序删除（避免索引偏移）
                    [Array]::Sort($toDelete)
                    [Array]::Reverse($toDelete)

                    foreach ($idx in $toDelete) {
                        try {
                            $wb.Sheets.Item($idx).Delete()
                            $fileDeleted++
                            $totalDeleted++
                        }
                        catch {
                            Write-Log "  !! 删除失败: $($_.Exception.Message)"
                        }
                    }

                    if ($fileDeleted -gt 0) {
                        $wb.Save()
                        Write-Log "  >> 本文件共删除: $fileDeleted 个空 Sheet"
                        $wb.Close($true)
                    }
                    else {
                        Write-Log "  >> 本文件无需清理"
                        $wb.Close($false)
                    }

                    # 记录结果
                    $fileResults += [PSCustomObject]@{
                        Name         = $file.Name
                        Type         = "Excel"
                        DeletedCount = $fileDeleted
                        RowCount     = $null
                    }

                }
                catch {
                    Write-Log "  !! 出错: $($_.Exception.Message)"
                    Write-Log "  !! 提示: 文件可能正在被 Excel/WPS 打开，请关闭后重试"
                    if ($wb) { try { $wb.Close($false) } catch { } }
                    $fileResults += [PSCustomObject]@{
                        Name         = $file.Name
                        Type         = "Excel"
                        DeletedCount = -1
                        RowCount     = $null
                    }
                }
            }
        }
        finally {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }

        Write-Log ""
        Write-Log "Excel 进程已安全释放"
    }
}
else {
    Write-Log ""
    Write-Log "未找到 Excel 文件 (.xlsx / .xls)"
}

# ==================== CSV 处理 ====================
if ($csvFiles.Count -gt 0) {
    Write-Log ""
    Write-Log "--- CSV 文件检查 ---"

    $csvIndex = 0
    foreach ($file in $csvFiles) {
        $csvIndex++
        $lineCount = $null
        try {
            $lines = Read-Csv -Path $file.FullName
            $lineCount = $lines.Count

            Write-Log "  [$csvIndex] $($file.Name): $lineCount 行"

            if ($lineCount -eq 0) {
                Write-Log "    -> 文件为空"
            }
            elseif ($lineCount -eq 1) {
                if ([string]::IsNullOrWhiteSpace($lines[0])) {
                    Write-Log "    -> 文件为空（仅有空行）"
                }
                else {
                    Write-Log "    -> 有数据 (1 行)"
                }
            }
            else {
                Write-Log "    -> 有数据 ($lineCount 行)"
            }
        }
        catch {
            Write-Log "  !! CSV 读取失败: $($_.Exception.Message)"
        }
        $fileResults += [PSCustomObject]@{
            Name         = $file.Name
            Type         = "CSV"
            DeletedCount = 0
            RowCount     = $lineCount
        }
    }
}

# ==================== 汇总 ====================
Write-Log ""
Write-Log "========================================"
Write-Log "  清理完成!"
Write-Log "========================================"

# 按文件列出详情
if ($fileResults.Count -gt 0) {
    Write-Log ""
    Write-Log "--- 文件处理明细 ---"
    $idx = 0
    foreach ($r in $fileResults) {
        $idx++
        if ($r.Type -eq "Excel") {
            if ($r.DeletedCount -eq -1) {
                Write-Log "  $idx. [$($r.Type)] $($r.Name) -> 处理失败"
            }
            elseif ($r.DeletedCount -eq 0) {
                Write-Log "  $idx. [$($r.Type)] $($r.Name) -> 无需清理"
            }
            else {
                Write-Log "  $idx. [$($r.Type)] $($r.Name) -> 删除 $($r.DeletedCount) 个空 Sheet"
            }
        }
        else {
            if ($null -eq $r.RowCount) {
                Write-Log "  $idx. [$($r.Type)] $($r.Name) -> 读取失败"
            }
            else {
                Write-Log "  $idx. [$($r.Type)] $($r.Name) -> $($r.RowCount) 行"
            }
        }
    }
}

Write-Log ""
Write-Log "--- 总计 ---"
Write-Log "  Excel: 处理 $($excelFiles.Count) 个, 删除 $totalDeleted 个空 Sheet"
Write-Log "  CSV  : 检查 $($csvFiles.Count) 个"
Write-Log "  日志 : $logFile"
Write-Log "========================================"

Write-Host ""
Read-Host "按 Enter 键退出"
