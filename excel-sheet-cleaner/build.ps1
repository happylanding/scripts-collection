# Excel Sheet Cleaner - One-click Build Script
# No Maven required. Downloads dependencies, compiles, and packages into a fat JAR.

param([switch]$clean)

$ErrorActionPreference = "Stop"

$projectDir = $PSScriptRoot
Set-Location $projectDir
$host.UI.RawUI.WindowTitle = "Building Excel Sheet Cleaner..."

Write-Host ""
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  Excel Sheet Cleaner - Fat JAR Builder" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""

# ========== Config ==========
$POI_VERSION = "4.1.2"
$workDir = Join-Path $projectDir "target"
$libDir = Join-Path $workDir "lib"
$classesDir = Join-Path $workDir "classes"
$srcDir = Join-Path $projectDir "src\main\java"
$outJar = Join-Path $workDir "excel-cleaner.jar"
$repoBase = "https://repo1.maven.org/maven2"

# Dependencies: (group, artifact, version)
$deps = @()
$deps += ,@("org.apache.poi", "poi", $POI_VERSION)
$deps += ,@("org.apache.poi", "poi-ooxml", $POI_VERSION)
$deps += ,@("org.apache.poi", "poi-ooxml-schemas", $POI_VERSION)
$deps += ,@("org.apache.commons", "commons-collections4", "4.4")
$deps += ,@("org.apache.commons", "commons-math3", "3.6.1")
$deps += ,@("org.apache.commons", "commons-compress", "1.19")
$deps += ,@("org.apache.xmlbeans", "xmlbeans", "3.1.0")
$deps += ,@("com.github.virtuald", "curvesapi", "1.06")

# ========== Clean ==========
if ($clean -and (Test-Path $workDir)) {
    Write-Host "[Clean] Removing old build files..." -ForegroundColor DarkYellow
    Remove-Item -Recurse -Force $workDir
}

# Create directories
foreach ($d in @($libDir, $classesDir)) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Force -Path $d | Out-Null
    }
}

# ========== Step 1: Download deps ==========
Write-Host "[1/4] Downloading dependencies..." -ForegroundColor Yellow

$downloaded = 0
$skipped = 0
$allJars = @()

foreach ($dep in $deps) {
    $g = $dep[0]; $a = $dep[1]; $v = $dep[2]
    $groupPath = $g.Replace(".", "/")
    $jarName = "$a-$v.jar"
    $url = "$repoBase/$groupPath/$a/$v/$jarName"
    $local = Join-Path $libDir $jarName
    $allJars += $local

    if (Test-Path $local) {
        $skipped++
    } else {
        Write-Host "      Download: $jarName" -ForegroundColor Gray
        try {
            Invoke-WebRequest -Uri $url -OutFile $local -UseBasicParsing -ErrorAction Stop
            $downloaded++
        } catch {
            Write-Host "      FAILED: $url" -ForegroundColor Red
            Write-Host "      Error: $_" -ForegroundColor Red
            Write-Host "      Tip: Check your network, or manually download and place in $libDir" -ForegroundColor Yellow
            exit 1
        }
    }
}
Write-Host "      Downloaded: $downloaded, Cached: $skipped" -ForegroundColor Green

# ========== Step 2: Compile ==========
Write-Host "[2/4] Compiling source code..." -ForegroundColor Yellow

$cp = [string]::Join(";", $allJars)
$javaSources = Get-ChildItem -Path $srcDir -Recurse -Filter '*.java' | Select-Object -ExpandProperty FullName

if ($javaSources.Count -eq 0) {
    Write-Host "      No .java source files found in $srcDir" -ForegroundColor Red
    exit 1
}

# Compile directly with javac
$javacArgs = @("-encoding", "UTF-8", "-cp", $cp, "-d", $classesDir)
$javacArgs += $javaSources
$compileOutput = & javac $javacArgs 2>&1 | Out-String

if ($LASTEXITCODE -ne 0) {
    Write-Host "      Compilation FAILED:" -ForegroundColor Red
    Write-Host $compileOutput
    exit 1
}
Write-Host "      Compiled $($javaSources.Count) source file(s)" -ForegroundColor Green

# ========== Step 3: Merge dependencies ==========
Write-Host "[3/4] Merging dependencies into classes..." -ForegroundColor Yellow

foreach ($jar in $allJars) {
    $name = Split-Path $jar -Leaf
    Write-Host "      $name" -ForegroundColor Gray
    Push-Location $classesDir
    & jar xf $jar 2>&1 | Out-Null
    Pop-Location
}
Write-Host "      Merged" -ForegroundColor Green

# ========== Step 4: Package JAR ==========
Write-Host "[4/4] Packaging fat JAR..." -ForegroundColor Yellow

$manifest = @"
Manifest-Version: 1.0
Main-Class: com.excelcleaner.ExcelSheetCleaner
Created-By: ExcelCleanerBuilder
"@

$manifestDir = Join-Path $workDir "META-INF"
if (-not (Test-Path $manifestDir)) {
    New-Item -ItemType Directory -Force -Path $manifestDir | Out-Null
}
$manifestPath = Join-Path $manifestDir "MANIFEST.MF"
[System.IO.File]::WriteAllText($manifestPath, $manifest, [System.Text.Encoding]::UTF8)

Push-Location $classesDir
& jar cfm "$outJar" "$manifestPath" . 2>&1 | Out-Null
Pop-Location

if (-not (Test-Path $outJar)) {
    Write-Host "      Packaging FAILED" -ForegroundColor Red
    exit 1
}

$size = [math]::Round((Get-Item $outJar).Length / 1MB, 2)
Write-Host "      Packaged: ${size} MB" -ForegroundColor Green

# ========== Done ==========
Write-Host ""
Write-Host "=================================================" -ForegroundColor Green
Write-Host "  BUILD SUCCESSFUL!" -ForegroundColor Green
Write-Host "=================================================" -ForegroundColor Green
Write-Host "  Output: $outJar" -ForegroundColor White
Write-Host "  Size:   ${size} MB" -ForegroundColor White
Write-Host ""
Write-Host "  How to use:" -ForegroundColor Cyan
Write-Host "    Copy excel-cleaner.jar to your target folder," -ForegroundColor White
Write-Host "    then double-click to clean all Excel files there." -ForegroundColor White
Write-Host ""
Write-Host "  Command line:" -ForegroundColor Cyan
Write-Host "    java -jar excel-cleaner.jar" -ForegroundColor White
Write-Host "    java -jar excel-cleaner.jar D:\MyFolder" -ForegroundColor White
Write-Host ""
