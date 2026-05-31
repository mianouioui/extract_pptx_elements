@echo off
setlocal EnableExtensions
set "VERSION=V1.2.2"
set "RESTORE_SCRIPT=%~f0"
set "RESTORE_PY_TEMP=%TEMP%\restore_pptx_elements_py_%RANDOM%%RANDOM%.py"
set "RESTORE_PS_TEMP=%TEMP%\restore_pptx_elements_ps_%RANDOM%%RANDOM%.ps1"
chcp 65001 >nul
REM ============================================================
REM  PPTX 元素还原器 %VERSION% - Windows 单文件启动器
REM  把本文件放在“原始 PPTX”和“pptx_extracted_elements”文件夹旁边，
REM  双击即可一键把修好的图片写回 PPTX（所有元素位置保持不变）。
REM ============================================================

set "RESTORE_INPUT="
if not "%~1"=="" goto choose_runtime
if exist "pptx_extracted_elements\manifest.csv" (
    echo.
    echo √ 已在当前文件夹检测到提取素材，开始一键还原 ...
    goto choose_runtime
)
if exist "manifest.csv" (
    echo.
    echo √ 已检测到 manifest.csv，开始一键还原 ...
    goto choose_runtime
)

echo.
echo 把修好的图片写回 PPTX，所有元素位置保持不变。
echo 请将“提取文件夹”（含 manifest.csv）拖到此窗口，或留空回车自动查找。
echo.
echo 常用选项: --images-only  --pptx 文件  -o 文件名  --overwrite  --dry-run
echo.
set /p "RESTORE_INPUT=路径: "
set "RESTORE_INPUT=%RESTORE_INPUT:"=%"

:choose_runtime
py -3 --version >nul 2>&1
if errorlevel 1 goto try_python

echo → 使用内嵌 Python 运行 ...
py -3 -c "from pathlib import Path; import os; text=Path(os.environ['RESTORE_SCRIPT']).read_text(encoding='utf-8').replace('\r\n','\n'); marker=chr(35)+' PYTHON_CODE_BELOW'; Path(os.environ['RESTORE_PY_TEMP']).write_text(text.split('\n'+marker+'\n',1)[1], encoding='utf-8')"
if errorlevel 1 goto python_extract_failed
if defined RESTORE_INPUT goto run_py_launcher_input
py -3 "%RESTORE_PY_TEMP%" %*
goto after_py_launcher_run

:run_py_launcher_input
py -3 "%RESTORE_PY_TEMP%" "%RESTORE_INPUT%"

:after_py_launcher_run
set "EXIT_CODE=%ERRORLEVEL%"
del "%RESTORE_PY_TEMP%" >nul 2>&1
goto after_run

:try_python
python --version >nul 2>&1
if errorlevel 1 goto try_python3

echo → 使用内嵌 Python 运行 ...
python -c "from pathlib import Path; import os; text=Path(os.environ['RESTORE_SCRIPT']).read_text(encoding='utf-8').replace('\r\n','\n'); marker=chr(35)+' PYTHON_CODE_BELOW'; Path(os.environ['RESTORE_PY_TEMP']).write_text(text.split('\n'+marker+'\n',1)[1], encoding='utf-8')"
if errorlevel 1 goto python_extract_failed
if defined RESTORE_INPUT goto run_python_input
python "%RESTORE_PY_TEMP%" %*
goto after_python_run

:run_python_input
python "%RESTORE_PY_TEMP%" "%RESTORE_INPUT%"

:after_python_run
set "EXIT_CODE=%ERRORLEVEL%"
del "%RESTORE_PY_TEMP%" >nul 2>&1
goto after_run

:try_python3
python3 --version >nul 2>&1
if errorlevel 1 goto powershell_fallback

echo → 使用内嵌 Python 运行 ...
python3 -c "from pathlib import Path; import os; text=Path(os.environ['RESTORE_SCRIPT']).read_text(encoding='utf-8').replace('\r\n','\n'); marker=chr(35)+' PYTHON_CODE_BELOW'; Path(os.environ['RESTORE_PY_TEMP']).write_text(text.split('\n'+marker+'\n',1)[1], encoding='utf-8')"
if errorlevel 1 goto python_extract_failed
if defined RESTORE_INPUT goto run_python3_input
python3 "%RESTORE_PY_TEMP%" %*
goto after_python3_run

:run_python3_input
python3 "%RESTORE_PY_TEMP%" "%RESTORE_INPUT%"

:after_python3_run
set "EXIT_CODE=%ERRORLEVEL%"
del "%RESTORE_PY_TEMP%" >nul 2>&1
goto after_run

:python_extract_failed
echo 内嵌 Python 提取失败，正在尝试 PowerShell 回退 ...
del "%RESTORE_PY_TEMP%" >nul 2>&1

:powershell_fallback
where powershell.exe >nul 2>&1
if errorlevel 1 (
    echo 未找到可用运行环境。
    echo 请安装 Python 3，或使用带 PowerShell 5.1+ 的 Windows 系统。
    echo 按任意键关闭此窗口 ...
    pause >nul
    exit /b 1
)

echo → 使用内嵌 PowerShell 运行 ...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$script=[IO.File]::ReadAllText($env:RESTORE_SCRIPT,[Text.Encoding]::UTF8); $psParts=[regex]::Split($script,'(?m)^# POWERSHELL_START\r?$',2); if($psParts.Count -lt 2){throw 'PowerShell section missing'}; $payload=[regex]::Split($psParts[1],'(?m)^# PYTHON_CODE_BELOW\r?$',2)[0]; $encoding=New-Object System.Text.UTF8Encoding $true; [IO.File]::WriteAllText($env:RESTORE_PS_TEMP,$payload,$encoding)"
if errorlevel 1 (
    echo 启动器损坏或 PowerShell 脚本提取失败。
    echo 按任意键关闭此窗口 ...
    pause >nul
    exit /b 1
)

if defined RESTORE_INPUT goto run_powershell_input
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%RESTORE_PS_TEMP%" %*
goto after_powershell_run

:run_powershell_input
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%RESTORE_PS_TEMP%" "%RESTORE_INPUT%"

:after_powershell_run
set "EXIT_CODE=%ERRORLEVEL%"
del "%RESTORE_PS_TEMP%" >nul 2>&1

:after_run
echo.
echo 完成！(退出码: %EXIT_CODE%)
echo 提醒：还原后的文件默认名为 ^<原名^>_restored.pptx，就在原始 PPTX 旁边。
echo 按任意键关闭此窗口 ...
pause >nul
exit /b %EXIT_CODE%

# POWERSHELL_START
param([string[]]$LauncherArgs)

$ErrorActionPreference = "Stop"

$Version = "V1.2.2"
$ManifestName = "manifest.csv"
$DefaultExtractDirName = "pptx_extracted_elements"
$DefaultOutputSuffix = "_restored"

# Kinds whose extracted file is the literal part inside the pptx and can be
# written back byte-for-byte. "text" is intentionally excluded.
$RestorableKinds = @("audio", "chart", "chart_colors", "chart_style", "diagram", "embedded", "image", "ole", "video")
$MediaKinds = @("image", "video", "audio")

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Show-Usage {
    Write-Host "用法:"
    Write-Host "  restore_pptx_elements.cmd"
    Write-Host "  restore_pptx_elements.cmd pptx_extracted_elements"
    Write-Host "  restore_pptx_elements.cmd pptx_extracted_elements --pptx deck.pptx -o deck_fixed.pptx"
    Write-Host "  restore_pptx_elements.cmd --images-only"
    Write-Host ""
    Write-Host "选项: --pptx 文件  -o 文件  --images-only  --media-only  --overwrite  --dry-run"
}

function Get-FullPath {
    param([string]$Path)
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $Path))
}

function Is-TempPptx {
    param([string]$Path)
    $name = [System.IO.Path]::GetFileName($Path)
    return ($name.StartsWith("~$") -or $name.StartsWith(".~"))
}

function Parse-Arguments {
    param([string[]]$Argv)

    if ($null -eq $Argv) { $Argv = @() }
    $options = [ordered]@{
        Folder = $null
        Pptx = $null
        Output = $null
        ImagesOnly = $false
        MediaOnly = $false
        Overwrite = $false
        DryRun = $false
    }

    for ($i = 0; $i -lt $Argv.Count; $i++) {
        $item = $Argv[$i]
        switch -Regex ($item) {
            '^(--help|-h|/\?)$' { Show-Usage; exit 0 }
            '^--version$' { Write-Host "restore_pptx_elements.cmd $Version"; exit 0 }
            '^--images-only$' { $options.ImagesOnly = $true; continue }
            '^--media-only$' { $options.MediaOnly = $true; continue }
            '^--overwrite$' { $options.Overwrite = $true; continue }
            '^--dry-run$' { $options.DryRun = $true; continue }
            '^--pptx$' {
                if ($i + 1 -ge $Argv.Count) { throw "缺少 --pptx 参数" }
                $i++
                $options.Pptx = $Argv[$i]
                continue
            }
            '^(-o|--output)$' {
                if ($i + 1 -ge $Argv.Count) { throw "缺少输出文件参数：$item" }
                $i++
                $options.Output = $Argv[$i]
                continue
            }
            default {
                if ($null -eq $options.Folder) { $options.Folder = $item }
                else { throw "多余的参数：$item" }
            }
        }
    }

    if ($options.ImagesOnly -and $options.MediaOnly) {
        throw "--images-only 与 --media-only 不能同时使用"
    }
    return [PSCustomObject]$options
}

function Find-ExtractDir {
    param([string]$FolderArg)

    $candidates = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($FolderArg)) {
        $candidates.Add((Get-FullPath $FolderArg)) | Out-Null
    }
    else {
        $cwd = (Get-Location).Path
        $candidates.Add((Join-Path $cwd $DefaultExtractDirName)) | Out-Null
        $candidates.Add($cwd) | Out-Null
    }

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath (Join-Path $candidate $ManifestName) -PathType Leaf) {
            return (Get-FullPath $candidate)
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($FolderArg)) {
        $base = Get-FullPath $FolderArg
        if (Test-Path -LiteralPath $base -PathType Container) {
            $subs = @(Get-ChildItem -LiteralPath $base -Directory | Where-Object {
                Test-Path -LiteralPath (Join-Path $_.FullName $ManifestName) -PathType Leaf
            })
            if ($subs.Count -eq 1) { return $subs[0].FullName }
        }
    }

    throw "找不到 manifest.csv（提取时生成的清单）。请把提取得到的文件夹拖到本工具，或在该文件夹内运行。"
}

function Read-Manifest {
    param([string]$ManifestPath)
    return @(Import-Csv -LiteralPath $ManifestPath -Encoding UTF8)
}

function Get-PptxStem {
    param([object[]]$Rows)

    $counts = @{}
    foreach ($row in $Rows) {
        $outputFile = $row.output_file
        if ([string]::IsNullOrWhiteSpace($outputFile)) { continue }
        $stem = [System.IO.Path]::GetFileNameWithoutExtension(($outputFile -replace '/', '\'))
        $stem = [regex]::Replace($stem, '_\d{3}(?:_\d{2})?$', '')
        if (-not [string]::IsNullOrWhiteSpace($stem)) {
            if ($counts.ContainsKey($stem)) { $counts[$stem]++ } else { $counts[$stem] = 1 }
        }
    }
    if ($counts.Count -eq 0) { return $null }
    $best = $counts.GetEnumerator() |
        Sort-Object @{ Expression = { $_.Value }; Descending = $true }, @{ Expression = { $_.Key }; Descending = $true } |
        Select-Object -First 1
    return $best.Key
}

function Find-SourcePptx {
    param([string]$ExtractDir, [string]$Explicit, [string]$Stem)

    if (-not [string]::IsNullOrWhiteSpace($Explicit)) {
        $path = Get-FullPath $Explicit
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "指定的 PPTX 不存在：$path"
        }
        return $path
    }

    $parent = [System.IO.Path]::GetDirectoryName($ExtractDir)
    $grand = if ($parent) { [System.IO.Path]::GetDirectoryName($parent) } else { $null }
    $searchDirs = @($parent, $ExtractDir, (Get-Location).Path, $grand) | Where-Object { $_ }

    if (-not [string]::IsNullOrWhiteSpace($Stem)) {
        $seen = @{}
        foreach ($dir in $searchDirs) {
            $full = Get-FullPath $dir
            if ($seen.ContainsKey($full)) { continue }
            $seen[$full] = $true
            $candidate = Join-Path $full ($Stem + ".pptx")
            if ((Test-Path -LiteralPath $candidate -PathType Leaf) -and -not (Is-TempPptx $candidate)) {
                return (Get-FullPath $candidate)
            }
        }
    }

    foreach ($dir in @($parent, $ExtractDir) | Where-Object { $_ }) {
        if (Test-Path -LiteralPath $dir -PathType Container) {
            $pptxs = @(Get-ChildItem -LiteralPath $dir -Filter "*.pptx" -File | Where-Object { -not (Is-TempPptx $_.FullName) })
            if ($pptxs.Count -eq 1) { return $pptxs[0].FullName }
        }
    }

    $hint = if ($Stem) { "$Stem.pptx" } else { "原始 PPTX" }
    throw "找不到原始 PPTX（$hint）。请用 --pptx 指定原始文件，或把它放在提取文件夹旁边。"
}

function Get-DefaultOutputPath {
    param([string]$SourcePptx)
    $dir = [System.IO.Path]::GetDirectoryName($SourcePptx)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($SourcePptx)
    return (Join-Path $dir ($stem + $DefaultOutputSuffix + ".pptx"))
}

function Get-ScopeKinds {
    param([object]$Options)
    if ($Options.ImagesOnly) { return @("image") }
    if ($Options.MediaOnly) { return $MediaKinds }
    return $RestorableKinds
}

function Get-StreamBytes {
    param([System.IO.Stream]$Stream)
    $ms = New-Object System.IO.MemoryStream
    try {
        $Stream.CopyTo($ms)
        return $ms.ToArray()
    }
    finally { $ms.Dispose() }
}

function Get-Sha1OfBytes {
    param([byte[]]$Bytes)
    $sha = [System.Security.Cryptography.SHA1]::Create()
    try { return [System.BitConverter]::ToString($sha.ComputeHash($Bytes)) }
    finally { $sha.Dispose() }
}

function Get-Sha1OfFile {
    param([string]$Path)
    $sha = [System.Security.Cryptography.SHA1]::Create()
    $stream = [System.IO.File]::OpenRead($Path)
    try { return [System.BitConverter]::ToString($sha.ComputeHash($stream)) }
    finally {
        $stream.Dispose()
        $sha.Dispose()
    }
}

function Build-Plan {
    param([string]$ExtractDir, [object[]]$Rows, [string[]]$Scope, [System.IO.Compression.ZipArchive]$Zin)

    $entriesByName = @{}
    foreach ($entry in $Zin.Entries) {
        if (-not $entriesByName.ContainsKey($entry.FullName)) { $entriesByName[$entry.FullName] = $entry }
    }

    $plan = [PSCustomObject]@{
        Replacements   = [ordered]@{}
        MissingFiles   = New-Object System.Collections.Generic.List[object]
        MissingInZip   = New-Object System.Collections.Generic.List[string]
        Unchanged      = New-Object System.Collections.Generic.List[string]
        Conflicts      = New-Object System.Collections.Generic.List[object]
        PresentTargets = 0
    }

    $groupOrder = New-Object System.Collections.Generic.List[string]
    $groups = @{}
    $kinds = @{}
    foreach ($row in $Rows) {
        $kind = $row.kind
        if ($kind -eq "text" -or ($Scope -notcontains $kind)) { continue }
        $target = $row.target_part
        $outputFile = $row.output_file
        if ([string]::IsNullOrWhiteSpace($target) -or [string]::IsNullOrWhiteSpace($outputFile)) { continue }
        $sourceFile = Join-Path $ExtractDir ($outputFile -replace '/', '\')
        if (-not (Test-Path -LiteralPath $sourceFile -PathType Leaf)) {
            $plan.MissingFiles.Add($row) | Out-Null
            continue
        }
        if (-not $groups.ContainsKey($target)) {
            $groups[$target] = New-Object System.Collections.Generic.List[string]
            $groupOrder.Add($target) | Out-Null
            $kinds[$target] = $kind
        }
        $groups[$target].Add($sourceFile) | Out-Null
    }

    foreach ($target in $groupOrder) {
        if (-not $entriesByName.ContainsKey($target)) {
            $plan.MissingInZip.Add($target) | Out-Null
            continue
        }
        $plan.PresentTargets++
        $entry = $entriesByName[$target]
        $originalSize = $entry.Length
        $originalHash = $null

        $changedFiles = New-Object System.Collections.Generic.List[string]
        foreach ($file in $groups[$target]) {
            $fileSize = (Get-Item -LiteralPath $file).Length
            $isChanged = $false
            if ($fileSize -ne $originalSize) {
                $isChanged = $true
            }
            else {
                if ($null -eq $originalHash) {
                    $stream = $entry.Open()
                    try { $originalHash = Get-Sha1OfBytes (Get-StreamBytes $stream) }
                    finally { $stream.Dispose() }
                }
                if ((Get-Sha1OfFile $file) -ne $originalHash) { $isChanged = $true }
            }
            if ($isChanged) { $changedFiles.Add($file) | Out-Null }
        }

        if ($changedFiles.Count -eq 0) {
            $plan.Unchanged.Add($target) | Out-Null
            continue
        }

        $winner = $changedFiles[0]
        if ($changedFiles.Count -gt 1) {
            # Only hash the winner when there are sibling copies to compare.
            $winnerHash = Get-Sha1OfFile $winner
            for ($k = 1; $k -lt $changedFiles.Count; $k++) {
                if ((Get-Sha1OfFile $changedFiles[$k]) -ne $winnerHash) {
                    $plan.Conflicts.Add([PSCustomObject]@{ Target = $target; First = $winner; Second = $changedFiles[$k] }) | Out-Null
                    break
                }
            }
        }
        $plan.Replacements[$target] = [PSCustomObject]@{ SourceFile = $winner; Kind = $kinds[$target] }
    }

    return $plan
}

function Write-RestoredPptx {
    param([string]$SourcePptx, [object]$Replacements, [string]$OutputPath)

    $outputDir = [System.IO.Path]::GetDirectoryName($OutputPath)
    if ($outputDir -and -not (Test-Path -LiteralPath $outputDir -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    $tempPath = $OutputPath + ".tmp"
    if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force }

    $zin = [System.IO.Compression.ZipFile]::OpenRead($SourcePptx)
    try {
        $fileStream = [System.IO.File]::Open($tempPath, [System.IO.FileMode]::Create)
        try {
            $zout = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Create)
            try {
                foreach ($entry in $zin.Entries) {
                    $newEntry = $zout.CreateEntry($entry.FullName)
                    $outStream = $newEntry.Open()
                    try {
                        if ($Replacements.Contains($entry.FullName)) {
                            $bytes = [System.IO.File]::ReadAllBytes($Replacements[$entry.FullName].SourceFile)
                            $outStream.Write($bytes, 0, $bytes.Length)
                        }
                        else {
                            $inStream = $entry.Open()
                            try { $inStream.CopyTo($outStream) }
                            finally { $inStream.Dispose() }
                        }
                    }
                    finally { $outStream.Dispose() }
                }
            }
            finally { $zout.Dispose() }
        }
        finally { $fileStream.Dispose() }
    }
    catch {
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force }
        throw
    }
    finally { $zin.Dispose() }

    if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }
    Move-Item -LiteralPath $tempPath -Destination $OutputPath
}

function Write-Report {
    param([string]$SourcePptx, [string]$OutputPath, [object]$Plan, [bool]$DryRun)

    $byKind = @{}
    foreach ($key in $Plan.Replacements.Keys) {
        $kind = $Plan.Replacements[$key].Kind
        if ($byKind.ContainsKey($kind)) { $byKind[$kind]++ } else { $byKind[$kind] = 1 }
    }

    $verb = if ($DryRun) { "将写回" } else { "已写回" }
    if ($byKind.Count -gt 0) {
        $parts = @()
        foreach ($kind in ($byKind.Keys | Sort-Object)) { $parts += "$kind $($byKind[$kind])" }
        Write-Host "$verb $($Plan.Replacements.Count) 个元素（$([string]::Join('，', $parts))）。"
    }
    else {
        Write-Host "未检测到与原始不同的素材（输出与原始内容一致）。"
    }

    if ($Plan.Unchanged.Count -gt 0) {
        Write-Host "  · $($Plan.Unchanged.Count) 个素材与原始一致，保持不变。"
    }
    if ($Plan.MissingFiles.Count -gt 0) {
        Write-Host "  · $($Plan.MissingFiles.Count) 个清单项在文件夹中找不到对应文件，保留原始内容。"
    }
    if ($Plan.MissingInZip.Count -gt 0) {
        Write-Host "  ⚠ $($Plan.MissingInZip.Count) 个部件在原始 PPTX 中不存在，已跳过。请确认 --pptx 指向的是同一个文件。"
    }
    foreach ($conflict in $Plan.Conflicts) {
        $first = [System.IO.Path]::GetFileName($conflict.First)
        $second = [System.IO.Path]::GetFileName($conflict.Second)
        Write-Host "  ⚠ 同一部件 $($conflict.Target) 有多个不同的修改版本：$first / $second。"
    }

    if ($DryRun) {
        Write-Host "（预演）将生成：$OutputPath"
    }
    else {
        Write-Host "完成 ✅  已生成：$OutputPath"
        Write-Host "原始文件未改动：$SourcePptx"
    }
}

function Write-ConflictErrors {
    param([object]$Plan)

    Write-Host "检测到同一 PPTX 部件有多个不同的修改版本，已停止还原。"
    Write-Host "请只保留一个修改版本，或让这些副本内容一致后再运行。"
    foreach ($conflict in $Plan.Conflicts) {
        Write-Host "  - $($conflict.Target): $($conflict.First) / $($conflict.Second)"
    }
}

function Invoke-Restore {
    param([string[]]$Argv)

    $options = Parse-Arguments $Argv
    $extractDir = Find-ExtractDir $options.Folder
    $manifestPath = Join-Path $extractDir $ManifestName
    $rows = @(Read-Manifest $manifestPath)
    if ($rows.Count -eq 0) {
        Write-Host "清单为空：$manifestPath"
        return 1
    }

    $stem = Get-PptxStem $rows
    $sourcePptx = Find-SourcePptx $extractDir $options.Pptx $stem

    try {
        $probe = [System.IO.Compression.ZipFile]::OpenRead($sourcePptx)
        $probe.Dispose()
    }
    catch {
        Write-Host "不是有效的 .pptx/zip 文件：$sourcePptx"
        return 1
    }

    if ($options.Output) {
        $outputPath = Get-FullPath $options.Output
    }
    else {
        $outputPath = Get-DefaultOutputPath $sourcePptx
    }

    if (-not $options.DryRun -and (Test-Path -LiteralPath $outputPath -PathType Leaf) -and -not $options.Overwrite) {
        Write-Host "输出文件已存在：$outputPath"
        Write-Host "使用 --overwrite 覆盖，或用 -o 指定其它文件名。"
        return 1
    }

    $scope = Get-ScopeKinds $options
    $zin = [System.IO.Compression.ZipFile]::OpenRead($sourcePptx)
    try {
        $plan = Build-Plan $extractDir $rows $scope $zin
    }
    finally { $zin.Dispose() }

    Write-Host "原始 PPTX：$sourcePptx"
    Write-Host "素材文件夹：$extractDir"

    $totalTargets = $plan.PresentTargets + $plan.MissingInZip.Count
    if ($totalTargets -eq 0) {
        Write-Host "没有找到任何可写回的素材文件。请确认筛选项正确，且文件名未被改动。"
        return 1
    }
    if ($plan.PresentTargets -eq 0) {
        Write-Host "这些素材（$($plan.MissingInZip.Count) 个）在该 PPTX 中都不存在，很可能不是同一个文件。请用 --pptx 指定正确的原始 PPTX。"
        return 1
    }
    if ($plan.Conflicts.Count -gt 0) {
        Write-ConflictErrors $plan
        return 1
    }

    if (-not $options.DryRun) {
        Write-RestoredPptx $sourcePptx $plan.Replacements $outputPath
    }

    Write-Report $sourcePptx $outputPath $plan $options.DryRun
    return 0
}

try {
    $code = Invoke-Restore -Argv $LauncherArgs
    exit $code
}
catch {
    Write-Host "还原失败：$($_.Exception.Message)"
    exit 1
}

# PYTHON_CODE_BELOW
#!/usr/bin/env python3
"""
PPTX 元素还原器 - Rebuild a .pptx from extracted media without moving anything.

This is the companion to extract_pptx_elements.py. It reads the manifest.csv
produced during extraction and writes each (possibly edited) media file back
into its exact part inside a *copy* of the original .pptx.

Why positions never change:
  An element's position and size live in the slide XML (<a:off>/<a:ext>), not
  in the image bytes. This tool only swaps the bytes of each media part and
  never touches any XML, so every element keeps its original position exactly.
  This is a strict guarantee for images and, in practice, for every other
  element too.

Typical workflow:
  1. extract_pptx_elements.py "deck.pptx"        -> pptx_extracted_elements/
  2. Batch-edit images in 图片/ (e.g. remove a watermark), keep the file names.
  3. restore_pptx_elements.py                     -> deck_restored.pptx

Examples:
  python3 restore_pptx_elements.py
  python3 restore_pptx_elements.py pptx_extracted_elements
  python3 restore_pptx_elements.py pptx_extracted_elements --pptx deck.pptx
  python3 restore_pptx_elements.py pptx_extracted_elements -o deck_fixed.pptx
  python3 restore_pptx_elements.py --images-only
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Iterable


VERSION = "1.2.2"

MANIFEST_NAME = "manifest.csv"
DEFAULT_EXTRACT_DIR_NAME = "pptx_extracted_elements"
DEFAULT_OUTPUT_SUFFIX = "_restored"

# Kinds whose extracted file is the literal part inside the pptx and can be
# safely written back byte-for-byte. "text" is intentionally excluded: its
# extracted .txt is a lossy plain-text dump of the slide XML and must never be
# written over the slide.
RESTORABLE_KINDS = {
    "audio",
    "chart",
    "chart_colors",
    "chart_style",
    "diagram",
    "embedded",
    "image",
    "ole",
    "video",
}
MEDIA_KINDS = {"image", "video", "audio"}

# Magic-number signatures used only for a soft "format changed" warning.
IMAGE_SIGNATURES = {
    ".png": (b"\x89PNG\r\n\x1a\n",),
    ".jpg": (b"\xff\xd8\xff",),
    ".jpeg": (b"\xff\xd8\xff",),
    ".gif": (b"GIF87a", b"GIF89a"),
    ".bmp": (b"BM",),
    ".tif": (b"II*\x00", b"MM\x00*"),
    ".tiff": (b"II*\x00", b"MM\x00*"),
    ".webp": (b"RIFF",),
}


@dataclass(frozen=True)
class ManifestRow:
    slide: str
    output_file: str
    kind: str
    source_part: str
    target_part: str
    rel_id: str
    rel_type: str


@dataclass(frozen=True)
class Replacement:
    target_part: str
    source_file: Path
    kind: str


@dataclass
class RestorePlan:
    # target_part -> file that should replace it (only parts that truly change)
    replacements: dict[str, Replacement] = field(default_factory=dict)
    # manifest rows whose file is no longer on disk (original bytes kept)
    missing_files: list[ManifestRow] = field(default_factory=list)
    # target parts present on disk but absent from the pptx (wrong pptx?)
    missing_in_zip: list[str] = field(default_factory=list)
    # parts whose on-disk file is byte-identical to the original (no edit)
    unchanged: list[str] = field(default_factory=list)
    # parts with several edited copies that disagree (target, winner, other)
    conflicts: list[tuple[str, Path, Path]] = field(default_factory=list)
    # distinct in-scope target parts that exist inside the pptx
    present_targets: int = 0

    @property
    def total_targets(self) -> int:
        return self.present_targets + len(self.missing_in_zip)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "PPTX 元素还原器 - Write extracted (and possibly edited) media back "
            "into a copy of the original .pptx, keeping every element in its "
            "original position."
        )
    )
    parser.add_argument(
        "folder",
        nargs="?",
        type=Path,
        default=None,
        help=(
            "Extracted folder containing manifest.csv (the output of "
            "extract_pptx_elements). If omitted, the tool looks for "
            f"'{DEFAULT_EXTRACT_DIR_NAME}' or a manifest.csv in the current folder."
        ),
    )
    parser.add_argument(
        "--pptx",
        type=Path,
        default=None,
        help=(
            "Original .pptx to rebuild from. If omitted, it is auto-detected "
            "next to the extracted folder using the file name recorded in the "
            "manifest."
        ),
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help=(
            "Output .pptx path. Default: '<original>_restored.pptx' next to the "
            "original file."
        ),
    )
    scope_group = parser.add_mutually_exclusive_group()
    scope_group.add_argument(
        "--images-only",
        action="store_true",
        help="Only write images back; leave every other element untouched.",
    )
    scope_group.add_argument(
        "--media-only",
        action="store_true",
        help="Only write images, videos, and audio back.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite the output .pptx if it already exists.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be written without creating the output file.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {VERSION}",
    )
    return parser.parse_args(argv)


def is_temp_pptx(path: Path) -> bool:
    return path.name.startswith("~$") or path.name.startswith(".~")


def find_extract_dir(folder_arg: Path | None) -> Path:
    """Resolve the folder that holds manifest.csv."""
    candidates: list[Path] = []
    if folder_arg is not None:
        candidates.append(folder_arg.expanduser())
    else:
        cwd = Path.cwd()
        candidates.append(cwd / DEFAULT_EXTRACT_DIR_NAME)
        candidates.append(cwd)

    for candidate in candidates:
        if (candidate / MANIFEST_NAME).is_file():
            return candidate.resolve()

    # A folder was given but the manifest is one level down (multi-input layout:
    # <out>/<deck-name>/manifest.csv with a single subfolder).
    if folder_arg is not None:
        base = folder_arg.expanduser()
        if base.is_dir():
            subdirs = [p for p in base.iterdir() if (p / MANIFEST_NAME).is_file()]
            if len(subdirs) == 1:
                return subdirs[0].resolve()

    searched = "\n".join(f"  - {c / MANIFEST_NAME}" for c in candidates)
    raise FileNotFoundError(
        "找不到 manifest.csv（提取时生成的清单）。已查找：\n"
        f"{searched}\n"
        "请把提取得到的文件夹拖到本工具，或在该文件夹内运行。"
    )


def read_manifest(manifest_path: Path) -> list[ManifestRow]:
    rows: list[ManifestRow] = []
    with manifest_path.open("r", newline="", encoding="utf-8-sig") as csv_file:
        reader = csv.DictReader(csv_file)
        for raw in reader:
            rows.append(
                ManifestRow(
                    slide=(raw.get("slide") or "").strip(),
                    output_file=(raw.get("output_file") or "").strip(),
                    kind=(raw.get("kind") or "").strip(),
                    source_part=(raw.get("source_part") or "").strip(),
                    target_part=(raw.get("target_part") or "").strip(),
                    rel_id=(raw.get("relationship_id") or "").strip(),
                    rel_type=(raw.get("relationship_type") or "").strip(),
                )
            )
    return rows


def derive_pptx_stem(rows: Iterable[ManifestRow]) -> str | None:
    """Recover the original pptx file stem from output file names.

    Output files are named '<stem>_<NNN>[_<NN>].ext'. Strip the trailing slide
    number (and optional duplicate index) and return the most common stem.
    """
    counts: dict[str, int] = {}
    pattern = re.compile(r"_\d{3}(?:_\d{2})?$")
    for row in rows:
        if not row.output_file:
            continue
        stem = pattern.sub("", PurePosixPath(row.output_file).stem)
        if stem:
            counts[stem] = counts.get(stem, 0) + 1
    if not counts:
        return None
    return max(counts, key=lambda key: (counts[key], key))


def find_source_pptx(
    extract_dir: Path,
    explicit: Path | None,
    stem: str | None,
) -> Path:
    if explicit is not None:
        path = explicit.expanduser()
        if not path.is_file():
            raise FileNotFoundError(f"指定的 PPTX 不存在：{path}")
        return path.resolve()

    search_dirs = [
        extract_dir.parent,         # default layout: folder sits next to the pptx
        extract_dir,
        Path.cwd(),
        extract_dir.parent.parent,  # multi-input layout
    ]
    seen: set[Path] = set()

    if stem:
        for directory in search_dirs:
            directory = directory.resolve()
            if directory in seen:
                continue
            seen.add(directory)
            candidate = directory / f"{stem}.pptx"
            if candidate.is_file() and not is_temp_pptx(candidate):
                return candidate.resolve()

    # Fallback: a single .pptx sitting next to the extracted folder.
    for directory in (extract_dir.parent, extract_dir):
        pptxs = [
            p for p in directory.glob("*.pptx") if p.is_file() and not is_temp_pptx(p)
        ]
        if len(pptxs) == 1:
            return pptxs[0].resolve()

    hint = f"{stem}.pptx" if stem else "原始 PPTX"
    raise FileNotFoundError(
        f"找不到原始 PPTX（{hint}）。\n"
        "请用 --pptx 指定原始文件，或把它放在提取文件夹旁边。"
    )


def default_output_path(source_pptx: Path) -> Path:
    return source_pptx.with_name(f"{source_pptx.stem}{DEFAULT_OUTPUT_SUFFIX}.pptx")


def kinds_in_scope(*, images_only: bool, media_only: bool) -> set[str]:
    if images_only:
        return {"image"}
    if media_only:
        return set(MEDIA_KINDS)
    return set(RESTORABLE_KINDS)


def signature_mismatch(target_part: str, data: bytes) -> bool:
    """Return True only when an image part's bytes clearly don't match its ext."""
    ext = PurePosixPath(target_part).suffix.lower()
    signatures = IMAGE_SIGNATURES.get(ext)
    if not signatures:
        return False
    return not any(data.startswith(sig) for sig in signatures)


def build_plan(
    extract_dir: Path,
    rows: list[ManifestRow],
    scope: set[str],
    zin: zipfile.ZipFile,
) -> RestorePlan:
    """Decide which parts to swap, reading the original pptx for comparison.

    The same part can be referenced by several slides (a shared logo/background).
    Among the on-disk copies of one part, the copy that actually differs from the
    original wins, so editing any single copy is enough. Copies identical to the
    original are reported as unchanged; genuinely conflicting edits are flagged.
    """
    names = set(zin.namelist())
    plan = RestorePlan()
    original_cache: dict[str, bytes] = {}

    # Group on-disk files by their target part, preserving manifest order.
    groups: dict[str, list[Path]] = {}
    kinds: dict[str, str] = {}
    for row in rows:
        if row.kind == "text" or row.kind not in scope:
            continue
        if not row.target_part or not row.output_file:
            continue
        source_file = extract_dir / row.output_file
        if not source_file.is_file():
            plan.missing_files.append(row)
            continue
        groups.setdefault(row.target_part, []).append(source_file)
        kinds.setdefault(row.target_part, row.kind)

    for target, files in groups.items():
        if target not in names:
            plan.missing_in_zip.append(target)
            continue
        plan.present_targets += 1
        original_size = zin.getinfo(target).file_size

        def is_changed(path: Path) -> bool:
            if path.stat().st_size != original_size:
                return True
            if target not in original_cache:
                original_cache[target] = zin.read(target)
            return path.read_bytes() != original_cache[target]

        changed_files = [path for path in files if is_changed(path)]
        if not changed_files:
            plan.unchanged.append(target)
            continue

        winner = changed_files[0]
        if len(changed_files) > 1:
            # Only read the winner's bytes when there are sibling copies to
            # check against; the common single-copy case needs no read here.
            winner_bytes = winner.read_bytes()
            for other in changed_files[1:]:
                if other.read_bytes() != winner_bytes:
                    plan.conflicts.append((target, winner, other))
                    break

        plan.replacements[target] = Replacement(target, winner, kinds[target])

    return plan


def rewrite_pptx(
    source_pptx: Path,
    replacements: dict[str, Replacement],
    output_path: Path,
) -> list[str]:
    """Copy source_pptx to output_path, swapping bytes for replaced parts.

    Every other part - including all slide XML - is copied verbatim, so every
    element keeps its original position and size.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_name(output_path.name + ".tmp")
    replaced: list[str] = []

    with zipfile.ZipFile(source_pptx) as zin:
        try:
            with zipfile.ZipFile(temp_path, "w") as zout:
                for item in zin.infolist():
                    replacement = replacements.get(item.filename)
                    if replacement is not None:
                        # Reuse the original ZipInfo so compress_type, date_time
                        # and flags are preserved for the swapped part. Read the
                        # replacement just-in-time so at most one media file is
                        # held in memory (matters for large video/audio parts).
                        zout.writestr(item, replacement.source_file.read_bytes())
                        replaced.append(item.filename)
                    else:
                        zout.writestr(item, zin.read(item.filename))
        except BaseException:
            temp_path.unlink(missing_ok=True)
            raise

    temp_path.replace(output_path)
    return replaced


def warn_signature_mismatches(plan: RestorePlan) -> None:
    for target, repl in plan.replacements.items():
        with repl.source_file.open("rb") as handle:
            head = handle.read(16)
        if signature_mismatch(target, head):
            print(
                f"  ⚠ 注意：{repl.source_file.name} 的实际格式可能与 "
                f"{PurePosixPath(target).suffix} 不一致，PowerPoint 仍会尝试显示。",
                file=sys.stderr,
            )


def print_report(
    *,
    source_pptx: Path,
    output_path: Path,
    plan: RestorePlan,
    dry_run: bool,
) -> None:
    by_kind: dict[str, int] = {}
    for repl in plan.replacements.values():
        by_kind[repl.kind] = by_kind.get(repl.kind, 0) + 1

    verb = "将写回" if dry_run else "已写回"
    if by_kind:
        summary = "，".join(f"{kind} {count}" for kind, count in sorted(by_kind.items()))
        print(f"{verb} {len(plan.replacements)} 个元素（{summary}）。")
    else:
        print("未检测到与原始不同的素材（输出与原始内容一致）。")

    warn_signature_mismatches(plan)

    if plan.unchanged:
        print(f"  · {len(plan.unchanged)} 个素材与原始一致，保持不变。")
    if plan.missing_files:
        print(
            f"  · {len(plan.missing_files)} 个清单项在文件夹中找不到对应文件，"
            "保留原始内容。"
        )
    if plan.missing_in_zip:
        print(
            f"  ⚠ {len(plan.missing_in_zip)} 个部件在原始 PPTX 中不存在，已跳过。"
            "请确认 --pptx 指向的是同一个文件。",
            file=sys.stderr,
        )
    for target, first, second in plan.conflicts:
        print(
            f"  ⚠ 同一部件 {target} 有多个不同的修改版本："
            f"{first.name} / {second.name}。",
            file=sys.stderr,
        )

    if dry_run:
        print(f"（预演）将生成：{output_path}")
    else:
        print(f"完成 ✅  已生成：{output_path}")
        print(f"原始文件未改动：{source_pptx}")


def print_conflict_errors(plan: RestorePlan) -> None:
    print(
        "检测到同一 PPTX 部件有多个不同的修改版本，已停止还原。\n"
        "请只保留一个修改版本，或让这些副本内容一致后再运行。",
        file=sys.stderr,
    )
    for target, first, second in plan.conflicts:
        print(
            f"  - {target}: {first} / {second}",
            file=sys.stderr,
        )


def restore(
    *,
    folder_arg: Path | None,
    pptx_arg: Path | None,
    output_arg: Path | None,
    images_only: bool,
    media_only: bool,
    overwrite: bool,
    dry_run: bool,
) -> int:
    extract_dir = find_extract_dir(folder_arg)
    manifest_path = extract_dir / MANIFEST_NAME
    rows = read_manifest(manifest_path)
    if not rows:
        print(f"清单为空：{manifest_path}", file=sys.stderr)
        return 1

    stem = derive_pptx_stem(rows)
    source_pptx = find_source_pptx(extract_dir, pptx_arg, stem)
    if not zipfile.is_zipfile(source_pptx):
        print(f"不是有效的 .pptx/zip 文件：{source_pptx}", file=sys.stderr)
        return 1

    if output_arg is not None:
        output_path = output_arg.expanduser()
        if not output_path.is_absolute():
            output_path = (Path.cwd() / output_path).resolve()
    else:
        output_path = default_output_path(source_pptx)

    if not dry_run and output_path.exists() and not overwrite:
        print(
            f"输出文件已存在：{output_path}\n"
            "使用 --overwrite 覆盖，或用 -o 指定其它文件名。",
            file=sys.stderr,
        )
        return 1

    scope = kinds_in_scope(images_only=images_only, media_only=media_only)
    with zipfile.ZipFile(source_pptx) as zin:
        plan = build_plan(extract_dir, rows, scope, zin)

    print(f"原始 PPTX：{source_pptx}")
    print(f"素材文件夹：{extract_dir}")

    if plan.total_targets == 0:
        print(
            "没有找到任何可写回的素材文件。请确认筛选项正确，且文件名未被改动。",
            file=sys.stderr,
        )
        return 1
    if plan.present_targets == 0:
        print(
            f"这些素材（{len(plan.missing_in_zip)} 个）在该 PPTX 中都不存在，"
            "很可能不是同一个文件。请用 --pptx 指定正确的原始 PPTX。",
            file=sys.stderr,
        )
        return 1
    if plan.conflicts:
        print_conflict_errors(plan)
        return 1

    if not dry_run:
        rewrite_pptx(source_pptx, plan.replacements, output_path)

    print_report(
        source_pptx=source_pptx,
        output_path=output_path,
        plan=plan,
        dry_run=dry_run,
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        return restore(
            folder_arg=args.folder,
            pptx_arg=args.pptx,
            output_arg=args.output,
            images_only=args.images_only,
            media_only=args.media_only,
            overwrite=args.overwrite,
            dry_run=args.dry_run,
        )
    except (FileNotFoundError, ValueError, zipfile.BadZipFile) as exc:
        print(f"还原失败：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
