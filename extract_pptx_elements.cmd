@echo off
setlocal EnableExtensions
set "VERSION=V1.1.2"
set "PPTX_EXTRACTOR_SCRIPT=%~f0"
set "PPTX_EXTRACTOR_PY_TEMP=%TEMP%\extract_pptx_elements_py_%RANDOM%%RANDOM%.py"
set "PPTX_EXTRACTOR_PS_TEMP=%TEMP%\extract_pptx_elements_ps_%RANDOM%%RANDOM%.ps1"
chcp 65001 >nul
REM ============================================================
REM  PPTX 元素提取工具 %VERSION% - Windows 单文件启动器
REM  直接双击或拖拽 .pptx 到本文件即可运行
REM ============================================================

set "PPTX_EXTRACTOR_INPUT="
if not "%~1"=="" goto choose_runtime

echo.
echo 请将 .pptx 文件拖到此窗口，或输入 .pptx 文件路径。
set /p "PPTX_EXTRACTOR_INPUT=路径: "
set "PPTX_EXTRACTOR_INPUT=%PPTX_EXTRACTOR_INPUT:"=%"

:choose_runtime
py -3 --version >nul 2>&1
if errorlevel 1 goto try_python

echo → 使用内嵌 Python 运行 ...
py -3 -c "from pathlib import Path; import os; text=Path(os.environ['PPTX_EXTRACTOR_SCRIPT']).read_text(encoding='utf-8').replace('\r\n','\n'); marker=chr(35)+' PYTHON_CODE_BELOW'; Path(os.environ['PPTX_EXTRACTOR_PY_TEMP']).write_text(text.split('\n'+marker+'\n',1)[1], encoding='utf-8')"
if errorlevel 1 goto python_extract_failed
if defined PPTX_EXTRACTOR_INPUT goto run_py_launcher_input
py -3 "%PPTX_EXTRACTOR_PY_TEMP%" %*
goto after_py_launcher_run

:run_py_launcher_input
py -3 "%PPTX_EXTRACTOR_PY_TEMP%" "%PPTX_EXTRACTOR_INPUT%"

:after_py_launcher_run
set "EXIT_CODE=%ERRORLEVEL%"
del "%PPTX_EXTRACTOR_PY_TEMP%" >nul 2>&1
goto after_run

:try_python
python --version >nul 2>&1
if errorlevel 1 goto try_python3

echo → 使用内嵌 Python 运行 ...
python -c "from pathlib import Path; import os; text=Path(os.environ['PPTX_EXTRACTOR_SCRIPT']).read_text(encoding='utf-8').replace('\r\n','\n'); marker=chr(35)+' PYTHON_CODE_BELOW'; Path(os.environ['PPTX_EXTRACTOR_PY_TEMP']).write_text(text.split('\n'+marker+'\n',1)[1], encoding='utf-8')"
if errorlevel 1 goto python_extract_failed
if defined PPTX_EXTRACTOR_INPUT goto run_python_input
python "%PPTX_EXTRACTOR_PY_TEMP%" %*
goto after_python_run

:run_python_input
python "%PPTX_EXTRACTOR_PY_TEMP%" "%PPTX_EXTRACTOR_INPUT%"

:after_python_run
set "EXIT_CODE=%ERRORLEVEL%"
del "%PPTX_EXTRACTOR_PY_TEMP%" >nul 2>&1
goto after_run

:try_python3
python3 --version >nul 2>&1
if errorlevel 1 goto powershell_fallback

echo → 使用内嵌 Python 运行 ...
python3 -c "from pathlib import Path; import os; text=Path(os.environ['PPTX_EXTRACTOR_SCRIPT']).read_text(encoding='utf-8').replace('\r\n','\n'); marker=chr(35)+' PYTHON_CODE_BELOW'; Path(os.environ['PPTX_EXTRACTOR_PY_TEMP']).write_text(text.split('\n'+marker+'\n',1)[1], encoding='utf-8')"
if errorlevel 1 goto python_extract_failed
if defined PPTX_EXTRACTOR_INPUT goto run_python3_input
python3 "%PPTX_EXTRACTOR_PY_TEMP%" %*
goto after_python3_run

:run_python3_input
python3 "%PPTX_EXTRACTOR_PY_TEMP%" "%PPTX_EXTRACTOR_INPUT%"

:after_python3_run
set "EXIT_CODE=%ERRORLEVEL%"
del "%PPTX_EXTRACTOR_PY_TEMP%" >nul 2>&1
goto after_run

:python_extract_failed
echo 内嵌 Python 提取失败，正在尝试 PowerShell 回退 ...
del "%PPTX_EXTRACTOR_PY_TEMP%" >nul 2>&1

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
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$script=[IO.File]::ReadAllText($env:PPTX_EXTRACTOR_SCRIPT,[Text.Encoding]::UTF8); $psParts=[regex]::Split($script,'(?m)^# POWERSHELL_START\r?$',2); if($psParts.Count -lt 2){throw 'PowerShell section missing'}; $payload=[regex]::Split($psParts[1],'(?m)^# PYTHON_CODE_BELOW\r?$',2)[0]; $encoding=New-Object System.Text.UTF8Encoding $true; [IO.File]::WriteAllText($env:PPTX_EXTRACTOR_PS_TEMP,$payload,$encoding)"
if errorlevel 1 (
    echo 启动器损坏或 PowerShell 脚本提取失败。
    echo 按任意键关闭此窗口 ...
    pause >nul
    exit /b 1
)

if defined PPTX_EXTRACTOR_INPUT goto run_powershell_input
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PPTX_EXTRACTOR_PS_TEMP%" %*
goto after_powershell_run

:run_powershell_input
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PPTX_EXTRACTOR_PS_TEMP%" "%PPTX_EXTRACTOR_INPUT%"

:after_powershell_run
set "EXIT_CODE=%ERRORLEVEL%"
del "%PPTX_EXTRACTOR_PS_TEMP%" >nul 2>&1

:after_run
echo.
echo 完成！(退出码: %EXIT_CODE%)
echo 提醒：默认输出文件夹会出现在 PPTX 文件旁边，名称为 pptx_extracted_elements。
echo 按任意键关闭此窗口 ...
pause >nul
exit /b %EXIT_CODE%

# POWERSHELL_START
param([string[]]$LauncherArgs)

$ErrorActionPreference = "Stop"

$Version = "V1.1.2"
$DefaultOutputDirName = "pptx_extracted_elements"

$PackageRelsNs = "http://schemas.openxmlformats.org/package/2006/relationships"
$PresentationNs = "http://schemas.openxmlformats.org/presentationml/2006/main"
$OfficeRelsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
$DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main"

$ImageExts = @(".bmp", ".emf", ".gif", ".jfif", ".jpeg", ".jpg", ".png", ".svg", ".tif", ".tiff", ".webp", ".wmf")
$VideoExts = @(".3gp", ".asf", ".avi", ".m4v", ".mkv", ".mov", ".mp4", ".mpeg", ".mpg", ".swf", ".webm", ".wmv")
$AudioExts = @(".aac", ".aif", ".aiff", ".m4a", ".mid", ".midi", ".mp3", ".oga", ".ogg", ".wav", ".wma")
$RelSkipWords = @(
    "/hyperlink",
    "/notesSlide",
    "/presProps",
    "/printerSettings",
    "/slideLayout",
    "/slideMaster",
    "/theme",
    "/viewProps"
)

$KindFolderNames = @{
    audio = "音频"
    chart = "图表"
    chart_colors = "图表"
    chart_style = "图表"
    diagram = "图示"
    embedded = "嵌入文件"
    image = "图片"
    ole = "嵌入文件"
    text = "文本"
    video = "视频"
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Show-Banner {
    Write-Host "╔══════════════════════════════════════════╗"
    Write-Host "║     PPTX 元素提取工具 $Version           ║"
    Write-Host "║     extract_pptx_elements               ║"
    Write-Host "╚══════════════════════════════════════════╝"
    Write-Host ""
    Write-Host "用法: 将 .pptx 文件拖拽到此窗口，然后按回车"
    Write-Host "      或直接输入 .pptx 文件路径"
    Write-Host ""
    Write-Host "常用选项:"
    Write-Host "  --no-text     不提取幻灯片文本"
    Write-Host "  --media-only  仅提取图片/视频/音频"
    Write-Host "  --overwrite   覆盖已有文件"
    Write-Host "  -o 目录名     指定输出目录"
    Write-Host ""
    Write-Host "──────────────────────────────────────────"
}

function Split-CommandLine {
    param([string]$Text)

    $items = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return @()
    }

    $matches = [regex]::Matches($Text, '("[^"]*"|''[^'']*''|\S+)')
    foreach ($match in $matches) {
        $value = $match.Value
        if (($value.StartsWith('"') -and $value.EndsWith('"')) -or ($value.StartsWith("'") -and $value.EndsWith("'"))) {
            $value = $value.Substring(1, $value.Length - 2)
        }
        $items.Add($value) | Out-Null
    }
    return $items.ToArray()
}

function Show-Usage {
    Write-Host "用法:"
    Write-Host "  extract_pptx_elements.cmd presentation.pptx"
    Write-Host "  extract_pptx_elements.cmd presentation.pptx -o my_assets"
    Write-Host "  extract_pptx_elements.cmd presentation.pptx --no-text"
    Write-Host "  extract_pptx_elements.cmd presentation.pptx --media-only"
    Write-Host "  extract_pptx_elements.cmd presentation.pptx --overwrite"
}

function Parse-Arguments {
    param([string[]]$Argv)

    $pptxList = New-Object System.Collections.Generic.List[string]
    $options = [ordered]@{
        Pptx = $pptxList
        Output = $null
        WithText = $true
        MediaOnly = $false
        Overwrite = $false
    }

    for ($i = 0; $i -lt $Argv.Count; $i++) {
        $item = $Argv[$i]
        switch -Regex ($item) {
            '^(--help|-h|/\?)$' {
                Show-Usage
                exit 0
            }
            '^--version$' {
                Write-Host "extract_pptx_elements.cmd $Version"
                exit 0
            }
            '^--with-text$' {
                $options.WithText = $true
                continue
            }
            '^--no-text$' {
                $options.WithText = $false
                continue
            }
            '^--media-only$' {
                $options.MediaOnly = $true
                continue
            }
            '^--overwrite$' {
                $options.Overwrite = $true
                continue
            }
            '^(-o|--output)$' {
                if ($i + 1 -ge $Argv.Count) {
                    throw "缺少输出目录参数：$item"
                }
                $i++
                $options.Output = $Argv[$i]
                continue
            }
            default {
                $options.Pptx.Add($item) | Out-Null
            }
        }
    }

    return [PSCustomObject]$options
}

function Is-TempPptx {
    param([string]$Path)
    $name = [System.IO.Path]::GetFileName($Path)
    return ($name.StartsWith("~$") -or $name.StartsWith(".~"))
}

function Get-FullPath {
    param([string]$Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $Path))
}

function Find-InputFiles {
    param([object]$Options)

    $files = New-Object System.Collections.Generic.List[string]
    if ($Options.Pptx.Count -gt 0) {
        foreach ($path in $Options.Pptx) {
            $files.Add((Get-FullPath $path)) | Out-Null
        }
        return $files.ToArray()
    }

    foreach ($file in Get-ChildItem -LiteralPath (Get-Location).Path -Filter "*.pptx" -File | Sort-Object Name) {
        if (-not (Is-TempPptx $file.FullName)) {
            $files.Add($file.FullName) | Out-Null
        }
    }
    return $files.ToArray()
}

function Normalize-PartPath {
    param([string]$Path)

    $stack = New-Object System.Collections.Generic.List[string]
    foreach ($segment in ($Path -replace '\\', '/').Split('/')) {
        if ([string]::IsNullOrEmpty($segment) -or $segment -eq ".") {
            continue
        }
        if ($segment -eq "..") {
            if ($stack.Count -gt 0) {
                $stack.RemoveAt($stack.Count - 1)
            }
            continue
        }
        $stack.Add($segment) | Out-Null
    }
    return ($stack -join "/")
}

function Resolve-TargetPart {
    param(
        [string]$SourcePart,
        [string]$Target
    )

    if ($Target.StartsWith("/")) {
        return (Normalize-PartPath $Target.TrimStart("/"))
    }

    $index = $SourcePart.LastIndexOf("/")
    if ($index -ge 0) {
        $sourceDir = $SourcePart.Substring(0, $index)
        return (Normalize-PartPath "$sourceDir/$Target")
    }
    return (Normalize-PartPath $Target)
}

function Get-RelsPartName {
    param([string]$Part)

    $index = $Part.LastIndexOf("/")
    if ($index -ge 0) {
        $dir = $Part.Substring(0, $index)
        $file = $Part.Substring($index + 1)
        return "$dir/_rels/$file.rels"
    }
    return "_rels/$Part.rels"
}

function Get-ZipEntry {
    param(
        [object]$Zip,
        [string]$Part
    )
    return $Zip.GetEntry($Part)
}

function Read-ZipText {
    param(
        [object]$Zip,
        [string]$Part
    )

    $entry = Get-ZipEntry $Zip $Part
    if ($null -eq $entry) {
        return $null
    }

    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Read-XmlPart {
    param(
        [object]$Zip,
        [string]$Part
    )

    $text = Read-ZipText $Zip $Part
    if ($null -eq $text) {
        return $null
    }

    try {
        $xml = New-Object System.Xml.XmlDocument
        $xml.PreserveWhitespace = $false
        $xml.LoadXml($text)
        return $xml
    }
    catch {
        Write-Warning "could not parse XML part $Part`: $($_.Exception.Message)"
        return $null
    }
}

function Get-Relationships {
    param(
        [object]$Zip,
        [string]$Part
    )

    $relsPart = Get-RelsPartName $Part
    $xml = Read-XmlPart $Zip $relsPart
    $relationships = New-Object System.Collections.Generic.List[object]
    if ($null -eq $xml) {
        return $relationships.ToArray()
    }

    foreach ($node in $xml.DocumentElement.ChildNodes) {
        if ($node.LocalName -ne "Relationship") {
            continue
        }
        $relationships.Add([PSCustomObject]@{
            Id = $node.GetAttribute("Id")
            Type = $node.GetAttribute("Type")
            Target = $node.GetAttribute("Target")
            TargetMode = $node.GetAttribute("TargetMode")
        }) | Out-Null
    }
    return $relationships.ToArray()
}

function Get-SlidePartsInOrder {
    param([object]$Zip)

    $presentation = Read-XmlPart $Zip "ppt/presentation.xml"
    $presentationRels = Get-Relationships $Zip "ppt/presentation.xml"
    $relsById = @{}
    foreach ($rel in $presentationRels) {
        if (-not [string]::IsNullOrEmpty($rel.Id)) {
            $relsById[$rel.Id] = $rel
        }
    }

    if ($null -ne $presentation -and $relsById.Count -gt 0) {
        $ordered = New-Object System.Collections.Generic.List[string]
        foreach ($node in $presentation.GetElementsByTagName("sldId", $PresentationNs)) {
            $relId = $node.GetAttribute("id", $OfficeRelsNs)
            if (-not [string]::IsNullOrEmpty($relId) -and $relsById.ContainsKey($relId)) {
                $target = Resolve-TargetPart "ppt/presentation.xml" $relsById[$relId].Target
                if ($null -ne (Get-ZipEntry $Zip $target)) {
                    $ordered.Add($target) | Out-Null
                }
            }
        }
        if ($ordered.Count -gt 0) {
            return $ordered.ToArray()
        }
    }

    return @(
        $Zip.Entries |
            Where-Object { $_.FullName.StartsWith("ppt/slides/") -and $_.FullName.EndsWith(".xml") } |
            Sort-Object `
                @{ Expression = { if ($_.FullName -match '/slide(\d+)\.xml$') { [int]$Matches[1] } else { [int]::MaxValue } } }, `
                @{ Expression = { $_.FullName } } |
            ForEach-Object { $_.FullName }
    )
}

function Get-Extension {
    param([string]$Part)
    return [System.IO.Path]::GetExtension($Part).ToLowerInvariant()
}

function Get-TagFor {
    param(
        [string]$Kind,
        [string]$Part
    )

    $ext = (Get-Extension $Part).TrimStart(".").ToUpperInvariant()
    if ($ext -eq "JPEG") { return "JPG" }
    if ($Kind -eq "chart") { return "CHART" }
    if ($Kind -eq "chart_style") { return "CHARTSTYLE" }
    if ($Kind -eq "chart_colors") { return "CHARTCOLORS" }
    if ($Kind -eq "diagram") { return "DIAGRAM" }
    if ($Kind -eq "ole") { return "OLE" }
    if ($Kind -eq "unknown") {
        if ($ext) { return $ext }
        return "FILE"
    }
    if ($ext) { return $ext }
    return $Kind.ToUpperInvariant()
}

function Get-OutputSuffix {
    param(
        [string]$Part,
        [string]$Tag
    )

    $ext = Get-Extension $Part
    if ($ext -eq ".jpeg") { return ".jpg" }
    if ($ext) { return $ext }
    return ".$($Tag.ToLowerInvariant())"
}

function Get-PartKind {
    param(
        [string]$RelType,
        [string]$Part
    )

    $lowerRelType = $RelType.ToLowerInvariant()
    $lowerPart = $Part.ToLowerInvariant()
    $ext = Get-Extension $Part

    foreach ($word in $RelSkipWords) {
        if ($lowerRelType.Contains($word.ToLowerInvariant())) {
            return $null
        }
    }

    if ($ImageExts -contains $ext -or $lowerPart.Contains("/media/image") -or $lowerRelType.EndsWith("/image")) { return "image" }
    if ($VideoExts -contains $ext -or $lowerRelType.Contains("video")) { return "video" }
    if ($AudioExts -contains $ext -or $lowerRelType.Contains("audio")) { return "audio" }
    if ($lowerPart.Contains("/embeddings/") -or $lowerRelType.EndsWith("/package")) {
        if ($ext -eq ".bin") { return "ole" }
        return "embedded"
    }
    if ($lowerPart.Contains("/charts/style")) { return "chart_style" }
    if ($lowerPart.Contains("/charts/colors")) { return "chart_colors" }
    if ($lowerPart.Contains("/charts/") -or $lowerRelType.EndsWith("/chart")) { return "chart" }
    if ($lowerPart.Contains("/diagrams/")) { return "diagram" }
    return $null
}

function Should-ExtractKind {
    param(
        [string]$Kind,
        [bool]$MediaOnly
    )

    if ($MediaOnly) {
        return @("image", "video", "audio") -contains $Kind
    }
    return @("audio", "chart", "chart_colors", "chart_style", "diagram", "embedded", "image", "ole", "video") -contains $Kind
}

function Collect-SlideResources {
    param(
        [object]$Zip,
        [string]$SlidePart,
        [int]$SlideNumber,
        [bool]$MediaOnly
    )

    $resources = New-Object System.Collections.Generic.List[object]
    $seenTargets = New-Object System.Collections.Generic.HashSet[string]

    function Walk-Relationships {
        param(
            [string]$SourcePart,
            [int]$Depth
        )

        if ($Depth -gt 2) {
            return
        }

        foreach ($rel in (Get-Relationships $Zip $SourcePart)) {
            if ([string]::IsNullOrEmpty($rel.Target) -or $rel.TargetMode.ToLowerInvariant() -eq "external") {
                continue
            }

            $targetPart = Resolve-TargetPart $SourcePart $rel.Target
            if ($null -eq (Get-ZipEntry $Zip $targetPart)) {
                continue
            }

            $kind = Get-PartKind $rel.Type $targetPart
            if ($kind -and (Should-ExtractKind $kind $MediaOnly) -and -not $seenTargets.Contains($targetPart)) {
                [void]$seenTargets.Add($targetPart)
                $resources.Add([PSCustomObject]@{
                    SlideNumber = $SlideNumber
                    Kind = $kind
                    Tag = Get-TagFor $kind $targetPart
                    SourcePart = $SourcePart
                    TargetPart = $targetPart
                    RelId = $rel.Id
                    RelType = $rel.Type
                }) | Out-Null
            }

            if (($kind -eq "chart" -or $kind -eq "diagram") -and -not $MediaOnly) {
                Walk-Relationships $targetPart ($Depth + 1)
            }
        }
    }

    Walk-Relationships $SlidePart 0
    return $resources.ToArray()
}

function Get-SlideText {
    param(
        [object]$Zip,
        [string]$SlidePart
    )

    $xml = Read-XmlPart $Zip $SlidePart
    if ($null -eq $xml) {
        return ""
    }

    $runs = New-Object System.Collections.Generic.List[string]
    foreach ($node in $xml.GetElementsByTagName("t", $DrawingNs)) {
        if ($node.InnerText) {
            $text = $node.InnerText.Trim()
            if ($text) {
                $runs.Add($text) | Out-Null
            }
        }
    }
    return ($runs -join "`n")
}

function Get-OutputDirForKind {
    param(
        [string]$OutputDir,
        [string]$Kind
    )

    $folder = "其他"
    if ($KindFolderNames.ContainsKey($Kind)) {
        $folder = $KindFolderNames[$Kind]
    }
    return (Join-Path $OutputDir $folder)
}

function Get-UniqueOutputPath {
    param(
        [string]$OutputDir,
        [int]$SlideNumber,
        [string]$Tag,
        [string]$Suffix,
        [hashtable]$Counters,
        [bool]$Overwrite
    )

    $key = "$OutputDir|$SlideNumber|$Tag"
    if (-not $Counters.ContainsKey($key)) {
        $Counters[$key] = 0
    }
    $Counters[$key] = [int]$Counters[$key] + 1
    $index = [int]$Counters[$key]

    if ($index -eq 1) {
        $stem = "{0:D3}_{1}" -f $SlideNumber, $Tag
    }
    else {
        $stem = "{0:D3}_{1}_{2:D2}" -f $SlideNumber, $Tag, $index
    }
    $candidate = Join-Path $OutputDir "$stem$Suffix"

    if ($Overwrite) {
        return $candidate
    }

    if (Test-Path -LiteralPath $candidate) {
        return $null
    }

    return $candidate
}

function Copy-ZipEntry {
    param(
        [object]$Zip,
        [string]$Part,
        [string]$Destination,
        [bool]$Overwrite
    )

    if ((Test-Path -LiteralPath $Destination) -and -not $Overwrite) {
        throw "输出文件已存在：$Destination"
    }

    $entry = Get-ZipEntry $Zip $Part
    if ($null -eq $entry) {
        throw "PPTX 内部资源不存在：$Part"
    }

    $parent = Split-Path -Parent $Destination
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    $source = $entry.Open()
    try {
        $target = [System.IO.File]::Open($Destination, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
        try {
            $source.CopyTo($target)
        }
        finally {
            $target.Dispose()
        }
    }
    finally {
        $source.Dispose()
    }
}

function Get-RelativePathText {
    param(
        [string]$Path,
        [string]$BaseDir
    )

    $baseUri = New-Object System.Uri(($BaseDir.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar))
    $pathUri = New-Object System.Uri($Path)
    return [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($pathUri).ToString())
}

function Escape-CsvField {
    param([string]$Value)

    if ($null -eq $Value) {
        $Value = ""
    }
    $Value = [string]$Value
    if ($Value.Contains('"') -or $Value.Contains(",") -or $Value.Contains("`r") -or $Value.Contains("`n")) {
        return '"' + $Value.Replace('"', '""') + '"'
    }
    return $Value
}

function Write-Manifest {
    param(
        [string]$Path,
        [object[]]$Rows
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $headers = @("slide", "output_file", "kind", "source_part", "target_part", "relationship_id", "relationship_type")
    $lines.Add(($headers -join ",")) | Out-Null

    foreach ($row in $Rows) {
        $fields = @(
            $row.Slide,
            $row.OutputFile,
            $row.Kind,
            $row.SourcePart,
            $row.TargetPart,
            $row.RelationshipId,
            $row.RelationshipType
        ) | ForEach-Object { Escape-CsvField $_ }
        $lines.Add(($fields -join ",")) | Out-Null
    }

    $encoding = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllLines($Path, $lines.ToArray(), $encoding)
}

function Get-OutputDirForPptx {
    param(
        [string]$PptxPath,
        [string]$OutputArg,
        [bool]$MultiInput
    )

    if ([string]::IsNullOrWhiteSpace($OutputArg)) {
        $baseOutput = Join-Path ([System.IO.Path]::GetDirectoryName($PptxPath)) $DefaultOutputDirName
    }
    else {
        $baseOutput = Get-FullPath $OutputArg
    }

    if ($MultiInput) {
        return (Join-Path $baseOutput ([System.IO.Path]::GetFileNameWithoutExtension($PptxPath)))
    }
    return $baseOutput
}

function Extract-Pptx {
    param(
        [string]$PptxPath,
        [string]$OutputDir,
        [bool]$MediaOnly,
        [bool]$WithText,
        [bool]$Overwrite
    )

    if (-not (Test-Path -LiteralPath $PptxPath -PathType Leaf)) {
        throw "文件不存在：$PptxPath"
    }

    if (-not (Test-Path -LiteralPath $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    $zip = [System.IO.Compression.ZipFile]::OpenRead($PptxPath)
    try {
        $slideParts = @(Get-SlidePartsInOrder $zip)
        $counters = @{}
        $manifestRows = New-Object System.Collections.Generic.List[object]
        $extractedCount = 0

        for ($index = 0; $index -lt $slideParts.Count; $index++) {
            $slideNumber = $index + 1
            $slidePart = $slideParts[$index]
            $resources = @(Collect-SlideResources $zip $slidePart $slideNumber $MediaOnly)

            foreach ($resource in $resources) {
                $suffix = Get-OutputSuffix $resource.TargetPart $resource.Tag
                $resourceOutputDir = Get-OutputDirForKind $OutputDir $resource.Kind
                $destination = Get-UniqueOutputPath $resourceOutputDir $resource.SlideNumber $resource.Tag $suffix $counters $Overwrite
                if ($null -eq $destination) {
                    continue
                }
                Copy-ZipEntry $zip $resource.TargetPart $destination $Overwrite
                $extractedCount++

                $manifestRows.Add([PSCustomObject]@{
                    Slide = "{0:D3}" -f $resource.SlideNumber
                    OutputFile = Get-RelativePathText $destination $OutputDir
                    Kind = $resource.Kind
                    SourcePart = $resource.SourcePart
                    TargetPart = $resource.TargetPart
                    RelationshipId = $resource.RelId
                    RelationshipType = $resource.RelType
                }) | Out-Null
            }

            if ($WithText) {
                $slideText = Get-SlideText $zip $slidePart
                if ($slideText) {
                    $textOutputDir = Get-OutputDirForKind $OutputDir "text"
                    $textPath = Get-UniqueOutputPath $textOutputDir $slideNumber "TXT" ".txt" $counters $Overwrite
                    if ($null -ne $textPath) {
                        $parent = Split-Path -Parent $textPath
                        if (-not (Test-Path -LiteralPath $parent)) {
                            New-Item -ItemType Directory -Path $parent -Force | Out-Null
                        }
                        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
                        [System.IO.File]::WriteAllText($textPath, ($slideText + "`n"), $utf8NoBom)
                        $extractedCount++

                        $manifestRows.Add([PSCustomObject]@{
                            Slide = "{0:D3}" -f $slideNumber
                            OutputFile = Get-RelativePathText $textPath $OutputDir
                            Kind = "text"
                            SourcePart = $slidePart
                            TargetPart = $slidePart
                            RelationshipId = ""
                            RelationshipType = ""
                        }) | Out-Null
                    }
                }
            }
        }

        Write-Manifest (Join-Path $OutputDir "manifest.csv") $manifestRows.ToArray()
        return [PSCustomObject]@{
            SlideCount = $slideParts.Count
            ExtractedCount = $extractedCount
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Main {
    param([string[]]$Argv)

    if ($Argv.Count -eq 0) {
        Show-Banner
        $inputText = Read-Host "请输入 .pptx 文件路径（可拖拽）"
        if ([string]::IsNullOrWhiteSpace($inputText)) {
            Write-Host "未输入文件路径。"
            return 0
        }
        $Argv = @(Split-CommandLine $inputText)

    }

    $options = Parse-Arguments $Argv
    $inputFiles = @(Find-InputFiles $options)
    if ($inputFiles.Count -eq 0) {
        Write-Error "未找到 .pptx 文件。请拖入/传入 PPTX 文件，或把本工具放到 PPTX 所在文件夹运行。"
        return 1
    }

    $multiInput = $inputFiles.Count -gt 1
    $outputDirs = New-Object System.Collections.Generic.List[string]
    $successCount = 0
    $failureCount = 0

    foreach ($pptxPath in $inputFiles) {
        if (Is-TempPptx $pptxPath) {
            Write-Host "跳过 PowerPoint 临时锁文件：$([System.IO.Path]::GetFileName($pptxPath))"
            continue
        }

        $destinationDir = Get-OutputDirForPptx $pptxPath $options.Output $multiInput
        try {
            $result = Extract-Pptx $pptxPath $destinationDir $options.MediaOnly ($options.WithText -and -not $options.MediaOnly) $options.Overwrite
            $successCount++
            $outputDirs.Add($destinationDir) | Out-Null
            Write-Host ("{0}: {1} slides, {2} files -> {3}" -f [System.IO.Path]::GetFileName($pptxPath), $result.SlideCount, $result.ExtractedCount, $destinationDir)
        }
        catch {
            $failureCount++
            Write-Error ("提取失败：{0} -> {1}" -f $pptxPath, $_.Exception.Message)
        }
    }

    if ($outputDirs.Count -gt 0) {
        $uniqueOutputDirs = @($outputDirs | Select-Object -Unique)
        if ($uniqueOutputDirs.Count -eq 1) {
            Write-Host "提醒：输出文件夹在这里：$($uniqueOutputDirs[0])"
        }
        else {
            Write-Host "提醒：输出文件夹在这里："
            foreach ($outputDir in $uniqueOutputDirs) {
                Write-Host "  - $outputDir"
            }
        }
    }

    if ($successCount -eq 0) {
        return 1
    }
    if ($failureCount -gt 0) {
        return 1
    }
    return 0
}

exit (Main -Argv $LauncherArgs)

# PYTHON_CODE_BELOW
#!/usr/bin/env python3
"""
Extract slide-level resources from PowerPoint .pptx files.

Examples:
  python3 extract_pptx_elements.py "deck.pptx"
  python3 extract_pptx_elements.py "deck.pptx" -o exported_assets
  python3 extract_pptx_elements.py --no-text

Output naming:
  Slide 1 JPG: 图片/001_JPG.jpg
  Slide 1 MP4: 视频/001_MP4.mp4
  Second JPG on slide 1: 图片/001_JPG_02.jpg
"""

from __future__ import annotations

import argparse
import csv
import posixpath
import re
import shutil
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Iterable
from xml.etree import ElementTree as ET


PACKAGE_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
OFFICE_RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

VERSION = "1.1.2"
DEFAULT_OUTPUT_DIR_NAME = "pptx_extracted_elements"

IMAGE_EXTS = {
    ".bmp",
    ".emf",
    ".gif",
    ".jfif",
    ".jpeg",
    ".jpg",
    ".png",
    ".svg",
    ".tif",
    ".tiff",
    ".webp",
    ".wmf",
}
VIDEO_EXTS = {
    ".3gp",
    ".asf",
    ".avi",
    ".m4v",
    ".mkv",
    ".mov",
    ".mp4",
    ".mpeg",
    ".mpg",
    ".swf",
    ".webm",
    ".wmv",
}
AUDIO_EXTS = {
    ".aac",
    ".aif",
    ".aiff",
    ".m4a",
    ".mid",
    ".midi",
    ".mp3",
    ".oga",
    ".ogg",
    ".wav",
    ".wma",
}
EMBED_EXTS = {
    ".bin",
    ".csv",
    ".doc",
    ".docx",
    ".html",
    ".pdf",
    ".ppt",
    ".pptx",
    ".rtf",
    ".txt",
    ".xls",
    ".xlsb",
    ".xlsm",
    ".xlsx",
    ".xml",
    ".zip",
}

REL_SKIP_WORDS = (
    "/hyperlink",
    "/notesSlide",
    "/presProps",
    "/printerSettings",
    "/slideLayout",
    "/slideMaster",
    "/theme",
    "/viewProps",
)

KIND_FOLDER_NAMES = {
    "audio": "音频",
    "chart": "图表",
    "chart_colors": "图表",
    "chart_style": "图表",
    "diagram": "图示",
    "embedded": "嵌入文件",
    "image": "图片",
    "ole": "嵌入文件",
    "text": "文本",
    "video": "视频",
}


@dataclass(frozen=True)
class Relationship:
    rel_id: str
    rel_type: str
    target: str
    target_mode: str


@dataclass(frozen=True)
class Resource:
    slide_number: int
    kind: str
    tag: str
    source_part: str
    target_part: str
    rel_id: str
    rel_type: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Extract images, videos, audio, embedded files, charts, and diagrams "
            "from .pptx files into Chinese type folders using slide-number-based names."
        )
    )
    parser.add_argument(
        "pptx",
        nargs="*",
        type=Path,
        help=(
            "PPTX file(s) to extract. If omitted, all non-temporary .pptx files "
            "in the current directory are processed."
        ),
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help=(
            "Output directory. Default: create pptx_extracted_elements next to "
            "the PPTX file being processed."
        ),
    )
    text_group = parser.add_mutually_exclusive_group()
    text_group.add_argument(
        "--with-text",
        dest="with_text",
        action="store_true",
        default=True,
        help="Export plain slide text as 001_TXT.txt, 002_TXT.txt, etc. Enabled by default.",
    )
    text_group.add_argument(
        "--no-text",
        dest="with_text",
        action="store_false",
        help="Do not export plain slide text.",
    )
    parser.add_argument(
        "--media-only",
        action="store_true",
        help="Only extract images, videos, and audio.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite files in the output directory if they already exist.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {VERSION}",
    )
    return parser.parse_args()


def is_temp_pptx(path: Path) -> bool:
    return path.name.startswith("~$") or path.name.startswith(".~")


def find_input_files(args: argparse.Namespace) -> list[Path]:
    if args.pptx:
        return [path.expanduser().resolve() for path in args.pptx]

    return sorted(
        path.resolve()
        for path in Path.cwd().glob("*.pptx")
        if path.is_file() and not is_temp_pptx(path)
    )


def read_xml(zip_file: zipfile.ZipFile, part: str) -> ET.Element | None:
    try:
        return ET.fromstring(zip_file.read(part))
    except KeyError:
        return None
    except ET.ParseError as exc:
        print(f"Warning: could not parse XML part {part}: {exc}", file=sys.stderr)
        return None


def rels_part_name(part: str) -> str:
    directory = posixpath.dirname(part)
    filename = posixpath.basename(part)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def parse_rels(zip_file: zipfile.ZipFile, part: str) -> dict[str, Relationship]:
    root = read_xml(zip_file, rels_part_name(part))
    if root is None:
        return {}

    rels: dict[str, Relationship] = {}
    for item in root.findall(f"{{{PACKAGE_RELS_NS}}}Relationship"):
        rel_id = item.get("Id", "")
        if not rel_id:
            continue

        rels[rel_id] = Relationship(
            rel_id=rel_id,
            rel_type=item.get("Type", ""),
            target=item.get("Target", ""),
            target_mode=item.get("TargetMode", ""),
        )
    return rels


def resolve_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")

    source_dir = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(source_dir, target))


def natural_slide_sort_key(part: str) -> tuple[int, str]:
    match = re.search(r"/slide(\d+)\.xml$", part)
    if match:
        return (int(match.group(1)), part)
    return (10**9, part)


def slide_parts_in_order(zip_file: zipfile.ZipFile) -> list[str]:
    presentation = read_xml(zip_file, "ppt/presentation.xml")
    presentation_rels = parse_rels(zip_file, "ppt/presentation.xml")
    names = set(zip_file.namelist())

    if presentation is not None and presentation_rels:
        ordered_slides: list[str] = []
        for slide_id in presentation.findall(
            f".//{{{PRESENTATION_NS}}}sldIdLst/{{{PRESENTATION_NS}}}sldId"
        ):
            rel_id = slide_id.get(f"{{{OFFICE_RELS_NS}}}id")
            if not rel_id or rel_id not in presentation_rels:
                continue
            target = resolve_target(
                "ppt/presentation.xml", presentation_rels[rel_id].target
            )
            if target in names:
                ordered_slides.append(target)

        if ordered_slides:
            return ordered_slides

    return sorted(
        (
            name
            for name in names
            if name.startswith("ppt/slides/") and name.endswith(".xml")
        ),
        key=natural_slide_sort_key,
    )


def extension_for(part: str) -> str:
    return PurePosixPath(part).suffix.lower()


def tag_for(kind: str, part: str) -> str:
    ext = extension_for(part).lstrip(".").upper()
    if ext == "JPEG":
        return "JPG"
    if kind == "chart":
        return "CHART"
    if kind == "chart_style":
        return "CHARTSTYLE"
    if kind == "chart_colors":
        return "CHARTCOLORS"
    if kind == "diagram":
        return "DIAGRAM"
    if kind == "ole":
        return "OLE"
    if kind == "unknown":
        return ext or "FILE"
    return ext or kind.upper()


def output_suffix_for(part: str, tag: str) -> str:
    ext = extension_for(part)
    if ext == ".jpeg":
        return ".jpg"
    if ext:
        return ext
    return f".{tag.lower()}"


def classify_part(rel_type: str, part: str) -> str | None:
    lower_rel_type = rel_type.lower()
    lower_part = part.lower()
    ext = extension_for(part)

    if any(word.lower() in lower_rel_type for word in REL_SKIP_WORDS):
        return None

    if ext in IMAGE_EXTS or "/media/image" in lower_part or lower_rel_type.endswith("/image"):
        return "image"
    if ext in VIDEO_EXTS or "video" in lower_rel_type:
        return "video"
    if ext in AUDIO_EXTS or "audio" in lower_rel_type:
        return "audio"
    if "/embeddings/" in lower_part or lower_rel_type.endswith("/package"):
        return "ole" if ext == ".bin" else "embedded"
    if "/charts/style" in lower_part:
        return "chart_style"
    if "/charts/colors" in lower_part:
        return "chart_colors"
    if "/charts/" in lower_part or lower_rel_type.endswith("/chart"):
        return "chart"
    if "/diagrams/" in lower_part:
        return "diagram"

    return None


def should_extract(kind: str, media_only: bool) -> bool:
    if media_only:
        return kind in {"image", "video", "audio"}
    return kind in {
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


def collect_slide_resources(
    zip_file: zipfile.ZipFile,
    slide_part: str,
    slide_number: int,
    *,
    media_only: bool,
) -> list[Resource]:
    names = set(zip_file.namelist())
    resources: list[Resource] = []
    seen_targets: set[str] = set()

    def walk_relationships(source_part: str, depth: int) -> None:
        if depth > 2:
            return

        for rel in parse_rels(zip_file, source_part).values():
            if not rel.target or rel.target_mode.lower() == "external":
                continue

            target_part = resolve_target(source_part, rel.target)
            if target_part not in names:
                continue

            kind = classify_part(rel.rel_type, target_part)
            if kind and should_extract(kind, media_only) and target_part not in seen_targets:
                seen_targets.add(target_part)
                resources.append(
                    Resource(
                        slide_number=slide_number,
                        kind=kind,
                        tag=tag_for(kind, target_part),
                        source_part=source_part,
                        target_part=target_part,
                        rel_id=rel.rel_id,
                        rel_type=rel.rel_type,
                    )
                )

            if kind in {"chart", "diagram"} and not media_only:
                walk_relationships(target_part, depth + 1)

    walk_relationships(slide_part, 0)
    return resources


def extract_slide_text(zip_file: zipfile.ZipFile, slide_part: str) -> str:
    root = read_xml(zip_file, slide_part)
    if root is None:
        return ""

    text_runs: list[str] = []
    for item in root.iter(f"{{{DRAWING_NS}}}t"):
        if item.text:
            text = item.text.strip()
            if text:
                text_runs.append(text)

    return "\n".join(text_runs)


def unique_output_path(
    output_dir: Path,
    slide_number: int,
    tag: str,
    suffix: str,
    counters: dict[tuple[Path, int, str], int],
    *,
    overwrite: bool,
) -> Path | None:
    key = (output_dir, slide_number, tag)
    counters[key] = counters.get(key, 0) + 1
    index = counters[key]
    stem = f"{slide_number:03d}_{tag}" if index == 1 else f"{slide_number:03d}_{tag}_{index:02d}"
    candidate = output_dir / f"{stem}{suffix}"

    if overwrite:
        return candidate

    if candidate.exists():
        return None

    return candidate


def extract_file(
    zip_file: zipfile.ZipFile,
    member: str,
    destination: Path,
    *,
    overwrite: bool,
) -> None:
    if destination.exists() and not overwrite:
        raise FileExistsError(destination)

    destination.parent.mkdir(parents=True, exist_ok=True)
    with zip_file.open(member) as source, destination.open("wb") as target:
        shutil.copyfileobj(source, target)


def output_dir_for(
    pptx_path: Path,
    output_arg: Path | None,
    multi_input: bool,
) -> Path:
    if output_arg is None:
        base_output_dir = pptx_path.parent / DEFAULT_OUTPUT_DIR_NAME
    else:
        base_output_dir = output_arg.expanduser()
        if not base_output_dir.is_absolute():
            base_output_dir = Path.cwd() / base_output_dir

    if multi_input:
        return base_output_dir / pptx_path.stem
    return base_output_dir


def kind_output_dir(output_dir: Path, kind: str) -> Path:
    return output_dir / KIND_FOLDER_NAMES.get(kind, "其他")


def relative_output_file(path: Path, output_dir: Path) -> str:
    return path.relative_to(output_dir).as_posix()


def write_manifest(manifest_path: Path, rows: list[dict[str, str]]) -> None:
    fieldnames = [
        "slide",
        "output_file",
        "kind",
        "source_part",
        "target_part",
        "relationship_id",
        "relationship_type",
    ]
    with manifest_path.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def extract_pptx(
    pptx_path: Path,
    output_dir: Path,
    *,
    media_only: bool,
    with_text: bool,
    overwrite: bool,
) -> tuple[int, int]:
    if not pptx_path.exists():
        raise FileNotFoundError(pptx_path)
    if not zipfile.is_zipfile(pptx_path):
        raise ValueError(f"不是有效的 .pptx/zip 文件：{pptx_path}")

    output_dir.mkdir(parents=True, exist_ok=True)
    counters: dict[tuple[Path, int, str], int] = {}
    manifest_rows: list[dict[str, str]] = []
    extracted_count = 0

    with zipfile.ZipFile(pptx_path) as pptx_zip:
        slide_parts = slide_parts_in_order(pptx_zip)

        for slide_number, slide_part in enumerate(slide_parts, start=1):
            resources = collect_slide_resources(
                pptx_zip,
                slide_part,
                slide_number,
                media_only=media_only,
            )

            for resource in resources:
                suffix = output_suffix_for(resource.target_part, resource.tag)
                resource_output_dir = kind_output_dir(output_dir, resource.kind)
                destination = unique_output_path(
                    resource_output_dir,
                    resource.slide_number,
                    resource.tag,
                    suffix,
                    counters,
                    overwrite=overwrite,
                )
                if destination is None:
                    continue
                extract_file(
                    pptx_zip,
                    resource.target_part,
                    destination,
                    overwrite=overwrite,
                )
                extracted_count += 1
                manifest_rows.append(
                    {
                        "slide": f"{resource.slide_number:03d}",
                        "output_file": relative_output_file(destination, output_dir),
                        "kind": resource.kind,
                        "source_part": resource.source_part,
                        "target_part": resource.target_part,
                        "relationship_id": resource.rel_id,
                        "relationship_type": resource.rel_type,
                    }
                )

            if with_text:
                slide_text = extract_slide_text(pptx_zip, slide_part)
                if slide_text:
                    text_path = unique_output_path(
                        kind_output_dir(output_dir, "text"),
                        slide_number,
                        "TXT",
                        ".txt",
                        counters,
                        overwrite=overwrite,
                    )
                    if text_path is not None:
                        text_path.parent.mkdir(parents=True, exist_ok=True)
                        text_path.write_text(slide_text + "\n", encoding="utf-8")
                        extracted_count += 1
                        manifest_rows.append(
                            {
                                "slide": f"{slide_number:03d}",
                                "output_file": relative_output_file(text_path, output_dir),
                                "kind": "text",
                                "source_part": slide_part,
                                "target_part": slide_part,
                                "relationship_id": "",
                                "relationship_type": "",
                            }
                        )

    write_manifest(output_dir / "manifest.csv", manifest_rows)
    return len(slide_parts), extracted_count


def main() -> int:
    args = parse_args()
    input_files = find_input_files(args)

    if not input_files:
        print(
            "未找到 .pptx 文件。请拖入/传入 PPTX 文件，或把本工具放到 PPTX 所在文件夹运行。",
            file=sys.stderr,
        )
        return 1

    multi_input = len(input_files) > 1
    output_dirs: list[Path] = []
    success_count = 0
    failure_count = 0

    for pptx_path in input_files:
        if is_temp_pptx(pptx_path):
            print(f"跳过 PowerPoint 临时锁文件：{pptx_path.name}")
            continue

        destination_dir = output_dir_for(pptx_path, args.output, multi_input).resolve()
        try:
            slide_count, extracted_count = extract_pptx(
                pptx_path,
                destination_dir,
                media_only=args.media_only,
                with_text=args.with_text and not args.media_only,
                overwrite=args.overwrite,
            )
        except (OSError, ValueError, zipfile.BadZipFile) as exc:
            failure_count += 1
            print(f"提取失败：{pptx_path} -> {exc}", file=sys.stderr)
            continue

        success_count += 1
        output_dirs.append(destination_dir)
        print(
            f"{pptx_path.name}: {slide_count} slides, "
            f"{extracted_count} files -> {destination_dir}"
        )

    if output_dirs:
        unique_output_dirs = list(dict.fromkeys(output_dirs))
        if len(unique_output_dirs) == 1:
            print(f"提醒：输出文件夹在这里：{unique_output_dirs[0]}")
        else:
            print("提醒：输出文件夹在这里：")
            for output_dir in unique_output_dirs:
                print(f"  - {output_dir}")

    if success_count == 0:
        return 1

    return 1 if failure_count else 0


if __name__ == "__main__":
    raise SystemExit(main())
