@echo off
setlocal EnableExtensions
set "VERSION=V1.1.2"
set "PPTX_EXTRACTOR_SCRIPT=%~f0"
chcp 65001 >nul
REM ============================================================
REM  PPTX 元素提取工具 %VERSION% - Windows PowerShell 启动器
REM  直接双击或拖拽 .pptx 到本文件即可运行
REM ============================================================

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$script=[IO.File]::ReadAllText($env:PPTX_EXTRACTOR_SCRIPT,[Text.Encoding]::UTF8); $parts=[regex]::Split($script,'(?m)^# POWERSHELL_START\r?$',2); if($parts.Count -lt 2){throw 'PowerShell section missing'}; $block=[scriptblock]::Create($parts[1]); & $block @args" %*
set "EXIT_CODE=%ERRORLEVEL%"
echo.
echo 完成！(退出码: %EXIT_CODE%)
echo 提醒：默认输出文件夹会出现在 PPTX 文件旁边，名称为 pptx_extracted_elements。
echo 按任意键关闭此窗口 ...
pause >nul
exit /b %EXIT_CODE%

# POWERSHELL_START
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
    Write-Host "  --with-text   同时提取幻灯片文本"
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
    Write-Host "  extract_pptx_elements.cmd presentation.pptx --with-text"
    Write-Host "  extract_pptx_elements.cmd presentation.pptx --media-only"
    Write-Host "  extract_pptx_elements.cmd presentation.pptx --overwrite"
}

function Parse-Arguments {
    param([string[]]$Argv)

    $pptxList = New-Object System.Collections.Generic.List[string]
    $options = [ordered]@{
        Pptx = $pptxList
        Output = $null
        WithText = $false
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
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$Part
    )
    return $Zip.GetEntry($Part)
}

function Read-ZipText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
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
        [System.IO.Compression.ZipArchive]$Zip,
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
        [System.IO.Compression.ZipArchive]$Zip,
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
    param([System.IO.Compression.ZipArchive]$Zip)

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
        [System.IO.Compression.ZipArchive]$Zip,
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
        [System.IO.Compression.ZipArchive]$Zip,
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

    $collisionIndex = $index
    while (Test-Path -LiteralPath $candidate) {
        $collisionIndex++
        $candidate = Join-Path $OutputDir ("{0:D3}_{1}_{2:D2}{3}" -f $SlideNumber, $Tag, $collisionIndex, $Suffix)
    }

    return $candidate
}

function Copy-ZipEntry {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
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
            $result = Extract-Pptx $pptxPath $destinationDir $options.MediaOnly $options.WithText $options.Overwrite
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

exit (Main -Argv $args)
