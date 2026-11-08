@echo off
setlocal EnableExtensions
chcp 65001 >nul
title Gifyu Album Downloader

REM ============================================================================
REM  Gifyu Album Downloader  -  powered by a global (pip) gallery-dl install
REM
REM  What it does:
REM    1. Asks you to paste a Gifyu.com album URL (e.g. https://gifyu.com/album/yKM)
REM    2. Auto-cleans messy links (strips ?sort=...&page=... filters and junk)
REM    3. Paginates the album and reads each image's embedded metadata
REM    4. Keeps ONLY original / full-resolution images whose real format is
REM       jpeg / jpg / png / tiff / webp  AND  are at least 300 x 300 pixels
REM    5. Downloads them into a NEW sub-folder of your Downloads folder
REM       (created next to wherever this .bat is run) using gallery-dl
REM
REM  Requirements:
REM    - gallery-dl installed globally (pip install gallery-dl)
REM    - Windows PowerShell (built into Windows)
REM
REM  This file is self-contained: the batch part below launches an embedded
REM  PowerShell block that does the heavy lifting. No other files needed.
REM ============================================================================

REM --- Launch the embedded PowerShell helper -----------------------------------
REM  The batch parser jumps straight to the ":POWERSHELL" label and PowerShell
REM  reads THIS SAME FILE, ignoring everything above the marker, then runs the
REM  script body. When PowerShell finishes, control returns to batch below.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $marker='#___PS_BODY_START___'; $lines=[IO.File]::ReadAllLines('%~f0'); $i=[Array]::IndexOf($lines,$marker); if($i -lt 0){Write-Host 'Internal error: PowerShell body marker not found.'; exit 1}; $body=($lines[($i+1)..($lines.Length-1)]) -join [Environment]::NewLine; Invoke-Expression $body"
set "EXITCODE=%ERRORLEVEL%"
echo.
if not "%EXITCODE%"=="0" (
  echo [Finished with errors - exit code %EXITCODE%]
) else (
  echo [Done]
)
echo.
pause
endlocal
exit /b %EXITCODE%

# The line below marks where the PowerShell body begins. Do not edit/remove it.
#___PS_BODY_START___


# ============================================================================
#  PowerShell body (executed from inside the batch file above)
# ============================================================================
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'   # speed up Invoke-WebRequest
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) gifyu-album-downloader'

# --- Settings ---------------------------------------------------------------
$MinWidth       = 300
$MinHeight      = 300
$AllowedExt     = @('jpg', 'jpeg', 'png', 'tiff', 'tif', 'webp')
$PageStep       = 20        # gifyu returns 20 items per page by default
$MaxPages       = 500       # hard safety cap so we can never loop forever

# --- Console helpers --------------------------------------------------------
function Write-Info  { param([string]$m) Write-Host $m -ForegroundColor Cyan }
function Write-Ok    { param([string]$m) Write-Host $m -ForegroundColor Green }
function Write-Warn  { param([string]$m) Write-Host $m -ForegroundColor Yellow }
function Write-Err   { param([string]$m) Write-Host $m -ForegroundColor Red }
function Write-Dim   { param([string]$m) Write-Host $m -ForegroundColor DarkGray }

# Build a filesystem-safe name (keeps letters/digits/spaces/dash/underscore).
function Make-SafeName {
    param([string]$Name, [int]$MaxLen = 120)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $s = $Name.Trim()
    # Replace any character that is illegal in Windows file names.
    $s = [regex]::Replace($s, '[<>:"/\\|?*]', ' ')
    # Collapse repeated whitespace.
    $s = [regex]::Replace($s, '\s+', ' ').Trim()
    # Drop trailing dots/spaces (illegal on Windows).
    $s = $s.TrimEnd('.', ' ')
    if ($s.Length -gt $MaxLen) { $s = $s.Substring(0, $MaxLen).TrimEnd('.', ' ') }
    return $s
}

# ============================================================================
#  STEP 1 - Prompt for the album URL
# ============================================================================
Write-Host ''
Write-Host '============================================================' -ForegroundColor Magenta
Write-Host '             GIFYU ALBUM DOWNLOADER (gallery-dl)            ' -ForegroundColor Magenta
Write-Host '============================================================' -ForegroundColor Magenta
Write-Dim   'Downloads ORIGINAL images only (jpeg/jpg/png/tiff/webp, >=300x300).'
Write-Host ''

$RawInput = ''
while ([string]::IsNullOrWhiteSpace($RawInput)) {
    Write-Host 'Paste a Gifyu.com album URL and press Enter:' -ForegroundColor White
    Write-Host '   e.g.  https://gifyu.com/album/yKM' -ForegroundColor DarkGray
    Write-Host -NoNewline 'URL> ' -ForegroundColor White
    $RawInput = Read-Host
    if ([string]::IsNullOrWhiteSpace($RawInput)) {
        Write-Warn 'Nothing entered. Please paste an album URL.'
    }
}
$RawInput = $RawInput.Trim().Trim('"').Trim("'")

# ============================================================================
#  STEP 2 - Clean / normalise the URL
#     https://gifyu.com/album/yKM/?sort=title_asc&page=1
#  becomes
#     https://gifyu.com/album/yKM
# ============================================================================
function Get-CleanAlbumUrl {
    param([string]$Url)

    # Pull out the album id from any common gifyu album URL shape, e.g.
    #   https://gifyu.com/album/yKM[/...][?query]
    #   http://www.gifyu.com/album/yKM
    #   gifyu.com/album/yKM/?sort=title_asc&page=1
    $m = [regex]::Match($Url, 'gifyu\.com/album/([A-Za-z0-9]+)', 'IgnoreCase')
    if ($m.Success) {
        return 'https://gifyu.com/album/' + $m.Groups[1].Value
    }
    return $null
}


$AlbumUrl = Get-CleanAlbumUrl -Url $RawInput
if (-not $AlbumUrl) {
    Write-Err 'That does not look like a Gifyu album URL.'
    Write-Err 'Expected something like: https://gifyu.com/album/yKM'
    exit 1
}
$AlbumId = ($AlbumUrl -split '/')[-1]
if ($RawInput -ne $AlbumUrl) {
    Write-Ok   ('Cleaned URL  ->  ' + $AlbumUrl)
} else {
    Write-Ok   ('Album URL    ->  ' + $AlbumUrl)
}
Write-Host ''

# ============================================================================
#  STEP 3 - Make sure gallery-dl is installed globally (pip)
# ============================================================================
$gd = Get-Command gallery-dl -ErrorAction SilentlyContinue
if (-not $gd) {
    Write-Err 'gallery-dl was not found on your PATH.'
    Write-Err 'Install it globally with:  pip install -U gallery-dl'
    exit 1
}
# Grab its version for the log (best effort).
$gdVersion = ''
try { $gdVersion = (& gallery-dl --version 2>$null | Select-Object -First 1) } catch { }
Write-Ok ('gallery-dl    ->  ' + $gd.Source + $(if ($gdVersion) { " (v$gdVersion)" } else { '' }))

# ============================================================================
#  STEP 4 - Work out the "Downloads" folder and a fresh sub-folder for this album
# ============================================================================
function Get-DownloadsFolder {
    # Prefer the user's real Downloads folder from the registry so relocated
    # Downloads directories (moved to another drive) are respected. Falls back
    # to %USERPROFILE%\Downloads.
    try {
        $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
        $val = (Get-ItemProperty -Path $key -Name '{374DE290-123F-4565-9164-39C4925E467B}' -ErrorAction Stop).'{374DE290-123F-4565-9164-39C4925E467B}'
        if ($val) {
            $expanded = [Environment]::ExpandEnvironmentVariables($val)
            if (Test-Path -LiteralPath $expanded) { return $expanded }
        }
    } catch { }
    $fallback = Join-Path $env:USERPROFILE 'Downloads'
    return $fallback
}

$DownloadsRoot = Get-DownloadsFolder
Write-Ok ('Downloads dir ->  ' + $DownloadsRoot)
Write-Host ''

# ============================================================================
#  STEP 5 - Fetch + paginate the album, collecting every image's metadata
#     Each album item carries a URL-encoded data-object JSON blob with:
#        image.url       -> original full-resolution image URL
#        image.extension -> real format (jpg / png / webp / tiff ...)
#        width, height   -> real pixel dimensions
#        title           -> human title (used for filenames)
#     Pages are additive (?page=1 shows items 1..20, ?page=2 shows 1..40, ...)
#     so we stop once a page brings no NEW image ids.
# ============================================================================
Write-Info 'Reading album pages...'

$AlbumTitle   = $null
$seenIds      = [System.Collections.Generic.HashSet[string]]::new()
$allImages    = [System.Collections.Generic.List[object]]::new()
$page         = 1
$emptyStreak  = 0

while ($page -le $MaxPages) {
    $pageUrl = $AlbumUrl + '/?page=' + $page
    try {
        $resp = Invoke-WebRequest -Uri $pageUrl -UseBasicParsing -Headers @{ 'User-Agent' = $UserAgent } -TimeoutSec 60
        $html = $resp.Content
    } catch {
        Write-Warn ("  page {0}: request failed ({1}) - stopping pagination." -f $page, $_.Exception.Message)
        break
    }

    # Grab the album title once (used for the output folder name).
    if (-not $AlbumTitle) {
        $tm = [regex]::Match($html, '<meta\s+property="og:title"\s+content="([^"]*)"', 'IgnoreCase')
        if ($tm.Success -and $tm.Groups[1].Value.Trim()) {
            $AlbumTitle = [System.Net.WebUtility]::HtmlDecode($tm.Groups[1].Value.Trim())
        }
    }

    # Every album item blob: data-object='URLENCODED_JSON'
    $blobs = [regex]::Matches($html, "data-object='([^']+)'")
    $newOnThisPage = 0

    foreach ($b in $blobs) {
        $json = [System.Uri]::UnescapeDataString($b.Groups[1].Value)
        try {
            $obj = $json | ConvertFrom-Json
        } catch {
            continue
        }
        $id = [string]$obj.id_encoded
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        if (-not $seenIds.Add($id)) { continue }   # already collected

        $imgUrl = [string]$obj.image.url
        $ext    = [string]$obj.image.extension
        if (-not $ext) { $ext = [IO.Path]::GetExtension(($imgUrl -split '\?')[0]).TrimStart('.') }

        # Decode HTML entities in titles (e.g. &#039; -> ') for clean filenames.
        $title = [string]$obj.title
        if ($title) { $title = [System.Net.WebUtility]::HtmlDecode($title) }

        $allImages.Add([pscustomobject]@{
            Id        = $id
            Title     = $title
            Url       = $imgUrl
            Extension = $ext.ToLowerInvariant()
            Width     = [int]($obj.width  | ForEach-Object { if ($_ -as [int]) { $_ } else { 0 } })
            Height    = [int]($obj.height | ForEach-Object { if ($_ -as [int]) { $_ } else { 0 } })
            Size      = [long]($obj.image.size | ForEach-Object { if ($_ -as [long]) { $_ } else { 0 } })
        })
        $newOnThisPage++
    }

    Write-Dim ("  page {0}: {1} item(s) on page, {2} new, {3} total so far" -f $page, $blobs.Count, $newOnThisPage, $allImages.Count)

    # End conditions: a page that adds nothing new twice, or an empty page.
    if ($blobs.Count -eq 0) { break }
    if ($newOnThisPage -eq 0) {
        $emptyStreak++
        if ($emptyStreak -ge 2) { break }
    } else {
        $emptyStreak = 0
    }
    $page++
}

if ($allImages.Count -eq 0) {
    Write-Err 'No images were found in this album (it may be private, password protected, or empty).'
    exit 1
}
Write-Ok ("Found {0} image(s) across the album." -f $allImages.Count)
Write-Host ''


# ============================================================================
#  STEP 6 - Filter: real format whitelist + minimum 300x300 resolution
# ============================================================================
$kept       = [System.Collections.Generic.List[object]]::new()
$skipFormat = 0
$skipSmall  = 0
$skipNoData = 0

foreach ($img in $allImages) {
    if ([string]::IsNullOrWhiteSpace($img.Url)) { $skipNoData++; continue }

    if ($AllowedExt -notcontains $img.Extension) {
        $skipFormat++
        Write-Dim ("  skip (format .{0})  {1}" -f $img.Extension, $img.Id)
        continue
    }
    if ($img.Width -lt $MinWidth -or $img.Height -lt $MinHeight) {
        $skipSmall++
        Write-Dim ("  skip (too small {0}x{1})  {2}" -f $img.Width, $img.Height, $img.Id)
        continue
    }
    $kept.Add($img)
}

Write-Host ''
Write-Ok   ("Kept {0} image(s) matching format + {1}x{2} minimum." -f $kept.Count, $MinWidth, $MinHeight)
Write-Dim  ("  Skipped: {0} wrong format, {1} below resolution, {2} missing data." -f $skipFormat, $skipSmall, $skipNoData)

if ($kept.Count -eq 0) {
    Write-Warn 'Nothing passed the filters - no folder created and nothing downloaded.'
    exit 0
}

# ============================================================================
#  STEP 7 - Create the output sub-folder inside Downloads
# ============================================================================
if ([string]::IsNullOrWhiteSpace($AlbumTitle)) { $AlbumTitle = 'gifyu_album_' + $AlbumId }
$folderName = Make-SafeName -Name $AlbumTitle -MaxLen 100
if ([string]::IsNullOrWhiteSpace($folderName)) { $folderName = 'gifyu_album_' + $AlbumId }

$outDir = Join-Path $DownloadsRoot $folderName
# If it already exists, add a numeric suffix instead of mixing into it.
if (Test-Path -LiteralPath $outDir) {
    $n = 2
    while (Test-Path -LiteralPath (Join-Path $DownloadsRoot ("{0} ({1})" -f $folderName, $n))) { $n++ }
    $folderName = "{0} ({1})" -f $folderName, $n
    $outDir = Join-Path $DownloadsRoot $folderName
}
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
Write-Ok ('Output folder ->  ' + $outDir)
Write-Host ''

# ============================================================================
#  STEP 8 - Download each original image with gallery-dl
#     -d <folder>            -> destination folder
#     -o directory=[]        -> flat (no nested sub-folders from gallery-dl)
#     -o filename="<name>"   -> exact filename incl. real extension
#     gallery-dl treats direct image URLs via its built-in 'directlink'
#     extractor, so this uses the global install exactly as required.
# ============================================================================
Write-Info 'Downloading originals via gallery-dl...'
Write-Host ''

$okCount      = 0
$failCount    = 0
$index        = 0
$usedNames    = @{}
$total        = $kept.Count

foreach ($img in $kept) {
    $index++
    $prefix = "[$index/$total]"

    # Build a clean, readable filename: "<title>_<id>.<ext>"
    $base = if ($img.Title) { $img.Title } else { 'image_' + $img.Id }
    $safe = Make-SafeName -Name $base -MaxLen 90
    $fileName = '{0}_{1}.{2}' -f $safe, $img.Id, $img.Extension
    # Guard against duplicate names within this run.
    $dedupe = $fileName
    $c = 2
    while ($usedNames.ContainsKey($dedupe.ToLowerInvariant())) {
        $dedupe = '{0}_{1}_{2}.{3}' -f $safe, $img.Id, $c, $img.Extension
        $c++
    }
    $fileName = $dedupe
    $usedNames[$fileName.ToLowerInvariant()] = $true

    $targetPath = Join-Path $outDir $fileName
    $sizeMb = if ($img.Size -gt 0) { ('{0:N1} MB' -f ($img.Size / 1MB)) } else { '? MB' }
    Write-Host ("{0} {1}  ({2})" -f $prefix, $fileName, $sizeMb) -ForegroundColor White

    try {
        # gallery-dl prints progress to stdout/stderr; capture quietly.
        & gallery-dl -d $outDir -o 'directory=[]' -o ('filename="' + $fileName + '"') $img.Url 2>&1 | Out-Null
        $gdExit = $LASTEXITCODE
        if ($gdExit -ne 0) { throw "gallery-dl exited with code $gdExit" }
        if (-not (Test-Path -LiteralPath $targetPath)) { throw 'file was not created' }
        $okCount++
        Write-Host '      ok' -ForegroundColor Green
    } catch {
        $failCount++
        Write-Warn ('      failed: ' + $_.Exception.Message)
    }
}

# ============================================================================
#  STEP 9 - Summary
# ============================================================================
Write-Host ''
Write-Host '============================================================' -ForegroundColor Magenta
Write-Host '                         SUMMARY                            ' -ForegroundColor Magenta
Write-Host '============================================================' -ForegroundColor Magenta
Write-Ok   ("Album     : {0}" -f $AlbumTitle)
Write-Ok   ("URL       : {0}" -f $AlbumUrl)
Write-Ok   ("Folder    : {0}" -f $outDir)
Write-Ok   ("Downloaded: {0}" -f $okCount)
if ($skipFormat -gt 0) { Write-Dim ("Skipped   : {0} wrong format" -f $skipFormat) }
if ($skipSmall  -gt 0) { Write-Dim ("Skipped   : {0} below {1}x{2}" -f $skipSmall, $MinWidth, $MinHeight) }
if ($failCount  -gt 0) { Write-Warn ("Failed    : {0}" -f $failCount) }
Write-Host ''

if ($failCount -gt 0) { exit 2 }
exit 0

