#Requires -Version 5.1
<#
  Copy metadata from Source.mp3 onto Target.mp3 (same folder as this script).
  Mirrors DOpus_ffmpeg.js MP3 restore: primary -all:all -unsafe, then supplement passes.

  Usage: place Source.mp3 and Target.mp3 next to this file, then:
    powershell -NoProfile -ExecutionPolicy Bypass -File .\exif_copy_mp3_SourceToTarget.ps1

  Official Windows build: https://exiftool.org/ — keep exiftool.exe and the exiftool_files folder together.
  If exiftool_files is missing or incomplete, reads may work but MP3 writes can fail oddly.

  Optional: env EXIFTOOL_EXE or -ExifTool "C:\path\to\exiftool.exe"
#>
param(
    [string]$ExifTool = $env:EXIFTOOL_EXE
)

if (-not $ExifTool) {
    $ExifTool = "C:\Users\WXP\Desktop\Tools\ExifTool-13.55\exiftool-13.55_64\exiftool.exe"
}

$here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$src = Join-Path $here "Source.mp3"
$dst = Join-Path $here "Target.mp3"

if (-not (Test-Path -LiteralPath $ExifTool)) {
    Write-Error "ExifTool not found: $ExifTool`nSet env EXIFTOOL_EXE or pass -ExifTool."
    exit 1
}
if (-not (Test-Path -LiteralPath $src)) {
    Write-Error "Missing: $src"
    exit 1
}
if (-not (Test-Path -LiteralPath $dst)) {
    Write-Error "Missing: $dst"
    exit 1
}

$ExifTool = (Resolve-Path -LiteralPath $ExifTool).Path
$exifDir = Split-Path -Parent $ExifTool
$exifBundle = Join-Path $exifDir "exiftool_files"

Write-Host "--- ExifTool binary ---"
Write-Host $ExifTool
Write-Host "--- exiftool_files (required next to exe for full write support) ---"
if (Test-Path -LiteralPath $exifBundle) {
    Write-Host "OK: $exifBundle"
} else {
    Write-Host "WARNING: folder not found. Re-unzip exiftool-13.55_64.zip so exiftool.exe and exiftool_files sit in the same directory."
}
Write-Host "--- Version ---"
& $ExifTool -ver 2>&1 | ForEach-Object { Write-Host $_ }
if ($LASTEXITCODE -ne 0) {
    Write-Error "exiftool -ver failed; check the executable above."
    exit 1
}

function Write-ExifArgsFile {
    param([string]$Path, [string[]]$Lines)
    $enc = New-Object System.Text.UTF8Encoding $false
    $body = ($Lines -join "`r`n") + "`r`n"
    [System.IO.File]::WriteAllText($Path, $body, $enc)
}

function Invoke-ExifArgFile {
    param(
        [string]$Label,
        [string[]]$ArgLines
    )
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("exif_copy_args_" + [Guid]::NewGuid().ToString("n") + ".args")
    try {
        Write-ExifArgsFile -Path $tmp -Lines $ArgLines
        Write-Host "[$Label] & `"$ExifTool`" -@ `"$tmp`""
        # Do not let exiftool stdout/stderr become function output (breaks $code = Run-Pass).
        & $ExifTool "-@", $tmp 2>&1 | ForEach-Object { Write-Host $_ }
        return [int]$LASTEXITCODE
    }
    finally {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
}

function Run-Pass {
    param([string]$Label, [string[]]$ExtraAfterTagsFromFile)
    # Write to a temp MP3 then replace Target — avoids some in-place / cloud / locker issues.
    $outTmp = Join-Path ([System.IO.Path]::GetTempPath()) ("exif_copy_out_" + [Guid]::NewGuid().ToString("n") + ".mp3")
    try {
        $lines = @(
            "-m",
            "-charset", "filename=UTF8",
            "-TagsFromFile", $src
        ) + $ExtraAfterTagsFromFile + @(
            "-o", $outTmp,
            $dst
        )
        $exitCode = Invoke-ExifArgFile -Label $Label -ArgLines $lines
        if ($exitCode -eq 0 -and (Test-Path -LiteralPath $outTmp)) {
            Move-Item -LiteralPath $outTmp -Destination $dst -Force
        }
        return $exitCode
    }
    finally {
        if (Test-Path -LiteralPath $outTmp) {
            Remove-Item -LiteralPath $outTmp -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-ExifReadLines {
    param([string[]]$TagArgs, [string]$File)
    & $ExifTool @TagArgs $File 2>&1 | ForEach-Object { Write-Host $_ }
}

Write-Host "--- Quick read (sanity) ---"
Invoke-ExifReadLines -TagArgs @("-s3", "-Title", "-Artist", "-Comment", "-Comment-xxx") -File $src
Write-Host "---"
Invoke-ExifReadLines -TagArgs @("-s3", "-Title", "-Artist", "-Comment", "-Comment-xxx") -File $dst
Write-Host "--- Target type ---"
Invoke-ExifReadLines -TagArgs @("-FileType", "-FileTypeExtension", "-MIMEType") -File $dst
Write-Host ""

$code = Run-Pass "1 primary all:all" @("-all:all", "-unsafe")
if ($code -ne 0) {
    Write-Host "Primary exit $code - trying ID3:All fallback..."
    $code = Run-Pass "2 fallback ID3:All" @("-ID3:All")
}

# -Popularimeter is not writable for TagsFromFile from many MP3s (POPM); skip for copy pass.
$null = Run-Pass "3 supplement TXXX/WXXX/PRIV/WOAR*" @(
    "-TXXX:All", "-WXXX:All", "-PRIV:All",
    "-WOAR", "-WOAS", "-WORS", "-WCOM"
)
$null = Run-Pass "4 supplement Comment + Comment-xxx" @("-Comment", "-Comment-xxx")
$null = Run-Pass "5 supplement APE" @("-APE:All")

Write-Host ("Done. Main pass exit code: " + $code)
