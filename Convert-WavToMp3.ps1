<#
.SYNOPSIS
    Converts WAV audio files to MP3 format using FFmpeg.

.DESCRIPTION
    This script converts one or more WAV files to MP3 format with configurable quality settings.
    Requires FFmpeg to be installed and accessible in the system PATH.

.PARAMETER InputPath
    Path to a WAV file or folder containing WAV files to convert.

.PARAMETER OutputPath
    Optional output directory. If not specified, MP3 files will be created in the same location as the input files.

.PARAMETER Bitrate
    MP3 bitrate in kbps. Default is 192. Common values: 128, 192, 256, 320.

.PARAMETER Quality
    VBR quality preset (0-9). Lower is better. Use this instead of bitrate for variable bitrate encoding.
    0 = best quality, 9 = worst quality. Default uses constant bitrate.

.PARAMETER Recursive
    Process WAV files in subfolders recursively.

.PARAMETER Overwrite
    Overwrite existing MP3 files without prompting.

.EXAMPLE
    .\Convert-WavToMp3.ps1 -InputPath "C:\Music\song.wav"
    Converts a single WAV file to MP3 with default settings (192 kbps).

.EXAMPLE
    .\Convert-WavToMp3.ps1 -InputPath "C:\Music" -Bitrate 320 -Recursive
    Converts all WAV files in the Music folder and subfolders to 320 kbps MP3.

.EXAMPLE
    .\Convert-WavToMp3.ps1 -InputPath "C:\Audio" -Quality 2 -OutputPath "C:\MP3"
    Converts WAV files using VBR quality preset 2 (high quality) to a specific output folder.

.NOTES
    Author: Audio Converter Script
    Requires: FFmpeg (https://ffmpeg.org/download.html)
    Version: 1.0
#>

param(
    [Parameter(Mandatory=$true, HelpMessage="Path to WAV file or folder")]
    [string]$InputPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(64, 320)]
    [int]$Bitrate = 192,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(0, 9)]
    [int]$Quality = -1,
    
    [Parameter(Mandatory=$false)]
    [switch]$Recursive,
    
    [Parameter(Mandatory=$false)]
    [switch]$Overwrite
)

# Check if FFmpeg is installed
function Test-FFmpeg {
    try {
        $null = & ffmpeg -version 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ FFmpeg found" -ForegroundColor Green
            return $true
        }
    }
    catch {
        Write-Host "✗ FFmpeg not found in system PATH" -ForegroundColor Red
        Write-Host "`nPlease install FFmpeg:" -ForegroundColor Yellow
        Write-Host "  1. Download from: https://ffmpeg.org/download.html" -ForegroundColor Cyan
        Write-Host "  2. Extract the files" -ForegroundColor Cyan
        Write-Host "  3. Add FFmpeg bin folder to system PATH" -ForegroundColor Cyan
        Write-Host "`nOr install via package manager:" -ForegroundColor Yellow
        Write-Host "  winget install FFmpeg" -ForegroundColor Cyan
        Write-Host "  choco install ffmpeg" -ForegroundColor Cyan
        return $false
    }
}

# Convert a single WAV file to MP3
function Convert-WavFile {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [int]$Bitrate,
        [int]$Quality,
        [bool]$Overwrite
    )
    
    # Check if input file exists
    if (-not (Test-Path $InputFile)) {
        Write-Warning "File not found: $InputFile"
        return $false
    }
    
    # Check if output file exists
    if ((Test-Path $OutputFile) -and -not $Overwrite) {
        $response = Read-Host "Output file exists: $OutputFile`nOverwrite? (Y/N)"
        if ($response -ne 'Y' -and $response -ne 'y') {
            Write-Host "Skipped: $InputFile" -ForegroundColor Yellow
            return $false
        }
    }
    
    # Build FFmpeg arguments
    $ffmpegArgs = @(
        "-i", "`"$InputFile`""
        "-vn"  # No video
    )
    
    # Use VBR or CBR encoding
    if ($Quality -ge 0) {
        # Variable Bitrate (VBR)
        $ffmpegArgs += @("-q:a", $Quality)
        $qualityText = "VBR Quality $Quality"
    }
    else {
        # Constant Bitrate (CBR)
        $ffmpegArgs += @("-b:a", "${Bitrate}k")
        $qualityText = "${Bitrate} kbps"
    }
    
    # Add output file
    if ($Overwrite) {
        $ffmpegArgs += "-y"
    }
    else {
        $ffmpegArgs += "-n"
    }
    
    $ffmpegArgs += "`"$OutputFile`""
    
    Write-Host "`nConverting: " -NoNewline
    Write-Host (Split-Path $InputFile -Leaf) -ForegroundColor Cyan
    Write-Host "Quality:    $qualityText" -ForegroundColor Gray
    Write-Host "Output:     " -NoNewline
    Write-Host (Split-Path $OutputFile -Leaf) -ForegroundColor Cyan
    
    # Execute FFmpeg
    $argString = $ffmpegArgs -join " "
    try {
        $process = Start-Process -FilePath "ffmpeg" -ArgumentList $argString -NoNewWindow -Wait -PassThru -RedirectStandardError "$env:TEMP\ffmpeg_error.log"
        
        if ($process.ExitCode -eq 0) {
            $outputInfo = Get-Item $OutputFile
            $sizeKB = [math]::Round($outputInfo.Length / 1KB, 2)
            Write-Host "✓ Success! Size: $sizeKB KB" -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "✗ Conversion failed!" -ForegroundColor Red
            $errorLog = Get-Content "$env:TEMP\ffmpeg_error.log" -Raw
            Write-Host "Error details: $errorLog" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "✗ Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Main script execution
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    WAV to MP3 Converter" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Verify FFmpeg is available
if (-not (Test-FFmpeg)) {
    exit 1
}

# Resolve input path
$InputPath = Resolve-Path $InputPath -ErrorAction SilentlyContinue
if (-not $InputPath) {
    Write-Host "✗ Input path not found!" -ForegroundColor Red
    exit 1
}

# Get list of WAV files
$wavFiles = @()
if (Test-Path $InputPath -PathType Leaf) {
    # Single file
    if ($InputPath -notmatch '\.wav$') {
        Write-Host "✗ Input file is not a WAV file!" -ForegroundColor Red
        exit 1
    }
    $wavFiles = @(Get-Item $InputPath)
}
else {
    # Directory
    if ($Recursive) {
        $wavFiles = Get-ChildItem -Path $InputPath -Filter "*.wav" -Recurse -File
    }
    else {
        $wavFiles = Get-ChildItem -Path $InputPath -Filter "*.wav" -File
    }
}

if ($wavFiles.Count -eq 0) {
    Write-Host "✗ No WAV files found!" -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $($wavFiles.Count) WAV file(s)" -ForegroundColor Green
Write-Host ""

# Process each file
$successCount = 0
$failCount = 0

foreach ($wavFile in $wavFiles) {
    # Determine output path
    if ($OutputPath) {
        # Ensure output directory exists
        if (-not (Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }
        
        # Preserve folder structure if recursive
        if ($Recursive -and (Test-Path $InputPath -PathType Container)) {
            $relativePath = $wavFile.FullName.Substring($InputPath.Path.Length + 1)
            $relativeDir = Split-Path $relativePath -Parent
            $outputDir = Join-Path $OutputPath $relativeDir
            
            if ($relativeDir -and -not (Test-Path $outputDir)) {
                New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
            }
        }
        else {
            $outputDir = $OutputPath
        }
        
        $mp3File = Join-Path $outputDir ($wavFile.BaseName + ".mp3")
    }
    else {
        # Output in same directory as input
        $mp3File = Join-Path $wavFile.DirectoryName ($wavFile.BaseName + ".mp3")
    }
    
    # Convert the file
    $result = Convert-WavFile -InputFile $wavFile.FullName -OutputFile $mp3File -Bitrate $Bitrate -Quality $Quality -Overwrite $Overwrite
    
    if ($result) {
        $successCount++
    }
    else {
        $failCount++
    }
}

# Summary
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Conversion Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total files:    $($wavFiles.Count)" -ForegroundColor White
Write-Host "Successful:     $successCount" -ForegroundColor Green
Write-Host "Failed:         $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Gray" })
Write-Host ""
