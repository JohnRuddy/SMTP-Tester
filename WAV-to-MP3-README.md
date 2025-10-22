# WAV to MP3 Converter

A PowerShell script to convert WAV audio files to MP3 format using FFmpeg.

## Features

- Convert single WAV files or entire folders
- Recursive folder processing
- Configurable bitrate (CBR) or quality (VBR) settings
- Batch conversion with progress tracking
- Preserves folder structure when using recursive mode
- Automatic output directory creation
- Overwrite protection with user prompts

## Requirements

- Windows PowerShell 5.1 or later
- FFmpeg installed and accessible in system PATH

### Installing FFmpeg

**Option 1: Package Manager (Recommended)**
```powershell
# Using winget
winget install FFmpeg

# Using Chocolatey
choco install ffmpeg
```

**Option 2: Manual Installation**
1. Download from https://ffmpeg.org/download.html
2. Extract the archive
3. Add the `bin` folder to your system PATH

## Usage

### Basic Conversion

Convert a single WAV file:
```powershell
.\Convert-WavToMp3.ps1 -InputPath "C:\Music\song.wav"
```

### Batch Conversion

Convert all WAV files in a folder:
```powershell
.\Convert-WavToMp3.ps1 -InputPath "C:\Music"
```

### Recursive Conversion

Convert all WAV files including subfolders:
```powershell
.\Convert-WavToMp3.ps1 -InputPath "C:\Music" -Recursive
```

### Custom Output Directory

Specify where to save MP3 files:
```powershell
.\Convert-WavToMp3.ps1 -InputPath "C:\Music" -OutputPath "C:\MP3"
```

### Quality Settings

**Constant Bitrate (CBR):**
```powershell
# High quality (320 kbps)
.\Convert-WavToMp3.ps1 -InputPath "C:\Music" -Bitrate 320

# Standard quality (192 kbps) - Default
.\Convert-WavToMp3.ps1 -InputPath "C:\Music" -Bitrate 192

# Lower quality (128 kbps)
.\Convert-WavToMp3.ps1 -InputPath "C:\Music" -Bitrate 128
```

**Variable Bitrate (VBR):**
```powershell
# Highest quality (VBR 0)
.\Convert-WavToMp3.ps1 -InputPath "C:\Music" -Quality 0

# High quality (VBR 2)
.\Convert-WavToMp3.ps1 -InputPath "C:\Music" -Quality 2

# Medium quality (VBR 4)
.\Convert-WavToMp3.ps1 -InputPath "C:\Music" -Quality 4
```

### Overwrite Existing Files

Automatically overwrite existing MP3 files:
```powershell
.\Convert-WavToMp3.ps1 -InputPath "C:\Music" -Overwrite
```

### Complete Example

Convert all WAV files recursively with high-quality VBR encoding:
```powershell
.\Convert-WavToMp3.ps1 -InputPath "C:\AudioProjects" -OutputPath "C:\MP3Output" -Quality 2 -Recursive -Overwrite
```

## Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| InputPath | String | Yes | - | Path to WAV file or folder |
| OutputPath | String | No | Same as input | Output directory for MP3 files |
| Bitrate | Integer | No | 192 | MP3 bitrate in kbps (64-320) |
| Quality | Integer | No | -1 | VBR quality (0-9, lower is better) |
| Recursive | Switch | No | False | Process subfolders |
| Overwrite | Switch | No | False | Overwrite without prompting |

## Quality Guide

### Bitrate (CBR) Recommendations
- **64-96 kbps**: Voice recordings, podcasts
- **128 kbps**: Acceptable music quality, smaller file size
- **192 kbps**: Good music quality (default)
- **256 kbps**: High-quality music
- **320 kbps**: Maximum quality, larger file size

### VBR Quality Settings
- **0-1**: Highest quality (~245 kbps average)
- **2-3**: High quality (~190 kbps average) - Recommended
- **4-5**: Medium quality (~165 kbps average)
- **6-7**: Lower quality (~130 kbps average)
- **8-9**: Lowest quality (~115 kbps average)

## Examples

### Convert album with high quality
```powershell
.\Convert-WavToMp3.ps1 -InputPath "C:\Album" -Quality 2
```

### Convert podcast episodes
```powershell
.\Convert-WavToMp3.ps1 -InputPath "C:\Podcasts" -Bitrate 96 -Recursive
```

### Prepare music library for portable device
```powershell
.\Convert-WavToMp3.ps1 -InputPath "C:\FLAC\Music" -OutputPath "D:\MP3Player" -Bitrate 256 -Recursive -Overwrite
```

## Troubleshooting

### "FFmpeg not found"
- Ensure FFmpeg is installed
- Verify FFmpeg is in your system PATH
- Restart PowerShell after installing FFmpeg

### "Conversion failed"
- Check that the input file is a valid WAV file
- Ensure you have write permissions to the output directory
- Check the error log at `%TEMP%\ffmpeg_error.log`

### Performance
- Converting large files or many files may take time
- VBR encoding is slightly slower than CBR
- Consider using lower quality settings for faster conversion

## Notes

- Output MP3 files are created in the same folder structure as input
- Original WAV files are preserved (not deleted)
- File metadata may not be preserved in conversion
- Script displays progress for each file being converted

## License

This script is provided as-is for personal and commercial use.
