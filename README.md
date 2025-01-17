# PSCopyFileEx

A PowerShell module for copying files with advanced progress reporting and Windows API support. Features include detailed progress bars, speed reporting, and support for both managed and Windows API file copy operations.

# Features

- Windows API support through CopyFileEx
- Managed file copy fallback when API is unavailable
- Detailed progress reporting with:
  - Current speed
  - Total progress
  - File-by-file progress
  - Total size of copied data
- Support for single files and directories
- Recursive copy support
- Force parameter for overwrite control
- Special character handling with LiteralPath
- PassThru parameter for returning copied file objects
- Pipeline support for copying multiple files

# Installation

1. Install the module
```powershell
Install-Module -Name PSCopyFileEx
```
2. Import the module:
```powershell
Import-Module PSCopyFileEx
```

# Usage

## Basic File Copy
```powershell
Copy-FileEx -Path "C:\source\file.txt" -Destination "D:\backup"
```

## Copy with Force Overwrite
```powershell
Copy-FileEx -Path "C:\source\file.txt" -Destination "D:\backup" -Force
```

## Copy Directory Recursively
```powershell
Copy-FileEx -Path "C:\source\folder" -Destination "D:\backup" -Recurse
```

## Copy with Verbose Output
```powershell
Copy-FileEx -Path "C:\source\file.txt" -Destination "D:\backup" -Verbose
```

## Copy Files with Special Characters
```powershell
Copy-FileEx -LiteralPath "C:\source\file[1].txt" -Destination "D:\backup"
```

## Copy Multiple Files using Wildcards
```powershell
Copy-FileEx -Path "C:\source\*.txt" -Destination "D:\backup"
```

## Copy Using Managed Method (No Win32API)
```powershell
Copy-FileEx -Path "C:\source\file.txt" -Destination "D:\backup" -UseWinApi $false
```

## Pipeline Support
```powershell
Get-ChildItem "C:\source" -Filter "*.txt" | Copy-FileEx -Destination "D:\backup"
```

## Special Filters
```powershell
Copy-FileEx -Path "C:\source\" -Destination "D:\backup" -Include "*.txt" -Recurse
Copy-FileEx -Path "C:\source\" -Destination "D:\backup" -Exclude ".git" -Recurse -Force
```

# Parameters

- `Path`: Path to source file(s) or directory (supports wildcards)
- `LiteralPath`: Path to source file(s) or directory (no wildcard interpretation)
- `Destination`: Destination path for copied files
- `Include`: Optional include filter
- `Exclude`: Optional exclude filter
- `Recurse`: Copy directories recursively
- `Force`: Overwrite existing files
- `PassThru`: Return copied file objects
- `UseWinApi`: Use Windows API for copying (default: $true)

# Notes

- The module will automatically fall back to managed copy if Win32API is unavailable
- Progress reporting works in both API and managed copy modes
- Directory structure is preserved in recursive copies
- Files are not overwritten unless -Force is specified

# Examples

For more detailed help and examples, run:
```powershell
Get-Help Copy-FileEx -Full
```