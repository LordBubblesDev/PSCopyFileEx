# .SYNOPSIS
# Copies files with advanced progress reporting and Windows API support.
# 
# .DESCRIPTION
# Copy-FileEx provides enhanced file copy capabilities with detailed progress reporting.
# It leverages the Windows CopyFileEx API when available and falls back to managed file
# copy operations when necessary. Features include speed reporting, progress bars,
# recursive copying, and special character handling.
#
# .PARAMETER Path
# Path to source file(s) or directory. Supports wildcards.
#
# .PARAMETER LiteralPath
# Path to source file(s) or directory. Does not support wildcards. Use this when path contains special characters.
#
# .PARAMETER Destination
# Destination path where files will be copied to.
#
# .PARAMETER Include
# Optional array of include filters (e.g., "*.txt", "file?.doc").
#
# .PARAMETER Exclude
# Optional array of exclude filters (e.g., "*.tmp", "~*").
#
# .PARAMETER Recurse
# If specified, copies subdirectories recursively. Required for directory copies.
#
# .PARAMETER Force
# If specified, overwrites existing files. Without this, existing files are skipped.
#
# .PARAMETER PassThru
# If specified, returns objects representing copied items.
#
# .PARAMETER UseWinApi
# If true (default), uses Windows CopyFileEx API. If false, uses managed file copy.
#
# .EXAMPLE
# Copy-FileEx -Path "C:\source\file.txt" -Destination "D:\backup"
# 
# Copies a single file with progress reporting.
#
# .EXAMPLE
# Copy-FileEx -Path "C:\source\folder" -Destination "D:\backup" -Recurse
# 
# Copies a directory and all its contents recursively.
#
# .EXAMPLE
# Copy-FileEx -Path "C:\source\*.txt" -Destination "D:\backup" -Force
# 
# Copies all .txt files, overwriting any existing files.
#
# .EXAMPLE
# Copy-FileEx -LiteralPath "C:\source\file[1].txt" -Destination "D:\backup"
# 
# Copies a file with special characters in the name.
#
# .EXAMPLE
# Get-ChildItem "C:\source" -Filter "*.txt" | Copy-FileEx -Destination "D:\backup"
# 
# Uses pipeline input for copying multiple files.
#
# .EXAMPLE
# Copy-FileEx -Path "C:\source\large.iso" -Destination "D:\backup" -UseWinApi $false
# 
# Forces use of managed copy method instead of Windows API.
#
# .NOTES
# Author: LordBubbles
# Module: PSCopyFileEx
# Version: 1.0.0
#
# Performance Notes:
# - Windows API method is generally faster
# - Large files benefit from API buffering
# - Network paths use compressed traffic when possible
#
# .LINK
# https://github.com/LordBubblesDev/PSCopyFileEx

Function Copy-FileEx {
    [CmdletBinding(SupportsShouldProcess=$true, DefaultParameterSetName='Path')]
    param(
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='Path')]
        [string[]]$Path,

        [Parameter(Mandatory=$true, Position=0, ValueFromPipelineByPropertyName=$true, ParameterSetName='LiteralPath')]
        [Alias('LP')]
        [string[]]$LiteralPath,

        [Parameter(Position=1, ValueFromPipelineByPropertyName=$true)]
        [string]$Destination,

        [Parameter()]
        [string[]]$Include,

        [Parameter()]
        [string[]]$Exclude,

        [Parameter()]
        [switch]$Recurse,

        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [switch]$PassThru,

        [Parameter()]
        [bool]$UseWinApi = $true
    )

    begin {
        # Generate a random progress ID to avoid conflicts
        $progressId = Get-Random -Minimum 0 -Maximum 1000
        $childProgressId = $progressId + 1

        # Initialize all variables that will be used across the function
        $script:speedSampleSize = 100  # Number of samples to average
        $script:speedSamples = @()
        $script:lastSpeedCheck = [DateTime]::Now
        $script:lastBytesForSpeed = 0
        $script:lastTime = [DateTime]::Now
        $script:lastBytes = 0
        $script:lastSpeedUpdate = [DateTime]::Now
        $script:currentSpeed = 0
        $script:lastProgressUpdate = [DateTime]::Now
        $script:progressThreshold = [TimeSpan]::FromMilliseconds(100)

        function Format-FileSize {
            param([long]$Size)
            
            switch ($Size) {
                { $_ -gt 1TB } { "{0:n2} TB" -f ($_ / 1TB); Break }
                { $_ -gt 1GB } { "{0:n2} GB" -f ($_ / 1GB); Break }
                { $_ -gt 1MB } { "{0:n2} MB" -f ($_ / 1MB); Break }
                { $_ -gt 1KB } { "{0:n2} KB" -f ($_ / 1KB); Break }
                default { "{0} B " -f $_ }
            }
        }
    
        function Get-CurrentSpeed {
            param (
                [DateTime]$now,
                [long]$currentBytes
            )
            
            $timeDiff = ($now - $lastSpeedCheck).TotalSeconds
            if ($timeDiff -gt 0) {
                $bytesDiff = $currentBytes - $lastBytesForSpeed
                $speed = $bytesDiff / $timeDiff
                
                # Add to rolling samples
                $speedSamples += $speed
                if ($speedSamples.Count -gt $speedSampleSize) {
                    $speedSamples = $speedSamples | Select-Object -Last $speedSampleSize
                }
                
                # Calculate average speed
                $avgSpeed = ($speedSamples | Measure-Object -Average).Average
                
                # Update last values
                $script:lastSpeedCheck = $now
                $script:lastBytesForSpeed = $currentBytes
                
                return $avgSpeed
            }
            return 0
        }

        # Attempt to use Windows API for CopyFileEx if UseWinApi is true
        if ($UseWinApi) {
            $useWin32Api = $true
            
            # Check if type already exists
            if (-not ([System.Management.Automation.PSTypeName]'Win32Helpers.Win32CopyFileEx').Type) {
                $signature = @'
                [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
                public static extern bool CopyFileEx(
                    string lpExistingFileName,
                    string lpNewFileName,
                    CopyProgressRoutine lpProgressRoutine,
                    IntPtr lpData,
                    ref bool pbCancel,
                    uint dwCopyFlags
                );

                public delegate uint CopyProgressRoutine(
                    long TotalFileSize,
                    long TotalBytesTransferred,
                    long StreamSize,
                    long StreamBytesTransferred,
                    uint dwStreamNumber,
                    uint dwCallbackReason,
                    IntPtr hSourceFile,
                    IntPtr hDestinationFile,
                    IntPtr lpData
                );
'@
                try {
                    Add-Type -MemberDefinition $signature -Name "Win32CopyFileEx" -Namespace "Win32Helpers"
                }
                catch {
                    Write-Warning "Failed to use Windows API for file copy operations, falling back to managed copy method"
                    $useWin32Api = $false
                }
            }
        }
    }

    process {
        # Handle both Path and LiteralPath parameters
        $pathsToProcess = @()
        if ($LiteralPath) {
            $pathsToProcess += $LiteralPath
            $useWildcards = $false
        } else {
            $pathsToProcess += $Path
            $useWildcards = $true
        }

        foreach ($currentPath in $pathsToProcess) {
            # Process paths based on whether they're literal or support wildcards
            try {
                # Handle path resolution differently for files with special characters
                if (Test-Path -LiteralPath $currentPath) {
                    $resolvedPaths = @([pscustomobject]@{
                        Path = $currentPath
                        ProviderPath = (Get-Item -LiteralPath $currentPath).FullName
                    })
                } else {
                    if ($useWildcards) {
                        $resolvedPaths = Resolve-Path -Path $currentPath -ErrorAction Stop
                    } else {
                        $resolvedPaths = Resolve-Path -LiteralPath $currentPath -ErrorAction Stop
                    }
                }

                foreach ($resolvedPath in $resolvedPaths) {
                    # Apply Include/Exclude filters
                    $shouldProcess = $true
                    if ($Include) {
                        $shouldProcess = $resolvedPath.Path | Where-Object { 
                            $item = $_
                            return ($Include | ForEach-Object { $item -like $_ }) -contains $true
                        }
                    }
                    if ($Exclude -and $shouldProcess) {
                        $shouldProcess = $resolvedPath.Path | Where-Object { 
                            $item = $_
                            return ($Exclude | ForEach-Object { $item -like $_ }) -notcontains $true
                        }
                    }

                    if ($shouldProcess) {
                        # Check if we should process this item
                        $targetPath = Join-Path $Destination (Split-Path -Leaf $resolvedPath.ProviderPath)
                        if ($Force -or $PSCmdlet.ShouldProcess($targetPath)) {
                            # Handle wildcards in path
                            $sourcePath = Split-Path -Path $currentPath -Parent
                            $sourceFilter = Split-Path -Path $currentPath -Leaf

                            try {
                                # Initialize variables
                                $isFile = $false
                                $relativePath = $null
                                
                                if ($sourceFilter.Contains('*')) {
                                    # Path contains wildcards
                                    $files = Get-ChildItem -Path $sourcePath -Filter $sourceFilter -File -Recurse:$Recurse -ErrorAction Stop
                                    $basePath = $sourcePath
                                } else {
                                    # Single file or directory
                                    $item = Get-Item -LiteralPath $currentPath -ErrorAction Stop
                                    $isFile = $item -is [System.IO.FileInfo]
                                    if ($isFile) {
                                        Write-Verbose "Single file: $($item.Name)"
                                        $files = @($item)
                                        $basePath = Split-Path -Path $item.FullName -Parent
                                        # For single files, use the file's directory as base path
                                        $relativePath = $item.Name
                                    } else {
                                        $files = Get-ChildItem -Path $currentPath -File -Recurse:$Recurse -ErrorAction Stop
                                        $basePath = $item.FullName
                                    }
                                }
                                Write-Verbose "Base Path: $basePath"
                            }
                            catch {
                                Write-Warning "Error accessing path: $_"
                                continue
                            }

                            if ($files.Count -eq 0) {
                                Write-Warning "No files found to copy"
                                continue
                            }

                            # Calculate total size
                            Write-Verbose "Calculating total size..."
                            try {
                                $totalSize = ($files | Measure-Object -Property Length -Sum).Sum
                                Write-Verbose "Total bytes to copy: $totalSize"
                            }
                            catch {
                                Write-Warning "Error calculating size: $_"
                                continue
                            }

                            $totalBytesCopied = 0
                            $startTime = [DateTime]::Now

                            # Initialize variables for managed copy method
                            if (-not $useWin32Api) {
                                # Set up buffer and timing
                                $bufferSize = 4MB
                                $buffer = New-Object byte[] $bufferSize
                                
                                # Reset speed calculation variables for new copy operation
                                $script:speedSamples = @()
                                $script:lastSpeedCheck = $startTime
                                $script:lastBytesForSpeed = 0
                                $script:lastProgressUpdate = $startTime
                            }

                            # Show initial progress only for multiple files
                            if ($files.Count -gt 1) {
                                Write-Progress -Activity "Copying files" `
                                    -Status "0 of $($files.Count) files (0 of $(Format-FileSize $totalSize))" `
                                    -PercentComplete 0 `
                                    -Id $progressId
                            }

                            $filesCopied = 0
                            $verboseOutput = @()  # Collect verbose messages
                            foreach ($file in $files) {
                                $filesCopied++
                                # Calculate relative path for destination
                                if ($isFile) {
                                    # Use pre-calculated relative path for single files
                                    $destPath = Join-Path $Destination $relativePath
                                } else {
                                    # Calculate relative path for files in directories
                                    $relativePath = $file.FullName.Substring($basePath.Length).TrimStart('\')
                                    $destPath = Join-Path $Destination $relativePath
                                }

                                # Create destination directory if it doesn't exist
                                $destDir = Split-Path -Path $destPath -Parent
                                if (-not (Test-Path -Path $destDir)) {
                                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                                }

                                # Check if destination file exists and handle Force parameter
                                if (Test-Path -Path $destPath -PathType Leaf) {
                                    if (-not $Force) {
                                        Write-Warning "Destination file already exists: $destPath. Use -Force to overwrite."
                                        continue
                                    }
                                    Write-Verbose "Overwriting existing file: $destPath"
                                }

                                # Store verbose message
                                $verboseOutput += "Copied '$($file.Name)' to '$destPath'"

                                try {
                                    if ($useWin32Api) {
                                        $cancel = $false
                                        Write-Verbose "Using Windows API for file copy operations"
                                        
                                        # Create script-scope variables for the callback
                                        $script:currentFile = $file
                                        $script:filesCount = $files.Count
                                        $script:filesCopied = $filesCopied
                                        $script:totalBytesCopied = $totalBytesCopied
                                        $script:totalSize = $totalSize
                                        $script:progressId = $progressId

                                        $callback = {
                                            param(
                                                [long]$TotalFileSize,
                                                [long]$TotalBytesTransferred,
                                                [long]$StreamSize,
                                                [long]$StreamBytesTransferred,
                                                [uint32]$StreamNumber,
                                                [uint32]$CallbackReason,
                                                [IntPtr]$SourceFile,
                                                [IntPtr]$DestinationFile,
                                                [IntPtr]$Data
                                            )
                                            
                                            try {
                                                # Use API values directly
                                                $percent = [math]::Min([math]::Round(($TotalBytesTransferred * 100) / [math]::Max($TotalFileSize, 1), 0), 100)

                                                # Update speed once per second
                                                $now = [DateTime]::Now
                                                if (($now - $script:lastSpeedUpdate).TotalSeconds -ge 1) {
                                                    $timeDiff = ($now - $script:lastTime).TotalSeconds
                                                    
                                                    # Detect new file start (when bytes transferred is less than last bytes)
                                                    if ($TotalBytesTransferred -lt $script:lastBytes) {
                                                        $script:lastBytes = 0
                                                        $script:lastTime = $now
                                                        $script:currentSpeed = 0
                                                    } else {
                                                        $bytesDiff = $TotalBytesTransferred - $script:lastBytes
                                                        # Ensure we never report negative speeds
                                                        $script:currentSpeed = if ($timeDiff -gt 0) { [math]::Max(0, [math]::Round($bytesDiff / $timeDiff)) } else { 0 }
                                                    }
                                                    
                                                    $script:lastTime = $now
                                                    $script:lastBytes = $TotalBytesTransferred
                                                    $script:lastSpeedUpdate = $now
                                                }

                                                if ($script:filesCount -gt 1) {
                                                    # Overall progress
                                                    $totalPercent = [math]::Min([math]::Round((($script:totalBytesCopied + $TotalBytesTransferred) / $script:totalSize * 100), 0), 100)
                                                    
                                                    Write-Progress -Activity "Copying files ($totalPercent%)" `
                                                        -Status "$($script:filesCopied) of $($script:filesCount) files ($(Format-FileSize $totalBytesCopied) of $(Format-FileSize $script:totalSize)) - $(Format-FileSize $script:currentSpeed)/s" `
                                                        -PercentComplete $totalPercent `
                                                        -Id $script:progressId

                                                    # File progress
                                                    Write-Progress -Activity "Copying $($script:currentFile.Name) ($percent%)" `
                                                        -Status "$(Format-FileSize $TotalBytesTransferred) of $(Format-FileSize $TotalFileSize)" `
                                                        -PercentComplete $percent `
                                                        -ParentId $script:progressId `
                                                        -Id ($script:progressId + 1)
                                                } else {
                                                    Write-Progress -Activity "Copying $($script:currentFile.Name) ($percent%)" `
                                                        -Status "$(Format-FileSize $TotalBytesTransferred) of $(Format-FileSize $TotalFileSize) - $(Format-FileSize $script:currentSpeed)/s" `
                                                        -PercentComplete $percent `
                                                        -Id $script:progressId
                                                }
                                            }
                                            catch {
                                                Write-Warning "Progress callback error: $_"
                                                return [uint32]3  # PROGRESS_QUIET
                                            }
                                            
                                            return [uint32]0  # PROGRESS_CONTINUE
                                        }

                                        # Determine optimal copy flags
                                        $copyFlags = 0
                                        if ($file.Length -gt 10MB) {
                                            $copyFlags = $copyFlags -bor 0x00001000  # COPY_FILE_NO_BUFFERING
                                        }
                                        if ($destPath -like "\\*") {  # Network path
                                            $copyFlags = $copyFlags -bor 0x10000000  # COPY_FILE_REQUEST_COMPRESSED_TRAFFIC
                                        }

                                        $result = [Win32Helpers.Win32CopyFileEx]::CopyFileEx(
                                            $file.FullName,
                                            $destPath,
                                            $callback,
                                            [IntPtr]::Zero,
                                            [ref]$cancel,
                                            $copyFlags
                                        )

                                        if (-not $result) {
                                            $errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
                                            throw "CopyFileEx failed with error code: $errorCode"
                                        }

                                        $totalBytesCopied += $file.Length
                                    }
                                    else {
                                        try {
                                            Write-Verbose "Using managed method for file copy operations"
                                            $sourceStream = [System.IO.File]::OpenRead($file.FullName)
                                            $destStream = [System.IO.File]::Create($destPath)
                                            $bytesRead = 0
                                            $fileSize = [Math]::Max($file.Length, 1)
                                            $fileBytesCopied = 0
                                
                                            while (($bytesRead = $sourceStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                                                $destStream.Write($buffer, 0, $bytesRead)
                                                $fileBytesCopied += $bytesRead
                                                $totalBytesCopied += $bytesRead
                                
                                                # Update progress less frequently
                                                $now = [DateTime]::Now
                                                if (($now - $lastProgressUpdate) -gt $progressThreshold) {
                                                    $totalPercent = [math]::Min([math]::Round(($totalBytesCopied / $totalSize * 100), 0), 100)
                                                    $filePercent = [math]::Min([math]::Round(($fileBytesCopied / $fileSize * 100), 0), 100)
                                                    
                                                    # Calculate current speed
                                                    $currentSpeed = Get-CurrentSpeed -now $now -currentBytes $totalBytesCopied
                                                    $speedText = if ($currentSpeed -gt 0) {
                                                        "$(Format-FileSize $currentSpeed)/s"
                                                    } else {
                                                        "0 B/s"
                                                    }
                                
                                                    if ($files.Count -gt 1) {
                                                        # Overall progress for multiple files
                                                        Write-Progress -Activity "Copying files ($totalPercent%)" `
                                                            -Status "$filesCopied of $($files.Count) files ($(Format-FileSize $totalBytesCopied) of $(Format-FileSize $totalSize)) - $speedText" `
                                                            -PercentComplete $totalPercent `
                                                            -Id $progressId
                                
                                                        # File progress as child
                                                        Write-Progress -Activity "Copying $($file.Name) ($filePercent%)" `
                                                            -Status "$(Format-FileSize $fileBytesCopied) of $(Format-FileSize $fileSize)" `
                                                            -PercentComplete $filePercent `
                                                            -ParentId $progressId `
                                                            -Id $childProgressId
                                                    } else {
                                                        # Single file progress
                                                        Write-Progress -Activity "Copying $($file.Name) ($filePercent%)" `
                                                            -Status "$(Format-FileSize $fileBytesCopied) of $(Format-FileSize $fileSize) - $speedText" `
                                                            -PercentComplete $filePercent `
                                                            -Id $progressId
                                                    }
                                
                                                    $lastProgressUpdate = $now
                                                }
                                            }
                                        }
                                        finally {
                                            if ($sourceStream) { $sourceStream.Close() }
                                            if ($destStream) { $destStream.Close() }
                                        }
                                    }
                                }
                                catch {
                                    Write-Warning "Error copying $($file.Name): $_"
                                }
                            }

                            # Calculate total elapsed time
                            $endTime = [DateTime]::Now
                            $elapsedTime = $endTime - $startTime
                            $elapsedText = if ($elapsedTime.TotalHours -ge 1) {
                                "{0:h'h 'm'm 's's'}" -f $elapsedTime
                            } elseif ($elapsedTime.TotalMinutes -ge 1) {
                                "{0:m'm 's's'}" -f $elapsedTime
                            } else {
                                "{0:s's'}" -f $elapsedTime
                            }

                            if ($files.Count -gt 1) {
                                $verboseOutput += "Total copied: $(Format-FileSize $totalSize) ($($files.Count) files)"
                            } else {
                                $verboseOutput += "Total size: $(Format-FileSize $totalSize)"
                            }

                            $verboseOutput += "Operation completed in $elapsedText"

                            # Write all verbose messages at once after copying is complete
                            if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {
                                if ($files.Count -gt 1) {
                                    $verboseOutput | ForEach-Object { Write-Verbose $_ }
                                } else {
                                    $verboseOutput | ForEach-Object { Write-Output $_ }
                                }
                            }

                            # Complete progress bars
                            if ($files.Count -gt 1) {
                                Write-Progress -Activity "Copying files" -Id $childProgressId -Completed
                            }
                            Write-Progress -Activity "Copying files" -Id $progressId -Completed

                            # Return copied item if PassThru is specified
                            if ($PassThru) {
                                Get-Item -LiteralPath $targetPath
                            }
                        }
                    }
                }
            }
            catch {
                Write-Error -ErrorRecord $_
            }
        }
    }
}