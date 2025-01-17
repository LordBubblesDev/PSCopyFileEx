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
        Write-Debug @"
`n=== Copy-FileEx Debug Information ===
Parameter Set: $($PSCmdlet.ParameterSetName)
Path: $($Path -join ', ')
LiteralPath: $($LiteralPath -join ', ')
Destination: $Destination
Include: $($Include -join ', ')
Exclude: $($Exclude -join ', ')
Recurse: $Recurse
Force: $Force
UseWinApi: $UseWinApi
=================================
"@

        # Generate a random progress ID to avoid conflicts
        $progressId = Get-Random -Minimum 0 -Maximum 1000
        $childProgressId = $progressId + 1

        # Initialize cancellation support and register CTRL+C handler
        $script:cancelRequested = $false
        $null = [Console]::TreatControlCAsInput = $true
        
        # Function to check for CTRL+C
        function Test-CancellationRequested {
            if ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($true)
                if ($key.Key -eq 'C' -and $key.Modifiers -eq 'Control') {
                    Write-Warning "Cancellation requested by user"
                    $script:cancelRequested = $true
                    return $true
                }
            }
            return $false
        }

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
        # Add cancellation check at the start of process block
        if ($script:cancelCheckJob.State -eq 'Completed' -and $script:cancelCheckJob.Output) {
            $script:cancelRequested = $true
            Write-Warning "Operation cancelled by user"
            return
        }

        Write-Debug "Process block started"
        # Handle both Path and LiteralPath parameters
        $pathsToProcess = @()
        if ($LiteralPath) {
            Write-Debug "Using LiteralPath: $($LiteralPath -join ', ')"
            $pathsToProcess += $LiteralPath
            $useWildcards = $false
        } else {
            Write-Debug "Using Path: $($Path -join ', ')"
            $pathsToProcess += $Path
            $useWildcards = $true
        }

        foreach ($currentPath in $pathsToProcess) {
            Write-Debug "Processing path: $currentPath"
            # Process paths based on whether they're literal or support wildcards
            try {
                Write-Debug "Testing path existence"
                # Handle path resolution differently for files with special characters
                if (Test-Path -LiteralPath $currentPath) {
                    Write-Debug "Path exists (LiteralPath): $currentPath"
                    $resolvedPaths = @([pscustomobject]@{
                        Path = $currentPath
                        ProviderPath = (Get-Item -LiteralPath $currentPath).FullName
                    })
                    Write-Debug "Resolved to: $($resolvedPaths.ProviderPath)"
                } else {
                    Write-Debug "Path does not exist directly, attempting resolution"
                    if ($useWildcards) {
                        Write-Debug "Resolving with wildcards"
                        $resolvedPaths = Resolve-Path -Path $currentPath -ErrorAction Stop
                    } else {
                        Write-Debug "Resolving without wildcards"
                        $resolvedPaths = Resolve-Path -LiteralPath $currentPath -ErrorAction Stop
                    }
                    Write-Debug "Resolved paths count: $($resolvedPaths.Count)"
                }

                foreach ($resolvedPath in $resolvedPaths) {
                    Write-Debug "Processing resolved path: $($resolvedPath.Path)"
                    
                    # Check if the current path is a directory
                    $isDirectory = (Get-Item -LiteralPath $resolvedPath.Path) -is [System.IO.DirectoryInfo]
                    Write-Debug "Is Directory: $isDirectory"
                    
                    # Only apply Include/Exclude filters to files, not directories
                    $shouldProcess = $true
                    if (-not $isDirectory) {
                        if ($Include) {
                            Write-Debug "Checking Include filters: $($Include -join ', ')"
                            $shouldProcess = $resolvedPath.Path | Where-Object { 
                                $item = $_
                                $matchResult = ($Include | ForEach-Object { $item -like $_ }) -contains $true
                                Write-Debug "Include match result for $item : $matchResult"
                                return $matchResult
                            }
                        }
                        if ($Exclude -and $shouldProcess) {
                            Write-Debug "Checking Exclude filters: $($Exclude -join ', ')"
                            $shouldProcess = $resolvedPath.Path | Where-Object { 
                                $item = $_
                                $matchResult = ($Exclude | ForEach-Object { $item -like $_ }) -notcontains $true
                                Write-Debug "Exclude match result for $item : $matchResult"
                                return $matchResult
                            }
                        }
                    }

                    Write-Debug "Should process path: $shouldProcess"
                    if ($shouldProcess) {
                        # Check if we should process this item
                        $targetPath = Join-Path $Destination (Split-Path -Leaf $resolvedPath.ProviderPath)
                        Write-Debug "Target path: $targetPath"
                        if ($Force -or $PSCmdlet.ShouldProcess($targetPath)) {
                            # Handle wildcards in path
                            $sourcePath = Split-Path -Path $currentPath -Parent
                            $sourceFilter = Split-Path -Path $currentPath -Leaf

                            try {
                                # Initialize variables
                                $isFile = $false
                                $relativePath = $null
                                
                                if ($sourceFilter.Contains('*')) {
                                    Write-Debug "Path contains wildcards: $sourceFilter"
                                    # Path contains wildcards
                                    Write-Debug "Getting child items with filter: $sourceFilter"
                                    $files = Get-ChildItem -Path $sourcePath -Filter $sourceFilter -File -Recurse:$Recurse -ErrorAction Stop
                                    Write-Debug "Found $($files.Count) files matching filter"
                                    $isFile = $false
                                    $basePath = $sourcePath
                                } else {
                                    Write-Debug "Path is direct: $currentPath"
                                    # Single file or directory
                                    $item = Get-Item -LiteralPath $currentPath -ErrorAction Stop
                                    $isFile = $item -is [System.IO.FileInfo]
                                    Write-Debug "Item is file: $isFile"
                                    if ($isFile) {
                                        Write-Debug "Single file: $($item.Name)"
                                        $files = @($item)
                                        $basePath = Split-Path -Path $item.FullName -Parent
                                        # For single files, use the file's directory as base path
                                        $relativePath = $item.Name
                                    } else {
                                        Write-Debug "Directory: $($item.FullName)"
                                        # For directories, apply Include/Exclude filters to Get-ChildItem
                                        $gciParams = @{
                                            Path = $currentPath
                                            File = $true
                                            Recurse = $Recurse
                                            ErrorAction = 'Stop'
                                        }
                                        if ($Include) {
                                            $gciParams['Include'] = $Include
                                            Write-Debug "Adding Include filter to Get-ChildItem: $($Include -join ', ')"
                                        }
                                        if ($Exclude) {
                                            $gciParams['Exclude'] = $Exclude
                                            Write-Debug "Adding Exclude filter to Get-ChildItem: $($Exclude -join ', ')"
                                        }
                                        Write-Debug "Getting child items with parameters: $($gciParams | ConvertTo-Json)"
                                        $files = Get-ChildItem @gciParams
                                        Write-Debug "Found $($files.Count) files in directory"
                                        $basePath = $item.FullName
                                    }
                                }
                                Write-Debug "Base Path: $basePath"
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
                                                # Check for cancellation
                                                if ($script:cancelRequested) {
                                                    Write-Verbose "Cancellation detected in callback"
                                                    return [uint32]1  # PROGRESS_CANCEL
                                                }

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
                                            if ($script:cancelRequested) {
                                                Write-Warning "File copy cancelled by user"
                                                break
                                            }
                                            throw "CopyFileEx failed with error code: $errorCode"
                                        }

                                        $totalBytesCopied += $file.Length
                                    }
                                    else {
                                        $continueProcessing = $true
                                        $sourceStream = $null
                                        $destStream = $null
                                        
                                        try {
                                            Write-Verbose "Using managed method for file copy operations"
                                            $sourceStream = [System.IO.FileStream]::new($file.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
                                            $destStream = [System.IO.FileStream]::new($destPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
                                            $bytesRead = 0
                                            $fileSize = [Math]::Max($file.Length, 1)
                                            $fileBytesCopied = 0
                                
                                            :copyLoop while ($continueProcessing -and -not $script:cancelRequested -and ($bytesRead = $sourceStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                                                # Check for cancellation before writing
                                                if (Test-CancellationRequested) {
                                                    Write-Warning "Cancelling file copy operation..."
                                                    $continueProcessing = $false
                                                    break copyLoop
                                                }

                                                try {
                                                    $destStream.Write($buffer, 0, $bytesRead)
                                                }
                                                catch {
                                                    Write-Warning "Error writing to destination: $_"
                                                    $continueProcessing = $false
                                                    break copyLoop
                                                }
                                                
                                                $fileBytesCopied += $bytesRead
                                                $totalBytesCopied += $bytesRead
                                
                                                # Update progress less frequently
                                                $now = [DateTime]::Now
                                                if (($now - $lastProgressUpdate) -gt $progressThreshold) {
                                                    if ($script:cancelRequested) {
                                                        $continueProcessing = $false
                                                        break copyLoop
                                                    }

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
                                        catch {
                                            Write-Warning "Error during file copy: $_"
                                            $continueProcessing = $false
                                        }
                                        finally {
                                            # Ensure streams are closed and disposed immediately
                                            if ($sourceStream) {
                                                try {
                                                    $sourceStream.Close()
                                                    $sourceStream.Dispose()
                                                }
                                                catch { }
                                                $sourceStream = $null
                                            }
                                            if ($destStream) {
                                                try {
                                                    $destStream.Close()
                                                    $destStream.Dispose()
                                                }
                                                catch { }
                                                $destStream = $null
                                            }

                                            # Clean up partial file if cancelled or failed
                                            if (-not $continueProcessing -or $script:cancelRequested) {
                                                Write-Verbose "Cleaning up partial file after cancellation/failure: $destPath"
                                                try {
                                                    [System.GC]::Collect()  # Force garbage collection
                                                    [System.GC]::WaitForPendingFinalizers()  # Wait for finalizers
                                                    
                                                    if (Test-Path -Path $destPath) {
                                                        # Try to open the file exclusively to ensure it's not locked
                                                        $fileInfo = [System.IO.FileInfo]::new($destPath)
                                                        try {
                                                            $stream = $fileInfo.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
                                                            $stream.Close()
                                                            $stream.Dispose()
                                                            
                                                            # Now we can safely delete the file
                                                            Remove-Item -Path $destPath -Force -ErrorAction Stop
                                                            Write-Verbose "Successfully removed partial file: $destPath"
                                                        }
                                                        catch {
                                                            Write-Warning "File is still locked, waiting for handles to be released..."
                                                            Start-Sleep -Milliseconds 500
                                                            Remove-Item -Path $destPath -Force -ErrorAction Stop
                                                        }
                                                    }
                                                }
                                                catch {
                                                    Write-Warning "Failed to remove partial file $destPath : $_"
                                                }
                                            }
                                        }
                                        
                                        # Exit immediately if cancelled
                                        if ($script:cancelRequested) {
                                            return
                                        }
                                    }
                                }
                                catch {
                                    Write-Warning "Error copying $($file.Name): $_"
                                    if ($script:cancelRequested) {
                                        break
                                    }
                                }

                                # Break the file loop if cancellation was requested
                                if ($script:cancelRequested) {
                                    break
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

                            # Only show completion messages if not cancelled
                            if (-not $script:cancelRequested) {
                                if ($files.Count -gt 1) {
                                    $verboseOutput += "Total copied: $(Format-FileSize $totalSize) ($($files.Count) files)"
                                } else {
                                    $verboseOutput += "Total size: $(Format-FileSize $totalSize)"
                                }

                                $verboseOutput += "Operation completed in $elapsedText"
                            }

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

    end {
        Write-Debug "End block started"
        # Restore console input handling
        [Console]::TreatControlCAsInput = $false
        Write-Debug "Function completed"
    }
}