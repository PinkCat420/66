#================================================================================
#                         Multi-Phase File Repair Script (V2)
#================================================================================
#
# V2 Update: Includes a new "Phase 0" to perform a deep validation scan on
#            images to detect internal data corruption before attempting repairs.
#
# DEPENDENCIES: exiftool.exe, 7z.exe (must be in system PATH or script folder)
#
#================================================================================

# --- Configuration ---
# V2.1 - Added parallel processing. Set how many files to process at once.
# Defaults to the number of logical processors on your system.
$throttleLimit = [System.Environment]::ProcessorCount
$ErrorActionPreference = "SilentlyContinue"
$startTime = Get-Date
$logPath = $PSScriptRoot # Log files will be saved in the same directory as the script.
$repairedLog = Join-Path $logPath "Repaired_Files.log"
$unrecoverableLog = Join-Path $logPath "Unrecoverable_Files.log"
$archiveExtensions = @('.zip', '.rar', '.7z')
$imageExtensions = @('.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp', '.tif', '.tiff')

# --- PowerShell Version Check ---
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "FATAL ERROR: This script requires PowerShell 7 or newer for parallel processing." -ForegroundColor Red
    Write-Host "Please upgrade your PowerShell version to run this script." -ForegroundColor Red
    Read-Host "Press Enter to exit."
    exit
}

# --- Initial Setup ---
Write-Host "--- Multi-Phase File Repair Script V2.1 (Parallel) ---" -ForegroundColor Yellow
Write-Host "Using a throttle limit of $throttleLimit processes."
if (Test-Path $repairedLog) { Remove-Item $repairedLog }
if (Test-Path $unrecoverableLog) { Remove-Item $unrecoverableLog }
"Repair log started at $startTime" | Out-File -FilePath $repairedLog -Encoding utf8
"Unrecoverable file log started at $startTime" | Out-File -FilePath $unrecoverableLog -Encoding utf8

# --- Dependency Check ---
$dependencies = "exiftool", "7z"
foreach ($dep in $dependencies) {
    if (-not (Get-Command $dep -ErrorAction SilentlyContinue)) {
        Write-Host "FATAL ERROR: Dependency '$dep.exe' not found." -ForegroundColor Red
        Write-Host "Please download it and ensure it is in your system PATH or the script's folder." -ForegroundColor Red
        Read-Host "Press Enter to exit."
        exit
    }
}
Write-Host "Dependencies found. Ready to proceed." -ForegroundColor Green
Write-Host "WARNING: This script will RENAME files in place. It is highly recommended to run this on a BACKUP of your data." -ForegroundColor Red
Read-Host "Press Enter to begin the repair process or Ctrl+C to abort..."

# --- Get Initial File List ---
$allFiles = Get-ChildItem -Path $PSScriptRoot -Recurse -File

# --- PHASE 0: Deep Image Corruption Scan ---
Write-Host "`n--- PHASE 0: Deep Scanning Images for Corruption ---" -ForegroundColor Cyan
# Use a thread-safe collection for the exclusion list.
$corruptFilesList = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
$imagesToScan = $allFiles | Where-Object { $imageExtensions -contains $_.Extension.ToLower() }
$totalImages = ($imagesToScan | Measure-Object).Count
$processedImages = 0
$corruptImageCount = 0

# Process images in parallel.
$phase0Results = $imagesToScan | ForEach-Object -ThrottleLimit $throttleLimit -Parallel {
    $image = $_
    $currentCount = [System.Threading.Interlocked]::Increment([ref]$using:processedImages)
    Write-Progress -Activity "Phase 0: Deep Scanning Images" -Id 0 -Status "$currentCount / $using:totalImages : $($image.Name)" -PercentComplete (($currentCount / $using:totalImages) * 100)

    # Use exiftool's built-in validation.
    $validationResult = exiftool -check -fast "$($image.FullName)" 2>&1

    if (-not [string]::IsNullOrWhiteSpace($validationResult)) {
        # Return a result object for logging later.
        [pscustomobject]@{
            Corrupt      = $true
            FullName     = $image.FullName
            Validation   = $validationResult
        }
    } else {
        [pscustomobject]@{ Corrupt = $false }
    }
}

Write-Progress -Activity "Phase 0: Deep Scanning Images" -Id 0 -Completed

# Now, process the results sequentially for safe logging.
$corruptImages = $phase0Results | Where-Object { $_.Corrupt }
foreach ($result in $corruptImages) {
    $corruptImageCount++
    $logEntry = "CORRUPT IMAGE DETECTED: '$($result.FullName)'"
    Write-Host $logEntry -ForegroundColor Red
    $logEntry | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
    $result.Validation | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
    $corruptFilesList.Add($result.FullName)
}
Write-Host "Phase 0 Complete. Found $corruptImageCount potentially corrupt images."

# --- PHASE 1: File Extension Repair ---
Write-Host "`n--- PHASE 1: Repairing File Extensions ---" -ForegroundColor Cyan
$repairedCount = 0

# Step 1: Identify all potential renames in parallel.
Write-Host "Phase 1, Step 1: Identifying necessary renames..."
$processedCount = 0
$filesToProcess = $allFiles | Where-Object {
    -not $corruptFilesList.Contains($_.FullName) -and
    -not ($archiveExtensions -contains $_.Extension.ToLower())
}
$totalFiles = ($filesToProcess | Measure-Object).Count

$renamePlan = $filesToProcess | ForEach-Object -ThrottleLimit $throttleLimit -Parallel {
    $file = $_
    $currentCount = [System.Threading.Interlocked]::Increment([ref]$using:processedCount)
    Write-Progress -Activity "Phase 1: Identifying Renames" -Id 1 -Status "$currentCount / $using:totalFiles : $($file.Name)" -PercentComplete (($currentCount / $using:totalFiles) * 100)

    $trueExtension = (exiftool -s3 -FileTypeExtension $file.FullName).Trim().ToLower()
    $currentExtension = $file.Extension.TrimStart('.').ToLower()

    if (-not [string]::IsNullOrEmpty($trueExtension) -and $trueExtension -ne $currentExtension) {
        # Improved logic for getting the base name
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $newFileName = "$baseName.$trueExtension"
        $newFilePath = Join-Path -Path $file.DirectoryName -ChildPath $newFileName

        [pscustomobject]@{
            OldFullName = $file.FullName
            NewFilePath = $newFilePath
            NewFileName = $newFileName
        }
    }
}
Write-Progress -Activity "Phase 1: Identifying Renames" -Id 1 -Completed

# Step 2: Resolve conflicts before renaming.
Write-Host "Phase 1, Step 2: Resolving renaming conflicts..."
$renameGroups = $renamePlan | Where-Object { $_ } | Group-Object NewFilePath
$goodRenames = [System.Collections.Generic.List[object]]::new()

foreach ($group in $renameGroups) {
    if ($group.Count -gt 1) {
        # Conflict: Multiple files would be renamed to the same new file.
        $sources = $group.Group.OldFullName -join "', '"
        $logEntry = "CONFLICT: Multiple files would be renamed to '$($group.Name)': '$sources'"
        Write-Host $logEntry -ForegroundColor Yellow
        $logEntry | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
    } elseif (Test-Path $group.Name) {
        # Conflict: A file with the new name already exists.
        $logEntry = "CONFLICT: Could not rename '$($group.Group[0].OldFullName)' to '$($group.Group[0].NewFileName)' because a file with that name already exists."
        Write-Host $logEntry -ForegroundColor Yellow
        $logEntry | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
    } else {
        # No conflicts found for this rename.
        $goodRenames.Add($group.Group[0])
    }
}

# Step 3: Perform the actual renames in parallel.
Write-Host "Phase 1, Step 3: Performing safe renames..."
$totalToRename = $goodRenames.Count
$processedRenames = 0
if ($totalToRename -gt 0) {
    $renameResults = $goodRenames | ForEach-Object -ThrottleLimit $throttleLimit -Parallel {
        $rename = $_
        $currentCount = [System.Threading.Interlocked]::Increment([ref]$using:processedRenames)
        Write-Progress -Activity "Phase 1: Renaming Files" -Id 1 -Status "$currentCount / $using:totalToRename : $($rename.OldFullName)" -PercentComplete (($currentCount / $using:totalToRename) * 100)

        try {
            Rename-Item -Path $rename.OldFullName -NewName $rename.NewFileName -ErrorAction Stop
            [pscustomobject]@{ Success = $true; Old = $rename.OldFullName; New = $rename.NewFileName }
        } catch {
            [pscustomobject]@{ Success = $false; Old = $rename.OldFullName; Error = $_.Exception.Message }
        }
    }
    Write-Progress -Activity "Phase 1: Renaming Files" -Id 1 -Completed

    # Step 4: Log the results of the rename operations.
    foreach ($result in $renameResults) {
        if ($result.Success) {
            $repairedCount++
            $logEntry = "REPAIRED: Renamed '$($result.Old)' -> '$($result.New)'"
            Write-Host $logEntry -ForegroundColor Green
            $logEntry | Out-File -FilePath $repairedLog -Encoding utf8 -Append
        } else {
            $logEntry = "ERROR: Failed to rename '$($result.Old)'. Details: $($result.Error)"
            Write-Host $logEntry -ForegroundColor Red
            $logEntry | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
        }
    }
}
Write-Host "Phase 1 Complete. Repaired $repairedCount file extensions."

# --- PHASE 2: Archive Integrity Testing ---
Write-Host "`n--- PHASE 2: Testing Archive Integrity ---" -ForegroundColor Cyan
$allArchives = Get-ChildItem -Path $PSScriptRoot -Recurse -File | Where-Object { $archiveExtensions -contains $_.Extension.ToLower() }
$totalArchives = ($allArchives | Measure-Object).Count
$processedArchives = 0
$corruptArchives = 0

# Test archives in parallel.
$phase2Results = $allArchives | ForEach-Object -ThrottleLimit $throttleLimit -Parallel {
    $archive = $_
    $currentCount = [System.Threading.Interlocked]::Increment([ref]$using:processedArchives)
    Write-Progress -Activity "Phase 2: Testing Archives" -Id 2 -Status "$currentCount / $using:totalArchives : $($archive.Name)" -PercentComplete (($currentCount / $using:totalArchives) * 100)

    $testResult = 7z t -scsUTF-8 "$($archive.FullName)" 2>&1
    if ($LASTEXITCODE -ne 0) {
        [pscustomobject]@{
            Corrupt     = $true
            FullName    = $archive.FullName
            TestResult  = $testResult
        }
    } else {
        [pscustomobject]@{ Corrupt = $false }
    }
}
Write-Progress -Activity "Phase 2: Testing Archives" -Id 2 -Completed

# Process results sequentially for safe logging.
$corruptArchiveResults = $phase2Results | Where-Object { $_.Corrupt }
foreach ($result in $corruptArchiveResults) {
    $corruptArchives++
    $logEntry = "CORRUPT ARCHIVE: '$($result.FullName)' failed integrity test."
    Write-Host $logEntry -ForegroundColor Red
    $logEntry | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
    $result.TestResult | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
}
Write-Host "Phase 2 Complete. Found $corruptArchives potentially corrupt archives."

# --- PHASE 3: Final Reporting ---
# ... (Phase 3 code is unchanged) ...
Write-Host "`n--- PHASE 3: Final Report ---" -ForegroundColor Yellow
$endTime = Get-Date
$duration = New-TimeSpan -Start $startTime -End $endTime

Write-Host "Repair process finished in $($duration.TotalSeconds) seconds."
Write-Host "Identified $corruptImageCount corrupt images in Phase 0." -ForegroundColor Red
Write-Host "Repaired $repairedCount file extensions in Phase 1." -ForegroundColor Green
Write-Host "Identified $corruptArchives corrupt archives in Phase 2." -ForegroundColor Red
Write-Host "Please review the log files for details:"
Write-Host " - Repaired Actions: $repairedLog"
Write-Host " - Unrecoverable/Problematic Files: $unrecoverableLog"

Read-Host "`nProcess complete. Press Enter to exit."
