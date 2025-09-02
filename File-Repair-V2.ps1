#================================================================================
#                  Advanced Multi-Phase File Repair Script (V3)
#================================================================================
#
# V3 Update: Complete logic overhaul for improved accuracy and performance.
#            - Sanitizes filenames before processing to handle special characters.
#            - Repairs extensions, fixes common PNG warnings, and tests archives.
#            - Deletes password-protected archives automatically.
#            - Performs a final validation scan to only log truly unrecoverable files.
#
# DEPENDENCIES: exiftool.exe, 7z.exe, optipng.exe (must be in system PATH)
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
Write-Host "--- Advanced Multi-Phase File Repair Script V3 ---" -ForegroundColor Yellow
Write-Host "Using a throttle limit of $throttleLimit processes."
if (Test-Path $repairedLog) { Remove-Item $repairedLog }
if (Test-Path $unrecoverableLog) { Remove-Item $unrecoverableLog }
"Repair log started at $startTime" | Out-File -FilePath $repairedLog -Encoding utf8
"Unrecoverable file log started at $startTime" | Out-File -FilePath $unrecoverableLog -Encoding utf8

# --- Dependency Check ---
$dependencies = "exiftool", "7z", "optipng"
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

# --- PHASE 1: Sanitize File and Directory Names ---
Write-Host "`n--- PHASE 1: Sanitizing File and Directory Names ---" -ForegroundColor Cyan
$sanitizedCount = 0
# We must get all items first, because renaming can interfere with the Get-ChildItem pipeline.
# Sanitize directories first, from deepest to shallowest, to avoid breaking paths.
$allDirs = Get-ChildItem -Path $PSScriptRoot -Recurse -Directory | Sort-Object { $_.FullName.Length } -Descending
foreach ($dir in $allDirs) {
    $originalName = $dir.Name
    # Regex to find any character that is NOT a letter, number, dot, hyphen, or underscore.
    $sanitizedName = $originalName -replace '[^a-zA-Z0-9._-]', '_'

    if ($originalName -ne $sanitizedName) {
        $newDirPath = Join-Path -Path $dir.Parent.FullName -ChildPath $sanitizedName
        if (Test-Path $newDirPath) {
            $logEntry = "CONFLICT (Dir): Could not rename '$($dir.FullName)' to '$sanitizedName' because a directory with that name already exists."
            Write-Host $logEntry -ForegroundColor Yellow
            $logEntry | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
        } else {
            try {
                Rename-Item -Path $dir.FullName -NewName $sanitizedName -ErrorAction Stop
                $logEntry = "SANITIZED (Dir): Renamed '$originalName' -> '$sanitizedName' in '$($dir.Parent.FullName)'"
                Write-Host $logEntry -ForegroundColor Green
                $logEntry | Out-File -FilePath $repairedLog -Encoding utf8 -Append
                $sanitizedCount++
            } catch {
                $logEntry = "ERROR (Dir): Failed to rename '$($dir.FullName)'. Details: $($_.Exception.Message)"
                Write-Host $logEntry -ForegroundColor Red
                $logEntry | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
            }
        }
    }
}

# Now, sanitize filenames. We re-fetch all files after directory renames have occurred.
$allFiles = Get-ChildItem -Path $PSScriptRoot -Recurse -File
foreach ($file in $allFiles) {
    $originalName = $file.Name
    $sanitizedName = $originalName -replace '[^a-zA-Z0-9._-]', '_'

    if ($originalName -ne $sanitizedName) {
        $newFilePath = Join-Path -Path $file.DirectoryName -ChildPath $sanitizedName
        if (Test-Path $newFilePath) {
            $logEntry = "CONFLICT (File): Could not rename '$($file.FullName)' to '$sanitizedName' because a file with that name already exists."
            Write-Host $logEntry -ForegroundColor Yellow
            $logEntry | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
        } else {
            try {
                Rename-Item -Path $file.FullName -NewName $sanitizedName -ErrorAction Stop
                $logEntry = "SANITIZED (File): Renamed '$($file.FullName)' -> '$newFilePath'"
                Write-Host $logEntry -ForegroundColor Green
                $logEntry | Out-File -FilePath $repairedLog -Encoding utf8 -Append
                $sanitizedCount++
            } catch {
                $logEntry = "ERROR (File): Failed to rename '$($file.FullName)'. Details: $($_.Exception.Message)"
                Write-Host $logEntry -ForegroundColor Red
                $logEntry | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
            }
        }
    }
}
Write-Host "Phase 1 Complete. Sanitized $sanitizedCount file and directory names."


# The file list is now updated after sanitization and available for subsequent phases.
$allFiles = Get-ChildItem -Path $PSScriptRoot -Recurse -File

# --- PHASE 2: File Extension Repair ---
Write-Host "`n--- PHASE 2: Repairing File Extensions ---" -ForegroundColor Cyan
$repairedCount = 0

# Step 1: Get true file types for all files in a single batch operation for performance.
Write-Host "Phase 2, Step 1: Batch-processing all files to identify necessary renames..."
$fileDataJson = exiftool -charset filename=UTF8 -json -r -SourceFile -FileTypeExtension "$PSScriptRoot"
$fileData = $fileDataJson | ConvertFrom-Json

# Step 2: Build the rename plan from the in-memory data. This is much faster than running exiftool per file.
Write-Host "Phase 2, Step 2: Building rename plan from batch results..."
$renamePlan = foreach ($file in $fileData) {
    # It's possible for exiftool to return no FileTypeExtension for some files (e.g. unsupported formats, directories).
    if (-not $file.PSObject.Properties.Contains('FileTypeExtension')) { continue }

    $currentPath = $file.SourceFile
    $currentExtWithDot = [System.IO.Path]::GetExtension($currentPath).ToLower()

    # Exclude archives from extension repair.
    if ($archiveExtensions -contains $currentExtWithDot) { continue }

    $currentExt = $currentExtWithDot.TrimStart('.')
    $trueExt = $file.FileTypeExtension.ToLower()

    if (-not [string]::IsNullOrEmpty($trueExt) -and $trueExt -ne $currentExt) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($currentPath)
        $newFileName = "$baseName.$trueExt"
        $newFilePath = Join-Path -Path ([System.IO.Path]::GetDirectoryName($currentPath)) -ChildPath $newFileName

        [pscustomobject]@{
            OldFullName = $currentPath
            NewFilePath = $newFilePath
            NewFileName = $newFileName
        }
    }
}

# Step 3: Resolve conflicts before renaming.
Write-Host "Phase 2, Step 3: Resolving renaming conflicts..."
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
Write-Host "Phase 2, Step 3: Performing safe renames..."
$totalToRename = $goodRenames.Count
$processedRenames = 0
if ($totalToRename -gt 0) {
    $renameResults = $goodRenames | ForEach-Object -ThrottleLimit $throttleLimit -Parallel {
        $rename = $_
        $currentCount = [System.Threading.Interlocked]::Increment([ref]$using:processedRenames)
        Write-Progress -Activity "Phase 2: Renaming Files" -Id 1 -Status "$currentCount / $using:totalToRename : $($rename.OldFullName)" -PercentComplete (($currentCount / $using:totalToRename) * 100)

        try {
            Rename-Item -Path $rename.OldFullName -NewName $rename.NewFileName -ErrorAction Stop
            [pscustomobject]@{ Success = $true; Old = $rename.OldFullName; New = $rename.NewFileName }
        } catch {
            [pscustomobject]@{ Success = $false; Old = $rename.OldFullName; Error = $_.Exception.Message }
        }
    }
    Write-Progress -Activity "Phase 2: Renaming Files" -Id 1 -Completed

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
Write-Host "Phase 2 Complete. Repaired $repairedCount file extensions."

# --- PHASE 3: PNG Structure Repair ---
Write-Host "`n--- PHASE 3: Repairing PNG File Structure ---" -ForegroundColor Cyan
$repairedPngCount = 0

# Use exiftool to find all PNGs that have the specific minor warning we can fix.
# The warning is "Text/EXIF chunk(s) found after PNG IDAT"
$pngsToFixPaths = exiftool -charset filename=UTF8 -r -if '($FileType eq "PNG") and ($Warning=~/Text\/EXIF chunk/)' -p '$Directory/$FileName' "$PSScriptRoot"
# The output can be a single string with newlines, so we split it into an array.
$pngsToFix = $pngsToFixPaths -split [System.Environment]::NewLine | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
$totalPngsToFix = ($pngsToFix | Measure-Object).Count

if ($totalPngsToFix -gt 0) {
    Write-Host "Found $totalPngsToFix PNGs with minor structure warnings to repair."

    $processedPngs = 0
    $pngRepairResults = $pngsToFix | ForEach-Object -ThrottleLimit $throttleLimit -Parallel {
        $pngPath = $_
        $currentCount = [System.Threading.Interlocked]::Increment([ref]$using:processedPngs)
        Write-Progress -Activity "Phase 3: Repairing PNGs" -Id 3 -Status "$currentCount / $using:totalPngsToFix : $pngPath" -PercentComplete (($currentCount / $using:totalPngsToFix) * 100)

        try {
            # optipng is very quiet on success. -fix will repair the structure.
            optipng -fix -o2 -quiet "$pngPath"
            # We assume success if optipng doesn't throw an error.
            [pscustomobject]@{ Success = $true; Path = $pngPath }
        } catch {
            [pscustomobject]@{ Success = $false; Path = $pngPath; Error = $_.Exception.Message }
        }
    }
    Write-Progress -Activity "Phase 3: Repairing PNGs" -Id 3 -Completed

    # Log results sequentially
    foreach ($result in $pngRepairResults) {
        if ($result.Success) {
            $repairedPngCount++
            $logEntry = "REPAIRED (STRUCTURE): Fixed minor chunk ordering in '$($result.Path)'"
            Write-Host $logEntry -ForegroundColor Green
            $logEntry | Out-File -FilePath $repairedLog -Encoding utf8 -Append
        } else {
            $logEntry = "ERROR (PNG Repair): optipng failed for '$($result.Path)'. Details: $($result.Error)"
            Write-Host $logEntry -ForegroundColor Red
            $logEntry | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
        }
    }
}
Write-Host "Phase 3 Complete. Repaired structure of $repairedPngCount PNG files."

# --- PHASE 4: Archive Integrity Testing & Deletion ---
Write-Host "`n--- PHASE 4: Testing Archive Integrity & Deleting Passworded Files ---" -ForegroundColor Cyan
$allArchives = Get-ChildItem -Path $PSScriptRoot -Recurse -File | Where-Object { $archiveExtensions -contains $_.Extension.ToLower() }
$totalArchives = ($allArchives | Measure-Object).Count
$processedArchives = 0
$corruptArchives = 0
$deletedArchives = 0

# Test archives in parallel.
$phase4Results = $allArchives | ForEach-Object -ThrottleLimit $throttleLimit -Parallel {
    $archive = $_
    $currentCount = [System.Threading.Interlocked]::Increment([ref]$using:processedArchives)
    Write-Progress -Activity "Phase 4: Testing Archives" -Id 4 -Status "$currentCount / $using:totalArchives : $($archive.Name)" -PercentComplete (($currentCount / $using:totalArchives) * 100)

    # We add '-p-' to provide an empty password, preventing 7z from halting the script with a prompt.
    $testResult = 7z t -scsUTF-8 -p- "$($archive.FullName)" 2>&1
    if ($LASTEXITCODE -ne 0) {
        if ($testResult -match "Wrong password") {
             [pscustomobject]@{ Status = 'Passworded'; FullName = $archive.FullName; TestResult = $testResult }
        } else {
             [pscustomobject]@{ Status = 'Corrupt'; FullName = $archive.FullName; TestResult = $testResult }
        }
    } else {
        [pscustomobject]@{ Status = 'OK' }
    }
}
Write-Progress -Activity "Phase 4: Testing Archives" -Id 4 -Completed

# Process results sequentially for safe logging and deletion.
foreach ($result in $phase4Results) {
    if ($result.Status -eq 'Passworded') {
        $deletedArchives++
        $logEntry = "DELETING (Passworded): '$($result.FullName)' appears to be password-protected."
        Write-Host $logEntry -ForegroundColor Magenta
        $logEntry | Out-File -FilePath $repairedLog -Encoding utf8 -Append
        try {
            Remove-Item -Path $result.FullName -Force -ErrorAction Stop
            Write-Host "  -> Successfully deleted." -ForegroundColor Magenta
        } catch {
            $deleteError = "ERROR (Deletion): Failed to delete '$($result.FullName)'. Details: $($_.Exception.Message)"
            Write-Host $deleteError -ForegroundColor Red
            $deleteError | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
        }
    } elseif ($result.Status -eq 'Corrupt') {
        $corruptArchives++
        $logEntry = "CORRUPT ARCHIVE: '$($result.FullName)' failed integrity test."
        Write-Host $logEntry -ForegroundColor Red
        $logEntry | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
        $result.TestResult | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
    }
}
Write-Host "Phase 4 Complete. Found $corruptArchives corrupt archives and deleted $deletedArchives password-protected archives."

# --- PHASE 5: Final Corruption Validation ---
Write-Host "`n--- PHASE 5: Final Deep Scan for Unrecoverable Images ---" -ForegroundColor Cyan
$unrecoverableImageCount = 0

# After all repairs, we do a final scan. We now only care about fatal errors, not warnings
# that we have already attempted to fix. exiftool is run once in batch mode for performance.
Write-Host "Scanning for images with fatal errors..."
# Note: The -args parameter is used to pass the list of extensions to the -ext option of exiftool.
$errorFilesOutput = exiftool -check -fast -r -if '$Error' -p '$Directory/$FileName' -args -ext $imageExtensions "$PSScriptRoot"
$errorFiles = $errorFilesOutput -split [System.Environment]::NewLine | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
$totalErrorFiles = ($errorFiles | Measure-Object).Count

if ($totalErrorFiles -gt 0) {
    "--- Final Scan: Unrecoverable Images Detected ---" | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
    foreach ($file in $errorFiles) {
        $unrecoverableImageCount++
        $logEntry = "UNRECOVERABLE IMAGE: '$file'"
        Write-Host $logEntry -ForegroundColor Red
        $logEntry | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
        # Get the specific error for the log
        $errorDetails = exiftool -check -fast "$file"
        $errorDetails | Out-File -FilePath $unrecoverableLog -Encoding utf8 -Append
    }
}
Write-Host "Phase 5 Complete. Found $unrecoverableImageCount images with fatal, unrecoverable errors."


# --- PHASE 6: Final Reporting ---
Write-Host "`n--- PHASE 6: Final Report ---" -ForegroundColor Yellow
$endTime = Get-Date
$duration = New-TimeSpan -Start $startTime -End $endTime

Write-Host "`nRepair process finished in $($duration.TotalSeconds) seconds."
Write-Host "--- Summary ---" -ForegroundColor Yellow
Write-Host "Phase 1: Sanitized $sanitizedCount file and directory names." -ForegroundColor Green
Write-Host "Phase 2: Repaired $repairedCount file extensions." -ForegroundColor Green
Write-Host "Phase 3: Repaired structure of $repairedPngCount PNG files." -ForegroundColor Green
Write-Host "Phase 4: Deleted $deletedArchives password-protected archives." -ForegroundColor Magenta
Write-Host "Phase 4: Found $corruptArchives corrupt archives." -ForegroundColor Red
Write-Host "Phase 5: Found $unrecoverableImageCount images with fatal errors." -ForegroundColor Red
Write-Host "`nPlease review the log files for details:"
Write-Host " - Repaired Actions & Deletions: $repairedLog"
Write-Host " - Unrecoverable/Problematic Files: $unrecoverableLog"

Read-Host "`nProcess complete. Press Enter to exit."
