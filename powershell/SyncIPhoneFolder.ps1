param (
    [Parameter(Mandatory=$true)]
    [string]$SourceFolderName, # e.g., 202506_a
    
    [string]$DestinationPath = "D:\dest"
)

Write-Host "--- Initializing iPhone Transfer Tool ---" -ForegroundColor Cyan
Write-Host "Target Folder: $SourceFolderName"
Write-Host "Destination:   $DestinationPath"

$shell = New-Object -ComObject Shell.Application
$thisPC = $shell.NameSpace(17) # "This PC" virtual folder

# 1. Prepare Destination Folder
if (-not (Test-Path $DestinationPath)) {
    try {
        New-Item -ItemType Directory -Path $DestinationPath -ErrorAction Stop | Out-Null
        Write-Host "[OK] Created destination: $DestinationPath"
    } catch {
        Write-Host "[!] FATAL ERROR: Could not create destination folder." -ForegroundColor Red
        return
    }
}
$destFolder = $shell.NameSpace($DestinationPath)

# 2. Find the iPhone
Write-Host "> Searching for iPhone..."
$iphone = $thisPC.Items() | Where-Object { $_.Name -match "iPhone" }
if (-not $iphone) { 
    Write-Host "[!] ERROR: iPhone not found. Please unlock and ensure it is plugged in." -ForegroundColor Red
    return 
}
Write-Host "[OK] Found: $($iphone.Name)"

# 3. Find Internal Storage
$storage = $iphone.GetFolder.Items() | Where-Object { $_.Name -match "Storage" -or $_.Name -match "Internal" }
if (-not $storage) {
    Write-Host "[!] ERROR: Internal Storage is invisible. Check 'Trust' settings on iPhone." -ForegroundColor Red
    return
}

# 4. Find the Specific Source Folder (e.g., 202506_a)
$source = $storage.GetFolder.Items() | Where-Object { $_.Name -eq $SourceFolderName }
if (-not $source) {
    Write-Host "[!] ERROR: Source folder '$SourceFolderName' not found in Internal Storage." -ForegroundColor Yellow
    Write-Host "Available folders are:"
    $storage.GetFolder.Items() | Select-Object Name
    return
}

# 5. Incremental Copy Logic
Write-Host "`n--- Starting Copy Operation ---" -ForegroundColor Green
$success = 0
$skipped = 0
$failed  = 0

$items = $source.GetFolder.Items()
Write-Host "Found $($items.Count) items in source."

foreach ($file in $items) {
    try {
        if ($file.IsFolder) { continue } # Only process files

        # Logic to ignore already copied files
        $targetFile = Join-Path $DestinationPath $file.Name
        if (Test-Path $targetFile) {
            Write-Host "Skipping: $($file.Name) (Exists)" -ForegroundColor Gray
            $skipped++
            continue 
        }

        # Perform Transfer
        Write-Host "Copying: $($file.Name)..." -NoNewline
        # Flag 16: "Yes to All" for prompts/overwrites
        $destFolder.CopyHere($file, 16)
        
        Write-Host " [DONE]" -ForegroundColor Green
        $success++
    }
    catch {
        Write-Host " [FAILED] - $($_.Exception.Message)" -ForegroundColor Red
        $failed++
    }
}

Write-Host "`n--- Transfer Summary ---" -ForegroundColor Cyan
Write-Host "Newly Copied: $success"
Write-Host "Skipped:      $skipped"
Write-Host "Failed:       $failed"
Write-Host "------------------------"
