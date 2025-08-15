#
# Enhanced Printer Migration Script v1.2
# 
# This script searches for printer installation commands from the Install_Renown_Printers.ps1
# and migrates them to a second PC.
#
# IMPORTANT NOTES:
# - Duplicate printers are filtered out, but may still appear from old PC configurations
# - Check for dated or offline printers - verify connectivity via File Explorer if needed  
# - Script copies printer configurations as-is, including any misconfigurations
# - Clear unnecessary printers to reduce EPIC compatibility issues
# - Review script output for troubleshooting information
#

#region Functions

function Test-ComputerConnectivity {
    <#
    .SYNOPSIS
    Tests computer connectivity using ping and traceroute validation
    
    .DESCRIPTION
    Tests if a computer is reachable via ping, then performs traceroute to verify
    the final destination matches the target computer name
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName           # Target computer name to test connectivity
    )
    
    Write-Host "`n=== Computer Status Check ===" -ForegroundColor Cyan
    Write-Host "Computer Name: $ComputerName" -ForegroundColor White
    Write-Host "Testing connectivity..." -ForegroundColor Yellow
    
    # Step 1: Basic connectivity test using Test-Connection
    try {
        $pingResults = Test-Connection -ComputerName $ComputerName -Count 4 -ErrorAction Stop
        
        Write-Host "Status: ONLINE" -ForegroundColor Green
        Write-Host "Packets Sent: 4" -ForegroundColor White
        Write-Host "Packets Received: $($pingResults.Count)" -ForegroundColor White
        
        # Calculate average response time
        $avgResponseTime = ($pingResults | Measure-Object -Property ResponseTime -Average).Average
        Write-Host "Average Response Time: $([math]::Round($avgResponseTime, 2)) ms" -ForegroundColor White
        
        # Step 2: Perform traceroute validation
        Write-Host "`nPerforming traceroute validation..." -ForegroundColor Yellow
        Test-TracerouteDestination -ComputerName $ComputerName
        
        return $true
    }
    catch {
        Write-Host "Status: OFFLINE or UNREACHABLE" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    finally {
        Write-Host "Check completed at: $(Get-Date)" -ForegroundColor Gray
    }
}

function Test-TracerouteDestination {
    <#
    .SYNOPSIS
    Performs traceroute and validates the final destination matches target computer
    
    .DESCRIPTION
    Uses tracert command to trace route to target computer and compares the final
    destination with the expected computer name
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName           # Target computer name to validate route to
    )
    
    try {
        Write-Host "Running traceroute to $ComputerName..." -ForegroundColor Gray
        
        # Execute tracert command and capture output
        $tracertOutput = & tracert -h 10 -w 2000 $ComputerName 2>&1 | Out-String
        
        # Parse traceroute output to find final destination
        $tracertLines = $tracertOutput -split "`n" | Where-Object { $_.Trim() -ne "" }
        $finalDestination = ""
        $routeFound = $false
        
        # Look through tracert output for the final successful hop
        foreach ($currentLine in $tracertLines) {
            # Skip header lines and error messages
            if ($currentLine -match "^\s*\d+" -and $currentLine -notmatch "\*\s*\*\s*\*") {
                # Extract the destination from lines with successful hops
                # Pattern matches: "  1    <1 ms    <1 ms    <1 ms  destination-name [ip-address]"
                if ($currentLine -match "\s+([^\[\s]+)(?:\s+\[[^\]]+\])?\s*$") {
                    $finalDestination = $matches[1].Trim()
                    $routeFound = $true
                }
                # Also handle lines that just show IP addresses
                elseif ($currentLine -match "\s+(\d+\.\d+\.\d+\.\d+)\s*$") {
                    $finalDestination = $matches[1].Trim()
                    $routeFound = $true
                }
            }
        }
        
        # Validate the final destination against expected computer name
        if ($routeFound -and -not [string]::IsNullOrWhiteSpace($finalDestination)) {
            Write-Host "Final destination: $finalDestination" -ForegroundColor White
            
            # Check if final destination matches target computer name (case-insensitive)
            if ($finalDestination -eq $ComputerName -or $finalDestination -like "*$ComputerName*") {
                Write-Host "Route validation: " -NoNewline -ForegroundColor White
                Write-Host "✓ MATCH" -ForegroundColor Green
                Write-Host "The traceroute destination matches the target computer." -ForegroundColor Green
            }
            else {
                Write-Host "Route validation: " -NoNewline -ForegroundColor White  
                Write-Host "⚠ WARNING" -ForegroundColor Yellow
                Write-Host "The traceroute destination ($finalDestination) does not match target ($ComputerName)." -ForegroundColor Yellow
                Write-Host "This might indicate DNS issues or network redirection." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "Route validation: " -NoNewline -ForegroundColor White
            Write-Host "⚠ UNABLE TO DETERMINE" -ForegroundColor Yellow
            Write-Host "Could not determine final destination from traceroute output." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Route validation: " -NoNewline -ForegroundColor White
        Write-Host "⚠ ERROR" -ForegroundColor Red
        Write-Host "Error during traceroute: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Search-PrinterFiles {
    <#
    .SYNOPSIS
    Recursively searches directories for .txt files containing printer commands
    
    .DESCRIPTION
    This function searches through folders and subfolders looking for .txt files
    that contain printer installation commands (/i or /id)
    #>
    param(
        [Parameter(Mandatory)]
        [string]$CurrentPath,        # The current directory being searched
        [Parameter(Mandatory)]
        [string]$BasePath,           # The original starting directory path
        [Parameter(Mandatory)]
        [ref]$AllResults,            # Reference to array collecting all found printer commands
        [Parameter(Mandatory)]
        [ref]$UniquePrinters,        # Reference to hashtable preventing duplicate printers
        [int]$Depth = 0              # Current folder depth for display indentation
    )
    
    # Create indentation for nested folder display
    $indent = "  " * $Depth
    
    try {
        Write-Host "$indent Searching: $CurrentPath" -ForegroundColor Gray
        
        # Get all items (files and folders) in current directory
        $allItems = Get-ChildItem -Path $CurrentPath -ErrorAction SilentlyContinue
        
        # Filter to only .txt files (not directories)
        $textFiles = $allItems | Where-Object { 
            $_.Extension -eq ".txt" -and -not $_.PSIsContainer 
        }
        
        # Process any .txt files found
        if ($textFiles.Count -gt 0) {
            Write-Host "$indent   Found $($textFiles.Count) .txt files" -ForegroundColor Green
            
            # Process each .txt file for printer commands
            foreach ($currentFile in $textFiles) {
                Write-Host "$indent   Processing: $($currentFile.Name)" -ForegroundColor Yellow
                Process-PrinterFile -File $currentFile -BasePath $BasePath -AllResults $AllResults -UniquePrinters $UniquePrinters -Indent $indent
            }
        }
        else {
            Write-Host "$indent   No .txt files found" -ForegroundColor DarkGray
        }
        
        # Get all subdirectories for recursive search
        $subDirectories = $allItems | Where-Object { $_.PSIsContainer }
        
        # Recursively search subdirectories
        if ($subDirectories.Count -gt 0) {
            Write-Host "$indent   Found $($subDirectories.Count) subdirectories" -ForegroundColor Cyan
            
            # Search each subdirectory
            foreach ($currentSubDir in $subDirectories) {
                Write-Host "$indent   Entering directory: $($currentSubDir.Name)" -ForegroundColor Cyan
                # Recursive call with increased depth
                Search-PrinterFiles -CurrentPath $currentSubDir.FullName -BasePath $BasePath -AllResults $AllResults -UniquePrinters $UniquePrinters -Depth ($Depth + 1)
            }
        }
        else {
            Write-Host "$indent   No subdirectories found" -ForegroundColor DarkGray
        }
    }
    catch {
        Write-Host "$indent ERROR accessing directory: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Process-PrinterFile {
    <#
    .SYNOPSIS
    Reads a .txt file and extracts printer installation commands
    
    .DESCRIPTION
    This function reads each line of a .txt file looking for printer installation
    commands that start with /i or /id followed by a printer path
    #>
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File,   # The .txt file to process
        [Parameter(Mandatory)]
        [string]$BasePath,           # Base directory path for relative path calculation
        [Parameter(Mandatory)]
        [ref]$AllResults,            # Reference to array collecting all results
        [Parameter(Mandatory)]
        [ref]$UniquePrinters,        # Reference to hashtable preventing duplicates
        [Parameter(Mandatory)]
        [string]$Indent              # Indentation string for display formatting
    )
    
    try {
        # Read all lines from the file
        $fileContent = Get-Content -Path $File.FullName -ErrorAction Stop
        $currentLineNumber = 0
        
        # Process each line in the file
        foreach ($currentLine in $fileContent) {
            $currentLineNumber++
            $cleanedLine = $currentLine.Trim()
            
            # Skip empty lines and comment lines
            if ([string]::IsNullOrWhiteSpace($cleanedLine) -or 
                $cleanedLine.StartsWith("REM") -or 
                $cleanedLine.StartsWith("#") -or
                $cleanedLine.StartsWith("===") -or 
                $cleanedLine.StartsWith("NOTE:")) {
                continue  # Skip to next line
            }
            
            # Look for printer installation commands using regex pattern
            # Pattern matches: /i or /id followed by space and \\server\printer path (case insensitive)
            if ($cleanedLine -imatch "^(/id?)\s+(\\\\[^\\]+\\[^\s]+)") {
                # Process the found printer command (using automatic $matches variable)
                Process-PrinterCommand -Line $currentLine -TrimmedLine $cleanedLine -LineNumber $currentLineNumber -File $File -BasePath $BasePath -AllResults $AllResults -UniquePrinters $UniquePrinters -Indent $Indent
            }
        }
    }
    catch {
        Write-Host "$Indent     ERROR reading file: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Process-PrinterCommand {
    <#
    .SYNOPSIS
    Processes a single printer installation command line
    
    .DESCRIPTION
    Takes a line containing /i or /id command and extracts printer information,
    checks for duplicates, and adds to results collection
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Line,                    # Original line from file (with spacing)
        [Parameter(Mandatory)]
        [string]$TrimmedLine,             # Cleaned/trimmed version of the line
        [Parameter(Mandatory)]
        [int]$LineNumber,                 # Line number in the file
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File,        # File object containing this command
        [Parameter(Mandatory)]
        [string]$BasePath,                # Base path for relative path calculation
        [Parameter(Mandatory)]
        [ref]$AllResults,                 # Reference to array collecting all results
        [Parameter(Mandatory)]
        [ref]$UniquePrinters,             # Reference to hashtable preventing duplicates
        [Parameter(Mandatory)]
        [string]$Indent                   # Display indentation string
    )
    
    # Re-run the regex match to populate $matches variable (case insensitive)
    if ($TrimmedLine -imatch "^(/id?)\s+(\\\\[^\\]+\\[^\s]+)") {
        # Extract command type (/i or /id) and printer path from regex matches
        $commandType = $matches[1]      # Either "/i" or "/id"
        $printerPath = $matches[2]      # Full UNC path like \\server\printer
        
        # Parse server and printer names from the UNC path
        $serverName = ""
        $printerName = ""
        
        # Use regex to split \\server\printer into parts
        if ($printerPath -match "\\\\([^\\]+)\\(.+)") {
            $serverName = $matches[1]       # Server name (after \\)
            $printerName = $matches[2]      # Printer name (after server\)
        }
        
        # Check if this printer path has already been found (prevent duplicates)
        if (-not $UniquePrinters.Value.ContainsKey($printerPath)) {
            # Create new printer result object
            $printerResult = [PSCustomObject]@{
                File = $File.FullName
                RelativePath = $File.FullName.Replace($BasePath, "")
                LineNumber = $LineNumber
                OriginalLine = $Line
                CleanCommand = $TrimmedLine
                PrinterPath = $printerPath
                PrintServer = $serverName
                PrinterName = $printerName
                CommandType = if ($commandType -eq "/id") { "Install and Set Default" } else { "Install Only" }
                IsCommented = $false
                DuplicateCount = 1
            }
            
            # Add to both collections
            $AllResults.Value += $printerResult
            $UniquePrinters.Value[$printerPath] = $printerResult
            
            # Display the found printer information
            Write-Host "$Indent     FOUND: Line $LineNumber" -ForegroundColor Green
            Write-Host "$Indent       Printer Path: $printerPath" -ForegroundColor White
            Write-Host "$Indent       Server: $serverName" -ForegroundColor White
            Write-Host "$Indent       Printer: $printerName" -ForegroundColor White
            Write-Host "$Indent       Type: $($printerResult.CommandType)" -ForegroundColor White
            Write-Host "$Indent       Command: $TrimmedLine" -ForegroundColor Gray
        }
        else {
            # Handle duplicate printer - increment duplicate count
            $UniquePrinters.Value[$printerPath].DuplicateCount++
            
            # Update if this one sets as default
            if ($commandType.ToLower() -eq "/id" -and $UniquePrinters.Value[$printerPath].CommandType -eq "Install Only") {
                $UniquePrinters.Value[$printerPath].CommandType = "Install and Set Default"
                Write-Host "$Indent     DUPLICATE ! (Updated to Default): $printerPath" -ForegroundColor Yellow
            }
            else {
                Write-Host "$Indent     DUPLICATE ! (Count: $($UniquePrinters.Value[$printerPath].DuplicateCount)): $printerPath" -ForegroundColor DarkYellow
            }
        }
    }
}

function Search-PrinterInstallCommands {
    <#
    .SYNOPSIS
    Main function to search for printer installation commands
    
    .DESCRIPTION
    Searches a specified directory and all subdirectories for .txt files containing
    printer installation commands (/i or /id) and returns unique printer results
    #>
    param(
        [Parameter(Mandatory)]
        [string]$SearchPath              # Root directory to start searching from
    )
    
    Write-Host "`nStarting comprehensive folder search..." -ForegroundColor Magenta
    Write-Host "Base path: $SearchPath" -ForegroundColor Cyan
    Write-Host "Looking for lines starting with '/i' or '/id' followed by printer paths..." -ForegroundColor Cyan
    
    # Verify the search path exists and is accessible
    if (-not (Test-Path $SearchPath)) {
        Write-Host "ERROR: Path does not exist or is not accessible: $SearchPath" -ForegroundColor Red
        return @()  # Return empty array
    }
    
    # Initialize collections to store results
    $allFoundResults = [ref]@()        # All printer commands found (including duplicates)
    $uniquePrinterList = [ref]@{}      # Hashtable to track unique printers by path
    
    # Start the recursive search through all folders
    Search-PrinterFiles -CurrentPath $SearchPath -BasePath $SearchPath -AllResults $allFoundResults -UniquePrinters $uniquePrinterList
    
    # Convert hashtable values to sorted array for final results
    $finalResults = $uniquePrinterList.Value.Values | Sort-Object PrinterPath
    
    # Display summary of search results
    Write-Host "`nSEARCH SUMMARY" -ForegroundColor Magenta
    Write-Host "Total printer commands found: $($allFoundResults.Value.Count)" -ForegroundColor White
    Write-Host "Unique printers found: $($finalResults.Count)" -ForegroundColor White
    
    # Count duplicates
    $duplicateCount = ($finalResults | Where-Object { $_.DuplicateCount -gt 1 }).Count
    if ($duplicateCount -gt 0) {
        Write-Host "Printers with duplicates: $duplicateCount (marked with !)" -ForegroundColor Yellow
    }
    
    # Display all unique printer commands found
    if ($finalResults.Count -gt 0) {
        Write-Host "`nUNIQUE PRINTER INSTALLATION COMMANDS:" -ForegroundColor Yellow
        
        foreach ($printerItem in $finalResults) {
            # Determine the correct command prefix
            $commandPrefix = if ($printerItem.CommandType -eq "Install and Set Default") { "/id" } else { "/i" }
            
            # Show duplicate indicator if more than one instance found
            $duplicateIndicator = if ($printerItem.DuplicateCount -gt 1) { " !" } else { "" }
            
            Write-Host "  $commandPrefix $($printerItem.PrinterPath)$duplicateIndicator" -ForegroundColor Green
        }
    }
    else {
        Write-Host "`nNo printer installation commands found." -ForegroundColor Yellow
        Write-Host "Searched for lines starting with '/i' or '/id' followed by UNC printer paths." -ForegroundColor Yellow
    }
    
    return $finalResults
}

function Start-PrinterMigration {
    param(
        [Parameter(Mandatory)]
        [array]$SearchResults
    )
    
    Write-Host "`n" + "=" * 80
    Write-Host "PRINTER MIGRATION TO SECOND PC" -ForegroundColor Magenta
    
    $secondPCName = Read-Host -Prompt "`nEnter the second PC name to copy printers to (or press Enter to skip migration)"
    
    if ([string]::IsNullOrWhiteSpace($secondPCName)) {
        Write-Host "Skipping migration..." -ForegroundColor Yellow
        return
    }
    
    Write-Host "`nTesting connectivity to second PC: $secondPCName" -ForegroundColor Cyan
    
    if (-not (Test-ComputerConnectivity -ComputerName $secondPCName)) {
        return
    }
    
    $secondPCPath = "\\$secondPCName\c$\ProgramData\Renown\Printers\"
    
    Write-Host "Testing path access: $secondPCPath" -ForegroundColor Cyan
    if (-not (Test-Path $secondPCPath)) {
        Write-Host "ERROR: Cannot access path on second PC: $secondPCPath" -ForegroundColor Red
        return
    }
    
    Write-Host "Path accessible: $secondPCPath" -ForegroundColor Green
    
    # Clean up directory
    Write-Host "`nCleaning up second PC printers directory..." -ForegroundColor Yellow
    try {
        $itemsToDelete = Get-ChildItem -Path $secondPCPath -ErrorAction Stop | Where-Object { $_.Name -ne "Install_Renown_Printers.ps1" }
        
        foreach ($item in $itemsToDelete) {
            Write-Host "  Deleting: $($item.Name)" -ForegroundColor Gray
            Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        Write-Host "Cleanup completed successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR during cleanup: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
    
    # Create printer configuration file
    Create-PrinterConfigFile -SearchResults $SearchResults -DestinationPath $secondPCPath
}

function Get-PrinterFileHeader {
    return @"
Printer queue repository file.
Enter each printer on a separate line.

Line Format:                      -Leading and trailing spaces ok, no quotes
\\PrintServer\PrinterName         -Skip
#/i \\PrintServer\PrinterName     -Install (omit #)
#/id \\PrintServer\PrinterName    -Install and make default (omit #)

NOTE: DO NOT use the # in the format: used here so the examples do not install.
===========================================================================================
Papercut Printer
"@
}

function Create-PrinterConfigFile {
    param(
        [Parameter(Mandatory)]
        [array]$SearchResults,
        [Parameter(Mandatory)]
        [string]$DestinationPath
    )
    
    Write-Host "`nCreating _SPPrinters.txt file..." -ForegroundColor Cyan
    
    $fileContent = Get-PrinterFileHeader
    
    foreach ($result in $SearchResults) {
        $command = if ($result.CommandType -eq "Install and Set Default") { "/id" } else { "/i" }
        $fileContent += "`n$command $($result.PrinterPath)"
    }
    
    $fileContent += "`n==========================================================================================="
    
    $newFilePath = Join-Path $DestinationPath "_SPPrinters.txt"
    
    try {
        $fileContent | Out-File -FilePath $newFilePath -Encoding UTF8 -ErrorAction Stop
        Write-Host "Successfully created: $newFilePath" -ForegroundColor Green
        
        Write-Host "`nPrinter commands added to _SPPrinters.txt:" -ForegroundColor Yellow
        foreach ($result in $SearchResults) {
            $command = if ($result.CommandType -eq "Install and Set Default") { "/id" } else { "/i" }
            $duplicateIndicator = if ($result.DuplicateCount -gt 1) { " !" } else { "" }
            Write-Host "  $command $($result.PrinterPath)$duplicateIndicator" -ForegroundColor White
        }
        
        Write-Host "`nMigration completed successfully!" -ForegroundColor Green
        Write-Host "Total unique printers migrated: $($SearchResults.Count)" -ForegroundColor White
    }
    catch {
        Write-Host "ERROR creating file: $($_.Exception.Message)" -ForegroundColor Red
    }
}

#endregion Functions

#region Main Script

Write-Host "Printer Migration Script v1.2" -ForegroundColor Magenta
Write-Host "===============================" -ForegroundColor Magenta

do {
    $computerName = Read-Host -Prompt "Enter the computer name (press Enter to exit)"
    
    if ([string]::IsNullOrWhiteSpace($computerName)) {
        Write-Host "Exiting script..." -ForegroundColor Yellow
        break
    }
    
    if (-not (Test-ComputerConnectivity -ComputerName $computerName)) {
        Write-Host "`nSkipping to next computer...`n" -ForegroundColor Yellow
        continue
    }
    
    $searchPath = "\\$computerName\c$\programdata\renown\printers\"
    $searchResults = Search-PrinterInstallCommands -SearchPath $searchPath
    
    if ($searchResults.Count -gt 0) {
        Start-PrinterMigration -SearchResults $searchResults
    }
    
    Write-Host "Processing complete for $computerName" -ForegroundColor Green
    
} while ($true)

Write-Host "Script execution completed." -ForegroundColor Magenta

#endregion Main Script
