#
#
#
# 1.1
# This script runs a modified printer search from the Install_Renown_Printers.ps1.
# Issues to watch out for while using this script:
# I removed DUPLICATE printers, but they may still show up when pulling printers from the old PC.
# Dated or offline printers: double-check if the printer script was able to add them or if you can connect to them using File Explorer.
# This just copies the printers from one PC to another, so if there's a misconfiguration on the original PC, expect it on the new PC too.
# Make sure to clear any unnecessary printers to reduce issues with EPIC.
# Read the information the scripts provides for troubleshooting please.
#
#
#



do {
    # Prompt user for computer name input
    $computerName = Read-Host -Prompt "Enter the computer name (press Enter to exit)"
    
    # Check if user wants to exit
    if ([string]::IsNullOrWhiteSpace($computerName)) {
        Write-Host "Exiting script..." -ForegroundColor Yellow
        break
    }

    Write-Host "`n=== Computer Status Check ===" -ForegroundColor Cyan
    Write-Host "Computer Name: $computerName" -ForegroundColor White

    # Perform ping test
    Write-Host "Testing connectivity..." -ForegroundColor Yellow


	# Need to change to to test route, I think there was a issue with that command before
    $pingResult = ping -n 4 -w 1000 $computerName 2>&1 | Out-String
    if ($pingResult -match "Reply from") {
        Write-Host "Status: ONLINE" -ForegroundColor Green
        # Extract packet statistics
        if ($pingResult -match "Packets: Sent = (\d+), Received = (\d+)") {
            Write-Host "Packets Sent: $($matches[1])" -ForegroundColor White
            Write-Host "Packets Received: $($matches[2])" -ForegroundColor White
        }
        if ($pingResult -match "Average = (\d+)ms") {
            Write-Host "Average Response Time: $($matches[1]) ms" -ForegroundColor White
        }
    } else {
        Write-Host "Status: OFFLINE or UNREACHABLE" -ForegroundColor Red
        Write-Host "`nSkipping to next computer...`n" -ForegroundColor Yellow
        continue
    }

    Write-Host "Check completed at: $(Get-Date)" -ForegroundColor Gray

    # Define the search path
    $searchPath = "\\$computerName\c$\programdata\renown\printers\"

    # Function to recursively search all folders
    function Search-AllFolders {
        param(
            [string]$CurrentPath,
            [string]$BasePath,
            [ref]$AllResults,
            [ref]$UniquePrinters,
            [int]$Depth = 0
        )
        
        $indent = "  " * $Depth
        
        try {
            Write-Host "$indent Searching: $CurrentPath" -ForegroundColor Gray
            
            # Get all items in current directory
            $items = Get-ChildItem -Path $CurrentPath -ErrorAction SilentlyContinue
            
            # Process .txt files in current directory
            $txtFiles = $items | Where-Object { $_.Extension -eq ".txt" -and -not $_.PSIsContainer }
            
            if ($txtFiles.Count -gt 0) {
                Write-Host "$indent   Found $($txtFiles.Count) .txt files" -ForegroundColor Green
                
                foreach ($file in $txtFiles) {
                    Write-Host "$indent   Processing: $($file.Name)" -ForegroundColor Yellow
                    
                    try {
                        $content = Get-Content -Path $file.FullName -ErrorAction Stop
                        $lineNumber = 0
                        
                        foreach ($line in $content) {
                            $lineNumber++
                            $trimmedLine = $line.Trim()
                            
                            # Skip empty lines and comment lines that start with REM
                            if ($trimmedLine -eq "" -or $trimmedLine.StartsWith("REM") -or $trimmedLine.StartsWith("===") -or $trimmedLine.StartsWith("NOTE:")) {
                                continue
                            }
                            
                            # Check for /i or /id commands (the actual install commands)
                            if ($trimmedLine -match "^(/id?)\s+(\\\\[^\\]+\\[^\s]+)") {
                                $command = $matches[1]
                                $printerPath = $matches[2]
                                
                                # Extract server and printer name from the path
                                $serverName = ""
                                $printerName = ""
                                if ($printerPath -match "\\\\([^\\]+)\\(.+)") {
                                    $serverName = $matches[1]
                                    $printerName = $matches[2]
                                }
                                
                                # Check if this printer has already been found
                                if (-not $UniquePrinters.Value.ContainsKey($printerPath)) {
                                    $result = [PSCustomObject]@{
                                        File = $file.FullName
                                        RelativePath = $file.FullName.Replace($BasePath, "")
                                        LineNumber = $lineNumber
                                        OriginalLine = $line
                                        CleanCommand = $trimmedLine
                                        PrinterPath = $printerPath
                                        PrintServer = $serverName
                                        PrinterName = $printerName
                                        CommandType = if ($command -eq "/id") { "Install and Set Default" } else { "Install Only" }
                                        IsCommented = $false
                                    }
                                    
                                    $AllResults.Value += $result
                                    $UniquePrinters.Value[$printerPath] = $result
                                    
                                    # Display immediate results
                                    Write-Host "$indent     FOUND: Line $lineNumber" -ForegroundColor Green
                                    Write-Host "$indent       Printer Path: $printerPath" -ForegroundColor White
                                    Write-Host "$indent       Server: $serverName" -ForegroundColor White
                                    Write-Host "$indent       Printer: $printerName" -ForegroundColor White
                                    Write-Host "$indent       Type: $($result.CommandType)" -ForegroundColor White
                                    Write-Host "$indent       Command: $trimmedLine" -ForegroundColor Gray
                                } else {
                                    # Update if this is a default command and previous wasn't
                                    if ($command -eq "/id" -and $UniquePrinters.Value[$printerPath].CommandType -eq "Install Only") {
                                        $UniquePrinters.Value[$printerPath].CommandType = "Install and Set Default"
                                        Write-Host "$indent     DUPLICATE (Updated to Default): $printerPath" -ForegroundColor Yellow
                                    } else {
                                        Write-Host "$indent     DUPLICATE (Skipped): $printerPath" -ForegroundColor DarkYellow
                                    }
                                }
                            }
                        }
                    } catch {
                        Write-Host "$indent     ERROR reading file: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            } else {
                Write-Host "$indent   No .txt files found" -ForegroundColor DarkGray
            }
            
            # Get all subdirectories and recursively search them
            $subDirs = $items | Where-Object { $_.PSIsContainer }
            
            if ($subDirs.Count -gt 0) {
                Write-Host "$indent   Found $($subDirs.Count) subdirectories" -ForegroundColor Cyan
                
                foreach ($subDir in $subDirs) {
                    Write-Host "$indent   Entering directory: $($subDir.Name)" -ForegroundColor Cyan
                    Search-AllFolders -CurrentPath $subDir.FullName -BasePath $BasePath -AllResults $AllResults -UniquePrinters $UniquePrinters -Depth ($Depth + 1)
                }
            } else {
                Write-Host "$indent   No subdirectories found" -ForegroundColor DarkGray
            }
            
        } catch {
            Write-Host "$indent ERROR accessing directory: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Main search function
    function Search-PrinterInstallCommands {
        param(
            [string]$SearchPath
        )
        
        $results = @()
        $uniquePrinters = @{}
        
        Write-Host "`nStarting comprehensive folder-by-folder search..." -ForegroundColor Magenta
        Write-Host "Base path: $SearchPath" -ForegroundColor Cyan
        Write-Host "Looking for lines starting with '/i' or '/id' followed by printer paths..." -ForegroundColor Cyan
        
        # Test if path exists
        if (-not (Test-Path $SearchPath)) {
            Write-Host "ERROR: Path does not exist or is not accessible: $SearchPath" -ForegroundColor Red
            return @()
        }
        
        # Use reference variables to collect results from recursive function
        $allResults = [ref]@()
        $allUniquePrinters = [ref]@{}
        
        # Start recursive search
        Search-AllFolders -CurrentPath $SearchPath -BasePath $SearchPath -AllResults $allResults -UniquePrinters $allUniquePrinters
        
        # Get unique results
        $results = $allUniquePrinters.Value.Values | Sort-Object PrinterPath
        
        # Final Summary
        Write-Host "`nCOMPREHENSIVE SEARCH SUMMARY" -ForegroundColor Magenta
        Write-Host "Total printer installation commands found: $($allResults.Value.Count)" -ForegroundColor White
        Write-Host "Unique printers found: $($results.Count)" -ForegroundColor White
        
        # Output all unique printer names found
        if ($results.Count -gt 0) {
            Write-Host "`nUNIQUE PRINTER INSTALLATION COMMANDS FOUND:" -ForegroundColor Yellow
            foreach ($printer in $results) {
                $command = if ($printer.CommandType -eq "Install and Set Default") { "/id" } else { "/i" }
                Write-Host "  $command $($printer.PrinterPath)" -ForegroundColor Green
            }
        } else {
            Write-Host "`nNo printer installation commands found." -ForegroundColor Yellow
            Write-Host "Searched for lines starting with '/i' or '/id' followed by printer paths." -ForegroundColor Yellow
        }
        
        return $results
    }

    # Execute the search
    $searchResults = Search-PrinterInstallCommands -SearchPath $searchPath

    # Copy printers to second PC if results found
    if ($searchResults.Count -gt 0) {
        Write-Host "`n" + "=" * 80
        Write-Host "PRINTER MIGRATION TO SECOND PC" -ForegroundColor Magenta
        
        # Prompt for second PC name
        $secondPCName = Read-Host -Prompt "`nEnter the second PC name to copy printers to (or press Enter to skip migration)"
        
        if (-not [string]::IsNullOrWhiteSpace($secondPCName)) {
            # Test connectivity to second PC
            Write-Host "`nTesting connectivity to second PC: $secondPCName" -ForegroundColor Cyan
            
            $secondPingResult = ping -n 2 -w 1000 $secondPCName 2>&1 | Out-String
            if ($secondPingResult -match "Reply from") {
                Write-Host "Second PC Status: ONLINE" -ForegroundColor Green
                
                # Define second PC printer path
                $secondPCPath = "\\$secondPCName\c$\ProgramData\Renown\Printers\"
                
                # Test if path exists on second PC
                Write-Host "Testing path access: $secondPCPath" -ForegroundColor Cyan
                if (Test-Path $secondPCPath) {
                    Write-Host "Path accessible: $secondPCPath" -ForegroundColor Green
                    
                    # Get all items in the second PC printers directory
                    Write-Host "`nCleaning up second PC printers directory..." -ForegroundColor Yellow
                    try {
                        $itemsToDelete = Get-ChildItem -Path $secondPCPath -ErrorAction Stop | Where-Object { $_.Name -ne "Install_Renown_Printers.ps1" }
                        
                        foreach ($item in $itemsToDelete) {
                            Write-Host "  Deleting: $($item.Name)" -ForegroundColor Gray
                            Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                        }
                        
                        Write-Host "Cleanup completed successfully." -ForegroundColor Green
                        
                        # Create the _SPPrinters.txt file
                        Write-Host "`nCreating _SPPrinters.txt file..." -ForegroundColor Cyan
                        
                        # Build the file content
                        $fileContent = @"
Printer queue repository file.
Enter each printer on a seperate line line.

Line Format:                      -Leading and trailing spaces ok, no quotes
\\PrintServer\PrinterName         -Skip
#/i \\PrintServer\PrinterName     -Install (omit #)
#/id \\PrintServer\PrinterName    -Install and make default (omit #)

NOTE: DO NOT use the # in the format: used here so the examples do not install.
===========================================================================================
Papercut Printer
"@
                        
                        # Add each unique printer found
                        foreach ($result in $searchResults) {
                            if ($result.CommandType -eq "Install and Set Default") {
                                $fileContent += "`n/id $($result.PrinterPath)"
                            } else {
                                $fileContent += "`n/i $($result.PrinterPath)"
                            }
                        }
                        
                        $fileContent += "`n==========================================================================================="
                        
                        # Write the file to the second PC
                        $newFilePath = Join-Path $secondPCPath "_SPPrinters.txt"
                        
                        try {
                            $fileContent | Out-File -FilePath $newFilePath -Encoding UTF8 -ErrorAction Stop
                            Write-Host "Successfully created: $newFilePath" -ForegroundColor Green
                            
                            # Display what was written
                            Write-Host "`nPrinter commands added to _SPPrinters.txt:" -ForegroundColor Yellow
                            foreach ($result in $searchResults) {
                                $command = if ($result.CommandType -eq "Install and Set Default") { "/id" } else { "/i" }
                                Write-Host "  $command $($result.PrinterPath)" -ForegroundColor White
                            }
                            
                            Write-Host "`nMigration completed successfully!" -ForegroundColor Green
                            Write-Host "Total unique printers migrated: $($searchResults.Count)" -ForegroundColor White
                            
                        } catch {
                            Write-Host "ERROR creating file: $($_.Exception.Message)" -ForegroundColor Red
                        }
                        
                    } catch {
                        Write-Host "ERROR during cleanup: $($_.Exception.Message)" -ForegroundColor Red
                    }
                } else {
                    Write-Host "ERROR: Cannot access path on second PC: $secondPCPath" -ForegroundColor Red
                }
                
            } else {
                Write-Host "Second PC Status: OFFLINE or UNREACHABLE" -ForegroundColor Red
            }
        } else {
            Write-Host "Skipping migration..." -ForegroundColor Yellow
        }
    }
    
    Write-Host "Processing complete for $computerName" -ForegroundColor Green

} while ($true)

Write-Host "Script execution completed." -ForegroundColor Magenta
