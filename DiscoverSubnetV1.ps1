#requires -Version 5.1
###################################################################################################
##### DiscoverSubnet.ps1 Version 1.0. Copyright MediaLinks, Inc. 2025
##### DiscoverSubnet program discovers MediaLinks devices in the network and produces a text or 
##### semicolon seperated CSV file of the discovered devices that can be used by other programs.
##### Author: Gary Faubert
##### Last Modified: September 16, 2025
##### Note: this PowerShell script can be compiled using the Visual Studio Code command Win-PS2EXE
##### Note: In Win-PS2EXE select the source and uncheck all boxes.
##################################################################################################
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Determine the script's execution path for saving files
try {
    # This works when running as a .ps1 file
    $scriptPath = $PSScriptRoot
    if (-not $scriptPath) {
        # This is a fallback for compiled .exe or ISE execution
        $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
}
catch {
    # Final fallback to the current working directory
    $scriptPath = Get-Location
}

# --- GUI: Input Form ---
$form = New-Object Windows.Forms.Form
$form.Text = "Subnet Discovery Configuration"
$form.Size = New-Object Drawing.Size(500, 380)
$form.StartPosition = "CenterScreen"

# IP Ranges Input
$labelRanges = New-Object Windows.Forms.Label
$labelRanges.Text = "Enter IP ranges (e.g., 192.168.1.10-20, 192.168.1.0):"
$labelRanges.AutoSize = $true
$labelRanges.Location = New-Object Drawing.Point(10, 20)
$form.Controls.Add($labelRanges)

$textBoxRanges = New-Object Windows.Forms.TextBox
$textBoxRanges.Size = New-Object Drawing.Size(460, 20)
$textBoxRanges.Location = New-Object Drawing.Point(10, 40)
$form.Controls.Add($textBoxRanges)

# Community String Input
$labelCommunity = New-Object Windows.Forms.Label
$labelCommunity.Text = "SNMP Community String:"
$labelCommunity.AutoSize = $true
$labelCommunity.Location = New-Object Drawing.Point(10, 80)
$form.Controls.Add($labelCommunity)

$textBoxCommunity = New-Object Windows.Forms.TextBox
$textBoxCommunity.Text = "medialinks"
$textBoxCommunity.Size = New-Object Drawing.Size(200, 20)
$textBoxCommunity.Location = New-Object Drawing.Point(10, 100)
$form.Controls.Add($textBoxCommunity)

# Retry Count Input
$labelRetries = New-Object Windows.Forms.Label
$labelRetries.Text = "Ping/SNMP Retries:"
$labelRetries.AutoSize = $true
$labelRetries.Location = New-Object Drawing.Point(270, 80)
$form.Controls.Add($labelRetries)

$comboRetries = New-Object Windows.Forms.ComboBox
$comboRetries.Items.AddRange(@("0", "1", "2", "3"))
$comboRetries.SelectedIndex = 0
$comboRetries.DropDownStyle = "DropDownList"
$comboRetries.Size = New-Object Drawing.Size(200, 20)
$comboRetries.Location = New-Object Drawing.Point(270, 100)
$form.Controls.Add($comboRetries)

# Output File Name Input
$labelFileName = New-Object Windows.Forms.Label
$labelFileName.Text = "Output File Name:"
$labelFileName.AutoSize = $true
$labelFileName.Location = New-Object Drawing.Point(10, 140)
$form.Controls.Add($labelFileName)

$textBoxFileName = New-Object Windows.Forms.TextBox
$textBoxFileName.Text = "DiscoveredDevices"
$textBoxFileName.Size = New-Object Drawing.Size(200, 20)
$textBoxFileName.Location = New-Object Drawing.Point(10, 160)
$form.Controls.Add($textBoxFileName)

# Output File Type Input
$labelFileType = New-Object Windows.Forms.Label
$labelFileType.Text = "Output File Type:"
$labelFileType.AutoSize = $true
$labelFileType.Location = New-Object Drawing.Point(270, 140)
$form.Controls.Add($labelFileType)

$comboFileType = New-Object Windows.Forms.ComboBox
$comboFileType.Items.AddRange(@("txt", "csv"))
$comboFileType.SelectedIndex = 0
$comboFileType.DropDownStyle = "DropDownList"
$comboFileType.Size = New-Object Drawing.Size(200, 20)
$comboFileType.Location = New-Object Drawing.Point(270, 160)
$form.Controls.Add($comboFileType)

# Save Unknown Devices Checkbox
$checkSaveUnknown = New-Object Windows.Forms.CheckBox
$checkSaveUnknown.Text = "Save 'UNKNOWN' device types"
$checkSaveUnknown.AutoSize = $true
$checkSaveUnknown.Location = New-Object Drawing.Point(10, 210)
$form.Controls.Add($checkSaveUnknown)

# Start Button
$okButton = New-Object Windows.Forms.Button
$okButton.Text = "Start Discovery"
$okButton.Size = New-Object Drawing.Size(460, 40)
$okButton.Location = New-Object Drawing.Point(10, 260)
$okButton.DialogResult = [Windows.Forms.DialogResult]::OK
$form.Controls.Add($okButton)
$form.AcceptButton = $okButton

# Show the form and capture user input
if ($form.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) {
    # Exit if user closes the form without clicking "Start"
    return
}

# Capture values from the form
$UserInput = @{
    Ranges        = $textBoxRanges.Text
    Community     = $textBoxCommunity.Text
    Retries       = [int]$comboRetries.SelectedItem
    FileName      = $textBoxFileName.Text
    FileType      = $comboFileType.SelectedItem
    SaveUnknown   = $checkSaveUnknown.Checked
}

# --- GUI: Log Window ---
$logForm = New-Object Windows.Forms.Form
$logForm.Text = "Discovery Progress"
$logForm.Size = New-Object Drawing.Size(700, 500)
$logForm.StartPosition = "CenterScreen"

$logBox = New-Object Windows.Forms.RichTextBox
$logBox.ReadOnly = $true
$logBox.Size = New-Object Drawing.Size(660, 380)
$logBox.Location = New-Object Drawing.Point(10, 10)
$logForm.Controls.Add($logBox)

$progressBar = New-Object Windows.Forms.ProgressBar
$progressBar.Size = New-Object Drawing.Size(660, 20)
$progressBar.Location = New-Object Drawing.Point(10, 400)
$progressBar.Minimum = 0
$logForm.Controls.Add($progressBar)

$deviceMap = @{}
$logFileName = Join-Path $scriptPath "DiscoveryLog_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

$saveButton = New-Object Windows.Forms.Button
$saveButton.Text = "Save Results"
$saveButton.Size = New-Object Drawing.Size(100, 30)
$saveButton.Location = New-Object Drawing.Point(570, 430)
$saveButton.Add_Click({
    # --- START: Diagnostic Logging ---
    Log "--- Save Button Clicked: Starting Diagnostics ---" "Magenta"
    Log "1. Checking the state of the 'Save UNKNOWN' checkbox..." "Magenta"
    Log "   - Checkbox value is: $($UserInput.SaveUnknown)" "DarkCyan"

    Log "2. Reporting all devices stored in memory BEFORE filtering..." "Magenta"
    Log "   - Total devices in memory: $($deviceMap.Count)" "DarkCyan"
    if ($deviceMap.Count -gt 0) {
        $deviceMap.Keys | ForEach-Object { Log "     - '$_'" "Gray" }
    } else {
        Log "   - No devices were found or stored in memory." "Gray"
    }
    # --- END: Diagnostic Logging ---

    $outputFile = Join-Path $scriptPath "$($UserInput.FileName).$($UserInput.FileType)"
    $header = "Name;Location;Type;IP(s)"
    
    $results = $deviceMap.GetEnumerator()
    
    # Filtering logic with added diagnostics
    if (-not $UserInput.SaveUnknown) {
        Log "3. Filter Decision: Checkbox is OFF. Applying filter to remove UNKNOWN devices." "Magenta"
        $preFilterCount = if ($results) { ($results | Measure-Object).Count } else { 0 }
        
        $results = $results | Where-Object { $_.Key -notlike "*;UNKNOWN" }
        
        $postFilterCount = if ($results) { ($results | Measure-Object).Count } else { 0 }
        Log "   - Devices before filtering: $preFilterCount" "DarkCyan"
        Log "   - Devices after filtering:  $postFilterCount" "DarkCyan"
    } else {
        Log "3. Filter Decision: Checkbox is ON. Skipping the filter." "Magenta"
    }

    $header | Out-File $outputFile -Encoding utf8

    $results | ForEach-Object {
        "$($_.Key);$($_.Value)" | Out-File $outputFile -Append -Encoding utf8
    }
    
    Log "4. Save complete. Results written to file." "Magenta"
    [System.Windows.Forms.MessageBox]::Show("Results saved to $outputFile")
})

$logForm.Controls.Add($saveButton)

$closeButton = New-Object Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Size = New-Object Drawing.Size(100, 30)
$closeButton.Location = New-Object Drawing.Point(460, 430)
$closeButton.Add_Click({ $logForm.Close() })
$logForm.Controls.Add($closeButton)

# --- Functions ---
function Log {
    param ($text, $color)
    $timestamp = Get-Date -Format 'HH:mm:ss'
    # Write to log file
    "$timestamp - $text" | Out-File -FilePath $logFileName -Append -Encoding utf8
    
    # Write to GUI log box
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.SelectionColor = [System.Drawing.Color]::$color
    $logBox.AppendText("[$timestamp] $text`r`n")
    $logBox.ScrollToCaret()
    $logBox.Refresh()
}

function Parse-Ranges {
    param ($Ranges)
    $IPGroups = @()
    # Sanitize input: remove quotes and extra spaces
    $cleanRanges = $Ranges.Trim().Trim('"').Trim()
    
    if ([string]::IsNullOrWhiteSpace($cleanRanges)) {
        Log "Input is empty. Please provide IP ranges." "Red"
        return $null
    }
    
    $parts = $cleanRanges -split ',' | ForEach-Object { $_.Trim() }

    foreach ($part in $parts) {
        $currentGroup = @()
        try {
            # Case 1: Full subnet (e.g., 192.168.12.0)
            if ($part -match '^(\d{1,3}\.\d{1,3}\.\d{1,3}\.)0$') {
                $base = $matches[1]
                $currentGroup += (2..254 | ForEach-Object { "$base$_" })
            }
            # Case 2: IP Range (e.g., 192.168.12.21-24)
            elseif ($part -match '^(\d{1,3}\.\d{1,3}\.\d{1,3}\.)(\d{1,3})-(\d{1,3})$') {
                $base = $matches[1]
                $start = [int]$matches[2]
                $end = [int]$matches[3]
                if ($start -gt $end -or $start -lt 1 -or $end -gt 254) {
                    throw "Invalid range numbers in '$part'. Start must be <= End, and between 1-254."
                }
                $currentGroup += ($start..$end | ForEach-Object { "$base$_" })
            }
            # Case 3: Single IP (e.g., 192.168.12.49)
            elseif ($part -match '^(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})$') {
                # Basic validation for IP format
                if ($part -as [ipaddress]) {
                    $currentGroup += $part
                } else {
                    throw "Invalid IP address format in '$part'."
                }
            }
            # Case 4: Invalid format
            else {
                # Remove any non-standard characters before reporting the error
                $invalidPart = $part -replace '[^\d\.\-,]', ''
                throw "Invalid input format for '$invalidPart'. Please use formats like '1.2.3.4', '1.2.3.10-20', or '1.2.3.0'."
            }
            
            if ($currentGroup.Count -gt 0) {
                $IPGroups += ,$currentGroup # The comma ensures it's added as a nested array
            }
        }
        catch {
            Log "Error parsing input: $($_.Exception.Message)" "Red"
            return $null # Stop processing on first error
        }
    }
    return $IPGroups
}

function Get-DeviceType {
    param ($OID)
    # Normalize OID by removing prefixes, quotes, and whitespace
    $cleanOID = $OID -replace '^OID=', '' -replace '"', '' -replace '\s', ''

    $oidMap = @{
        ".iso.org.dod.internet.private.enterprises.17186.1.10"      = "1.3.6.1.4.1.17186.1.10"
        ".iso.org.dod.internet.private.enterprises.21839.1.2.17"    = "1.3.6.1.4.1.21839.1.2.17"
        ".iso.org.dod.internet.private.enterprises.17186.1.24"      = "1.3.6.1.4.1.17186.1.24"
        ".iso.org.dod.internet.private.enterprises.17186.3.1.1.1.0" = "1.3.6.1.4.1.17186.3.1.1.1.0"
    }

    if ($oidMap.ContainsKey($cleanOID)) {
        $cleanOID = $oidMap[$cleanOID]
    }

    switch ($cleanOID) {
        "1.3.6.1.4.1.17186.1.10"      { return "MD8000" }
        "1.3.6.1.4.1.21839.1.2.17"    { return "MDX2040" }
        "1.3.6.1.4.1.17186.1.24"      { return "MDP3020" }
        "1.3.6.1.4.1.17186.3.1.1.1.0" { return "MDX32C/48X6C" }
        default {
            Log "Unrecognized device type OID: '$cleanOID'" "DarkRed"
            return "UNKNOWN"
        }
    }
}

function Get-SNMP {
    param ($IP, $OID, $community, $retries)
    try {
        $SNMP = New-Object -ComObject olePrn.OleSNMP
        # The 3rd parameter is Retries, 4th is Timeout in ms
        $SNMP.open($IP, $community, $retries, 1000)
        $value = $SNMP.get($OID)
        $SNMP.Close()
        return $value
    } catch {
        Log "SNMP query failed for $IP OID $OID" "Red"
        return $null
    }
}

function Discover-Devices {
    param ($IPGroups, $community, $retries)
    
    $totalIPs = ($IPGroups | ForEach-Object { $_.Count }) | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    $progressBar.Maximum = $totalIPs
    $progressBar.Value = 0

    foreach ($group in $IPGroups) {
        $rangeDisplay = if ($group.Count -gt 1) { "$($group[0]) - $($group[-1])" } else { $group[0] }
        Log "Starting scan for range: $rangeDisplay" "Blue"

        foreach ($IP in $group) {
            $progressBar.Value += 1
            $progressBar.Refresh()

            $isReachable = $false
            for ($i = 0; $i -le $retries; $i++) {
                if (Test-Connection -ComputerName $IP -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                    $isReachable = $true
                    break 
                }
            }

            if ($isReachable) {
                Log "Discovering $IP..." "Green"

                $oidMap = @{
                    ".1.3.6.1.2.1.1.2.0" = Get-SNMP -IP $IP -OID ".1.3.6.1.2.1.1.2.0" -community $community -retries $retries
                    ".1.3.6.1.2.1.1.5.0" = Get-SNMP -IP $IP -OID ".1.3.6.1.2.1.1.5.0" -community $community -retries $retries
                    ".1.3.6.1.2.1.1.6.0" = Get-SNMP -IP $IP -OID ".1.3.6.1.2.1.1.6.0" -community $community -retries $retries
                }

                $typeOID = $oidMap[".1.3.6.1.2.1.1.2.0"]
                $name = $oidMap[".1.3.6.1.2.1.1.5.0"]
                $location = $oidMap[".1.3.6.1.2.1.1.6.0"]

                # --- CORRECTED LOGIC ---
                # No longer discard devices that fail all SNMP. Instead, create an entry with placeholders.
                if ([string]::IsNullOrWhiteSpace($name)) { $name = "[No Name Found]" }
                if ([string]::IsNullOrWhiteSpace($location)) { $location = "[No Location Found]" }
                
                # If the type OID is missing, the type is UNKNOWN.
                $type = if ($typeOID) { Get-DeviceType $typeOID } else { "UNKNOWN" }
                
                $key = "$name;$location;$type"
                if ($deviceMap.ContainsKey($key)) {
                    $deviceMap[$key] += ",$IP"
                } else {
                    $deviceMap[$key] = $IP
                }
                Log "Added/Updated device in memory: $key" "Purple"

            } else {
                Log "Address $IP is not reachable." "Gray"
            }
        }
        Log "Scan complete for range: $rangeDisplay" "Blue"
    }

    $progressBar.Value = $progressBar.Maximum
    $progressBar.Refresh()

    Log "`nDiscovery complete. Devices found:" "Cyan"
    $deviceMap.GetEnumerator() | ForEach-Object {
        Log "$($_.Key);$($_.Value)" "Black"
    }
}

# --- Main Execution ---
$logForm.Show()
Log "Starting Discovery Process. Log file at: $logFileName" "DarkBlue"

try {
    $IPGroupsList = Parse-Ranges $UserInput.Ranges
    if ($null -ne $IPGroupsList) {
        Discover-Devices -IPGroups $IPGroupsList -community $UserInput.Community -retries $UserInput.Retries
    } else {
        Log "Discovery halted due to invalid input." "Red"
    }
}
catch {
    Log "A critical error occurred: $_" "Red"
}

# Keep the log form open until the user closes it
[System.Windows.Forms.Application]::Run($logForm)
