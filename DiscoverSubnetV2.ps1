#requires -Version 5.1
<#
.SYNOPSIS
    DiscoverSubnet is a network discovery tool to identify MediaLinks hardware on a network.
.DESCRIPTION
    This PowerShell script provides a graphical user interface to define IP ranges and other
    scan parameters. It performs network discovery in a background job to keep the UI responsive,
    using parallel jobs to scan IPs efficiently. It identifies devices using SNMP queries and
    produces a summary report in the GUI, a log file, and a user-specified output file (.txt or .csv).
.VERSION
    2.3
.AUTHOR
    Gary Faubert - Gemini
.DATE
    2025-09-25
#>

#region Global Variables & Initial Setup
$scriptVersion = "2.3"
# This robustly determines the script's execution directory for both .ps1 and .exe files.
if ($PSScriptRoot) {
    # This variable is reliably populated when running as a .ps1 file.
    $scriptDir = $PSScriptRoot
}
else {
    # This is the fallback for compiled .exe files where $PSScriptRoot is not available.
    # $MyInvocation.MyCommand.Path will contain the full path to the .exe file.
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}

$settingsFilePath = Join-Path -Path $scriptDir -ChildPath "DiscoverSubnet.settings.json"

# Define the default settings structure. This is used if no settings file is found.
$defaultSettings = [PSCustomObject]@{
    IpRanges              = "192.168.1.0, 10.0.0.10-20"
    SnmpCommunity         = "medialinks"
    Retries               = 0
    OutputFileName        = "DiscoveredDevices"
    OutputFileExtension   = "csv"
    SaveUnresponsive      = $false
    MaxParallelScans      = 20
    DiagnosticLevel       = "Standard"
}
#endregion

#region Core Helper Functions

# Safely add required .NET assemblies for GUI components.
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}
catch {
    Write-Error "Failed to load required .NET Assemblies for GUI. Please ensure you are running in a Windows environment with .NET Framework."
    exit 1
}

function Load-Settings {
    if (Test-Path $settingsFilePath) {
        try { return Get-Content -Path $settingsFilePath | ConvertFrom-Json }
        catch {
            Write-Warning "Could not parse settings file. Using defaults."
            $defaultSettings | ConvertTo-Json | Set-Content -Path $settingsFilePath
            return $defaultSettings
        }
    }
    else {
        $defaultSettings | ConvertTo-Json | Set-Content -Path $settingsFilePath
        return $defaultSettings
    }
}

function Save-Settings {
    param([Parameter(Mandatory = $true)][PSCustomObject]$Settings)
    try {
        $Settings | ConvertTo-Json | Set-Content -Path $settingsFilePath
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to save settings to `'$settingsFilePath`'. Check permissions.", "Error", "OK", "Error")
    }
}

function Parse-IpRanges {
    param([Parameter(Mandatory = $true)][string]$IpRangeString)
    $allIps = New-Object System.Collections.Generic.List[string]
    $ranges = $IpRangeString -split ',' | ForEach-Object { $_.Trim() }
    foreach ($range in $ranges) {
        if ($range -match '^(\d{1,3}\.\d{1,3}\.\d{1,3})\.0$') {
            $base = $matches[1]; 2..254 | ForEach-Object { [void]$allIps.Add("$base.$_") }
        }
        elseif ($range -match '^(\d{1,3}\.\d{1,3}\.\d{1,3})\.(\d{1,3})-(\d{1,3})$') {
            $base = $matches[1]; ([int]$matches[2])..([int]$matches[3]) | ForEach-Object { [void]$allIps.Add("$base.$_") }
        }
        elseif ($range -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
            [void]$allIps.Add($range)
        }
    }
    return $allIps
}

function Validate-Inputs {
    param([Parameter(Mandatory = $true)][PSCustomObject]$inputs)
    $ipRanges = $inputs.IpRanges.Replace(" ", "")
    if ([string]::IsNullOrWhiteSpace($ipRanges)) {
        [System.Windows.Forms.MessageBox]::Show("IP Address Ranges cannot be empty.", "Validation Error", "OK", "Warning"); return $false
    }
    foreach ($range in ($ipRanges -split ',')) {
        if ($range -notmatch '^(\d{1,3}\.\d{1,3}\.\d{1,3})\.(\d{1,3}(?:-\d{1,3})?|0)$') {
            [System.Windows.Forms.MessageBox]::Show("Invalid IP range format: `'$range`'. Use formats like '192.168.1.5', '192.168.1.10-20', or '192.168.1.0'.", "Validation Error", "OK", "Warning"); return $false
        }
        $parts = $range -split '\.'
        for ($i = 0; $i -lt 3; $i++) {
            if ([int]$parts[$i] -lt 1 -or [int]$parts[$i] -gt 254) {
                [System.Windows.Forms.MessageBox]::Show("Invalid octet value in `'$range`'. The first three octets must be between 1 and 254.", "Validation Error", "OK", "Warning"); return $false
            }
        }
        if ($parts[3] -match '(\d+)-(\d+)') {
            if ([int]$matches[1] -lt 2 -or [int]$matches[2] -gt 254 -or [int]$matches[1] -ge [int]$matches[2]) {
                [System.Windows.Forms.MessageBox]::Show("Invalid range in `'$range`'. Range must be between 2 and 254, and the start must be less than the end.", "Validation Error", "OK", "Warning"); return $false
            }
        }
        elseif ($parts[3] -ne '0') {
            if ([int]$parts[3] -lt 2 -or [int]$parts[3] -gt 254) {
                 [System.Windows.Forms.MessageBox]::Show("Invalid host value in `'$range`'. The fourth octet must be between 2 and 254 (or 0 for a full range).", "Validation Error", "OK", "Warning"); return $false
            }
        }
    }
    if ($inputs.SnmpCommunity -notmatch '^[a-zA-Z0-9@#$%\&\*]{1,32}$') {
        [System.Windows.Forms.MessageBox]::Show("Community String must be 1-32 characters and can only contain letters, numbers, and the symbols: @#$%&*.", "Validation Error", "OK", "Warning"); return $false
    }
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars() -join ''; $regexInvalid = "[{0}]" -f [System.Text.RegularExpressions.Regex]::Escape($invalidChars)
    if ($inputs.OutputFileName -match $regexInvalid) {
        [System.Windows.Forms.MessageBox]::Show("Output File Name contains invalid characters.", "Validation Error", "OK", "Warning"); return $false
    }
    return $true
}
#endregion

#region GUI Creation

function Create-ConfigForm {
    param([Parameter(Mandatory = $true)][PSCustomObject]$Settings)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "DiscoverSubnet v$scriptVersion - Configuration"; $form.Size = New-Object System.Drawing.Size(420, 480); $form.FormBorderStyle = 'FixedDialog'
    $form.StartPosition = 'CenterScreen'; $form.MaximizeBox = $false; $form.MinimizeBox = $false
    $yPos = 15; $labelWidth = 160; $controlWidth = 210
    $label = New-Object System.Windows.Forms.Label; $label.Text = "IP Address Ranges:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $ipRangesBox = New-Object System.Windows.Forms.TextBox; $ipRangesBox.Location = New-Object System.Drawing.Point(180, $yPos); $ipRangesBox.Size = New-Object System.Drawing.Size($controlWidth, 20); $ipRangesBox.Text = $Settings.IpRanges; $form.Controls.Add($ipRangesBox); $yPos += 30
    $label = New-Object System.Windows.Forms.Label; $label.Text = "SNMP Community String:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $communityBox = New-Object System.Windows.Forms.TextBox; $communityBox.Location = New-Object System.Drawing.Point(180, $yPos); $communityBox.Size = New-Object System.Drawing.Size($controlWidth, 20); $communityBox.Text = $Settings.SnmpCommunity; $form.Controls.Add($communityBox); $yPos += 30
    $label = New-Object System.Windows.Forms.Label; $label.Text = "Ping/SNMP Retries:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $retriesDropdown = New-Object System.Windows.Forms.ComboBox; $retriesDropdown.Location = New-Object System.Drawing.Point(180, $yPos); $retriesDropdown.Size = New-Object System.Drawing.Size($controlWidth, 20); $retriesDropdown.DropDownStyle = 'DropDownList'; [void]$retriesDropdown.Items.AddRange(@(0, 1, 2, 3)); $retriesDropdown.SelectedItem = $Settings.Retries; $form.Controls.Add($retriesDropdown); $yPos += 30
    $label = New-Object System.Windows.Forms.Label; $label.Text = "Output File Name:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $fileNameBox = New-Object System.Windows.Forms.TextBox; $fileNameBox.Location = New-Object System.Drawing.Point(180, $yPos); $fileNameBox.Size = New-Object System.Drawing.Size($controlWidth, 20); $fileNameBox.Text = $Settings.OutputFileName; $form.Controls.Add($fileNameBox); $yPos += 30
    $label = New-Object System.Windows.Forms.Label; $label.Text = "Output File Type:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $fileTypeDropdown = New-Object System.Windows.Forms.ComboBox; $fileTypeDropdown.Location = New-Object System.Drawing.Point(180, $yPos); $fileTypeDropdown.Size = New-Object System.Drawing.Size($controlWidth, 20); $fileTypeDropdown.DropDownStyle = 'DropDownList'; [void]$fileTypeDropdown.Items.AddRange(@('txt', 'csv')); $fileTypeDropdown.SelectedItem = $Settings.OutputFileExtension; $form.Controls.Add($fileTypeDropdown); $yPos += 30
    $label = New-Object System.Windows.Forms.Label; $label.Text = "Max Parallel Scans:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $parallelScansUpDown = New-Object System.Windows.Forms.NumericUpDown; $parallelScansUpDown.Location = New-Object System.Drawing.Point(180, $yPos); $parallelScansUpDown.Size = New-Object System.Drawing.Size($controlWidth, 20); $parallelScansUpDown.Minimum = 1; $parallelScansUpDown.Maximum = 100; $parallelScansUpDown.Value = $Settings.MaxParallelScans; $form.Controls.Add($parallelScansUpDown); $yPos += 30
    $label = New-Object System.Windows.Forms.Label; $label.Text = "Diagnostic Level:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $diagDropdown = New-Object System.Windows.Forms.ComboBox; $diagDropdown.Location = New-Object System.Drawing.Point(180, $yPos); $diagDropdown.Size = New-Object System.Drawing.Size($controlWidth, 20); $diagDropdown.DropDownStyle = 'DropDownList'; [void]$diagDropdown.Items.AddRange(@('Off', 'Standard', 'Verbose')); $diagDropdown.SelectedItem = $Settings.DiagnosticLevel; $form.Controls.Add($diagDropdown); $yPos += 40
    $saveUnresponsiveCheck = New-Object System.Windows.Forms.CheckBox; $saveUnresponsiveCheck.Text = "Save SNMP-unresponsive devices to output file"; $saveUnresponsiveCheck.Location = New-Object System.Drawing.Point(20, $yPos); $saveUnresponsiveCheck.Size = New-Object System.Drawing.Size(370, 20); $saveUnresponsiveCheck.Checked = $Settings.SaveUnresponsive; $form.Controls.Add($saveUnresponsiveCheck); $yPos += 40
    $startButton = New-Object System.Windows.Forms.Button; $startButton.Text = "Start Discovery"; $startButton.Location = New-Object System.Drawing.Point(150, $yPos); $startButton.Size = New-Object System.Drawing.Size(120, 40); $startButton.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Controls.Add($startButton); $form.AcceptButton = $startButton
    return @{ Form = $form; IpRangesBox = $ipRangesBox; CommunityBox = $communityBox; RetriesDropdown = $retriesDropdown; FileNameBox = $fileNameBox; FileTypeDropdown = $fileTypeDropdown; ParallelScans = $parallelScansUpDown; DiagDropdown = $diagDropdown; SaveUnresponsive = $saveUnresponsiveCheck }
}

function Create-LogForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "DiscoverSubnet v$scriptVersion - In Progress"; $form.Size = New-Object System.Drawing.Size(800, 600); $form.StartPosition = 'CenterScreen'
    $logBox = New-Object System.Windows.Forms.TextBox; $logBox.Multiline = $true; $logBox.ScrollBars = 'Vertical'; $logBox.ReadOnly = $true; $logBox.Dock = 'Fill'; $logBox.Font = New-Object System.Drawing.Font("Consolas", 9); $form.Controls.Add($logBox)
    $bottomPanel = New-Object System.Windows.Forms.Panel; $bottomPanel.Dock = 'Bottom'; $bottomPanel.Height = 50; $form.Controls.Add($bottomPanel)
    $progressBar = New-Object System.Windows.Forms.ProgressBar; $progressBar.Location = New-Object System.Drawing.Point(10, 15); $progressBar.Size = New-Object System.Drawing.Size(420, 23); $bottomPanel.Controls.Add($progressBar)
    $saveButton = New-Object System.Windows.Forms.Button; $saveButton.Text = "Save Results"; $saveButton.Location = New-Object System.Drawing.Point(440, 12); $saveButton.Enabled = $false; $bottomPanel.Controls.Add($saveButton)
    $cancelButton = New-Object System.Windows.Forms.Button; $cancelButton.Text = "Cancel Scan"; $cancelButton.Location = New-Object System.Drawing.Point(530, 12); $bottomPanel.Controls.Add($cancelButton)
    $closeButton = New-Object System.Windows.Forms.Button; $closeButton.Text = "Close"; $closeButton.Location = New-Object System.Drawing.Point(620, 12); $closeButton.Enabled = $false; $bottomPanel.Controls.Add($closeButton)
    return @{ Form = $form; LogBox = $logBox; ProgressBar = $progressBar; SaveButton = $saveButton; CancelButton = $cancelButton; CloseButton = $closeButton }
}
#endregion

#region Main Script Execution

$currentSettings = Load-Settings
$configFormElements = Create-ConfigForm -Settings $currentSettings
$configFormElements.Form.ShowDialog() | Out-Null

if ($configFormElements.Form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {

    $scanSettings = [PSCustomObject]@{
        IpRanges              = $configFormElements.IpRangesBox.Text; SnmpCommunity = $configFormElements.CommunityBox.Text
        Retries               = [int]$configFormElements.RetriesDropdown.SelectedItem; OutputFileName = $configFormElements.FileNameBox.Text
        OutputFileExtension   = $configFormElements.FileTypeDropdown.SelectedItem; SaveUnresponsive = $configFormElements.SaveUnresponsive.Checked
        MaxParallelScans      = [int]$configFormElements.ParallelScans.Value; DiagnosticLevel = $configFormElements.DiagDropdown.SelectedItem
    }

    if (-not (Validate-Inputs -inputs $scanSettings)) { exit }
    Save-Settings -Settings $scanSettings

    $logFormElements = Create-LogForm
    $logForm = $logFormElements.Form; $script:logTextBox = $logFormElements.LogBox; $progressBar = $logFormElements.ProgressBar
    $cancelButton = $logFormElements.CancelButton; $saveButton = $logFormElements.SaveButton; $closeButton = $logFormElements.CloseButton

    $global:scanJob = $null; $global:allDiscoveredDevices = New-Object System.Collections.Generic.List[PSCustomObject]
    $logFileName = "DiscoverSubnet-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
    $logFilePath = Join-Path -Path $scriptDir -ChildPath $logFileName
    $outputFilePath = Join-Path -Path $scriptDir -ChildPath "$($scanSettings.OutputFileName).$($scanSettings.OutputFileExtension)"

    # Initialize log file with version header as first line
    $versionHeader = "DiscoverSubnet Version $scriptVersion"
    Add-Content -Path $logFilePath -Value $versionHeader

    function Write-Log {
        param($Message)
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; $logEntry = "[$timestamp] $Message"
        if ($script:logTextBox -and -not $script:logTextBox.IsDisposed) {
            try {
                # Force handle creation if it doesn't exist
                if (-not $script:logTextBox.IsHandleCreated) {
                    $script:logTextBox.CreateControl()
                }
                
                if ($script:logTextBox.IsHandleCreated) {
                    $script:logTextBox.BeginInvoke([Action[string]]{ param($text)
                        if ($script:logTextBox -and -not $script:logTextBox.IsDisposed) { 
                            $script:logTextBox.AppendText($text + [Environment]::NewLine) 
                        }
                    }, $logEntry)
                }
            }
            catch {
                # If BeginInvoke fails, fall back to direct access (less thread-safe but works)
                Write-Warning "BeginInvoke failed, using direct access: $($_.Exception.Message)"
                if ($script:logTextBox -and -not $script:logTextBox.IsDisposed) {
                    $script:logTextBox.AppendText($logEntry + [Environment]::NewLine)
                }
            }
        }
        Add-Content -Path $logFilePath -Value $logEntry
    }

    $controllerScriptBlock = {
        param($Settings)
        function New-JobMessage { param($Type, $Value) [PSCustomObject]@{Type = $Type; Value = $Value} }
        
        # Add a unique identifier to detect multiple controller instances
        $controllerId = [System.Guid]::NewGuid().ToString().Substring(0,8)
        $global:controllerExecuting = $true
        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Starting scan process...")
        
        # Log the scan parameters for troubleshooting
        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: === SCAN PARAMETERS ===")
        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: IP Ranges: $($Settings.IpRanges)")
        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: SNMP Community: $($Settings.SnmpCommunity)")
        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Retries: $($Settings.Retries)")
        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Max Parallel Scans: $($Settings.MaxParallelScans)")
        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Output File: $($Settings.OutputFileName).$($Settings.OutputFileExtension)")
        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Save Unresponsive: $($Settings.SaveUnresponsive)")
        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: ========================")
        
        # Safety check to prevent infinite recursion
        if ($global:controllerAlreadyRan) {
            Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: ERROR - Controller already executed, preventing duplicate run")
            Write-Output (New-JobMessage -Type "Status" -Value "Complete")
            return
        }
        $global:controllerAlreadyRan = $true
        
        # Test OleSNMP COM object availability first
        try {
            $testSNMP = New-Object -ComObject olePrn.OleSNMP
            $testSNMP = $null  # Release the test object
            Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Successfully verified OleSNMP COM object availability")
        } catch {
            $errorMsg = "Failed to create OleSNMP COM object: $($_.Exception.Message)"
            Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: $errorMsg")
            Write-Output (New-JobMessage -Type "Status" -Value "Complete")
            return
        }
        
        $workerScriptBlock = {
            param($CurrentIP, $ScanSettings)
            
            # Function to create worker messages that can be captured by the controller
            function New-WorkerMessage { param($Type, $Value) [PSCustomObject]@{Type = $Type; Value = $Value; IP = $CurrentIP} }
            
            #region SNMP COM Object Availability Check for Worker
            $snmpAvailable = $false
            try {
                $testSNMP = New-Object -ComObject olePrn.OleSNMP
                $testSNMP = $null  # Release the test object
                $snmpAvailable = $true
                Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - OleSNMP COM object available")
            } catch {
                # SNMP not available, will do ping-only scan
                $snmpAvailable = $false
                $errorMsg = $_.Exception.Message
                Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - OleSNMP COM object unavailable - $errorMsg")
            }
            #endregion

            function Get-SnmpValue {
                param([string]$IP, [string]$Community, [string[]]$OIDs, [int]$Retries)
                $results = @{}
                $errorOccurred = $false
                $lastError = ""
                for ($attempt = 0; $attempt -le $Retries; $attempt++) {
                    try {
                        foreach ($oid in $OIDs) {
                            try {
                                $SNMP = New-Object -ComObject olePrn.OleSNMP
                                $SNMP.open($IP, $Community, $Retries, 1000)
                                $value = $SNMP.get($oid)
                                $SNMP.Close()
                                Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - SNMP raw value for $oid : '$value' (Type: $($value.GetType().Name))")
                                if ($value -and $value -ne "") {
                                    $results[$oid] = $value
                                } else {
                                    $results[$oid] = $null
                                }
                            } catch {
                                $results[$oid] = $null
                            } finally {
                                if ($SNMP) { 
                                    try { $SNMP.Close() } catch { }
                                    $SNMP = $null
                                }
                            }
                        }
                        $errorOccurred = $false
                        break
                    } catch { 
                        $errorOccurred = $true
                        $lastError = $_.Exception.Message
                        Start-Sleep -Milliseconds 500 
                    }
                }
                if ($errorOccurred) { 
                    return $null 
                }
                return $results
            }

            function Get-DeviceType {
                param([string]$OID, [string]$SysName = "", [string]$IP = $CurrentIP)
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
                    "1.3.6.1.4.1.17186.1.10" { 
                        # MD8000-Series Refinement: Check for EX/SX variants
                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - Detected MD8000, checking for EX/SX variant...")
                        try {
                            $SNMP = New-Object -ComObject olePrn.OleSNMP
                            $SNMP.open($IP, $ScanSettings.SnmpCommunity, $ScanSettings.Retries, 1000)
                            $variantValue = $SNMP.get(".1.3.6.1.4.1.17186.1.10.1.1.3.0")
                            $SNMP.Close()
                            Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - MD8000 variant OID value: '$variantValue'")
                            switch ($variantValue) {
                                "1" { return "MD8000EX" }
                                "2" { return "MD8000SX" }
                                default { return "MD8000" }
                            }
                        } catch {
                            Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - Failed to query MD8000 variant OID: $($_.Exception.Message)")
                            return "MD8000"
                        } finally {
                            if ($SNMP) { 
                                try { $SNMP.Close() } catch { }
                                $SNMP = $null
                            }
                        }
                    }
                    "1.3.6.1.4.1.21839.1.2.17"    { return "MDX2040" }
                    "1.3.6.1.4.1.17186.1.24"      { return "MDP3020" }
                    "1.3.6.1.4.1.17186.3.1.1.1.0" { 
                        # MDX-Series Refinement: Check sysName for 32C vs 48X6C
                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - Detected MDX series, checking sysName '$SysName' for refinement...")
                        if ($SysName -match "32C") {
                            Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - sysName contains '32C', refined type: MDX-32C")
                            return "MDX-32C"
                        } elseif ($SysName -match "48X") {
                            Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - sysName contains '48X', refined type: MDX-48X6C")
                            return "MDX-48X6C"
                        } else {
                            Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - sysName does not contain '32C' or '48X', keeping generic type: MDX32C/48X6C")
                            return "MDX32C/48X6C"
                        }
                    }
                    default {
                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - Unrecognized device type OID: '$cleanOID'")
                        return "UNKNOWN"
                    }
                }
            }
            
            $pingSuccess = $false
            Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - Starting ping test")
            for ($i = 0; $i -le $ScanSettings.Retries; $i++) { 
                if (Test-Connection -ComputerName $CurrentIP -Count 1 -Quiet -ErrorAction SilentlyContinue) { 
                    $pingSuccess = $true
                    Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - Ping successful on attempt $($i + 1)")
                    break 
                } 
            }
            if (-not $pingSuccess) {
                Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - All ping attempts failed")
            }
            if ($pingSuccess) {
                if ($snmpAvailable) {
                    Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - Starting SNMP queries with community '$($ScanSettings.SnmpCommunity)'")
                    $oidsToGet = @( ".1.3.6.1.2.1.1.2.0", ".1.3.6.1.2.1.1.5.0", ".1.3.6.1.2.1.1.6.0" )
                    try {
                        $snmpResult = Get-SnmpValue -IP $CurrentIP -Community $ScanSettings.SnmpCommunity -OIDs $oidsToGet -Retries $ScanSettings.Retries
                        # Handle the case where the result might be mixed with other outputs in an array
                        if ($snmpResult -is [array]) {
                            $snmpResult = $snmpResult | Where-Object { $_ -is [hashtable] } | Select-Object -First 1
                        }
                        if ($snmpResult -and $snmpResult -is [hashtable]) {
                            Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - SNMP returned hashtable with $($snmpResult.Keys.Count) keys")
                        } else {
                            Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - SNMP returned null or invalid result")
                            $snmpResult = $null
                        }
                    } catch {
                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - Exception in Get-SnmpValue: $($_.Exception.Message)")
                        $snmpResult = $null
                    }
                    if ($snmpResult -and $snmpResult.ContainsKey(".1.3.6.1.2.1.1.2.0")) {
                        $typeOID = $snmpResult[".1.3.6.1.2.1.1.2.0"]
                        $name = $snmpResult[".1.3.6.1.2.1.1.5.0"]
                        $location = $snmpResult[".1.3.6.1.2.1.1.6.0"]
                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - SNMP data retrieved: typeOID='$typeOID', name='$name', location='$location'")
                        
                        # Handle empty/null values like v1 does
                        if ([string]::IsNullOrWhiteSpace($name)) { $name = "[No Name Found]" }
                        if ([string]::IsNullOrWhiteSpace($location)) { $location = "[No Location Found]" }
                        
                        # If the type OID is missing, the type is UNKNOWN.
                        $type = if ($typeOID) { Get-DeviceType -OID $typeOID -SysName $name -IP $CurrentIP } else { "UNKNOWN" }
                        
                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - Device type determined: '$type'")
                        Write-Output ([PSCustomObject]@{ IP = $CurrentIP; Name = $name; Location = $location; Type = $type; Status = "Responsive" })
                    } else { 
                        Write-Output ([PSCustomObject]@{ IP = $CurrentIP; Name = "[No Name Found]"; Location = "[No Location Found]"; Type = "UNKNOWN"; Status = "Unresponsive" })
                    }
                } else {
                    # SNMP not available, return ping-only result
                    Write-Output ([PSCustomObject]@{ IP = $CurrentIP; Name = "[SNMP Unavailable]"; Location = "[Ping Only]"; Type = "PING_ONLY"; Status = "Responsive" })
                }
            } else {
                Write-Output ([PSCustomObject]@{ IP = $CurrentIP; Name = "[No Response]"; Location = "[No Response]"; Type = "NO_PING"; Status = "Unresponsive" })
            }
        } # End of $workerScriptBlock

        function Parse-IpRangesJob {
            param([string]$IpRangeString)
            $allIps = New-Object System.Collections.Generic.List[string]; $ranges = $IpRangeString -split ',' | ForEach-Object { $_.Trim() }
            foreach ($range in $ranges) {
                if ($range -match '^(\d{1,3}\.\d{1,3}\.\d{1,3})\.0$') { $base = $matches[1]; 2..254 | ForEach-Object { [void]$allIps.Add("$base.$_") } }
                elseif ($range -match '^(\d{1,3}\.\d{1,3}\.\d{1,3})\.(\d{1,3})-(\d{1,3})$') { $base = $matches[1]; ([int]$matches[2])..([int]$matches[3]) | ForEach-Object { [void]$allIps.Add("$base.$_") } }
                elseif ($range -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') { [void]$allIps.Add($range) }
            }
            return $allIps
        }

        $allIpsToScan = Parse-IpRangesJob -IpRangeString $Settings.IpRanges; $totalIpCount = $allIpsToScan.Count
        $processedIpCount = 0; $allResults = New-Object System.Collections.Generic.List[object]; $runningJobs = @()
        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Starting scan of $totalIpCount IP addresses...")

        foreach($ip in $allIpsToScan) {
            while ($runningJobs.Count -ge $Settings.MaxParallelScans) {
                try {
                    $completedJob = Wait-Job -Job $runningJobs -Any -Timeout 30
                    if ($completedJob) {
                        $runningJobs = $runningJobs | Where-Object { $_.Id -ne $completedJob.Id }
                        $jobResults = Receive-Job -Job $completedJob
                        
                        # Process all outputs from the job
                        Write-Output (New-JobMessage -Type "Log" -Value "Controller [\$controllerId]: Processing $($jobResults.Count) results from final cleanup job")
                        foreach ($jobResult in $jobResults) {
                            if ($jobResult -and $jobResult.Type -eq "WorkerLog") {
                                # Forward worker log messages to main log
                                Write-Output (New-JobMessage -Type "Log" -Value $jobResult.Value)
                            } elseif ($jobResult -and $jobResult.IP) {
                                # This is a device result
                                Write-Output (New-JobMessage -Type "Log" -Value "Controller [\$controllerId]: Found device result for IP $($jobResult.IP)")
                                [void]$allResults.Add($jobResult)
                            } else {
                                Write-Output (New-JobMessage -Type "Log" -Value "Controller [\$controllerId]: Unrecognized final job result: $($jobResult | Out-String)")
                            }
                        }
                        
                        Remove-Job -Job $completedJob
                        $processedIpCount++
                        if ($totalIpCount -gt 0) { $progress = ($processedIpCount / $totalIpCount) * 100; Write-Output (New-JobMessage -Type "Progress" -Value $progress) }
                    } else {
                        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Timeout waiting for jobs, stopping remaining jobs")
                        $runningJobs | Stop-Job
                        $runningJobs | Remove-Job -Force
                        $runningJobs = @()
                        break
                    }
                } catch {
                    Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Error in job management: $($_.Exception.Message)")
                    break
                }
            }
            
            try {
                $newJob = Start-Job -ScriptBlock $workerScriptBlock -ArgumentList $ip, $Settings
                $runningJobs += $newJob
                Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Started job for IP $ip (Job ID: $($newJob.Id))")
            } catch {
                Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Failed to start job for IP $ip : $($_.Exception.Message)")
            }
        }
        while ($runningJobs.Count -gt 0) {
            try {
                $completedJob = Wait-Job -Job $runningJobs -Any -Timeout 30
                if ($completedJob) {
                    $runningJobs = $runningJobs | Where-Object { $_.Id -ne $completedJob.Id }
                    $jobResults = Receive-Job -Job $completedJob
                    
                    # Process all outputs from the job
                    Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Processing $($jobResults.Count) results from completed job")
                    foreach ($jobResult in $jobResults) {
                        if ($jobResult -and $jobResult.Type -eq "WorkerLog") {
                            # Forward worker log messages to main log
                            Write-Output (New-JobMessage -Type "Log" -Value $jobResult.Value)
                        } elseif ($jobResult -and $jobResult.IP) {
                            # This is a device result
                            Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Found device result for IP $($jobResult.IP)")
                            [void]$allResults.Add($jobResult)
                        } else {
                            Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Unrecognized job result: $($jobResult | Out-String)")
                        }
                    }
                    
                    Remove-Job -Job $completedJob
                    $processedIpCount++
                    if ($totalIpCount -gt 0) { $progress = ($processedIpCount / $totalIpCount) * 100; Write-Output (New-JobMessage -Type "Progress" -Value $progress) }
                } else {
                    Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Timeout waiting for remaining jobs, stopping all")
                    $runningJobs | Stop-Job
                    $runningJobs | Remove-Job -Force
                    break
                }
            } catch {
                Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Error in final job cleanup: $($_.Exception.Message)")
                $runningJobs | Stop-Job
                $runningJobs | Remove-Job -Force
                break
            }
        }

        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Sending results - found $($allResults.Count) devices")
        Write-Output (New-JobMessage -Type "RangeResult" -Value $allResults)
        Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Scan completed successfully")
        Write-Output (New-JobMessage -Type "Status" -Value "Complete")
        $global:controllerExecuting = $false
    } # End Controller ScriptBlock

    # --- GUI Event Handlers & Timer ---
    $guiTimer = New-Object System.Windows.Forms.Timer; $guiTimer.Interval = 500
    $guiTimer.add_Tick({
        if ($global:scanJob) {
            # Check if job still exists and has valid state
            try {
                $jobState = $global:scanJob.State
                # Don't preemptively stop on job completion - wait for proper "Status Complete" message
                # This prevents race conditions where the job completes before the controller sends final results
                if ($jobState -eq 'Completed') {
                    Write-Log -Message "Job completed - waiting for controller to send final results..."
                    # Continue processing messages, don't stop here
                }
                if ($jobState -eq 'Failed') {
                    Write-Log -Message "Job failed - checking for error details"
                    $jobErrors = Receive-Job -Job $global:scanJob -ErrorAction SilentlyContinue
                    if ($jobErrors) {
                        Write-Log -Message "Job error: $jobErrors"
                    }
                    $guiTimer.Stop()
                    $cancelButton.Enabled = $false; $saveButton.Enabled = $true; $closeButton.Enabled = $true
                    $global:scanJob = $null
                    return
                } elseif ($jobState -eq 'Stopped') {
                    Write-Log -Message "Job was stopped - stopping timer"
                    $guiTimer.Stop()
                    $global:scanJob = $null
                    return
                }
                # Always consume messages (never keep them) to prevent reprocessing
                # This prevents the infinite loop of processing the same messages repeatedly
                $messages = Receive-Job -Job $global:scanJob -Keep:$false
            } catch {
                Write-Log -Message "Error accessing job: $($_.Exception.Message)"
                $guiTimer.Stop()
                $cancelButton.Enabled = $false; $saveButton.Enabled = $true; $closeButton.Enabled = $true
                $global:scanJob = $null
                return
            }
            foreach ($msg in $messages) {
                switch ($msg.Type) {
                    "Log" { Write-Log -Message $msg.Value }
                    "WorkerLog" { Write-Log -Message $msg.Value }
                    "Progress" { if($msg.Value -le 100) {$progressBar.Value = [int]$msg.Value } }
                    "RangeResult" {
                        $responsive = $msg.Value | Where-Object { $_.Status -eq 'Responsive' }; $unresponsive = $msg.Value | Where-Object { $_.Status -eq 'Unresponsive' }
                        foreach ($item in $msg.Value) { 
                            [void]$global:allDiscoveredDevices.Add($item)
                            Write-Log -Message "Added device to collection: $($item.IP) - $($item.Name) - Status: $($item.Status)"
                        }
                        $grouped = $responsive | Group-Object Name, Location, Type
                        if($grouped) {
                            Write-Log -Message "--- Discovered Devices ---"
                            foreach ($group in $grouped) {
                                $ips = ($group.Group.IP | Sort-Object) -join ', '; $name = $group.Group[0].Name; $location = $group.Group[0].Location; $type = $group.Group[0].Type
                                Write-Log -Message "  Name=$name Location=$location Type=$type Address=$ips"
                            }
                        }
                        if ($unresponsive) { $unresponsiveIPs = ($unresponsive.IP | Sort-Object) -join ', '; Write-Log -Message "  SNMP Unresponsive: $unresponsiveIPs" }
                    }
                    "Status" {
                        if ($msg.Value -eq 'Complete') {
                            Write-Log -Message "Overall Discovery Complete."; $progressBar.Value = 100
                            $cancelButton.Enabled = $false; $saveButton.Enabled = $true; $closeButton.Enabled = $true
                            $guiTimer.Stop()
                            if ($global:scanJob) {
                                Receive-Job -Job $global:scanJob | Out-Null
                                Remove-Job -Job $global:scanJob -Force
                                $global:scanJob = $null
                            }
                        }
                    }
                }
            }
        }
    })

    $cancelButton.add_Click({
        Write-Log -Message "Scan cancellation requested by user."
        if ($global:scanJob) {
            Get-Job | Where-Object { $_.Id -ge $global:scanJob.Id } | Stop-Job
            Get-Job | Where-Object { $_.Id -ge $global:scanJob.Id } | Remove-Job -Force
            $global:scanJob = $null
        }
        $guiTimer.Stop(); $cancelButton.Enabled = $false; $saveButton.Enabled = $true; $closeButton.Enabled = $true
        Write-Log -Message "Scan cancelled."
    })

    $saveButton.add_Click({
        try {
            Write-Log -Message "Save button clicked. Total devices in collection: $($global:allDiscoveredDevices.Count)"
            if ($global:allDiscoveredDevices.Count -gt 0) {
                Write-Log -Message "Sample devices in collection:"
                $global:allDiscoveredDevices | Select-Object -First 3 | ForEach-Object { 
                    Write-Log -Message "  Device: $($_.IP) - $($_.Name) - Status: $($_.Status)" 
                }
            }
            $devicesToSave = if ($scanSettings.SaveUnresponsive) { $global:allDiscoveredDevices } else { $global:allDiscoveredDevices | Where-Object { $_.Status -eq 'Responsive' } }
            
            # Group devices by Name, Location, Type and combine their IP addresses
            $groupedDevices = $devicesToSave | Group-Object Name, Location, Type
            $outputData = $groupedDevices | ForEach-Object {
                $group = $_
                $combinedIPs = ($group.Group.IP | Sort-Object) -join ', '
                [PSCustomObject]@{
                    Name = $group.Group[0].Name
                    Location = $group.Group[0].Location
                    Type = $group.Group[0].Type
                    IPs = $combinedIPs
                }
            }
            if ($outputData) {
                 $outputData | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ';'; Write-Log -Message "Results successfully saved to $outputFilePath"
                 [System.Windows.Forms.MessageBox]::Show("Results saved successfully.", "Save Complete", "OK", "Information")
            } else {
                 Write-Log -Message "No devices to save."; [System.Windows.Forms.MessageBox]::Show("There are no discovered devices to save.", "Save Complete", "OK", "Information")
            }
        } catch {
            Write-Log -Message "ERROR: Failed to save results. $_"; [System.Windows.Forms.MessageBox]::Show("An error occurred while saving the file: `n$($_.Exception.Message)", "Save Error", "OK", "Error")
        }
    })

    $closeButton.add_Click({ $logForm.Close() })

    $logForm.add_FormClosing({
        if ($global:scanJob) {
            Get-Job | Where-Object { $_.Id -ge $global:scanJob.Id } | Stop-Job
            Get-Job | Where-Object { $_.Id -ge $global:scanJob.Id } | Remove-Job -Force
        }
        $guiTimer.Stop()
    })
    
    # Flag to prevent multiple job starts
    $script:jobStarted = $false
    $script:formShownCount = 0
    
    # This event handler ensures the job and timer only start AFTER the form is fully visible.
    $logForm.add_Shown({
        $script:formShownCount++
        Write-Log -Message "Form shown event #$($script:formShownCount)"
        
        if (-not $script:jobStarted -and -not $global:scanJob -and $script:formShownCount -eq 1) {
            $script:jobStarted = $true
            Write-Log -Message "Form shown - starting scan job..."
            try {
                $global:scanJob = Start-Job -ScriptBlock $controllerScriptBlock -ArgumentList @($scanSettings)
                $guiTimer.Start()
                Write-Log -Message "Scan job started successfully with ID: $($global:scanJob.Id)"
            } catch {
                Write-Log -Message "Error starting scan job: $($_.Exception.Message)"
                $script:jobStarted = $false
            }
        } else {
            Write-Log -Message "Form shown but conditions not met - jobStarted:$script:jobStarted, scanJobExists:$($global:scanJob -ne $null), formShownCount:$script:formShownCount"
        }
    })

    [void]$logForm.ShowDialog()
}
# End of script.
