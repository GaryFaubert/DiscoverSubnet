#requires -Version 5.1
<#
.SYNOPSIS
    DiscoverSubnet is a network discovery tool specifically designed to identify MediaLinks hardware on a network.

.DESCRIPTION
    This PowerShell script provides a comprehensive network discovery solution with the following features:
    
    GUI INTERFACE:
    - User-friendly Windows Forms interface for configuration
    - Real-time progress monitoring and logging display
    - Automatic system performance analysis and parallel scan recommendations
    
    NETWORK SCANNING:
    - Supports multiple IP range formats (single IPs, ranges, subnets)
    - Parallel scanning architecture for optimal performance
    - SNMP device identification with fallback to ping-only mode
    - Configurable retry logic for unreliable network conditions
    
    DEVICE IDENTIFICATION:
    - Specialized recognition of MediaLinks hardware (MD8000, MDX series, etc.)
    - SNMP OID-based device type classification
    - Enhanced device variant detection (EX/SX models, 32C/48X6C variants)
    
    OUTPUT OPTIONS:
    - Real-time GUI display with verbosity control
    - Detailed log files with timestamps
    - Exportable results in CSV or TXT format
    - Option to include/exclude unresponsive devices
    
    PS2EXE COMPATIBILITY:
    - Fully compatible with PS2EXE compilation
    - Robust path resolution for both .ps1 and .exe execution
    - Fallback mechanisms for COM object availability

.PARAMETER None
    This script runs interactively and does not accept command-line parameters.

.EXAMPLE
    PS C:\> .\DiscoverSubnet.ps1
    Launches the GUI interface for network discovery configuration.

.EXAMPLE
    PS C:\> .\DiscoverSubnet.exe
    Runs the compiled version with identical functionality.

.NOTES
    REQUIREMENTS:
    - Windows operating system with .NET Framework
    - PowerShell 5.1 or higher
    - OleSNMP COM object for full SNMP functionality (optional - degrades gracefully)
    
    COMPILATION:
    - Compatible with PS2EXE for standalone executable creation
    - Settings file and logs are created in the same directory as the script/exe
    
    PERFORMANCE:
    - Automatically analyzes system capabilities for optimal parallel scan count
    - Typical scan speed: 2-4 seconds per IP address depending on responsiveness
    - Memory usage scales with parallel job count (typically 50-200MB)

.AUTHOR
    Gary Faubert - Assisted by Gemini and Copilot
    Copyright "Medialinks.inc 2025"

.DATE
    2025-10-27

.CHANGELOG
    v2.21 - Fixed XSCEND device detection: now properly detects XSCEND devices when SNMP queries fail (not just when SNMP COM object is unavailable)
    v2.20 - Added MDX2049 and XSCEND device detection: when SNMP is not available, checks device via HTTP and identifies XSCEND devices by searching for "XSCEND" keyword in web response
    v2.19 - Added support for multiple SNMP community strings: enter comma-separated values (e.g., "medialinks, custom, private") for mixed-device environments
    v2.18 - Enhanced SNMP querying: try "public" community first, then user-entered; added verbose file-only logging for detailed SNMP values when DiagnosticLevel=Verbose
    v2.17 - Improved GUI display: changed DISCOVERED DEVICES REPORT to white, removed SCAN COMPLETION SUMMARY from GUI, moved total count under DISCOVERED DEVICES REPORT
    v2.16 - Fixed IP range validation to allow scanning of .1 addresses (gateways); changed minimum range from 2 to 1
    v2.15 - Fixed GUI verbosity filtering to ensure summary sections always appear in discovery window
    v2.14 - Fixed missing summary sections; ensured discovered devices report and scan completion summary always appear
    v2.13 - Added visual section separators with distinct colors: dark cyan for scan completion, dark green for results report
    v2.12 - Suppressed unreachable IP messages during scan to reduce noise; show gray summary at end instead
    v2.11 - Added color-coded logging in discovery window to distinguish message types (errors=red, success=green, warnings=yellow, etc.)
    v2.10 - Added support for SWCNT9-100G device type detection in MD8000 series hardware
    v2.9 - Updated subnet scanning to include gateway addresses (.1) for more comprehensive network discovery
    v2.8 - Enhanced PS2EXE compatibility, improved error handling for null paths
    v2.7 - Added system performance analysis and automatic parallel scan recommendations
    v2.6 - Improved GUI verbosity controls and professional logging format
    v2.5 - Enhanced device type detection for MediaLinks hardware variants
#>

#region Global Variables & Initial Setup

# =============================================================================
# SCRIPT METADATA AND PATH RESOLUTION
# =============================================================================

# Version identifier used throughout the application for logging and display
$scriptVersion = "2.21"

# Robust script directory resolution that works for both .ps1 and compiled .exe files
# This is critical for PS2EXE compatibility where standard PowerShell variables may not be available

if ($PSScriptRoot) {
    # Standard PowerShell execution: $PSScriptRoot is reliably populated when running as .ps1 file
    $scriptDir = $PSScriptRoot
}
elseif ($null -ne $MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne "") {
    # PS2EXE compilation fallback: Handle cases where $PSScriptRoot is not available
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}
else {
    # Simple fallback for PS2EXE double-click: Use current directory without warning
    # This is perfectly acceptable for most use cases
    $scriptDir = Get-Location
}

# =============================================================================
# SETTINGS MANAGEMENT
# =============================================================================

# Path to the persistent settings file (JSON format) stored alongside the script/executable
$settingsFilePath = Join-Path -Path $scriptDir -ChildPath "DiscoverSubnet.settings.json"

# Global variables for tracking scan progress and results
$script:unreachableIPs = @()  # Track unreachable IPs to show summary at end instead of cluttering log

# Default configuration structure - used when no settings file exists or parsing fails
# These values represent sensible defaults for most network environments
$defaultSettings = [PSCustomObject]@{
    # Network scanning parameters
    IpRanges              = "192.168.1.0, 10.0.0.10-20"    # Example IP ranges for initial configuration
    SnmpCommunity         = "medialinks"                    # Default SNMP community string for MediaLinks devices
    Retries               = 0                               # Number of retry attempts for ping/SNMP failures
    
    # Output configuration
    OutputFileName        = "DiscoveredDevices"             # Base filename for results (extension added separately)
    OutputFileExtension   = "csv"                           # File format: 'csv' or 'txt'
    SaveUnresponsive      = $false                          # Whether to include unresponsive devices in output
    
    # Performance and system settings
    MaxParallelScans      = 20                              # Maximum concurrent scanning jobs
    DiagnosticLevel       = "Standard"                      # Logging detail: 'Off', 'Standard', 'Verbose'
    GuiVerbosity          = "Standard"                      # GUI display level: 'Minimal', 'Standard', 'Verbose'
}

#endregion

#region Core Helper Functions

# =============================================================================
# SYSTEM INITIALIZATION
# =============================================================================

# Load required .NET assemblies for Windows Forms GUI components
# This must be done early in the script execution to enable GUI functionality
try {
    Add-Type -AssemblyName System.Windows.Forms  # Windows Forms controls (buttons, textboxes, etc.)
    Add-Type -AssemblyName System.Drawing        # Drawing objects (fonts, colors, sizing)
}
catch {
    # Fatal error - GUI cannot function without these assemblies
    Write-Error "Failed to load required .NET Assemblies for GUI. Please ensure you are running in a Windows environment with .NET Framework."
    exit 1
}

# =============================================================================
# SETTINGS PERSISTENCE FUNCTIONS
# =============================================================================

function Load-Settings {
    <#
    .SYNOPSIS
        Loads user settings from the JSON configuration file or creates default settings.
    
    .DESCRIPTION
        Attempts to read and parse the settings JSON file. If the file doesn't exist or is corrupted,
        creates a new settings file with default values. This ensures the application always has
        valid configuration to work with.
    
    .OUTPUTS
        PSCustomObject containing all application settings
    
    .NOTES
        File location is determined by $settingsFilePath global variable.
        Uses error-tolerant approach - corruption results in defaults, not failure.
    
    .EXAMPLE
        $settings = Load-Settings
        # Returns settings object with all configuration properties
    #>
    
    if (Test-Path $settingsFilePath) {
        try { 
            # Attempt to parse existing settings file
            return Get-Content -Path $settingsFilePath | ConvertFrom-Json 
        }
        catch {
            # Settings file exists but is corrupted - recreate with defaults
            Write-Warning "Could not parse settings file. Using defaults."
            $defaultSettings | ConvertTo-Json | Set-Content -Path $settingsFilePath
            return $defaultSettings
        }
    }
    else {
        # No settings file exists - create one with default values
        $defaultSettings | ConvertTo-Json | Set-Content -Path $settingsFilePath
        return $defaultSettings
    }
}

function Save-Settings {
    <#
    .SYNOPSIS
        Persists current settings to the JSON configuration file.
    
    .DESCRIPTION
        Converts the settings object to JSON format and saves it to the configuration file.
        Handles write permissions errors gracefully by displaying a user-friendly message.
    
    .PARAMETER Settings
        PSCustomObject containing all settings to be saved
    
    .NOTES
        File location is determined by $settingsFilePath global variable.
        Write failures are handled gracefully with GUI error dialogs.
    
    .EXAMPLE
        Save-Settings -Settings $currentSettings
        # Saves settings to JSON file
    #>
    
    param([Parameter(Mandatory = $true)][PSCustomObject]$Settings)
    
    try {
        # Convert settings object to JSON and write to file
        $Settings | ConvertTo-Json | Set-Content -Path $settingsFilePath
    }
    catch {
        # Display user-friendly error message for write failures (permissions, disk full, etc.)
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to save settings to '$settingsFilePath'. Check permissions.", 
            "Error", 
            "OK", 
            "Error"
        )
    }
}

function Parse-IpRanges {
    <#
    .SYNOPSIS
        Parses IP range strings into individual IP addresses for scanning.
    
    .DESCRIPTION
        Converts user-friendly IP range notation into a list of individual IP addresses.
        Supports multiple formats:
        - Single IP: "192.168.1.5"
        - Subnet notation: "192.168.1.0" (expands to .1-.254)
        - Range notation: "192.168.1.10-20" (expands to .10-.20)
        - Mixed formats: "192.168.1.5, 10.0.0.0, 172.16.1.100-110"
    
    .PARAMETER IpRangeString
        Comma-separated string containing IP ranges in various formats
    
    .OUTPUTS
        System.Collections.Generic.List[string] containing individual IP addresses
    
    .NOTES
        Subnet notation (.0) automatically excludes .255 (broadcast) but includes .1 (gateway)
        Range notation is inclusive of both start and end addresses
        Invalid formats are silently ignored (validation handled elsewhere)
    
    .EXAMPLE
        Parse-IpRanges -IpRangeString "192.168.1.0, 10.0.0.5-10"
        # Returns: 192.168.1.1, 192.168.1.2, ..., 192.168.1.254, 10.0.0.5, 10.0.0.6, ..., 10.0.0.10
    #>
    
    param([Parameter(Mandatory = $true)][string]$IpRangeString)
    
    # Use generic list for better performance than array concatenation
    $allIps = New-Object System.Collections.Generic.List[string]
    
    # Split comma-separated ranges and clean whitespace
    $ranges = $IpRangeString -split ',' | ForEach-Object { $_.Trim() }
    
    foreach ($range in $ranges) {
        if ($range -match '^(\d{1,3}\.\d{1,3}\.\d{1,3})\.0$') {
            # Subnet notation: "192.168.1.0" expands to 192.168.1.1 through 192.168.1.254
            # Includes .1 (gateway) but excludes .255 (broadcast address)
            $base = $matches[1]
            1..254 | ForEach-Object { [void]$allIps.Add("$base.$_") }
        }
        elseif ($range -match '^(\d{1,3}\.\d{1,3}\.\d{1,3})\.(\d{1,3})-(\d{1,3})$') {
            # Range notation: "192.168.1.10-20" expands to all IPs in the range
            $base = $matches[1]
            ([int]$matches[2])..([int]$matches[3]) | ForEach-Object { [void]$allIps.Add("$base.$_") }
        }
        elseif ($range -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
            # Single IP address: add directly to list
            [void]$allIps.Add($range)
        }
        # Invalid formats are silently ignored - validation occurs in Validate-Inputs function
    }
    
    return $allIps
}

function Validate-Inputs {
    <#
    .SYNOPSIS
        Validates all user input from the configuration form before starting the scan.
    
    .DESCRIPTION
        Performs comprehensive validation of user-provided settings including:
        - IP range format validation (supports single IPs, ranges, and subnets)
        - Network address boundary checking (excludes broadcast addresses)
        - SNMP community string format validation
        - Output filename character validation
        
        Displays user-friendly error messages for validation failures.
    
    .PARAMETER inputs
        PSCustomObject containing all user input from the configuration form
    
    .OUTPUTS
        Boolean - True if all inputs are valid, False if any validation fails
    
    .NOTES
        Validation rules:
        - IP octets 1-3: Must be 1-254 (excludes 0.x.x.x and 255.x.x.x networks)
        - IP octet 4: Must be 2-254 for specific IPs, or 0 for subnet notation
        - SNMP community: 1-32 chars, alphanumeric plus @#$%&* symbols
        - Filename: Must not contain filesystem-invalid characters
    
    .EXAMPLE
        $isValid = Validate-Inputs -inputs $userSettings
        if ($isValid) { Start-NetworkScan }
    #>
    
    param([Parameter(Mandatory = $true)][PSCustomObject]$inputs)
    
    # =============================================================================
    # IP RANGE VALIDATION
    # =============================================================================
    
    # Remove spaces and check for empty input
    $ipRanges = $inputs.IpRanges.Replace(" ", "")
    if ([string]::IsNullOrWhiteSpace($ipRanges)) {
        [System.Windows.Forms.MessageBox]::Show(
            "IP Address Ranges cannot be empty.", 
            "Validation Error", 
            "OK", 
            "Warning"
        )
        return $false
    }
    
    # Validate each comma-separated range
    foreach ($range in ($ipRanges -split ',')) {
        # Check basic IP range format using regex
        if ($range -notmatch '^(\d{1,3}\.\d{1,3}\.\d{1,3})\.(\d{1,3}(?:-\d{1,3})?|0)$') {
            [System.Windows.Forms.MessageBox]::Show(
                "Invalid IP range format: '$range'. Use formats like '192.168.1.5', '192.168.1.10-20', or '192.168.1.0'.", 
                "Validation Error", 
                "OK", 
                "Warning"
            )
            return $false
        }
        
        # Validate individual octets (first three must be 1-254)
        $parts = $range -split '\.'
        for ($i = 0; $i -lt 3; $i++) {
            if ([int]$parts[$i] -lt 1 -or [int]$parts[$i] -gt 254) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Invalid octet value in '$range'. The first three octets must be between 1 and 254.", 
                    "Validation Error", 
                    "OK", 
                    "Warning"
                )
                return $false
            }
        }
        
        # Validate fourth octet (different rules for ranges vs single IPs)
        if ($parts[3] -match '(\d+)-(\d+)') {
            # Range format: validate start and end values
            if ([int]$matches[1] -lt 1 -or [int]$matches[2] -gt 254 -or [int]$matches[1] -ge [int]$matches[2]) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Invalid range in '$range'. Range must be between 1 and 254, and the start must be less than the end.", 
                    "Validation Error", 
                    "OK", 
                    "Warning"
                )
                return $false
            }
        }
        elseif ($parts[3] -ne '0') {
            # Single IP: validate host portion
            if ([int]$parts[3] -lt 1 -or [int]$parts[3] -gt 254) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Invalid host value in '$range'. The fourth octet must be between 1 and 254 (or 0 for a full range).", 
                    "Validation Error", 
                    "OK", 
                    "Warning"
                )
                return $false
            }
        }
    }
    
    # =============================================================================
    # SNMP COMMUNITY STRING VALIDATION
    # =============================================================================
    
    # SNMP community strings have specific character and length restrictions
    # Support single community string or comma-separated multiple strings
    # Each community string must be 1-32 characters with allowed characters
    if ($inputs.SnmpCommunity -notmatch '^[a-zA-Z0-9@#$%\&\*]{1,32}(\s*,\s*[a-zA-Z0-9@#$%\&\*]{1,32})*$') {
        [System.Windows.Forms.MessageBox]::Show(
            "Community String(s) must be 1-32 characters each and can only contain letters, numbers, and the symbols: @#$%&*. Multiple communities can be separated by commas (e.g., 'public, medialinks, custom').", 
            "Validation Error", 
            "OK", 
            "Warning"
        )
        return $false
    }
    
    # =============================================================================
    # OUTPUT FILENAME VALIDATION
    # =============================================================================
    
    # Check for filesystem-invalid characters in the output filename
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $regexInvalid = "[{0}]" -f [System.Text.RegularExpressions.Regex]::Escape($invalidChars)
    if ($inputs.OutputFileName -match $regexInvalid) {
        [System.Windows.Forms.MessageBox]::Show(
            "Output File Name contains invalid characters.", 
            "Validation Error", 
            "OK", 
            "Warning"
        )
        return $false
    }
    
    # All validations passed
    return $true
}

function Get-SystemCapabilities {
    <#
    .SYNOPSIS
        Evaluates system hardware capabilities for optimal parallel scan recommendations.
    .DESCRIPTION
        Assesses CPU cores, logical processors, available memory, and system performance
        characteristics to determine the system's capacity for parallel network scanning.
    .OUTPUTS
        PSCustomObject with system capability metrics
    #>
    try {
        # Get CPU information
        $cpu = Get-WmiObject -Class Win32_Processor -ErrorAction SilentlyContinue
        $logicalProcessors = if ($cpu) { ($cpu | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum } else { 4 }
        $physicalCores = if ($cpu) { ($cpu | Measure-Object -Property NumberOfCores -Sum).Sum } else { 2 }
        
        # Get memory information (in GB)
        $memory = try {
            [math]::Round((Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue | 
                         Measure-Object -Property Capacity -Sum).Sum / 1GB)
        } catch { 8 } # Default fallback
        
        # Calculate performance category
        $performanceCategory = "Low"
        if ($physicalCores -ge 8 -and $memory -ge 16 -and $logicalProcessors -ge 12) {
            $performanceCategory = "High"
        } elseif ($physicalCores -ge 4 -and $memory -ge 8 -and $logicalProcessors -ge 6) {
            $performanceCategory = "Medium"
        }
        
        return [PSCustomObject]@{
            LogicalProcessors = $logicalProcessors
            PhysicalCores = $physicalCores
            MemoryGB = $memory
            PerformanceCategory = $performanceCategory
            MaxRecommendedJobs = [math]::Min($logicalProcessors * 2, 50) # Cap at 50 for safety
        }
    }
    catch {
        # Return conservative defaults on any error
        return [PSCustomObject]@{
            LogicalProcessors = 4
            PhysicalCores = 2
            MemoryGB = 8
            PerformanceCategory = "Low"
            MaxRecommendedJobs = 8
        }
    }
}

function Get-ScanComplexity {
    <#
    .SYNOPSIS
        Analyzes IP range complexity and scan scope to determine resource requirements.
    .PARAMETER IpRanges
        String containing IP ranges to analyze (same format as GUI input)
    .OUTPUTS
        PSCustomObject with scan complexity metrics
    #>
    param([string]$IpRanges)
    
    try {
        $totalIPs = 0
        $rangeCount = 0
        $complexityLevel = "Low"
        
        if ([string]::IsNullOrWhiteSpace($IpRanges)) {
            return [PSCustomObject]@{
                TotalIPs = 0
                RangeCount = 0
                ComplexityLevel = "Low"
                EstimatedScanTime = 0
            }
        }
        
        # Parse ranges using existing Parse-IpRanges function
        $parsedRanges = Parse-IpRanges -IpRanges $IpRanges
        if ($parsedRanges) {
            $totalIPs = $parsedRanges.Count
            $rangeCount = ($IpRanges -split ',').Count
        }
        
        # Determine complexity based on IP count and range distribution
        if ($totalIPs -gt 500) {
            $complexityLevel = "High"
        } elseif ($totalIPs -gt 100 -or $rangeCount -gt 5) {
            $complexityLevel = "Medium"
        }
        
        # Estimate scan time (rough calculation: ~0.5-4 seconds per IP depending on responsiveness)
        $estimatedScanTime = [math]::Ceiling($totalIPs * 2.0) # Conservative estimate
        
        return [PSCustomObject]@{
            TotalIPs = $totalIPs
            RangeCount = $rangeCount
            ComplexityLevel = $complexityLevel
            EstimatedScanTime = $estimatedScanTime
        }
    }
    catch {
        return [PSCustomObject]@{
            TotalIPs = 0
            RangeCount = 0
            ComplexityLevel = "Low"
            EstimatedScanTime = 0
        }
    }
}

function Get-RecommendedParallelScans {
    <#
    .SYNOPSIS
        Recommends optimal parallel scan count based on system capabilities and scan complexity.
    .PARAMETER IpRanges
        String containing IP ranges to scan
    .OUTPUTS
        PSCustomObject with recommendation details
    #>
    param([string]$IpRanges)
    
    try {
        $systemCaps = Get-SystemCapabilities
        $scanComplexity = Get-ScanComplexity -IpRanges $IpRanges
        
        # Base recommendation on system performance category
        $baseRecommendation = switch ($systemCaps.PerformanceCategory) {
            "High"   { 20 }
            "Medium" { 12 }
            "Low"    { 6 }
            default  { 8 }
        }
        
        # Adjust based on scan complexity
        $complexityMultiplier = switch ($scanComplexity.ComplexityLevel) {
            "High"   { 1.2 }  # Increase for large scans (more parallel processing beneficial)
            "Medium" { 1.0 }  # No adjustment
            "Low"    { 0.8 }  # Decrease for small scans (overhead not worth it)
            default  { 1.0 }
        }
        
        # Calculate recommended value
        $recommended = [math]::Round($baseRecommendation * $complexityMultiplier)
        
        # Apply safety constraints
        $recommended = [math]::Max($recommended, 2)  # Minimum of 2
        $recommended = [math]::Min($recommended, $systemCaps.MaxRecommendedJobs)  # Cap at system max
        $recommended = [math]::Min($recommended, [math]::Max($scanComplexity.TotalIPs / 2, 1))  # Don't exceed half the IP count
        
        # Create explanation text
        $explanation = "System: $($systemCaps.PerformanceCategory) ($($systemCaps.PhysicalCores) cores, $($systemCaps.MemoryGB)GB RAM)`n"
        $explanation += "Scan: $($scanComplexity.TotalIPs) IPs, $($scanComplexity.ComplexityLevel) complexity`n"
        $explanation += "Est. time with $recommended parallel: $([math]::Round($scanComplexity.EstimatedScanTime / $recommended / 60, 1)) minutes"
        
        return [PSCustomObject]@{
            RecommendedCount = $recommended
            SystemCapabilities = $systemCaps
            ScanComplexity = $scanComplexity
            Explanation = $explanation
            PerformanceGain = if ($scanComplexity.TotalIPs -gt 0) { 
                [math]::Round(([math]::Min($recommended, $scanComplexity.TotalIPs) / 1) * 100, 0) 
            } else { 100 }
        }
    }
    catch {
        return [PSCustomObject]@{
            RecommendedCount = 8
            SystemCapabilities = $null
            ScanComplexity = $null
            Explanation = "Unable to analyze system - using safe default of 8 parallel scans"
            PerformanceGain = 100
        }
    }
}
#endregion

#region GUI Creation

# =============================================================================
# GUI FORM CREATION AND LAYOUT
# =============================================================================

function Create-ConfigForm {
    <#
    .SYNOPSIS
        Creates the main configuration form with all user input controls.
    
    .DESCRIPTION
        Builds a Windows Forms dialog containing all configuration options for the network scan.
        Uses a consistent layout pattern with labels on the left and controls on the right.
        Includes input validation, tooltips, and automatic recommendations for optimal settings.
    
    .PARAMETER Settings
        PSCustomObject containing current settings to populate the form controls
    
    .OUTPUTS
        Hashtable containing references to the form and all its controls for easy access
    
    .NOTES
        Form Layout:
        - Fixed size dialog (420x480) that cannot be resized
        - Centered on screen with consistent 30px vertical spacing
        - All controls aligned for professional appearance
        - OK button serves as default action (Enter key)
    
    .EXAMPLE
        $formElements = Create-ConfigForm -Settings $currentSettings
        $result = $formElements.Form.ShowDialog()
        if ($result -eq 'OK') { $ipRanges = $formElements.IpRangesBox.Text }
    #>
    
    param([Parameter(Mandatory = $true)][PSCustomObject]$Settings)
    
    # =============================================================================
    # MAIN FORM CONFIGURATION
    # =============================================================================
    
    # Create the main form window with fixed dimensions and behavior
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "DiscoverSubnet v$scriptVersion - Configuration"
    $form.Size = New-Object System.Drawing.Size(420, 480)
    $form.FormBorderStyle = 'FixedDialog'    # Prevents resizing
    $form.StartPosition = 'CenterScreen'     # Center on user's screen
    $form.MaximizeBox = $false               # Disable maximize button
    $form.MinimizeBox = $false               # Disable minimize button
    
    # Layout constants for consistent control positioning
    $yPos = 15                               # Current vertical position (incremented for each row)
    $labelWidth = 160                        # Width of all label controls
    $controlWidth = 210                      # Width of all input controls
    
    # =============================================================================
    # NETWORK CONFIGURATION CONTROLS
    # =============================================================================
    
    # IP Address Ranges input - supports multiple formats (single, range, subnet)
    $label = New-Object System.Windows.Forms.Label; $label.Text = "IP Address Ranges:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $ipRangesBox = New-Object System.Windows.Forms.TextBox; $ipRangesBox.Location = New-Object System.Drawing.Point(180, $yPos); $ipRangesBox.Size = New-Object System.Drawing.Size($controlWidth, 20); $ipRangesBox.Text = $Settings.IpRanges; $ipRangesBox.Tag = "IpRanges"; $form.Controls.Add($ipRangesBox); $yPos += 30
    
    # SNMP Community String - supports multiple comma-separated values for device identification queries
    $label = New-Object System.Windows.Forms.Label; $label.Text = "SNMP Community Strings:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $communityBox = New-Object System.Windows.Forms.TextBox; $communityBox.Location = New-Object System.Drawing.Point(180, $yPos); $communityBox.Size = New-Object System.Drawing.Size($controlWidth, 20); $communityBox.Text = $Settings.SnmpCommunity; $communityBox.Tag = "Separate multiple communities with commas (e.g., medialinks, custom, private)"; $form.Controls.Add($communityBox); $yPos += 30
    
    # Retry count for failed ping/SNMP attempts
    $label = New-Object System.Windows.Forms.Label; $label.Text = "Ping/SNMP Retries:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $retriesDropdown = New-Object System.Windows.Forms.ComboBox; $retriesDropdown.Location = New-Object System.Drawing.Point(180, $yPos); $retriesDropdown.Size = New-Object System.Drawing.Size($controlWidth, 20); $retriesDropdown.DropDownStyle = 'DropDownList'; [void]$retriesDropdown.Items.AddRange(@(0, 1, 2, 3)); $retriesDropdown.SelectedItem = $Settings.Retries; $form.Controls.Add($retriesDropdown); $yPos += 30
    
    # =============================================================================
    # OUTPUT CONFIGURATION CONTROLS  
    # =============================================================================
    
    # Output filename (without extension)
    $label = New-Object System.Windows.Forms.Label; $label.Text = "Output File Name:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $fileNameBox = New-Object System.Windows.Forms.TextBox; $fileNameBox.Location = New-Object System.Drawing.Point(180, $yPos); $fileNameBox.Size = New-Object System.Drawing.Size($controlWidth, 20); $fileNameBox.Text = $Settings.OutputFileName; $form.Controls.Add($fileNameBox); $yPos += 30
    
    # Output file format selection
    $label = New-Object System.Windows.Forms.Label; $label.Text = "Output File Type:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $fileTypeDropdown = New-Object System.Windows.Forms.ComboBox; $fileTypeDropdown.Location = New-Object System.Drawing.Point(180, $yPos); $fileTypeDropdown.Size = New-Object System.Drawing.Size($controlWidth, 20); $fileTypeDropdown.DropDownStyle = 'DropDownList'; [void]$fileTypeDropdown.Items.AddRange(@('txt', 'csv')); $fileTypeDropdown.SelectedItem = $Settings.OutputFileExtension; $form.Controls.Add($fileTypeDropdown); $yPos += 30
    
    # =============================================================================
    # DISPLAY AND PERFORMANCE CONTROLS
    # =============================================================================
    
    # GUI verbosity level - controls amount of information displayed during scan
    $label = New-Object System.Windows.Forms.Label; $label.Text = "GUI Display Level:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $guiVerbosityDropdown = New-Object System.Windows.Forms.ComboBox; $guiVerbosityDropdown.Location = New-Object System.Drawing.Point(180, $yPos); $guiVerbosityDropdown.Size = New-Object System.Drawing.Size($controlWidth, 20); $guiVerbosityDropdown.DropDownStyle = 'DropDownList'; [void]$guiVerbosityDropdown.Items.AddRange(@('Standard', 'Minimal')); $guiVerbosityDropdown.SelectedItem = $Settings.GuiVerbosity; $form.Controls.Add($guiVerbosityDropdown); $yPos += 30
    
    # Parallel scan count with automatic recommendation button
    $label = New-Object System.Windows.Forms.Label; $label.Text = "Max Parallel Scans:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $parallelScansUpDown = New-Object System.Windows.Forms.NumericUpDown; $parallelScansUpDown.Location = New-Object System.Drawing.Point(180, $yPos); $parallelScansUpDown.Size = New-Object System.Drawing.Size(($controlWidth - 70), 20); $parallelScansUpDown.Minimum = 1; $parallelScansUpDown.Maximum = 100; $parallelScansUpDown.Value = $Settings.MaxParallelScans; $parallelScansUpDown.Tag = "ParallelScans"; $form.Controls.Add($parallelScansUpDown)
    $recommendButton = New-Object System.Windows.Forms.Button; $recommendButton.Text = "Auto"; $recommendButton.Location = New-Object System.Drawing.Point((180 + $controlWidth - 65), $yPos); $recommendButton.Size = New-Object System.Drawing.Size(60, 22); $recommendButton.UseVisualStyleBackColor = $true; $form.Controls.Add($recommendButton)
    # =============================================================================
    # AUTOMATIC RECOMMENDATION EVENT HANDLER
    # =============================================================================
    
    # Configure the "Auto" button to analyze system capabilities and IP ranges
    # for optimal parallel scan count recommendation
    $recommendButton.Add_Click({
        param($clickSender, $clickEvent)
        try {
            # Locate the parent form and find tagged controls for recommendation analysis
            $parentForm = $clickSender.FindForm()
            $ipRangesTextBox = $null
            $parallelScansControl = $null
            
            # Search through all form controls to find our tagged controls
            # This approach is more robust than direct variable references in closures
            foreach ($control in $parentForm.Controls) {
                if ($control.Tag -eq "IpRanges") {
                    $ipRangesTextBox = $control
                }
                elseif ($control.Tag -eq "ParallelScans") {
                    $parallelScansControl = $control
                }
            }
            
            if ($ipRangesTextBox -and $parallelScansControl) {
                # Generate recommendation based on current IP ranges and system capabilities
                $recommendation = Get-RecommendedParallelScans -IpRanges $ipRangesTextBox.Text
                $parallelScansControl.Value = $recommendation.RecommendedCount
                
                # Display explanation of the recommendation to the user
                [System.Windows.Forms.MessageBox]::Show(
                    $recommendation.Explanation, 
                    "Parallel Scans Recommendation", 
                    "OK", 
                    "Information"
                )
            } else {
                # Debug information for troubleshooting control location issues
                $debugInfo = "Could not locate form controls. Found controls: "
                foreach ($control in $parentForm.Controls) {
                    $debugInfo += "$($control.GetType().Name)($($control.Tag)), "
                }
                [System.Windows.Forms.MessageBox]::Show("$debugInfo", "Debug Info", "OK", "Information")
                [System.Windows.Forms.MessageBox]::Show(
                    "Could not locate form controls for recommendation.", 
                    "Recommendation Error", 
                    "OK", 
                    "Warning"
                )
            }
        }
        catch {
            # Handle any errors in the recommendation process
            $errorMessage = "Error details: $($_.Exception.Message)`nStack: $($_.ScriptStackTrace)"
            [System.Windows.Forms.MessageBox]::Show(
                "Unable to generate recommendation.$([Environment]::NewLine)$errorMessage", 
                "Recommendation Error", 
                "OK", 
                "Warning"
            )
        }
    })
    
    # =============================================================================
    # DIAGNOSTIC AND OPTION CONTROLS
    # =============================================================================
    
    $yPos += 30
    
    # Diagnostic logging level - controls amount of technical detail in logs
    $label = New-Object System.Windows.Forms.Label; $label.Text = "Diagnostic Level:"; $label.Location = New-Object System.Drawing.Point(20, $yPos); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $form.Controls.Add($label)
    $diagDropdown = New-Object System.Windows.Forms.ComboBox; $diagDropdown.Location = New-Object System.Drawing.Point(180, $yPos); $diagDropdown.Size = New-Object System.Drawing.Size($controlWidth, 20); $diagDropdown.DropDownStyle = 'DropDownList'; [void]$diagDropdown.Items.AddRange(@('Off', 'Standard', 'Verbose')); $diagDropdown.SelectedItem = $Settings.DiagnosticLevel; $form.Controls.Add($diagDropdown); $yPos += 40
    
    # Option to include unresponsive devices in output file
    $saveUnresponsiveCheck = New-Object System.Windows.Forms.CheckBox; $saveUnresponsiveCheck.Text = "Save SNMP-unresponsive devices to output file"; $saveUnresponsiveCheck.Location = New-Object System.Drawing.Point(20, $yPos); $saveUnresponsiveCheck.Size = New-Object System.Drawing.Size(370, 20); $saveUnresponsiveCheck.Checked = $Settings.SaveUnresponsive; $form.Controls.Add($saveUnresponsiveCheck); $yPos += 40
    
    # =============================================================================
    # FORM ACTION BUTTON
    # =============================================================================
    
    # Main action button - starts the network discovery process
    $startButton = New-Object System.Windows.Forms.Button; $startButton.Text = "Start Discovery"; $startButton.Location = New-Object System.Drawing.Point(150, $yPos); $startButton.Size = New-Object System.Drawing.Size(120, 40); $startButton.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Controls.Add($startButton); $form.AcceptButton = $startButton
    
    # Return hashtable containing form and all controls for easy access by calling code
    return @{ Form = $form; IpRangesBox = $ipRangesBox; CommunityBox = $communityBox; RetriesDropdown = $retriesDropdown; FileNameBox = $fileNameBox; FileTypeDropdown = $fileTypeDropdown; GuiVerbosityDropdown = $guiVerbosityDropdown; ParallelScans = $parallelScansUpDown; DiagDropdown = $diagDropdown; SaveUnresponsive = $saveUnresponsiveCheck }
}

function Create-LogForm {
    <#
    .SYNOPSIS
        Creates the progress monitoring and logging form displayed during network scanning.
    
    .DESCRIPTION
        Builds a resizable Windows Forms dialog that displays real-time scan progress and logging.
        Features a large scrollable text area for log messages, a progress bar, and action buttons
        for controlling the scan and saving results.
    
    .OUTPUTS
        Hashtable containing references to the form and all its controls
    
    .NOTES
        Form Layout:
        - Resizable window (800x600 default) for maximum log visibility
        - Docked text area that expands with window resizing
        - Fixed bottom panel with progress bar and action buttons
        - Monospace font (Consolas) for better log readability
        
        Button States:
        - Save Results: Disabled during scan, enabled when complete
        - Cancel Scan: Enabled during scan, disabled when complete
        - Close: Disabled during scan, enabled when complete
    
    .EXAMPLE
        $logFormElements = Create-LogForm
        $logFormElements.Form.ShowDialog()
    #>
    
    # =============================================================================
    # MAIN PROGRESS FORM CONFIGURATION
    # =============================================================================
    
    # Create resizable progress monitoring window
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "DiscoverSubnet v$scriptVersion - In Progress"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = 'CenterScreen'
    # Note: Form is resizable (default) to allow users to expand for better log viewing
    
    # =============================================================================
    # LOG DISPLAY AREA
    # =============================================================================
    
    # Create main log text area with professional monospace formatting and color support
    $logBox = New-Object System.Windows.Forms.RichTextBox  # Changed from TextBox to RichTextBox for color support
    $logBox.Multiline = $true                                    # Enable multi-line display
    $logBox.ScrollBars = 'Vertical'                             # Add vertical scrollbar
    $logBox.ReadOnly = $true                                    # Prevent user editing
    $logBox.Dock = 'Fill'                                       # Expand to fill available space
    $logBox.Font = New-Object System.Drawing.Font("Consolas", 9) # Monospace font for aligned output
    $logBox.BackColor = [System.Drawing.Color]::Black           # Dark background for better contrast
    $logBox.ForeColor = [System.Drawing.Color]::White           # Default white text
    $form.Controls.Add($logBox)
    
    # =============================================================================
    # BOTTOM CONTROL PANEL
    # =============================================================================
    
    # Create fixed bottom panel for progress bar and action buttons
    $bottomPanel = New-Object System.Windows.Forms.Panel
    $bottomPanel.Dock = 'Bottom'                                # Dock to bottom of form
    $bottomPanel.Height = 50                                    # Fixed height for consistent layout
    $form.Controls.Add($bottomPanel)
    
    # Progress bar - shows scan completion percentage
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 15)
    $progressBar.Size = New-Object System.Drawing.Size(420, 23)
    $bottomPanel.Controls.Add($progressBar)
    
    # =============================================================================
    # ACTION BUTTONS
    # =============================================================================
    
    # Save Results button - disabled during scan, enabled when complete
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save Results"
    $saveButton.Location = New-Object System.Drawing.Point(440, 12)
    $saveButton.Enabled = $false                                # Disabled until scan completes
    $bottomPanel.Controls.Add($saveButton)
    
    # Cancel Scan button - allows user to stop scan in progress
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel Scan"
    $cancelButton.Location = New-Object System.Drawing.Point(530, 12)
    # Enabled by default - user can cancel at any time during scan
    $bottomPanel.Controls.Add($cancelButton)
    
    # Close button - disabled during scan, enabled when complete
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(620, 12)
    $closeButton.Enabled = $false                               # Disabled until scan completes
    $bottomPanel.Controls.Add($closeButton)
    
    # Return hashtable with form and control references for event handler binding
    return @{ 
        Form = $form
        LogBox = $logBox
        ProgressBar = $progressBar
        SaveButton = $saveButton
        CancelButton = $cancelButton
        CloseButton = $closeButton 
    }
}
#endregion

#region Main Script Execution

# =============================================================================
# MAIN APPLICATION WORKFLOW
# =============================================================================

# This section orchestrates the complete application workflow:
# 1. Load user settings and display configuration form
# 2. Validate inputs and show progress form
# 3. Start background scanning jobs
# 4. Monitor progress and handle user interactions
# 5. Save results and cleanup

# Load persistent settings from JSON file (or create defaults)
$currentSettings = Load-Settings

# Create and display the main configuration form
$configFormElements = Create-ConfigForm -Settings $currentSettings

# =============================================================================
# INTELLIGENT PERFORMANCE RECOMMENDATIONS
# =============================================================================

# Automatically analyze system capabilities and provide optimal parallel scan recommendations
# This helps users achieve best performance without manual performance tuning
try {
    $initialRecommendation = Get-RecommendedParallelScans -IpRanges $configFormElements.IpRangesBox.Text
    if ($initialRecommendation.RecommendedCount -ne $configFormElements.ParallelScans.Value) {
        # Update the form control with the recommended value
        $configFormElements.ParallelScans.Value = $initialRecommendation.RecommendedCount
        
        # Add a tooltip to explain the reasoning behind the recommendation
        $tooltip = New-Object System.Windows.Forms.ToolTip
        $tooltip.SetToolTip($configFormElements.ParallelScans, "Auto-recommended: $($initialRecommendation.Explanation)")
        
        # Add tooltip for community strings to explain multiple community support
        $communityTooltip = New-Object System.Windows.Forms.ToolTip
        $communityTooltip.SetToolTip($configFormElements.CommunityBox, "Enter multiple SNMP community strings separated by commas. Tool tries 'public' first, then your specified communities in order. Example: medialinks, custom, private")
    }
}
catch {
    # If recommendation analysis fails, silently continue with user's current settings
    # This ensures the application remains functional even if performance analysis fails
}

# =============================================================================
# USER CONFIGURATION AND VALIDATION
# =============================================================================

# Display configuration form and wait for user input
$configFormElements.Form.ShowDialog() | Out-Null

# Process user configuration only if they clicked "Start Discovery"
if ($configFormElements.Form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {

    # Extract all settings from form controls into a structured configuration object
    $scanSettings = [PSCustomObject]@{
        IpRanges              = $configFormElements.IpRangesBox.Text
        SnmpCommunity         = $configFormElements.CommunityBox.Text
        Retries               = [int]$configFormElements.RetriesDropdown.SelectedItem
        OutputFileName        = $configFormElements.FileNameBox.Text
        OutputFileExtension   = $configFormElements.FileTypeDropdown.SelectedItem
        SaveUnresponsive      = $configFormElements.SaveUnresponsive.Checked
        MaxParallelScans      = [int]$configFormElements.ParallelScans.Value
        DiagnosticLevel       = $configFormElements.DiagDropdown.SelectedItem
        GuiVerbosity          = $configFormElements.GuiVerbosityDropdown.SelectedItem
    }

    # Validate all user inputs before proceeding with scan
    if (-not (Validate-Inputs -inputs $scanSettings)) { exit }
    
    # Persist settings for future use
    Save-Settings -Settings $scanSettings

    $logFormElements = Create-LogForm
    $logForm = $logFormElements.Form; $script:logTextBox = $logFormElements.LogBox; $progressBar = $logFormElements.ProgressBar
    $cancelButton = $logFormElements.CancelButton; $saveButton = $logFormElements.SaveButton; $closeButton = $logFormElements.CloseButton

    $global:scanJob = $null; $global:allDiscoveredDevices = New-Object System.Collections.Generic.List[PSCustomObject]
    $script:unreachableIPs = @()  # Reset unreachable IPs tracking for new scan
    $logFileName = "DiscoverSubnet-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
    $logFilePath = Join-Path -Path $scriptDir -ChildPath $logFileName
    $outputFilePath = Join-Path -Path $scriptDir -ChildPath "$($scanSettings.OutputFileName).$($scanSettings.OutputFileExtension)"

    # =============================================================================
    # LOGGING AND ERROR HANDLING INFRASTRUCTURE
    # =============================================================================

    # Initialize log file with version header for troubleshooting and audit purposes
    $versionHeader = "DiscoverSubnet Version $scriptVersion"
    Add-Content -Path $logFilePath -Value $versionHeader

    function Write-Log {
        <#
        .SYNOPSIS
            Thread-safe logging function that writes to both file and GUI with intelligent filtering.
        
        .DESCRIPTION
            Implements a robust logging strategy with multiple error handling fallbacks:
            
            ERROR HANDLING STRATEGIES:
            1. Thread-safe GUI updates using BeginInvoke for cross-thread operations
            2. Fallback to direct GUI access if BeginInvoke fails
            3. Silent continuation if GUI controls are disposed or unavailable
            4. Complete logging to file regardless of GUI state
            
            PS2EXE COMPATIBILITY:
            - Uses file paths resolved during script initialization
            - Handles control disposal gracefully in compiled executables
            - Thread-safe operations work correctly in PS2EXE environment
        
        .PARAMETER Message
            The message text to log
        
        .PARAMETER MessageType
            Classification for GUI filtering: General, Diagnostic, RangeStart, RangeEnd, PingResult, ScanResult
        #>
        
        param(
            [string]$Message,
            [string]$MessageType = "General"  # General, Diagnostic, RangeStart, RangeEnd, PingResult, ScanResult
        )
        
        # Create timestamped log entry
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] $Message"
        
        # ALWAYS write to log file (complete audit trail regardless of GUI state)
        Add-Content -Path $logFilePath -Value $logEntry
        
        # =============================================================================
        # INTELLIGENT GUI DISPLAY FILTERING
        # =============================================================================
        
        # Apply user-selected verbosity filtering to reduce GUI clutter
        $showInGui = $true
        $addBlankLine = $false
        
        if ($scanSettings.GuiVerbosity -eq "Minimal") {
            # Minimal: only show range start/end messages, general status, and discovered devices report
            $showInGui = ($MessageType -in @("RangeStart", "RangeEnd", "General", "ResultsSeparator", "DeviceResult", "Success"))
            if ($MessageType -eq "RangeEnd") { $addBlankLine = $true }
        } elseif ($scanSettings.GuiVerbosity -eq "Standard") {
            # Standard: show scan parameters, range info, device results, and discovered devices report
            $showInGui = ($MessageType -in @("ScanParameters", "RangeStart", "RangeEnd", "ScanResult", "General", "ResultsSeparator", "DeviceResult", "Success"))
            if ($MessageType -eq "RangeEnd") { $addBlankLine = $true }
        }
        # Default: show everything (full verbosity)
        
        # =============================================================================
        # THREAD-SAFE GUI UPDATES WITH COLOR CODING AND FALLBACK ERROR HANDLING
        # =============================================================================
        
        if ($showInGui -and $script:logTextBox -and -not $script:logTextBox.IsDisposed) {
            try {
                # Ensure control handle is created before attempting cross-thread operations
                if (-not $script:logTextBox.IsHandleCreated) {
                    $script:logTextBox.CreateControl()
                }
                
                if ($script:logTextBox.IsHandleCreated) {
                    # PRIMARY: Use BeginInvoke for thread-safe GUI updates with color coding
                    $script:logTextBox.BeginInvoke([Action[string,string]]{ param($text, $msgType)
                        if ($script:logTextBox -and -not $script:logTextBox.IsDisposed) { 
                            # Determine color based on message type
                            $textColor = switch ($msgType) {
                                "General"           { [System.Drawing.Color]::White }         # Default white
                                "RangeStart"        { [System.Drawing.Color]::Yellow }        # Yellow for range start
                                "RangeEnd"          { [System.Drawing.Color]::Cyan }          # Cyan for range end  
                                "ScanResult"        { [System.Drawing.Color]::LightGreen }    # Light green for successful scans
                                "PingResult"        { [System.Drawing.Color]::LightBlue }     # Light blue for ping results
                                "Diagnostic"        { [System.Drawing.Color]::Gray }          # Gray for diagnostic info
                                "ScanParameters"    { [System.Drawing.Color]::Orange }        # Orange for scan parameters
                                "UnreachableSummary"{ [System.Drawing.Color]::Gray }          # Gray for unreachable summary (less distraction)
                                "SectionSeparator"  { [System.Drawing.Color]::DarkCyan }      # Dark cyan for section separators
                                "ResultsSeparator"  { [System.Drawing.Color]::White }         # White for discovered devices report header
                                "DeviceResult"      { [System.Drawing.Color]::LightGreen }    # Light green for discovered devices
                                "Error"             { [System.Drawing.Color]::Red }           # Red for errors
                                "Warning"           { [System.Drawing.Color]::Yellow }        # Yellow for warnings
                                "Success"           { [System.Drawing.Color]::Green }         # Green for success messages
                                default             { [System.Drawing.Color]::White }         # Default white
                            }
                            
                            # Add colored text to RichTextBox
                            $script:logTextBox.SelectionStart = $script:logTextBox.TextLength
                            $script:logTextBox.SelectionLength = 0
                            $script:logTextBox.SelectionColor = $textColor
                            $script:logTextBox.AppendText($text + [Environment]::NewLine)
                            $script:logTextBox.SelectionColor = $script:logTextBox.ForeColor  # Reset to default
                            $script:logTextBox.ScrollToCaret()  # Auto-scroll to bottom
                        }
                    }, $logEntry, $MessageType)
                    
                    # Add visual spacing for better readability
                    if ($addBlankLine) {
                        $script:logTextBox.BeginInvoke([Action]{ 
                            if ($script:logTextBox -and -not $script:logTextBox.IsDisposed) { 
                                $script:logTextBox.AppendText([Environment]::NewLine) 
                            }
                        })
                    }
                }
            }
            catch {
                # FALLBACK: If BeginInvoke fails, use direct access (less thread-safe but functional)  
                # This handles edge cases in PS2EXE environments or unusual threading scenarios
                Write-Warning "BeginInvoke failed, using direct access: $($_.Exception.Message)"
                
                if ($script:logTextBox -and -not $script:logTextBox.IsDisposed) {
                    # Fallback to plain text without colors
                    $script:logTextBox.AppendText($logEntry + [Environment]::NewLine)
                    if ($addBlankLine) {
                        $script:logTextBox.AppendText([Environment]::NewLine)
                    }
                }
            }
        }
        # NOTE: If GUI controls are unavailable/disposed, logging continues silently to file
        # This ensures the application remains functional even with GUI issues
    }

    # =============================================================================
    # BACKGROUND JOB CONTROLLER SCRIPT BLOCK
    # =============================================================================
    
    # This script block runs in a separate PowerShell job to perform the actual network scanning
    # It coordinates multiple worker jobs and communicates progress back to the GUI via structured messages
    # 
    # ARCHITECTURE OVERVIEW:
    # ┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
    # │   Main GUI      │◄──►│   Controller     │◄──►│  Worker Jobs    │
    # │   Thread        │    │   Script Block   │    │  (Per IP)       │
    # │                 │    │  (Background)    │    │                 │
    # │ - Progress Bar  │    │ - Job Management │    │ - Ping Test     │
    # │ - Log Display   │    │ - Result Agg.    │    │ - SNMP Query    │
    # │ - User Controls │    │ - Status Updates │    │ - Device ID     │
    # └─────────────────┘    └──────────────────┘    └─────────────────┘
    
    $controllerScriptBlock = {
        param($Settings)
        
        # =============================================================================
        # CONTROLLER INITIALIZATION AND SAFETY CHECKS
        # =============================================================================
        
        # Helper function to create structured messages for GUI communication
        function New-JobMessage { 
            param($Type, $Value) 
            [PSCustomObject]@{Type = $Type; Value = $Value} 
        }
        
        # Generate unique controller ID for debugging multiple instances
        $controllerId = [System.Guid]::NewGuid().ToString().Substring(0,8)
        $global:controllerExecuting = $true
        
        # Display professional scan initialization header with all parameters
        Write-Output (New-JobMessage -Type "Log" -Value "=== SCAN PARAMETERS ===")
        Write-Output (New-JobMessage -Type "Log" -Value "IP Ranges: $($Settings.IpRanges)")
        $communityDisplay = if ($Settings.SnmpCommunity -match ',') { 
            "public, $($Settings.SnmpCommunity)" 
        } else { 
            "public, $($Settings.SnmpCommunity)" 
        }
        Write-Output (New-JobMessage -Type "Log" -Value "SNMP Communities: $communityDisplay")
        Write-Output (New-JobMessage -Type "Log" -Value "Retries: $($Settings.Retries)")
        Write-Output (New-JobMessage -Type "Log" -Value "Max Parallel Scans: $($Settings.MaxParallelScans)")
        Write-Output (New-JobMessage -Type "Log" -Value "Output File: $($Settings.OutputFileName).$($Settings.OutputFileExtension)")
        Write-Output (New-JobMessage -Type "Log" -Value "Save Unresponsive: $($Settings.SaveUnresponsive)")
        Write-Output (New-JobMessage -Type "Log" -Value " ========================")
        
        # Safety check to prevent infinite recursion or multiple controller instances
        if ($global:controllerAlreadyRan) {
            Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: ERROR - Controller already executed, preventing duplicate run")
            Write-Output (New-JobMessage -Type "Status" -Value "Complete")
            return
        }
        $global:controllerAlreadyRan = $true
        
        # =============================================================================
        # SYSTEM CAPABILITY VERIFICATION  
        # =============================================================================
        
        # Test SNMP COM object availability before starting worker jobs
        # This prevents workers from failing due to missing dependencies
        try {
            $testSNMP = New-Object -ComObject olePrn.OleSNMP
            $testSNMP = $null  # Release the test object immediately
            Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: Successfully verified OleSNMP COM object availability")
        } catch {
            # SNMP COM object not available - workers will fall back to ping-only mode
            $errorMsg = "Failed to create OleSNMP COM object: $($_.Exception.Message)"
            Write-Output (New-JobMessage -Type "Log" -Value "Controller [$controllerId]: $errorMsg")
            Write-Output (New-JobMessage -Type "Status" -Value "Complete")
            return
        }
        
        # =============================================================================
        # WORKER SCRIPT BLOCK - INDIVIDUAL IP SCANNING
        # =============================================================================
        
        # This script block is executed as a separate PowerShell job for each IP address to be scanned
        # It performs ping testing, SNMP queries, and device identification for a single IP
        # 
        # WORKER WORKFLOW:
        # 1. Check SNMP COM object availability (per-worker verification)
        # 2. Perform ping test with retry logic
        # 3. If ping successful and SNMP available: Query device information
        # 4. Identify device type using SNMP OID mapping
        # 5. Return structured device information to controller
        
        $workerScriptBlock = {
            param($CurrentIP, $ScanSettings)
            
            # Helper function to create structured messages for controller communication
            function New-WorkerMessage { 
                param($Type, $Value) 
                [PSCustomObject]@{Type = $Type; Value = $Value; IP = $CurrentIP} 
            }
            
            # Helper function for file-only verbose logging (bypasses GUI completely)
            function Write-VerboseLog {
                param([string]$Message)
                if ($ScanSettings.DiagnosticLevel -eq "Verbose") {
                    Write-Output (New-WorkerMessage -Type "VerboseLog" -Value $Message)
                }
            }
            
            # =============================================================================
            # WORKER-LEVEL SNMP CAPABILITY CHECK
            # =============================================================================
            
            # Each worker independently verifies SNMP COM object availability
            # This handles cases where SNMP might be available to controller but not workers
            $snmpAvailable = $false
            try {
                $testSNMP = New-Object -ComObject olePrn.OleSNMP
                $testSNMP = $null  # Release the test object immediately
                $snmpAvailable = $true
                Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - OleSNMP COM object available")
            } catch {
                # SNMP not available - this worker will perform ping-only scanning
                $snmpAvailable = $false
                $errorMsg = $_.Exception.Message
                Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - OleSNMP COM object unavailable - $errorMsg")
            }

            function Get-SnmpValue {
                param([string]$IP, [string[]]$Communities, [string[]]$OIDs, [int]$Retries)
                $results = @{}
                $successfulCommunity = $null
                
                # Try each community string in order until one works
                foreach ($community in $Communities) {
                    Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - Trying SNMP with community '$community'")
                    Write-VerboseLog "Worker for $CurrentIP - Starting SNMP query attempt with community string '$community'"
                    $results = @{}
                    $lastError = ""
                    
                    for ($attempt = 0; $attempt -le $Retries; $attempt++) {
                        try {
                            foreach ($oid in $OIDs) {
                                try {
                                    $SNMP = New-Object -ComObject olePrn.OleSNMP
                                    $SNMP.open($IP, $community, $Retries, 1000)
                                    $value = $SNMP.get($oid)
                                    $SNMP.Close()
                                    
                                    # Log raw SNMP values to file only when in verbose mode
                                    Write-VerboseLog "Worker for $CurrentIP - SNMP GET $oid with community '$community': '$value' (Type: $($value.GetType().Name))"
                                    
                                    if ($value -and $value -ne "") {
                                        $results[$oid] = $value
                                        Write-VerboseLog "Worker for $CurrentIP - Successfully stored OID $oid = '$value'"
                                    } else {
                                        $results[$oid] = $null
                                        Write-VerboseLog "Worker for $CurrentIP - OID $oid returned null or empty value"
                                    }
                                } catch {
                                    $results[$oid] = $null
                                    Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - SNMP query failed for $oid with community '$community': $($_.Exception.Message)")
                                } finally {
                                    if ($SNMP) { 
                                        try { $SNMP.Close() } catch { }
                                        $SNMP = $null
                                    }
                                }
                            }
                            
                            # If we got at least one successful OID result, consider this community successful
                            if ($results.Values | Where-Object { $_ -ne $null }) {
                                $successfulCommunity = $community
                                Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - SNMP successful with community '$community'")
                                Write-VerboseLog "Worker for $CurrentIP - SNMP queries completed successfully using community string '$community'"
                                break
                            }
                            break
                        } catch { 
                            $lastError = $_.Exception.Message
                            Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - SNMP attempt $($attempt + 1) failed with community '$community': $lastError")
                            Start-Sleep -Milliseconds 500 
                        }
                    }
                    
                    # If this community worked, stop trying others
                    if ($successfulCommunity) {
                        break
                    }
                }
                
                if (-not $successfulCommunity) { 
                    Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - All SNMP community strings failed")
                    return $null 
                }
                return $results
            }

            function Get-DeviceType {
                param([string]$OID, [string]$SysName = "", [string]$IP = $CurrentIP, [string[]]$Communities = @("public"))
                # Normalize OID by removing prefixes, quotes, and whitespace
                $cleanOID = $OID -replace '^OID=', '' -replace '"', '' -replace '\s', ''

                $oidMap = @{
                    ".iso.org.dod.internet.private.enterprises.17186.1.10"      = "1.3.6.1.4.1.17186.1.10"
                    ".iso.org.dod.internet.private.enterprises.21839.1.2.17"    = "1.3.6.1.4.1.21839.1.2.17"
                    ".iso.org.dod.internet.private.enterprises.21839.1.2.20"    = "1.3.6.1.4.1.21839.1.2.20"
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
                        
                        foreach ($community in $Communities) {
                            try {
                                Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - Trying MD8000 variant detection with community '$community'")
                                $SNMP = New-Object -ComObject olePrn.OleSNMP
                                $SNMP.open($IP, $community, 3, 1000)
                                $variantValue = $SNMP.get(".1.3.6.1.4.1.17186.1.10.1.1.3.0")
                                
                                # Log detailed variant detection values to file only when in verbose mode
                                Write-VerboseLog "Worker for $IP - MD8000 variant OID (.1.3.6.1.4.1.17186.1.10.1.1.3.0) with community '$community': '$variantValue'"
                                
                                Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - MD8000 variant OID value with community '$community': '$variantValue'")
                                if ($variantValue -eq "1") {
                                    Write-VerboseLog "Worker for $IP - Variant value '1' detected - returning MD8000EX"
                                    $SNMP.Close()
                                    return "MD8000EX"
                                } elseif ($variantValue -eq "2") {
                                    Write-VerboseLog "Worker for $IP - Variant value '2' detected - returning MD8000SX"
                                    $SNMP.Close()
                                    return "MD8000SX"
                                } else {
                                    # Check for SWCNT9-100G if variantValue is not 1 or 2
                                    try {
                                        $swcntOid = ".1.3.6.1.4.1.17186.1.10.1.1.6.1.2.13"
                                        $swcntValue = $SNMP.get($swcntOid)
                                        
                                        # Log detailed SWCNT detection values to file only when in verbose mode
                                        Write-VerboseLog "Worker for $IP - SWCNT9-100G OID ($swcntOid) with community '$community': '$swcntValue'"
                                        
                                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - SWCNT9-100G OID value with community '$community': '$swcntValue'")
                                        $SNMP.Close()
                                        if ($swcntValue -eq 69 -or $swcntValue -eq "69") {
                                            Write-VerboseLog "Worker for $IP - SWCNT value '69' detected - returning SWCNT9-100G"
                                            return "SWCNT9-100G"
                                        } else {
                                            Write-VerboseLog "Worker for $IP - No specific variant detected - returning generic MD8000"
                                            return "MD8000"
                                        }
                                    } catch {
                                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - Failed to query SWCNT9-100G OID with community '$community': $($_.Exception.Message)")
                                        Write-VerboseLog "Worker for $IP - SWCNT9-100G query failed: $($_.Exception.Message)"
                                        $SNMP.Close()
                                        return "MD8000"
                                    }
                                }
                                # If we got here, we successfully queried but didn't get variant info
                                $SNMP.Close()
                                return "MD8000"
                            } catch {
                                Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - Failed to query MD8000 variant OID with community '$community': $($_.Exception.Message)")
                                Write-VerboseLog "Worker for $IP - MD8000 variant detection failed with community '$community': $($_.Exception.Message)"
                                # Try next community string
                                continue
                            } finally {
                                if ($SNMP) { 
                                    try { $SNMP.Close() } catch { }
                                    $SNMP = $null
                                }
                            }
                        }
                        # If all community strings failed
                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - All community strings failed for MD8000 variant detection")
                        return "MD8000"
                    }
                    "1.3.6.1.4.1.21839.1.2.17"    { return "MDX2040" }
                    "1.3.6.1.4.1.21839.1.2.20"    { return "MDX2049" }
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

            function Test-XscendDevice {
                param([string]$IP)
                try {
                    Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - Testing for XSCEND device via HTTP")
                    
                    # Create web request with timeout
                    $webRequest = [System.Net.WebRequest]::Create("http://$IP")
                    $webRequest.Timeout = 5000  # 5 second timeout
                    $webRequest.Method = "GET"
                    
                    # Get response
                    $response = $webRequest.GetResponse()
                    $stream = $response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($stream)
                    $content = $reader.ReadToEnd()
                    
                    # Clean up
                    $reader.Close()
                    $stream.Close()
                    $response.Close()
                    
                    # Check if content contains "XSCEND"
                    if ($content -match "XSCEND") {
                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - XSCEND keyword found in HTTP response")
                        return $true
                    } else {
                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - XSCEND keyword not found in HTTP response")
                        return $false
                    }
                } catch {
                    Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $IP - HTTP test failed: $($_.Exception.Message)")
                    return $false
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
                    # Parse comma-separated community strings and prepare the attempt list
                    $userCommunities = $ScanSettings.SnmpCommunity -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" -and $_ -ne "public" }
                    $communityList = if ($userCommunities.Count -gt 0) { $userCommunities -join ", " } else { "none" }
                    
                    Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - Starting SNMP queries, trying 'public' first, then: $communityList")
                    $oidsToGet = @( ".1.3.6.1.2.1.1.2.0", ".1.3.6.1.2.1.1.5.0", ".1.3.6.1.2.1.1.6.0" )
                    
                    # Create array of community strings to try: "public" first, then user-entered communities
                    $communityStrings = @("public")
                    if ($userCommunities.Count -gt 0) {
                        $communityStrings += $userCommunities
                    }
                    
                    Write-VerboseLog "Worker for $CurrentIP - Community strings to try in order: $($communityStrings -join ', ')"
                    
                    try {
                        $snmpResult = Get-SnmpValue -IP $CurrentIP -Communities $communityStrings -OIDs $oidsToGet -Retries $ScanSettings.Retries
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
                        
                        # Log detailed SNMP values to file only when in verbose mode
                        Write-VerboseLog "Worker for $CurrentIP - SNMP query results:"
                        Write-VerboseLog "  Device Type OID (.1.3.6.1.2.1.1.2.0): '$typeOID'"
                        Write-VerboseLog "  System Name (.1.3.6.1.2.1.1.5.0): '$name'"
                        Write-VerboseLog "  System Location (.1.3.6.1.2.1.1.6.0): '$location'"
                        
                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - SNMP data retrieved: typeOID='$typeOID', name='$name', location='$location'")
                        
                        # Handle empty/null values like v1 does
                        if ([string]::IsNullOrWhiteSpace($name)) { $name = "[No Name Found]" }
                        if ([string]::IsNullOrWhiteSpace($location)) { $location = "[No Location Found]" }
                        
                        # If the type OID is missing, the type is UNKNOWN.
                        $type = if ($typeOID) { Get-DeviceType -OID $typeOID -SysName $name -IP $CurrentIP -Communities $communityStrings } else { "UNKNOWN" }
                        
                        Write-VerboseLog "Worker for $CurrentIP - Final processed values: Name='$name', Location='$location', Type='$type'"
                        
                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - Device type determined: '$type'")
                        Write-Output ([PSCustomObject]@{ IP = $CurrentIP; Name = $name; Location = $location; Type = $type; Status = "Responsive" })
                    } else { 
                        # SNMP failed, check for XSCEND device via HTTP before marking as UNKNOWN
                        Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - SNMP query failed, checking for XSCEND device")
                        
                        if (Test-XscendDevice -IP $CurrentIP) {
                            Write-Output ([PSCustomObject]@{ IP = $CurrentIP; Name = "[HTTP Detected]"; Location = "[SNMP Failed]"; Type = "XSCEND"; Status = "Responsive" })
                        } else {
                            Write-Output ([PSCustomObject]@{ IP = $CurrentIP; Name = "[No Name Found]"; Location = "[No Location Found]"; Type = "UNKNOWN"; Status = "Unresponsive" })
                        }
                    }
                } else {
                    # SNMP not available, check for XSCEND device via HTTP before returning ping-only result
                    Write-Output (New-WorkerMessage -Type "WorkerLog" -Value "Worker for $CurrentIP - SNMP not available, checking for XSCEND device")
                    
                    if (Test-XscendDevice -IP $CurrentIP) {
                        Write-Output ([PSCustomObject]@{ IP = $CurrentIP; Name = "[HTTP Detected]"; Location = "[No SNMP]"; Type = "XSCEND"; Status = "Responsive" })
                    } else {
                        Write-Output ([PSCustomObject]@{ IP = $CurrentIP; Name = "[SNMP Unavailable]"; Location = "[Ping Only]"; Type = "PING_ONLY"; Status = "Responsive" })
                    }
                }
            } else {
                Write-Output ([PSCustomObject]@{ IP = $CurrentIP; Name = "[No Response]"; Location = "[No Response]"; Type = "NO_PING"; Status = "Unresponsive" })
            }
        } # End of $workerScriptBlock

        function Parse-IpRangesJob {
            param([string]$IpRangeString)
            $allIps = New-Object System.Collections.Generic.List[string]; $ranges = $IpRangeString -split ',' | ForEach-Object { $_.Trim() }
            foreach ($range in $ranges) {
                if ($range -match '^(\d{1,3}\.\d{1,3}\.\d{1,3})\.0$') { $base = $matches[1]; 1..254 | ForEach-Object { [void]$allIps.Add("$base.$_") } }
                elseif ($range -match '^(\d{1,3}\.\d{1,3}\.\d{1,3})\.(\d{1,3})-(\d{1,3})$') { $base = $matches[1]; ([int]$matches[2])..([int]$matches[3]) | ForEach-Object { [void]$allIps.Add("$base.$_") } }
                elseif ($range -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') { [void]$allIps.Add($range) }
            }
            return $allIps
        }

        # Parse IP ranges and group by range for better reporting
        $ranges = $Settings.IpRanges -split ',' | ForEach-Object { $_.Trim() }
        $allIpsToScan = Parse-IpRangesJob -IpRangeString $Settings.IpRanges; $totalIpCount = $allIpsToScan.Count
        $processedIpCount = 0; $allResults = New-Object System.Collections.Generic.List[object]; $runningJobs = @()
        
        Write-Output (New-JobMessage -Type "Log" -Value "Starting scan of $totalIpCount IP addresses in $($ranges.Count) range(s)...")
        foreach ($range in $ranges) {
            $rangeIps = Parse-IpRangesJob -IpRangeString $range
            Write-Output (New-JobMessage -Type "Log" -Value "Starting scan of range '$range' ($($rangeIps.Count) addresses)")
        }
        Write-Output (New-JobMessage -Type "Log" -Value "========================")

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
                            } elseif ($jobResult -and $jobResult.Type -eq "VerboseLog") {
                                # Verbose log messages: write to log file only, never to GUI
                                Write-Output (New-JobMessage -Type "VerboseFileOnly" -Value $jobResult.Value)
                            } elseif ($jobResult -and $jobResult.IP) {
                                # This is a device result - format for professional display
                                if ($jobResult.Status -eq "Unresponsive") {
                                    Write-Output (New-JobMessage -Type "Log" -Value "$($jobResult.IP): Unreachable")
                                } else {
                                    $deviceName = if ($jobResult.Name -and $jobResult.Name -notin @("[SNMP Unavailable]", "[No Name Found]")) { $jobResult.Name } else { "UNKNOWN" }
                                    $deviceLocation = if ($jobResult.Location -and $jobResult.Location -notin @("[Ping Only]", "[No Location Found]")) { $jobResult.Location } else { "UNKNOWN" }
                                    $deviceType = if ($jobResult.Type -and $jobResult.Type -ne "PING_ONLY") { $jobResult.Type } else { "UNKNOWN" }
                                    Write-Output (New-JobMessage -Type "Log" -Value "Discovered $($jobResult.IP): Name=$deviceName, Location=$deviceLocation, Type=$deviceType")
                                }
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
                    foreach ($jobResult in $jobResults) {
                        if ($jobResult -and $jobResult.Type -eq "WorkerLog") {
                            # Skip worker log messages from GUI display (keep in full log only)
                        } elseif ($jobResult -and $jobResult.Type -eq "VerboseLog") {
                            # Verbose log messages: write to log file only, never to GUI
                            Write-Output (New-JobMessage -Type "VerboseFileOnly" -Value $jobResult.Value)
                        } elseif ($jobResult -and $jobResult.IP) {
                            # This is a device result - format for professional display
                            if ($jobResult.Status -eq "Unresponsive") {
                                Write-Output (New-JobMessage -Type "Log" -Value "$($jobResult.IP): Unreachable")
                            } else {
                                $deviceName = if ($jobResult.Name -and $jobResult.Name -notin @("[SNMP Unavailable]", "[No Name Found]")) { $jobResult.Name } else { "UNKNOWN" }
                                $deviceLocation = if ($jobResult.Location -and $jobResult.Location -notin @("[Ping Only]", "[No Location Found]")) { $jobResult.Location } else { "UNKNOWN" }
                                $deviceType = if ($jobResult.Type -and $jobResult.Type -ne "PING_ONLY") { $jobResult.Type } else { "UNKNOWN" }
                                Write-Output (New-JobMessage -Type "Log" -Value "Discovered $($jobResult.IP): Name=$deviceName, Location=$deviceLocation, Type=$deviceType")
                            }
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

        Write-Output (New-JobMessage -Type "RangeResult" -Value $allResults)
        Write-Output (New-JobMessage -Type "Log" -Value "Scan completed successfully")
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
                    Write-Log -Message "Job failed - checking for error details" -MessageType "Error"
                    $jobErrors = Receive-Job -Job $global:scanJob -ErrorAction SilentlyContinue
                    if ($jobErrors) {
                        Write-Log -Message "Job error: $jobErrors" -MessageType "Error"
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
                Write-Log -Message "Error accessing job: $($_.Exception.Message)" -MessageType "Error"
                $guiTimer.Stop()
                $cancelButton.Enabled = $false; $saveButton.Enabled = $true; $closeButton.Enabled = $true
                $global:scanJob = $null
                return
            }
            foreach ($msg in $messages) {
                switch ($msg.Type) {
                    "Log" { 
                        # Check if this is an unreachable message and handle specially
                        if ($msg.Value -match "^([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+): Unreachable$") {
                            # Track unreachable IP but don't display during scan to reduce noise
                            $unreachableIP = $matches[1]
                            $script:unreachableIPs += $unreachableIP
                            # Skip displaying this message during scanning
                            continue
                        }
                        
                        # Classify message type based on content for all other messages
                        $messageType = "General"
                        if ($msg.Value -match "Starting scan.*range|Starting scan of.*IP addresses") { $messageType = "RangeStart" }
                        elseif ($msg.Value -match "Completed scan.*range|All scans completed|scan process completed|Scan completed successfully") { 
                            $messageType = "RangeEnd" 
                            
                            # If this is the final scan completion, show completion summary
                            if ($msg.Value -match "Scan completed successfully") {
                                Write-Log -Message $msg.Value -MessageType $messageType
                                
                                # Always show scan completion separator and unreachable summary if there are any
                                Write-Log -Message "" -MessageType "General"  # Blank line for spacing
                                Write-Log -Message "===============================================" -MessageType "SectionSeparator"
                                Write-Log -Message "===         SCAN COMPLETION SUMMARY        ===" -MessageType "SectionSeparator"
                                Write-Log -Message "===============================================" -MessageType "SectionSeparator"
                                
                                # Show unreachable summary if there are unreachable IPs
                                if ($script:unreachableIPs.Count -gt 0) {
                                    Write-Log -Message "The following $($script:unreachableIPs.Count) IP addresses did not respond to ping:" -MessageType "UnreachableSummary"
                                    
                                    # Group IPs for compact display (show in groups of 8 per line)
                                    $ipGroups = @()
                                    for ($i = 0; $i -lt $script:unreachableIPs.Count; $i += 8) {
                                        $ipGroup = $script:unreachableIPs[$i..([Math]::Min($i + 7, $script:unreachableIPs.Count - 1))] -join ", "
                                        $ipGroups += $ipGroup
                                    }
                                    
                                    foreach ($ipGroup in $ipGroups) {
                                        Write-Log -Message "  $ipGroup" -MessageType "UnreachableSummary"
                                    }
                                } else {
                                    Write-Log -Message "All scanned IP addresses were reachable via ping." -MessageType "Success"
                                }
                                
                                Write-Log -Message "===============================================" -MessageType "SectionSeparator"
                                continue  # Skip the normal processing since we already logged the completion message
                            }
                        }
                        elseif ($msg.Value -match "Discovered.*:") { $messageType = "ScanResult" }
                        elseif ($msg.Value -match "=== SCAN PARAMETERS ===|IP Ranges:|SNMP Community:|Retries:|Max Parallel Scans:|Output File:|Save Unresponsive:|========================") { $messageType = "ScanParameters" }
                        elseif ($msg.Value -match "Worker for.*-|SNMP.*raw value|COM object|OID.*value|Exception|Controller.*Started job|Controller.*Processing.*results|Controller.*Found device result|Found device result for IP|Started job for IP|Processing.*results from") { $messageType = "Diagnostic" }
                        Write-Log -Message $msg.Value -MessageType $messageType
                    }
                    "WorkerLog" { 
                        # Worker logs are typically diagnostic information
                        $messageType = "Diagnostic"
                        if ($msg.Value -match "Ping successful|Ping failed") { $messageType = "PingResult" }
                        elseif ($msg.Value -match "Device type determined|Found device") { $messageType = "ScanResult" }
                        Write-Log -Message $msg.Value -MessageType $messageType
                    }
                    "VerboseFileOnly" {
                        # Write directly to log file only, bypassing GUI completely
                        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                        $logEntry = "[$timestamp] $($msg.Value)"
                        Add-Content -Path $logFilePath -Value $logEntry
                    }
                    "Progress" { if($msg.Value -le 100) {$progressBar.Value = [int]$msg.Value } }
                    "RangeResult" {
                        $responsive = $msg.Value | Where-Object { $_.Status -eq 'Responsive' }; $unresponsive = $msg.Value | Where-Object { $_.Status -eq 'Unresponsive' }
                        foreach ($item in $msg.Value) { 
                            [void]$global:allDiscoveredDevices.Add($item)
                            Write-Log -Message "Added device to collection: $($item.IP) - $($item.Name) - Status: $($item.Status)" -MessageType "Diagnostic"
                        }
                        
                        # Always show the discovered devices report section
                        Write-Log -Message "" -MessageType "General"  # Blank line for spacing
                        Write-Log -Message "===============================================" -MessageType "ResultsSeparator"
                        Write-Log -Message "===        DISCOVERED DEVICES REPORT       ===" -MessageType "ResultsSeparator"
                        Write-Log -Message "===============================================" -MessageType "ResultsSeparator"
                        Write-Log -Message "Total devices in collection: $($global:allDiscoveredDevices.Count)" -MessageType "Success"
                        
                        $grouped = $responsive | Group-Object Name, Location, Type
                        if($grouped) {
                            foreach ($group in $grouped) {
                                $ips = ($group.Group.IP | Sort-Object) -join ', '; $name = $group.Group[0].Name; $location = $group.Group[0].Location; $type = $group.Group[0].Type
                                Write-Log -Message "  Name=$name Location=$location Type=$type Address=$ips" -MessageType "DeviceResult"
                            }
                        } else {
                            Write-Log -Message "  No MediaLinks devices discovered via SNMP." -MessageType "UnreachableSummary"
                        }
                        
                        if ($unresponsive) { 
                            $unresponsiveIPs = ($unresponsive.IP | Sort-Object) -join ', '
                            Write-Log -Message "  SNMP Unresponsive: $unresponsiveIPs" -MessageType "UnreachableSummary" 
                        }
                        
                        Write-Log -Message "===============================================" -MessageType "ResultsSeparator"
                    }
                    "Status" {
                        if ($msg.Value -eq 'Complete') {
                            $progressBar.Value = 100
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
                # Create header information
                $currentDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $headerLines = @(
                    "# DiscoverSubnet Version $scriptVersion Report",
                    "# Generated: $currentDateTime", 
                    "# IP Ranges Scanned: $($scanSettings.IpRanges)",
                    "# SNMP Communities: public, $($scanSettings.SnmpCommunity)",
                    "# Max Parallel Scans: $($scanSettings.MaxParallelScans)",
                    "# Total Devices Found: $($outputData.Count)",
                    "#"
                )
                
                # Write header lines to file first
                $headerLines | Out-File -FilePath $outputFilePath -Encoding UTF8
                
                # Format and append data based on selected file type
                if ($scanSettings.OutputFileExtension -eq 'csv') {
                    # CSV format with semicolon delimiters - manually format to avoid Export-Csv append issues
                    '"Name";"Location";"Type";"IPs"' | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
                    foreach ($device in $outputData) {
                        $csvLine = "`"$($device.Name)`";`"$($device.Location)`";`"$($device.Type)`";`"$($device.IPs)`""
                        $csvLine | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
                    }
                } else {
                    # TXT format - tab-delimited for better readability
                    $txtHeader = "Name".PadRight(20) + "Location".PadRight(18) + "Type".PadRight(12) + "IPs"
                    $txtHeader | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
                    "-" * 70 | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
                    
                    foreach ($device in $outputData) {
                        # Convert properties to strings to avoid PadRight issues with deserialized objects
                        $nameStr = [string]$device.Name
                        $locationStr = [string]$device.Location
                        $typeStr = [string]$device.Type
                        $ipsStr = [string]$device.IPs
                        
                        # Truncate long names/locations to fit better
                        $name = if ($nameStr.Length -gt 19) { $nameStr.Substring(0,16) + "..." } else { $nameStr }
                        $location = if ($locationStr.Length -gt 17) { $locationStr.Substring(0,14) + "..." } else { $locationStr }
                        $type = if ($typeStr.Length -gt 11) { $typeStr.Substring(0,8) + "..." } else { $typeStr }
                        
                        $line = $name.PadRight(20) + $location.PadRight(18) + $type.PadRight(12) + $ipsStr
                        $line | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
                    }
                }
                
                Write-Log -Message "Results successfully saved to $outputFilePath" -MessageType "Success"
                [System.Windows.Forms.MessageBox]::Show("Results saved successfully.", "Save Complete", "OK", "Information")
            } else {
                Write-Log -Message "No devices to save."; [System.Windows.Forms.MessageBox]::Show("There are no discovered devices to save.", "Save Complete", "OK", "Information")
            }
        } catch {
            Write-Log -Message "ERROR: Failed to save results. $_" -MessageType "Error"; [System.Windows.Forms.MessageBox]::Show("An error occurred while saving the file: `n$($_.Exception.Message)", "Save Error", "OK", "Error")
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
        
        if (-not $script:jobStarted -and -not $global:scanJob -and $script:formShownCount -eq 1) {
            # Display version header in GUI
            Write-Log -Message "DiscoverSubnet Version $scriptVersion" -MessageType "General"
            $script:jobStarted = $true
            Write-Log -Message "Form shown - starting scan job..."
            try {
                $global:scanJob = Start-Job -ScriptBlock $controllerScriptBlock -ArgumentList @($scanSettings)
                $guiTimer.Start()
                Write-Log -Message "Scan job started successfully with ID: $($global:scanJob.Id)" -MessageType "Success"
            } catch {
                Write-Log -Message "Error starting scan job: $($_.Exception.Message)" -MessageType "Error"
                $script:jobStarted = $false
            }
        } else {
            Write-Log -Message "Form shown but conditions not met - jobStarted:$script:jobStarted, scanJobExists:$($global:scanJob -ne $null), formShownCount:$script:formShownCount"
        }
    })

    [void]$logForm.ShowDialog()
}
# End of script.
