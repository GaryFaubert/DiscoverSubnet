# DiscoverSubnet v2.19

A powerful network discovery tool specifically designed to identify and catalog MediaLinks hardware on your network infrastructure.

## 📋 Table of Contents

- [Overview](#overview)
- [System Requirements](#system-requirements)
- [Installation](#installation)
- [Quick Start Guide](#quick-start-guide)
- [Configuration Options](#configuration-options)
- [Supported IP Range Formats](#supported-ip-range-formats)
- [Device Identification](#device-identification)
- [Output Formats](#output-formats)
- [Performance Optimization](#performance-optimization)
- [Troubleshooting](#troubleshooting)
- [File Locations](#file-locations)
- [Version History](#version-history)

## 🔍 Overview

DiscoverSubnet is a Windows-based network discovery application that combines ping testing with SNMP queries to identify and classify MediaLinks network devices. The tool features a user-friendly GUI interface, intelligent performance optimization, and comprehensive logging capabilities.

### Key Features

- **Enhanced SNMP Discovery**: Automatically tries "public" community first, then falls back to multiple user-specified community strings for maximum compatibility with mixed-device environments
- **Smart Device Detection**: Automatically identifies MediaLinks hardware models (MD8000EX/SX, SWCNT9-100G, MDX series, MDP3020, etc.)
- **Verbose Diagnostic Logging**: Optional file-only logging of detailed SNMP values and queries for troubleshooting
- **Flexible IP Scanning**: Supports single IPs, IP ranges, and subnet notation including gateway addresses
- **Performance Optimization**: Automatically analyzes your system to recommend optimal parallel scan settings
- **Real-time Monitoring**: Live progress updates and color-coded logging during scans
- **Multiple Output Formats**: Export results in CSV or TXT format
- **Graceful Degradation**: Works with ping-only if SNMP is unavailable

## 💻 System Requirements

### Minimum Requirements
- **Operating System**: Windows 7/8/10/11 or Windows Server 2012+
- **Framework**: .NET Framework 4.5 or higher
- **Memory**: 100 MB available RAM
- **Disk Space**: 10 MB free space for application and logs

### Recommended Requirements
- **Operating System**: Windows 10/11 or Windows Server 2016+
- **Memory**: 500 MB available RAM for large network scans
- **Network**: Administrative access for SNMP queries (optional)

### Optional Components
- **OleSNMP COM Object**: For full SNMP device identification
  - If unavailable, the application automatically falls back to ping-only mode
  - Device identification will be limited but functional

## 📦 Installation

### Option 1: Standalone Executable (Recommended)
1. Download `DiscoverSubnet.exe` from the release
2. Place the executable in a dedicated folder (e.g., `C:\Tools\DiscoverSubnet\`)
3. No additional installation required - the application is self-contained

### Option 2: PowerShell Script
1. Download `DiscoverSubnet.ps1`
2. Ensure PowerShell 5.1 or higher is installed
3. Run from PowerShell: `.\DiscoverSubnet.ps1`

> **Note**: The executable version is recommended for most users as it requires no PowerShell configuration and includes all dependencies.

## 🚀 Quick Start Guide

### Basic Usage
1. **Launch**: Double-click `DiscoverSubnet.exe`
2. **Configure**: Enter your network's IP ranges in the configuration dialog
3. **Start**: Click "Start Discovery" to begin the scan
4. **Monitor**: Watch real-time progress in the logging window
5. **Save**: Export results when the scan completes

### Example: Scanning a Subnet
1. In the "IP Address Ranges" field, enter: `192.168.1.0`
2. Set SNMP Communities to: `medialinks, custom, private` (the tool will automatically try "public" first, then these in order)
3. Set Diagnostic Level to: `Verbose` (for detailed SNMP logging) or `Standard` (for normal operation)
4. Click "Auto" next to "Max Parallel Scans" for optimal performance
5. Click "Start Discovery"

## ⚙️ Configuration Options

### Network Configuration

| Setting | Description | Example |
|---------|-------------|---------|
| **IP Address Ranges** | Target IP ranges to scan | `192.168.1.0, 10.0.0.10-20` |
| **SNMP Community Strings** | Multiple community strings (tool tries "public" first automatically) | `medialinks, custom, private` |
| **Ping/SNMP Retries** | Number of retry attempts for failed connections | `0-3` (0 = no retries) |

### Output Configuration

| Setting | Description | Options |
|---------|-------------|---------|
| **Output File Name** | Base name for result files | `DiscoveredDevices` |
| **Output File Type** | Format for exported results | `CSV`, `TXT` |
| **Save Unresponsive** | Include unresponsive devices in output | Checked/Unchecked |

### Display and Performance

| Setting | Description | Options |
|---------|-------------|---------|
| **GUI Display Level** | Amount of information shown during scan | `Minimal`, `Standard`, `Verbose` |
| **Max Parallel Scans** | Number of concurrent IP scans | `1-100` (Auto-recommended) |
| **Diagnostic Level** | Detail level for logging | `Off`, `Standard`, `Verbose` |

> **Note**: When Diagnostic Level is set to "Verbose", detailed SNMP values and queries are logged to the log file only (not displayed in GUI) for comprehensive troubleshooting.

## 🌐 Supported IP Range Formats

The application supports multiple IP range formats for flexible network scanning:

### Format Examples

| Format | Description | Example | Result |
|--------|-------------|---------|--------|
| **Single IP** | Individual IP address | `192.168.1.100` | Scans only 192.168.1.100 |
| **IP Range** | Range of IP addresses | `192.168.1.10-20` | Scans 192.168.1.10 through 192.168.1.20 |
| **Subnet** | Entire subnet (includes .1 gateway) | `192.168.1.0` | Scans 192.168.1.1 through 192.168.1.254 |
| **Mixed** | Combination of formats | `192.168.1.5, 10.0.0.0, 172.16.1.100-110` | Scans all specified addresses |

### Valid IP Range Rules
- First three octets must be between 1-254
- Fourth octet must be between 1-254 for specific IPs
- Subnet notation (.0) automatically includes gateway (.1) addresses for comprehensive discovery
- Ranges must have start value less than end value

## 🔧 Device Identification

### Supported MediaLinks Hardware

| Device Model | Detection Method | Variants Detected |
|--------------|------------------|-------------------|
| **MD8000 Series** | SNMP OID identification | Standard, EX, SX, SWCNT9-100G |
| **MDX Series** | SNMP OID + sysName parsing | 32C, 48X6C variants |
| **MDX2040** | SNMP OID identification | Standard |
| **MDP3020** | SNMP OID identification | Standard |

### SNMP Community Strategy
The application uses an intelligent multi-community approach:
1. **First Attempt**: Always tries "public" community string
2. **Fallback**: If "public" fails, tries each user-specified community string in order
3. **Multiple Communities**: Supports comma-separated communities (e.g., "medialinks, custom, private")
4. **Logging**: In Verbose mode, logs which community string succeeded for each device

This strategy maximizes device discovery success in mixed-device environments where different devices may use different community strings.

### Device Information Collected
- **IP Address**: Network address of the device
- **Device Name**: SNMP sysName or hostname
- **Location**: SNMP sysLocation field
- **Device Type**: Specific MediaLinks model identification
- **Status**: Responsive (SNMP), Responsive (Ping-only), or Unresponsive

## 📄 Output Formats

### CSV Format
Semicolon-delimited format suitable for Excel import:
```csv
"Name";"Location";"Type";"IPs"
"MD8000-Lab";"Server Room A";"MD8000EX";"192.168.1.100"
"MDX-Switch-01";"Wiring Closet B";"MDX-32C";"192.168.1.150, 192.168.1.151"
```

### TXT Format
Human-readable tabular format:
```
Name                 Location           Type         IPs
------------------------------------------------------------
MD8000-Lab           Server Room A      MD8000EX     192.168.1.100
MDX-Switch-01        Wiring Closet B    MDX-32C      192.168.1.150, 192.168.1.151
```

## 🚀 Performance Optimization

### Automatic Performance Analysis
The application automatically analyzes your system to recommend optimal settings:

- **System Assessment**: Evaluates CPU cores, memory, and performance category
- **Scan Complexity**: Analyzes IP range size and distribution
- **Intelligent Recommendations**: Suggests optimal parallel scan count
- **Performance Estimation**: Provides estimated scan completion time

### Manual Performance Tuning

| System Type | Recommended Parallel Scans | Typical Performance |
|-------------|---------------------------|-------------------|
| **Low-end** | 6-8 concurrent scans | 50-100 IPs/minute |
| **Mid-range** | 12-16 concurrent scans | 150-300 IPs/minute |
| **High-end** | 20-30 concurrent scans | 400-600 IPs/minute |

### Performance Tips
1. **Use the "Auto" button** for optimal parallel scan recommendations
2. **Reduce GUI verbosity** to "Minimal" for large scans
3. **Increase parallel scans** on systems with more CPU cores
4. **Monitor memory usage** during large subnet scans

## 🔧 Troubleshooting

### Common Issues

#### Application Won't Start
- **Cause**: Missing .NET Framework
- **Solution**: Install .NET Framework 4.5 or higher from Microsoft

#### SNMP Devices Not Detected
- **Cause**: Incorrect community string or SNMP disabled
- **Solution**: 
  - The tool automatically tries "public" first, then each of your specified community strings in order
  - Try multiple community strings separated by commas: `medialinks, custom, private`
  - Verify SNMP community strings with network administrator
  - Check if SNMP is enabled on target devices
  - Enable Verbose diagnostic logging to see detailed SNMP attempts and which community succeeded
  - Application will continue with ping-only detection if all SNMP communities fail

#### Slow Scanning Performance
- **Cause**: Too few parallel scans or network congestion
- **Solution**:
  - Click "Auto" for recommended parallel scan count
  - Reduce scan range to test connectivity
  - Check network bandwidth and latency

#### Application Crashes During Scan
- **Cause**: Insufficient memory or system resources
- **Solution**:
  - Reduce parallel scan count
  - Scan smaller IP ranges
  - Close other applications to free memory

### Error Messages

| Error | Possible Cause | Solution |
|-------|---------------|----------|
| "Cannot bind argument to parameter 'Path'" | Path resolution issue | Run from executable's directory |
| "Failed to load .NET Assemblies" | Missing framework components | Install/repair .NET Framework |
| "SNMP COM object unavailable" | Missing SNMP components | Continue with ping-only mode |

## 📁 File Locations

The application creates files in the same directory as the executable:

### Generated Files

| File | Purpose | Location |
|------|---------|----------|
| **DiscoverSubnet.settings.json** | Persistent user settings | Same as executable |
| **DiscoverSubnet-YYYYMMDD-HHMMSS.log** | Detailed scan logs | Same as executable |
| **DiscoveredDevices.csv/.txt** | Scan results export | Same as executable |

### Settings File Example
```json
{
  "IpRanges": "192.168.1.0, 10.0.0.10-20",
  "SnmpCommunity": "medialinks, custom, private",
  "Retries": 0,
  "OutputFileName": "DiscoveredDevices",
  "OutputFileExtension": "csv",
  "SaveUnresponsive": false,
  "MaxParallelScans": 20,
  "DiagnosticLevel": "Verbose",
  "GuiVerbosity": "Standard"
}
```

### Verbose Logging
When DiagnosticLevel is set to "Verbose", the log file includes detailed information such as:
- Raw SNMP OID values retrieved from each device
- Multiple community string attempts and which one succeeded for each device
- Detailed device type detection logic
- SNMP query timing and retry information

This verbose information is written only to the log file and never displayed in the GUI to maintain clean operation.

## 📈 Version History

### v2.19 (2025-10-09)
- **Multiple SNMP Community Support**: Added support for multiple comma-separated community strings for mixed-device environments
- **Enhanced GUI**: Updated interface to support multiple communities with helpful tooltips and validation
- **Improved Discovery**: Better success rate in environments with devices using different community strings

### v2.18 (2025-10-09)
- **Enhanced SNMP Querying**: Worker jobs now try "public" community first, then user-entered community string for improved device discovery
- **Verbose File-Only Logging**: Added detailed SNMP value logging to file only (not GUI) when DiagnosticLevel is "Verbose"
- **Improved Troubleshooting**: Comprehensive logging of SNMP queries, community string attempts, and device detection logic

### v2.17 (2025-09-28)
- Improved GUI display: changed DISCOVERED DEVICES REPORT to white text
- Removed SCAN COMPLETION SUMMARY from GUI, moved total count under DISCOVERED DEVICES REPORT
- Enhanced visual clarity and reduced GUI clutter

### v2.16
- Fixed IP range validation to allow scanning of .1 addresses (gateways)
- Changed minimum range validation from 2 to 1 for single IP scanning

### v2.15
- Fixed GUI verbosity filtering to ensure summary sections always appear
- Improved discovery window display consistency

### v2.14
- Fixed missing summary sections in scan results
- Ensured discovered devices report and scan completion summary always appear

### v2.13
- Added visual section separators with distinct colors
- Dark cyan for scan completion, dark green for results report

### v2.12
- Suppressed unreachable IP messages during scan to reduce noise
- Show gray summary at end instead of cluttering real-time display

### v2.11
- Added color-coded logging in discovery window
- Distinguished message types (errors=red, success=green, warnings=yellow, etc.)

### v2.10
- Added support for SWCNT9-100G device type detection in MD8000 series hardware
- Enhanced MediaLinks device variant identification

### v2.9
- Updated subnet scanning to include gateway addresses (.1)
- More comprehensive network discovery capabilities

### v2.8
- Enhanced PS2EXE compatibility with improved path resolution
- Fixed null path errors when running as compiled executable
- Improved error handling for edge cases

### v2.7
- Added system performance analysis
- Automatic parallel scan recommendations
- Enhanced system capability detection

### v2.6
- Improved GUI verbosity controls
- Professional logging format
- Better real-time progress display

### v2.5
- Enhanced device type detection for MediaLinks variants
- Improved SNMP OID mapping
- Better device identification accuracy

---

## 📞 Support

For technical support or feature requests, please refer to the application's source repository or contact your system administrator.

**Version**: 2.19  
**Last Updated**: October 9, 2025  
**Author**: Gary Faubert