# DiscoverSubnet v2.8

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

- **Smart Device Detection**: Automatically identifies MediaLinks hardware models (MD8000, MDX series, MDP3020, etc.)
- **Flexible IP Scanning**: Supports single IPs, IP ranges, and subnet notation
- **Performance Optimization**: Automatically analyzes your system to recommend optimal parallel scan settings
- **Real-time Monitoring**: Live progress updates and detailed logging during scans
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
2. Set SNMP Community to: `public` (or your network's community string)
3. Click "Auto" next to "Max Parallel Scans" for optimal performance
4. Click "Start Discovery"

## ⚙️ Configuration Options

### Network Configuration

| Setting | Description | Example |
|---------|-------------|---------|
| **IP Address Ranges** | Target IP ranges to scan | `192.168.1.0, 10.0.0.10-20` |
| **SNMP Community String** | Community string for SNMP queries | `public`, `medialinks` |
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
| **GUI Display Level** | Amount of information shown during scan | `Standard`, `Minimal` |
| **Max Parallel Scans** | Number of concurrent IP scans | `1-100` (Auto-recommended) |
| **Diagnostic Level** | Detail level for logging | `Off`, `Standard`, `Verbose` |

## 🌐 Supported IP Range Formats

The application supports multiple IP range formats for flexible network scanning:

### Format Examples

| Format | Description | Example | Result |
|--------|-------------|---------|--------|
| **Single IP** | Individual IP address | `192.168.1.100` | Scans only 192.168.1.100 |
| **IP Range** | Range of IP addresses | `192.168.1.10-20` | Scans 192.168.1.10 through 192.168.1.20 |
| **Subnet** | Entire subnet (excludes .1 and .255) | `192.168.1.0` | Scans 192.168.1.2 through 192.168.1.254 |
| **Mixed** | Combination of formats | `192.168.1.5, 10.0.0.0, 172.16.1.100-110` | Scans all specified addresses |

### Valid IP Range Rules
- First three octets must be between 1-254
- Fourth octet must be between 1-254 for specific IPs
- Subnet notation (.0) automatically excludes network and broadcast addresses
- Ranges must have start value less than end value

## 🔧 Device Identification

### Supported MediaLinks Hardware

| Device Model | Detection Method | Variants Detected |
|--------------|------------------|-------------------|
| **MD8000 Series** | SNMP OID identification | Standard, EX, SX |
| **MDX Series** | SNMP OID + sysName parsing | 32C, 48X6C |
| **MDX2040** | SNMP OID identification | Standard |
| **MDP3020** | SNMP OID identification | Standard |

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
  - Verify SNMP community string with network administrator
  - Check if SNMP is enabled on target devices
  - Application will continue with ping-only detection

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
  "SnmpCommunity": "medialinks",
  "Retries": 0,
  "OutputFileName": "DiscoveredDevices",
  "OutputFileExtension": "csv",
  "SaveUnresponsive": false,
  "MaxParallelScans": 20,
  "DiagnosticLevel": "Standard",
  "GuiVerbosity": "Standard"
}
```

## 📈 Version History

### v2.8 (2025-09-26)
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

**Version**: 2.9  
**Last Updated**: September 28, 2025  
**Author**: Gary Faubert