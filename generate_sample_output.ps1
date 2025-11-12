# Simulate a complete device file with XSCEND and other devices
# This shows the expected output format after the v2.30 fixes

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Create a simulated device file content
$deviceFileContent = @"
# DiscoverSubnet Version 2.30 Report
# Generated: $timestamp
# IP Ranges Scanned: 192.168.12.0
# SNMP Communities: public, medialinks
# Max Parallel Scans: 12
# Total Devices Found: 11
#
Status;Name;Location;Type;IP(s)
OK;XSX-21-192.168.12.21;Windsor;XSCEND;192.168.12.21,192.168.12.22
OK;XSX-21-192.168.12.22;Windsor;XSCEND;192.168.12.21,192.168.12.22
OK;XSX-23-192.168.12.23;Windsor;XSCEND;192.168.12.23,192.168.12.24
OK;XSX-23-192.168.12.24;Windsor;XSCEND;192.168.12.23,192.168.12.24
OK;48X6C-1;MLI-LAB Windor;MDX-48X6C;192.168.12.91
OK;48X6C-2;MLI-LAB Windor;MDX-48X6C;192.168.12.92
OK;MDP3020_0000000115;[No Location Found];MDP3020;192.168.12.121
OK;MDP3020_M310000251;[No Location Found];MDP3020;192.168.12.122
OK;ESPN;No Location Found;MD8000SX;192.168.12.150
OK;CNN;MLI-LAB Windor;MD8000EX;192.168.12.170,192.168.12.171
OK;ABC;MLI-LAB Windor Windor;MD8000;192.168.12.175,192.168.12.176
OK;32C-1;MLI-LAB Windor;MDX-32C;192.168.12.181
OK;32C-2;MLI-LAB Windor;MDX-32C;192.168.12.182
OK;2040 1;Switches;MDX2040;192.168.12.173
"@

# Write to a sample output file
$outputFile = "c:\DiscoverSubnet\Sample_DiscoveredDevices_v2.30.csv"
$deviceFileContent | Out-File -FilePath $outputFile -Encoding UTF8

Write-Host "=== DiscoverSubnet v2.30 - Sample Output Generated ===" -ForegroundColor White
Write-Host ""
Write-Host "Sample device file created: $outputFile" -ForegroundColor Green
Write-Host ""
Write-Host "Key improvements in v2.30:" -ForegroundColor Yellow
Write-Host "  ✅ XSCEND devices now appear as separate entries" -ForegroundColor Green
Write-Host "  ✅ Each XSCEND has unique name (chassis-alias + IP)" -ForegroundColor Green
Write-Host "  ✅ No more device merging due to identical names" -ForegroundColor Green
Write-Host "  ✅ Maintains peer management IP collection" -ForegroundColor Green
Write-Host ""

# Show the differences between old and new behavior
Write-Host "=== BEFORE (v2.29) vs AFTER (v2.30) ===" -ForegroundColor White
Write-Host ""
Write-Host "BEFORE (Problem):" -ForegroundColor Red
Write-Host "OK;XSCEND-Device;[API Detected];XSCEND;192.168.12.21,192.168.12.22,192.168.12.23,192.168.12.24" -ForegroundColor Red
Write-Host "  ^ All XSCEND devices merged into one entry!" -ForegroundColor Red
Write-Host ""
Write-Host "AFTER (Fixed):" -ForegroundColor Green
Write-Host "OK;XSX-21-192.168.12.21;Windsor;XSCEND;192.168.12.21,192.168.12.22" -ForegroundColor Green
Write-Host "OK;XSX-21-192.168.12.22;Windsor;XSCEND;192.168.12.21,192.168.12.22" -ForegroundColor Green
Write-Host "OK;XSX-23-192.168.12.23;Windsor;XSCEND;192.168.12.23,192.168.12.24" -ForegroundColor Green
Write-Host "OK;XSX-23-192.168.12.24;Windsor;XSCEND;192.168.12.23,192.168.12.24" -ForegroundColor Green
Write-Host "  ^ Each XSCEND device has its own entry with proper alias names!" -ForegroundColor Green
Write-Host ""

Write-Host "=== Sample Device File Content ===" -ForegroundColor Cyan
Get-Content $outputFile | ForEach-Object { 
    if ($_ -match "^OK;XSX-") {
        Write-Host $_ -ForegroundColor Green
    } elseif ($_ -match "^#") {
        Write-Host $_ -ForegroundColor Gray
    } elseif ($_ -match "^Status;") {
        Write-Host $_ -ForegroundColor Cyan
    } else {
        Write-Host $_ -ForegroundColor White
    }
}