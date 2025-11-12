# Quick test to verify the v2.31 fix is working correctly
# This checks if the XSCEND naming logic has been properly updated in the script

Write-Host "Verifying v2.31 XSCEND Device Naming Fix" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green
Write-Host ""

# Check if the script file contains the correct naming logic
$scriptContent = Get-Content "c:\DiscoverSubnet\DiscoverSubnet.ps1" -Raw

# Look for the key lines that should show the clean naming logic
$keyPattern = 'if \(\$actualDeviceName\) \{\s*# Use the clean device name from API \(e\.g\., "XSX-21", "XSX-23"\)\s*\$deviceName = \$actualDeviceName'

if ($scriptContent -match $keyPattern) {
    Write-Host "✓ Found correct naming logic in script:" -ForegroundColor Green
    Write-Host "  → Uses clean device names without IP suffixes" -ForegroundColor Cyan
} else {
    Write-Host "✗ Naming logic not found - checking for old pattern..." -ForegroundColor Red
    $oldPattern = '\$deviceName = "\$actualDeviceName-\$IP"'
    if ($scriptContent -match $oldPattern) {
        Write-Host "  → Found old naming pattern (still has IP suffixes)" -ForegroundColor Yellow
    }
}

# Check version number
if ($scriptContent -match '\$scriptVersion = "2\.31"') {
    Write-Host "✓ Script version updated to 2.31" -ForegroundColor Green
} else {
    Write-Host "✗ Script version not updated" -ForegroundColor Red
}

# Check changelog
if ($scriptContent -match 'v2\.31.*REFINED XSCEND device naming.*clean device names') {
    Write-Host "✓ Changelog updated for v2.31 fix" -ForegroundColor Green
} else {
    Write-Host "✗ Changelog not updated" -ForegroundColor Red
}

Write-Host ""
Write-Host "Expected behavior:" -ForegroundColor Yellow
Write-Host "  - XSX-21 device → Name: 'XSX-21' (clean, no IP suffix)" -ForegroundColor White
Write-Host "  - XSX-23 device → Name: 'XSX-23' (clean, no IP suffix)" -ForegroundColor White
Write-Host "  - No API name   → Name: 'XSCEND-Device-192.168.x.x' (fallback)" -ForegroundColor Gray
Write-Host ""
Write-Host "Device file output should be:" -ForegroundColor Yellow
Write-Host "  OK;XSX-21;Windsor;XSCEND;192.168.12.21,192.168.12.22" -ForegroundColor White
Write-Host "  OK;XSX-23;Windsor;XSCEND;192.168.12.23,192.168.12.24" -ForegroundColor White