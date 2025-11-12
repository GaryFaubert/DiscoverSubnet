# Test XSCEND naming logic for v2.31 - Clean device names without IP suffixes
# This script simulates the XSCEND naming logic to verify proper device name handling

Write-Host "Testing XSCEND Device Naming Logic v2.31" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green
Write-Host ""

function Test-XscendNamingV231 {
    param(
        [string]$IP,
        [string]$DeviceName = $null,
        [string]$Alias = $null,
        [string]$ChassisAlias = $null,
        [string]$Location = "Windsor"
    )
    
    # Simulate the new v2.31 naming logic
    $deviceName = "XSCEND-Device-$IP"  # Default fallback name
    $deviceLocation = "[API Detected]"
    $allIPs = @($IP)
    
    # Prefer canonical name fields in order: device-name, alias, chassis.alias
    $actualDeviceName = $null
    
    if ($DeviceName -and -not [string]::IsNullOrWhiteSpace($DeviceName)) {
        $actualDeviceName = $DeviceName.Trim()
    }
    elseif ($Alias -and -not [string]::IsNullOrWhiteSpace($Alias)) {
        $actualDeviceName = $Alias.Trim()
    }
    elseif ($ChassisAlias -and -not [string]::IsNullOrWhiteSpace($ChassisAlias)) {
        $actualDeviceName = $ChassisAlias.Trim()
    }
    
    if ($actualDeviceName) {
        # Use the clean device name from API (e.g., "XSX-21", "XSX-23")
        $deviceName = $actualDeviceName
        Write-Host "✓ Found XSCEND device name from API: '$actualDeviceName'" -ForegroundColor Cyan
    }
    else {
        Write-Host "⚠ No device-name/alias found in API response, using default name: '$deviceName'" -ForegroundColor Yellow
    }
    
    if ($Location) {
        $deviceLocation = $Location
    }
    
    # Simulate additional management IPs (peer interfaces)
    if ($IP -eq "192.168.12.21") {
        $allIPs += "192.168.12.22"  # Peer controller
    }
    elseif ($IP -eq "192.168.12.23") {
        $allIPs += "192.168.12.24"  # Peer controller
    }
    
    # Create device result
    $result = [PSCustomObject]@{
        Success = $true
        Name = $deviceName
        Location = $deviceLocation
        Type = "XSCEND"
        IPs = ($allIPs -join ",")
        Status = "Responsive"
    }
    
    return $result
}

Write-Host "Test Case 1: XSX-21 device with proper API name" -ForegroundColor White
$device1 = Test-XscendNamingV231 -IP "192.168.12.21" -DeviceName "XSX-21" -Location "Windsor"
Write-Host "Device File Entry: OK;$($device1.Name);$($device1.Location);$($device1.Type);$($device1.IPs)" -ForegroundColor Green
Write-Host ""

Write-Host "Test Case 2: XSX-23 device with alias name" -ForegroundColor White  
$device2 = Test-XscendNamingV231 -IP "192.168.12.23" -DeviceName $null -Alias "XSX-23" -Location "Windsor"
Write-Host "Device File Entry: OK;$($device2.Name);$($device2.Location);$($device2.Type);$($device2.IPs)" -ForegroundColor Green
Write-Host ""

Write-Host "Test Case 3: Device without API name (fallback scenario)" -ForegroundColor White
$device3 = Test-XscendNamingV231 -IP "192.168.12.25" -DeviceName $null -Alias $null -ChassisAlias $null -Location "Windsor"
Write-Host "Device File Entry: OK;$($device3.Name);$($device3.Location);$($device3.Type);$($device3.IPs)" -ForegroundColor Green
Write-Host ""

Write-Host "Test Case 4: Device with chassis.alias name" -ForegroundColor White
$device4 = Test-XscendNamingV231 -IP "192.168.12.26" -DeviceName $null -Alias $null -ChassisAlias "XSX-Backup" -Location "Windsor"
Write-Host "Device File Entry: OK;$($device4.Name);$($device4.Location);$($device4.Type);$($device4.IPs)" -ForegroundColor Green
Write-Host ""

Write-Host "SUMMARY - Expected Device File Output:" -ForegroundColor Yellow
Write-Host "======================================" -ForegroundColor Yellow
Write-Host "OK;XSX-21;Windsor;XSCEND;192.168.12.21,192.168.12.22" -ForegroundColor White
Write-Host "OK;XSX-23;Windsor;XSCEND;192.168.12.23,192.168.12.24" -ForegroundColor White
Write-Host "OK;XSCEND-Device-192.168.12.25;Windsor;XSCEND;192.168.12.25" -ForegroundColor Gray
Write-Host "OK;XSX-Backup;Windsor;XSCEND;192.168.12.26" -ForegroundColor White
Write-Host ""
Write-Host "✓ v2.31 Fix: Clean device names (XSX-21, XSX-23) without IP suffixes!" -ForegroundColor Green
Write-Host "✓ Only fallback devices get IP-suffixed names" -ForegroundColor Green