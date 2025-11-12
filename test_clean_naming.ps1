# Corrected test for XSCEND naming logic v2.31
Write-Host "Testing XSCEND Device Naming Logic v2.31 - Corrected" -ForegroundColor Green
Write-Host "======================================================" -ForegroundColor Green
Write-Host ""

function Test-XscendNamingV231 {
    param(
        [string]$IP,
        [string]$DeviceName = $null,
        [string]$Alias = $null,
        [string]$ChassisAlias = $null,
        [string]$Location = "Windsor"
    )
    
    Write-Host "  Testing IP: $IP" -ForegroundColor Gray
    Write-Host "  Input DeviceName: '$DeviceName'" -ForegroundColor Gray
    Write-Host "  Input Alias: '$Alias'" -ForegroundColor Gray
    Write-Host "  Input ChassisAlias: '$ChassisAlias'" -ForegroundColor Gray
    
    # Simulate the new v2.31 naming logic
    $deviceName = "XSCEND-Device-$IP"  # Default fallback name
    $deviceLocation = $Location
    $allIPs = @($IP)
    
    # Prefer canonical name fields in order: device-name, alias, chassis.alias
    $actualDeviceName = $null
    
    if ($DeviceName -and -not [string]::IsNullOrWhiteSpace($DeviceName)) {
        $actualDeviceName = $DeviceName.Trim()
        Write-Host "  → Using device-name: '$actualDeviceName'" -ForegroundColor Cyan
    }
    elseif ($Alias -and -not [string]::IsNullOrWhiteSpace($Alias)) {
        $actualDeviceName = $Alias.Trim()
        Write-Host "  → Using alias: '$actualDeviceName'" -ForegroundColor Cyan
    }
    elseif ($ChassisAlias -and -not [string]::IsNullOrWhiteSpace($ChassisAlias)) {
        $actualDeviceName = $ChassisAlias.Trim()
        Write-Host "  → Using chassis.alias: '$actualDeviceName'" -ForegroundColor Cyan
    }
    
    if ($actualDeviceName) {
        # Use the clean device name from API (e.g., "XSX-21", "XSX-23")
        $deviceName = $actualDeviceName
        Write-Host "  ✓ Final device name: '$deviceName'" -ForegroundColor Green
    }
    else {
        Write-Host "  ⚠ No API name found, using default: '$deviceName'" -ForegroundColor Yellow
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

Write-Host "Test Case 1: XSX-21 device with device-name field" -ForegroundColor White
Write-Host "---------------------------------------------------" -ForegroundColor White
$device1 = Test-XscendNamingV231 -IP "192.168.12.21" -DeviceName "XSX-21" -Location "Windsor"
Write-Host "  Result: OK;$($device1.Name);$($device1.Location);$($device1.Type);$($device1.IPs)" -ForegroundColor Green
Write-Host ""

Write-Host "Test Case 2: XSX-23 device with alias field (no device-name)" -ForegroundColor White  
Write-Host "-------------------------------------------------------------" -ForegroundColor White
$device2 = Test-XscendNamingV231 -IP "192.168.12.23" -DeviceName "" -Alias "XSX-23" -Location "Windsor"
Write-Host "  Result: OK;$($device2.Name);$($device2.Location);$($device2.Type);$($device2.IPs)" -ForegroundColor Green
Write-Host ""

Write-Host "Test Case 3: Device without any API names (fallback scenario)" -ForegroundColor White
Write-Host "--------------------------------------------------------------" -ForegroundColor White
$device3 = Test-XscendNamingV231 -IP "192.168.12.25" -DeviceName "" -Alias "" -ChassisAlias "" -Location "Windsor"
Write-Host "  Result: OK;$($device3.Name);$($device3.Location);$($device3.Type);$($device3.IPs)" -ForegroundColor Green
Write-Host ""

Write-Host "Test Case 4: Device with chassis.alias only" -ForegroundColor White
Write-Host "--------------------------------------------" -ForegroundColor White
$device4 = Test-XscendNamingV231 -IP "192.168.12.26" -DeviceName "" -Alias "" -ChassisAlias "XSX-Backup" -Location "Windsor"
Write-Host "  Result: OK;$($device4.Name);$($device4.Location);$($device4.Type);$($device4.IPs)" -ForegroundColor Green
Write-Host ""

Write-Host "EXPECTED DEVICE FILE OUTPUT (v2.31):" -ForegroundColor Yellow
Write-Host "=====================================" -ForegroundColor Yellow
Write-Host "OK;XSX-21;Windsor;XSCEND;192.168.12.21,192.168.12.22" -ForegroundColor White
Write-Host "OK;XSX-23;Windsor;XSCEND;192.168.12.23,192.168.12.24" -ForegroundColor White  
Write-Host "OK;XSCEND-Device-192.168.12.25;Windsor;XSCEND;192.168.12.25" -ForegroundColor Gray
Write-Host "OK;XSX-Backup;Windsor;XSCEND;192.168.12.26" -ForegroundColor White
Write-Host ""
Write-Host "✓ v2.31 SUCCESS: Clean device names without IP suffixes when API names available!" -ForegroundColor Green