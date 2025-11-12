# Test script to demonstrate XSCEND device naming fixes
# This simulates the XSCEND detection logic without the GUI

# Simulate the XSCEND API response data based on the ShowEquip CSV
$simulatedXscendDevices = @(
    @{
        IP = "192.168.12.21"
        SystemMgmt = @{
            'device-name' = 'XSX-21'
            'location' = 'Windsor'
            'outband-interface' = @{
                '1' = @{ 'ip-address' = '192.168.12.21' }
                '2' = @{ 'ip-address' = '192.168.12.22' }
            }
        }
    },
    @{
        IP = "192.168.12.22"
        SystemMgmt = @{
            'device-name' = 'XSX-21'  # Same chassis name as .21
            'location' = 'Windsor'
            'outband-interface' = @{
                '1' = @{ 'ip-address' = '192.168.12.21' }
                '2' = @{ 'ip-address' = '192.168.12.22' }
            }
        }
    },
    @{
        IP = "192.168.12.23"
        SystemMgmt = @{
            'device-name' = 'XSX-23'
            'location' = 'Windsor'
            'outband-interface' = @{
                '1' = @{ 'ip-address' = '192.168.12.23' }
                '2' = @{ 'ip-address' = '192.168.12.24' }
            }
        }
    },
    @{
        IP = "192.168.12.24"
        SystemMgmt = @{
            'device-name' = 'XSX-23'  # Same chassis name as .23
            'location' = 'Windsor'
            'outband-interface' = @{
                '1' = @{ 'ip-address' = '192.168.12.23' }
                '2' = @{ 'ip-address' = '192.168.12.24' }
            }
        }
    }
)

# Simulate the updated XSCEND naming logic from DiscoverSubnet v2.30
function Test-XscendNaming {
    param($IP, $SystemMgmt)
    
    Write-Host "Processing IP: $IP" -ForegroundColor Cyan
    
    # Extract device information
    $deviceName = "XSCEND-Device-$IP"  # Default unique name per IP
    $deviceLocation = "[API Detected]"
    $allIPs = @($IP)

    # Prefer canonical name fields in order: device-name, alias, chassis.alias
    $actualDeviceName = $null
    if ($systemMgmt) {
        if ($systemMgmt.'device-name' -and -not [string]::IsNullOrWhiteSpace($systemMgmt.'device-name')) {
            $actualDeviceName = $systemMgmt.'device-name'.Trim()
        }
        elseif ($systemMgmt.'alias' -and -not [string]::IsNullOrWhiteSpace($systemMgmt.'alias')) {
            $actualDeviceName = $systemMgmt.'alias'.Trim()
        }
        elseif ($systemMgmt.'chassis' -and $systemMgmt.'chassis'.'alias' -and -not [string]::IsNullOrWhiteSpace($systemMgmt.'chassis'.'alias')) {
            $actualDeviceName = $systemMgmt.'chassis'.'alias'.Trim()
        }
    }

    if ($actualDeviceName) {
        # Use the discovered name but append IP to keep grouping unique when necessary
        $deviceName = "$actualDeviceName-$IP"
        Write-Host "  Found XSCEND device name: '$actualDeviceName', using unique name: '$deviceName'" -ForegroundColor Green
    }
    else {
        Write-Host "  No device-name/alias found in API response, using default unique name: '$deviceName'" -ForegroundColor Yellow
    }
    
    # Try to get location information
    if ($systemMgmt -and $systemMgmt.'location') {
        $deviceLocation = $systemMgmt.'location'.Trim()
        Write-Host "  Found XSCEND device location: '$deviceLocation'" -ForegroundColor Green
    }
    
    # Check for peer management IP (dual-controller detection)
    if ($systemMgmt -and $systemMgmt.'outband-interface') {
        $outband1 = $systemMgmt.'outband-interface'.'1'
        $outband2 = $systemMgmt.'outband-interface'.'2'
        
        $ip1 = if ($outband1 -and $outband1.'ip-address') { $outband1.'ip-address'.Trim() } else { $null }
        $ip2 = if ($outband2 -and $outband2.'ip-address') { $outband2.'ip-address'.Trim() } else { $null }
        
        # Add peer IP if different from current IP
        if ($ip1 -and $ip1 -ne $IP -and $ip1 -ne "0.0.0.0") {
            $allIPs += $ip1
            Write-Host "  Found peer management interface: $ip1" -ForegroundColor Cyan
        }
        if ($ip2 -and $ip2 -ne $IP -and $ip2 -ne "0.0.0.0" -and $ip2 -ne $ip1) {
            $allIPs += $ip2
            Write-Host "  Found peer management interface: $ip2" -ForegroundColor Cyan
        }
    }
    
    # Return device information
    return [PSCustomObject]@{
        Success = $true
        Name = $deviceName
        Location = $deviceLocation
        Type = "XSCEND"
        IPs = ($allIPs -join ",")
        Status = "Responsive"
    }
}

Write-Host "=== XSCEND Device Naming Test - DiscoverSubnet v2.30 ===" -ForegroundColor White
Write-Host "Testing the updated XSCEND naming logic that prefers device-name/alias fields" -ForegroundColor White
Write-Host ""

$allDevices = @()

# Process each simulated XSCEND device
foreach ($device in $simulatedXscendDevices) {
    $result = Test-XscendNaming -IP $device.IP -SystemMgmt $device.SystemMgmt
    $allDevices += $result
    Write-Host ""
}

Write-Host "=== DEVICE GROUPING RESULTS ===" -ForegroundColor White
Write-Host "This shows how devices will be grouped in the device file:" -ForegroundColor White
Write-Host ""

# Group devices by Name, Location, Type (same logic as DiscoverSubnet)
$grouped = $allDevices | Group-Object Name, Location, Type

foreach ($group in $grouped) {
    $ips = ($group.Group.IPs | ForEach-Object { $_ -split ',' } | Sort-Object -Unique) -join ', '
    $name = $group.Group[0].Name
    $location = $group.Group[0].Location
    $type = $group.Group[0].Type
    
    Write-Host "Device Entry:" -ForegroundColor Yellow
    Write-Host "  Status: OK" -ForegroundColor Green
    Write-Host "  Name: $name" -ForegroundColor Green
    Write-Host "  Location: $location" -ForegroundColor Green
    Write-Host "  Type: $type" -ForegroundColor Green
    Write-Host "  IP(s): $ips" -ForegroundColor Green
    Write-Host ""
}

Write-Host "=== DEVICE FILE OUTPUT PREVIEW ===" -ForegroundColor White
Write-Host "Status;Name;Location;Type;IP(s)" -ForegroundColor Cyan
foreach ($group in $grouped) {
    $ips = ($group.Group.IPs | ForEach-Object { $_ -split ',' } | Sort-Object -Unique) -join ','
    $name = $group.Group[0].Name
    $location = $group.Group[0].Location
    $type = $group.Group[0].Type
    
    Write-Host "OK;$name;$location;$type;$ips" -ForegroundColor White
}

Write-Host ""
Write-Host "=== SUMMARY ===" -ForegroundColor White
Write-Host "✅ Each XSCEND device now has a unique name (device-name + IP)" -ForegroundColor Green
Write-Host "✅ Devices will appear as separate entries in the device file" -ForegroundColor Green
Write-Host "✅ Peer management IPs are still properly collected" -ForegroundColor Green
Write-Host "✅ No more grouping conflicts due to identical names" -ForegroundColor Green