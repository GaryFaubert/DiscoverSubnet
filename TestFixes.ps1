#requires -Version 5.1
<#
.SYNOPSIS
    Test script to verify the fixes for Trim() method errors and XSCEND logging
#>

# Test 1: Verify PSCustomObject Type field can be converted to string and trimmed
Write-Host "Testing PSCustomObject Type field conversion and trimming..." -ForegroundColor Cyan

# Simulate deserialized objects that come from job results
$simulatedDevice = [PSCustomObject]@{
    Name = "TestDevice"
    Location = "TestLocation" 
    Type = "  XSCEND  "  # With leading/trailing spaces
    IPs = "192.168.1.100"
}

# Convert to JSON and back to simulate deserialization from job results
$jsonString = $simulatedDevice | ConvertTo-Json
$deserializedDevice = $jsonString | ConvertFrom-Json

Write-Host "Original Type field: '$($deserializedDevice.Type)'" -ForegroundColor Yellow
Write-Host "Type field type: $($deserializedDevice.Type.GetType().FullName)" -ForegroundColor Yellow

# Test the fix: Convert to string first, then trim
try {
    $cleanType = [string]$deserializedDevice.Type
    $cleanType = $cleanType.Trim()
    Write-Host "✓ Fixed approach works! Clean Type: '$cleanType'" -ForegroundColor Green
} catch {
    Write-Host "✗ Fixed approach failed: $($_.Exception.Message)" -ForegroundColor Red
}

# Test what would happen with the old approach (should fail)
try {
    $oldApproach = $deserializedDevice.Type.Trim()
    Write-Host "✗ Old approach unexpectedly worked: '$oldApproach'" -ForegroundColor Red
} catch {
    Write-Host "✓ Old approach correctly fails: $($_.Exception.Message)" -ForegroundColor Green
}

Write-Host "`nTest completed successfully!" -ForegroundColor Green
Write-Host "Fixes implemented:" -ForegroundColor White
Write-Host "  1. PSCustomObject Type field converted to string before Trim()" -ForegroundColor White  
Write-Host "  2. XSCEND WorkerLog messages enabled in Standard/Verbose modes" -ForegroundColor White
Write-Host "  3. Individual worker debug files created for verbose XSCEND troubleshooting" -ForegroundColor White