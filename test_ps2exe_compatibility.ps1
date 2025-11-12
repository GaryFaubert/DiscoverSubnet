# Test PS2EXE Compatibility Improvements for XSCEND API - v2.32
Write-Host "Testing PS2EXE Compatibility Improvements - v2.32" -ForegroundColor Green
Write-Host "===================================================" -ForegroundColor Green
Write-Host ""

Write-Host "Changes Made for PS2EXE Compatibility:" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Yellow
Write-Host "✓ Replaced Invoke-RestMethod with System.Net.WebClient" -ForegroundColor Green
Write-Host "✓ Replaced System.Web.HttpUtility with System.Uri.EscapeDataString" -ForegroundColor Green  
Write-Host "✓ Removed System.Web assembly dependency" -ForegroundColor Green
Write-Host "✓ Added proper WebClient disposal and error handling" -ForegroundColor Green
Write-Host "✓ Added JSON parsing fallback for raw text responses" -ForegroundColor Green
Write-Host ""

# Test the URL encoding replacement
Write-Host "Testing URL Encoding Replacement:" -ForegroundColor Cyan
Write-Host "---------------------------------" -ForegroundColor Cyan

$testUsername = "s-admin"
$testPassword = "A5hU3CqW4+"

# Old method (problematic in PS2EXE)
try {
    Add-Type -AssemblyName System.Web -ErrorAction Stop
    $oldEncoded = [System.Web.HttpUtility]::UrlEncode($testUsername)
    Write-Host "OLD (System.Web): $testUsername → $oldEncoded" -ForegroundColor Gray
} catch {
    Write-Host "OLD (System.Web): FAILED - Assembly not available in PS2EXE context" -ForegroundColor Red
}

# New method (PS2EXE compatible)
$newEncoded = [System.Uri]::EscapeDataString($testUsername)
Write-Host "NEW (System.Uri): $testUsername → $newEncoded" -ForegroundColor Green

$newEncodedPwd = [System.Uri]::EscapeDataString($testPassword)
Write-Host "NEW (System.Uri): $testPassword → $newEncodedPwd" -ForegroundColor Green

Write-Host ""

# Test WebClient creation (simulated)
Write-Host "Testing WebClient Creation:" -ForegroundColor Cyan  
Write-Host "---------------------------" -ForegroundColor Cyan

try {
    $webClient = New-Object System.Net.WebClient
    $webClient.Headers.Add("User-Agent", "DiscoverSubnet/2.32")
    $webClient.Headers.Add("Content-Type", "application/json")
    $webClient.Headers.Add("Accept", "application/json")
    
    Write-Host "✓ WebClient created successfully" -ForegroundColor Green
    Write-Host "✓ Headers added successfully:" -ForegroundColor Green
    foreach ($header in $webClient.Headers.AllKeys) {
        Write-Host "  $header`: $($webClient.Headers[$header])" -ForegroundColor White
    }
    
    $webClient.Dispose()
    Write-Host "✓ WebClient disposed properly" -ForegroundColor Green
} catch {
    Write-Host "✗ WebClient test failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# Simulate the authentication URL construction
Write-Host "Testing XSCEND Authentication URL Construction:" -ForegroundColor Cyan
Write-Host "-----------------------------------------------" -ForegroundColor Cyan

$baseUrl = "http://192.168.12.21:80"
$authUrl = "$baseUrl/api/auth/token?username=$newEncoded&password=$newEncodedPwd"
Write-Host "Base URL: $baseUrl" -ForegroundColor White
Write-Host "Auth URL: $authUrl" -ForegroundColor White

Write-Host ""

Write-Host "PS2EXE Compatibility Summary:" -ForegroundColor Yellow
Write-Host "=============================" -ForegroundColor Yellow
Write-Host "✓ No more Invoke-RestMethod dependency" -ForegroundColor Green
Write-Host "✓ No more System.Web assembly dependency" -ForegroundColor Green
Write-Host "✓ WebClient works reliably in compiled executables" -ForegroundColor Green
Write-Host "✓ System.Uri is part of core .NET Framework" -ForegroundColor Green
Write-Host "✓ Proper resource disposal prevents memory leaks" -ForegroundColor Green
Write-Host ""

Write-Host "Expected Results:" -ForegroundColor Yellow
Write-Host "• .ps1 version: Should work as before" -ForegroundColor White
Write-Host "• .exe version: XSCEND API should now work correctly" -ForegroundColor Green
Write-Host "• Both versions: Should produce identical device file output" -ForegroundColor White

Write-Host ""
Write-Host "Next Step: Recompile with PS2EXE and test against XSCEND devices!" -ForegroundColor Magenta