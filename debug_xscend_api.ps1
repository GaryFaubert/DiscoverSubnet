# Debug XSCEND API Response - v2.32
# This script helps diagnose why device names and locations aren't being extracted

param(
    [string]$IP = "192.168.12.21",
    [string]$Username = "s-admin", 
    [string]$Password = "A5hU3CqW4+",
    [int]$Port = 80,
    [string]$Protocol = "http"
)

Write-Host "XSCEND API Response Debug Tool - v2.32" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green
Write-Host "Target: $Protocol`://$IP`:$Port" -ForegroundColor Yellow
Write-Host ""

try {
    # Step 1: Authentication
    Write-Host "Step 1: Authentication" -ForegroundColor Cyan
    Write-Host "---------------------" -ForegroundColor Cyan
    
    $encodedUsername = [System.Uri]::EscapeDataString($Username)
    $encodedPassword = [System.Uri]::EscapeDataString($Password)
    $baseUrl = "${Protocol}://${IP}:${Port}"
    $authUrl = "$baseUrl/api/auth/token?username=$encodedUsername&password=$encodedPassword"
    
    Write-Host "Auth URL: $authUrl" -ForegroundColor Gray
    
    $webClient = New-Object System.Net.WebClient
    try {
        $webClient.Headers.Add("User-Agent", "DiscoverSubnet/2.32")
        $responseText = $webClient.DownloadString($authUrl)
        
        Write-Host "✓ Authentication successful" -ForegroundColor Green
        Write-Host "Raw Response: $responseText" -ForegroundColor White
        
        # Parse token
        $token = $null
        try {
            $response = $responseText | ConvertFrom-Json
            if ($response.token) {
                $token = $response.token
                Write-Host "✓ Token extracted from JSON: $($token.Substring(0,20))..." -ForegroundColor Green
            }
        } catch {
            if ($responseText.StartsWith("eyJ")) {
                $token = $responseText
                Write-Host "✓ Token extracted from raw text: $($token.Substring(0,20))..." -ForegroundColor Green
            }
        }
        
        if (-not $token) {
            Write-Host "✗ Failed to extract token" -ForegroundColor Red
            return
        }
        
    } finally {
        $webClient.Dispose()
    }
    
    Write-Host ""
    
    # Step 2: System Management Query
    Write-Host "Step 2: System Management Query" -ForegroundColor Cyan
    Write-Host "-------------------------------" -ForegroundColor Cyan
    
    $sysMgmtUrl = "$baseUrl/api/properties/chassis/system-management"
    Write-Host "System Mgmt URL: $sysMgmtUrl" -ForegroundColor Gray
    
    $webClient2 = New-Object System.Net.WebClient
    try {
        $webClient2.Headers.Add("Authorization", "Bearer $token")
        $webClient2.Headers.Add("Content-Type", "application/json")
        $webClient2.Headers.Add("Accept", "application/json")
        $webClient2.Headers.Add("User-Agent", "DiscoverSubnet/2.32")
        
        $sysMgmtResponseText = $webClient2.DownloadString($sysMgmtUrl)
        $systemMgmt = $sysMgmtResponseText | ConvertFrom-Json
        
        Write-Host "✓ System management query successful" -ForegroundColor Green
        Write-Host ""
        
        Write-Host "Raw JSON Response:" -ForegroundColor Yellow
        Write-Host "==================" -ForegroundColor Yellow
        Write-Host $sysMgmtResponseText -ForegroundColor White
        Write-Host ""
        
        Write-Host "Parsed JSON Object:" -ForegroundColor Yellow
        Write-Host "===================" -ForegroundColor Yellow
        $systemMgmt | ConvertTo-Json -Depth 5 | Write-Host -ForegroundColor White
        Write-Host ""
        
        # Check specific fields
        Write-Host "Field Analysis:" -ForegroundColor Yellow
        Write-Host "===============" -ForegroundColor Yellow
        
        $fields = @(
            @{Name="device-name"; Value=$systemMgmt.'device-name'},
            @{Name="alias"; Value=$systemMgmt.'alias'},
            @{Name="location"; Value=$systemMgmt.'location'},
            @{Name="chassis.alias"; Value=$systemMgmt.'chassis'.'alias'}
        )
        
        foreach ($field in $fields) {
            $value = $field.Value
            $status = if ([string]::IsNullOrWhiteSpace($value)) { "❌ NULL/EMPTY" } else { "✅ '$value'" }
            Write-Host "  $($field.Name): $status" -ForegroundColor $(if ($status.StartsWith("✅")) { "Green" } else { "Red" })
        }
        
        Write-Host ""
        
        # Check outband interfaces
        Write-Host "Outband Interfaces:" -ForegroundColor Yellow
        Write-Host "==================" -ForegroundColor Yellow
        if ($systemMgmt.'outband-interface') {
            $systemMgmt.'outband-interface' | ConvertTo-Json -Depth 3 | Write-Host -ForegroundColor White
            
            Write-Host ""
            Write-Host "Interface IPs:" -ForegroundColor Yellow
            $outband1 = $systemMgmt.'outband-interface'.'1'
            $outband2 = $systemMgmt.'outband-interface'.'2'
            
            if ($outband1.'ip-address') {
                Write-Host "  Interface 1: $($outband1.'ip-address')" -ForegroundColor Green
            }
            if ($outband2.'ip-address') {
                Write-Host "  Interface 2: $($outband2.'ip-address')" -ForegroundColor Green
            }
        } else {
            Write-Host "  No outband interfaces found" -ForegroundColor Red
        }
        
    } finally {
        $webClient2.Dispose()
    }
    
} catch {
    Write-Host "✗ Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Debug completed. Check the field values above to see what data is available." -ForegroundColor Magenta