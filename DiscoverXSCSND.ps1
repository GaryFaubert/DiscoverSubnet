function Invoke-DeviceApiRequest {
    param (
        [string]$BaseUrl = "http://192.168.12.21:8080/api",
        [string]$Endpoint,
        [string]$Method = "GET",  # GET, POST, PUT, DELETE
        [hashtable]$Body = $null,
        [string]$Token = $null
    )

    # Construct full URL
    $url = "$BaseUrl/$Endpoint"

    # Prepare headers
    $headers = @{}
    if ($Token) {
        $headers["Authorization"] = "Bearer $Token"
    }

    # Convert body to JSON if present
    $jsonBody = if ($Body) { $Body | ConvertTo-Json -Depth 5 } else { $null }

    try {
        $response = Invoke-RestMethod -Uri $url -Method $Method -Headers $headers `
            -Body $jsonBody -ContentType "application/json" -ErrorAction Stop

        return $response
    }
    catch {
        Write-Warning "API call failed: $($_.Exception.Message)"
        return $null
    }
}


function Get-DeviceApiToken {
    param (
        [string]$AuthUrl = "http://192.168.12.21:8080/api/auth",
        [string]$Username,
        [string]$Password
    )

    $body = @{
        username = $Username
        password = $Password
    } | ConvertTo-Json

    try {
        $response = Invoke-RestMethod -Uri $AuthUrl -Method Post -Body $body -ContentType "application/json"
        return $response.token
    }
    catch {
        Write-Warning "Token retrieval failed: $($_.Exception.Message)"
        return $null
    }
}
