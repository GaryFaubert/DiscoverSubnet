try {
    [System.Management.Automation.Language.Parser]::ParseFile('C:\DiscoverSubnet\DiscoverSubnet.ps1',[ref]$null,[ref]$null)
    Write-Output 'Syntax check: PASS'
}
catch {
    Write-Output ('Syntax check: FAIL: ' + $_.Exception.Message)
    exit 1
}
