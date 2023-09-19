[CmdletBinding()]
param(
    [string]$Env = "LAB"
)

$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
try {
    . ("$ScriptDirectory\_Helpers.ps1")
}
catch {
    Write-Error "Error while loading PowerShell scripts" 
    Write-Error $_.Exception.Message
}

Invoke-Start $MyInvocation.MyCommand.Name $ScriptDirectory

try {
    $config = Get-Config $Env
    $config

    $appPoolAccount = CheckOrCreate-AppPoolAccount $config

    $appPool = CheckOrCreate-AppPool $config $appPoolAccount
    
    Write-Host "Checking if [$($config.SearchServiceAppName)] not exists..." -NoNewline:$True
    $ServiceApplication = Get-SPServiceApplication -Name $config.SearchServiceAppName -ErrorAction SilentlyContinue
    if ($null -ne $ServiceApplication) {
        Write-Host " [TO DELETE]" -ForegroundColor Yellow
        Remove-SPServiceApplication $ServiceApplication -Confirm:$False
    }
    else {
        Write-Host " [OK]" -ForegroundColor Green
    }

    Write-Host "Checking if [$($config.SearchServiceProxyName)] not exists..." -NoNewline:$True
    $ServiceAppProxy = Get-SPEnterpriseSearchServiceApplicationProxy -Identity $config.SearchServiceProxyName -ErrorAction SilentlyContinue
    if ($null -ne $ServiceAppProxy) {
        Write-Host " [TO DELETE]" -ForegroundColor Yellow
        Remove-SPServiceApplicationProxy $ServiceAppProxy.Id -Confirm:$False
    }
    else {
        Write-Host " [OK]" -ForegroundColor Green
    }

    $hostApp = Get-SPEnterpriseSearchServiceInstance -Identity $config.SearchServer

    Start-SPEnterpriseSearchServiceInstance -Identity $hostApp
    
    Get-SPEnterpriseSearchServiceInstance -Identity $hostApp
    
    Write-Host "Restoring Search Service Application database '$($config.SearchDB)' to '$($config.DatabaseServer)'..." -NoNewline:$True
    Restore-SPEnterpriseSearchServiceApplication -Name $config.SearchDB -ApplicationPool $AppPool -AdminSearchServiceInstance $hostApp -DatabaseName $config.SearchDB -DatabaseServer $config.DatabaseServer
    Write-Host " [OK]" -ForegroundColor Green

    Write-Host "Renaming the Search Service Application '$($config.SearchDB)' to '$($config.SearchServiceAppName)'..." -NoNewline:$True
    $service = Get-SPServiceApplication -Name $config.SearchDB
    $service.Name = $config.SearchServiceAppName
    $service.Update()
    Write-Host " [OK]" -ForegroundColor Green
    
    $ServiceApplication = Get-SPEnterpriseSearchServiceApplication -Identity $config.SearchServiceAppName

    Write-Host "Checking if [$($config.SearchServiceProxyName)] exists..." -NoNewline:$True
    $ServiceAppProxy = Get-SPEnterpriseSearchServiceApplicationProxy -Identity $config.SearchServiceProxyName -ErrorAction SilentlyContinue
    if ($null -eq $ServiceAppProxy) {
        Write-Host " [TO CREATE]" -ForegroundColor Yellow
        $ServiceAppProxy = New-SPEnterpriseSearchServiceApplicationProxy -Name $config.SearchServiceProxyName -SearchApplication $ServiceApplication     
    }
    else {
        Write-Host " [OK]" -ForegroundColor Green
    }
    
    Write-Host -ForegroundColor Green "Search Service Application created successfully!"
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}