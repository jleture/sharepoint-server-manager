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
   
    Write-Host "Checking if [$($config.AppSubscriptionServiceAppName)] exists..." -NoNewline:$True
    $ServiceApplication = Get-SPServiceApplication -Name $config.AppSubscriptionServiceAppName -ErrorAction SilentlyContinue
    if ($null -eq $ServiceApplication) {
        Write-Host " [TO CREATE]" -ForegroundColor Yellow
        $ServiceApplication = New-SPSubscriptionSettingsServiceApplication -Name $config.AppSubscriptionServiceAppName -ApplicationPool $config.AppPoolName -DatabaseName $config.AppSubscriptionDB
    }
    else {
        Write-Host " [OK]" -ForegroundColor Green
    }

    Write-Host "Checking if the proxy exists..." -NoNewline:$True
    $ServiceAppProxy = Get-SPServiceApplicationProxy | Where-Object { $_.TypeName -eq "Microsoft SharePoint Foundation Subscription Settings Service Application Proxy" }
    if ($null -eq $ServiceAppProxy) {
        Write-Host " [TO CREATE]" -ForegroundColor Yellow
        $ServiceAppProxy = New-SPSubscriptionSettingsServiceApplicationProxy -ServiceApplication $ServiceApplication     
    }
    else {
        Write-Host " [OK]" -ForegroundColor Green
    }

    Write-Host "Checking if [$($config.AppManagementServiceAppName)] exists..." -NoNewline:$True
    $ServiceApplication = Get-SPServiceApplication -Name $config.AppManagementServiceAppName -ErrorAction SilentlyContinue
    if ($null -eq $ServiceApplication) {
        Write-Host " [TO CREATE]" -ForegroundColor Yellow
        $ServiceApplication = New-SPAppManagementServiceApplication -Name $config.AppManagementServiceAppName -ApplicationPool $config.AppPoolName -DatabaseName $config.AppManagementDB
    }
    else {
        Write-Host " [OK]" -ForegroundColor Green
    }

    Write-Host "Checking if [$($config.AppManagementServiceProxyName)] exists..." -NoNewline:$True
    $ServiceAppProxy = Get-SPServiceApplicationProxy | Where-Object { $_.Name -eq $config.AppManagementServiceProxyName }
    if ($null -eq $ServiceAppProxy) {
        Write-Host " [TO CREATE]" -ForegroundColor Yellow
        $ServiceAppProxy = New-SPAppManagementServiceApplicationProxy -Name $config.AppManagementServiceProxyName -ServiceApplication $ServiceApplication       
    }
    else {
        Write-Host " [OK]" -ForegroundColor Green
    }

    $ServiceInstance = Get-SPServiceInstance | Where-Object { $_.TypeName -eq "App Management Service" }
    
    if ($ServiceInstance.Status -ne "Online") {
        Write-Host -ForegroundColor Yellow "Starting the App Management Service Instance..."
        Start-SPServiceInstance $ServiceInstance
    }
    
    Write-Host -ForegroundColor Green "App Management Service Application created successfully!"
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}