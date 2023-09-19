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
    
    Write-Host "Checking if [$($config.UserProfileServiceAppName)] exists..." -NoNewline:$True
    $ServiceApplication = Get-SPServiceApplication -Name $config.UserProfileServiceAppName -ErrorAction SilentlyContinue
    if ($null -eq $ServiceApplication) {
        Write-Host " [TO CREATE]" -ForegroundColor Yellow
        $ServiceApplication = New-SPProfileServiceApplication -Name $config.UserProfileServiceAppName -ApplicationPool $config.AppPoolName -ProfileDBName $config.UserProfileDB -ProfileSyncDBName $config.UserProfileSyncDB -SocialDBName $config.UserProfileSocialDB
    }
    else {
        Write-Host " [OK]" -ForegroundColor Green
    }

    Write-Host "Checking if [$($config.UserProfileServiceProxyName)] exists..." -NoNewline:$True
    $ServiceAppProxy = Get-SPServiceApplicationProxy | Where-Object { $_.Name -eq $config.UserProfileServiceProxyName }
    if ($null -eq $ServiceAppProxy) {
        Write-Host " [TO CREATE]" -ForegroundColor Yellow
        $ServiceAppProxy = New-SPProfileServiceApplicationProxy -Name $config.UserProfileServiceProxyName -ServiceApplication $ServiceApplication -DefaultProxyGroup       
    }
    else {
        Write-Host " [OK]" -ForegroundColor Green
    }

    $ServiceInstance = Get-SPServiceInstance | Where-Object { $_.TypeName -eq "User Profile Service" }
    
    if ($ServiceInstance.Status -ne "Online") {
        Write-Host -ForegroundColor Yellow "Starting the User Profile Service Instance..."
        Start-SPServiceInstance $ServiceInstance
    }
    
    Write-Host -ForegroundColor Green "User Profile Service Application created successfully!"
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}