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

    Write-Host "Checking if [$($config.ManagedMetadataServiceAppName)] exists..." -NoNewline:$True
    $ServiceApplication = Get-SPServiceApplication -Name $config.ManagedMetadataServiceAppName -ErrorAction SilentlyContinue
    if ($null -eq $ServiceApplication) {
        Write-Host " [TO CREATE]" -ForegroundColor Yellow
        $ServiceApplication = New-SPMetadataServiceApplication -Name $config.ManagedMetadataServiceAppName -ApplicationPool $config.AppPoolName -DatabaseName $config.ManagedMetadataDB
    }
    else {
        Write-Host " [OK]" -ForegroundColor Green
    }

    Write-Host "Checking if [$($config.ManagedMetadataServiceProxyName)] exists..." -NoNewline:$True
    $ServiceAppProxy = Get-SPServiceApplicationProxy | Where-Object { $_.Name -eq $config.ManagedMetadataServiceProxyName }
    if ($null -eq $ServiceAppProxy) {
        Write-Host " [TO CREATE]" -ForegroundColor Yellow
        $ServiceAppProxy = New-SPMetadataServiceApplicationProxy -Name $config.ManagedMetadataServiceProxyName -ServiceApplication $ServiceApplication       
    }
    else {
        Write-Host " [OK]" -ForegroundColor Green
    }

    $ServiceInstance = Get-SPServiceInstance | Where-Object { $_.TypeName -eq "Managed Metadata Web Service" }
    
    if ($ServiceInstance.Status -ne "Online") {
        Write-Host -ForegroundColor Yellow "Starting the Managed Metadata Web Service Instance..."
        Start-SPServiceInstance $ServiceInstance
    }
    
    Write-Host -ForegroundColor Green "Managed Metadata Service Application created successfully!"
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}