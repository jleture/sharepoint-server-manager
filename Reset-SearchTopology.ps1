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

function DisableContinuousCrawl {
    Write-Host "# Begin Disabling continuous crawl" -ForegroundColor Green

    $ssa = Get-SPEnterpriseSearchServiceApplication
    if ($null -ne $ssa) {
        Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa | ForEach-Object { 
            $csname = $_.Name
            $cs = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa | Where-Object { $_.Name -eq $csname }
            try {
                $cs = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa | Where-Object { $_.Name -eq $csname }
                $cs.StopCrawl()
                Write-Host "## Stop running crawl on $($cs.Name)" -ForegroundColor Yellow
                Write-Host ""
            }
            catch {
                Write-Error $_
            }
        }

        Write-Host ""

        Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa | ForEach-Object { 
            $csname = $_.Name
            $cs = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa | Where-Object { $_.Name -eq $csname }
            Write-Host "## Disable ContinuousCrawls on $($cs.Name)" -ForegroundColor Yellow
                
            if ($cs.EnableContinuousCrawls -ne $false) {
                While (($cs.ContinuousCrawlStatus -eq "Crawling") -eq $false) {
                    Write-Host "Before Waiting for the content source to be 'Idle'..." -ForegroundColor DarkYellow
                    Write-Host "Current status : CrawlStatus='$($cs.CrawlStatus)' ContinuousCrawlStatus='$($cs.ContinuousCrawlStatus)'" -ForegroundColor DarkYellow
                    Start-Sleep 10
                    $cs = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa | Where-Object { $_.Name -eq $csname }
                }
                    
                    
                $cs.EnableContinuousCrawls = $false
                $cs.Update()
                Write-Host "## ContinuousCrawls is now disabled" -ForegroundColor DarkYellow
                Start-Sleep 10
            }
            else {
                Write-Host "## ContinuousCrawls already disabled" -ForegroundColor DarkYellow
            }
                
            While (($cs.ContinuousCrawlStatus -eq "Idle") -eq $false) {
                Write-Host "## After Waiting for the content source to be Idle..." -ForegroundColor DarkYellow
                Write-Host "## Current status : CrawlStatus='$($cs.CrawlStatus)' ContinuousCrawlStatus='$($cs.ContinuousCrawlStatus)'" -ForegroundColor DarkYellow
                Start-Sleep 10
                $cs = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa | Where-Object { $_.Name -eq $csname }
            }
            Write-Host ""
        }
    }

    Write-Host "# Ended Disabling continuous crawl" -ForegroundColor Green
    Write-Host ""
}


function ResetIndex { 
    $ssa = Get-SPEnterpriseSearchServiceApplication

    Write-Host "# Starting reseting Search index" -ForegroundColor Green
        
    $disableAlerts = $true

    $ignoreUnreachableServer = $true

    Write-Host ""
    Write-Host "## Reseting Search running" -ForegroundColor Yellow
    Write-Host ""

    $ssa.reset($disableAlerts, $ignoreUnreachableServer)

    if (-not $?) {
        Write-Error " An error has occured while trying to reset the search index" -ForegroundColor Red
    }

    Write-Host "# Reseting Search index Done" -ForegroundColor Green
    Write-Host ""
}

function EnableContinuousCrawl {
    Write-Host "# Enabling continuous crawl" -ForegroundColor Green
    Write-Host ""

    $ssa = Get-SPEnterpriseSearchServiceApplication

    Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa | ForEach-Object { 
        $csname = $_.Name
        $cs = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa | Where-Object { $_.Name -eq $csname }
        try {
            $cs = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa | Where-Object { $_.Name -eq $csname }
            $cs.StopCrawl()
                
            Write-Host "## Stop running crawl on $($cs.Name)" -ForegroundColor Yellow
            Write-Host ""

        }
        catch {
            Write-Error $_
        }
    }

    Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa | ForEach-Object { 
        $csname = $_.Name
        $cs = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa | Where-Object { $_.Name -eq $csname }
        Write-Host "## Enable ContinuousCrawls on $($cs.Name)" -ForegroundColor Yellow

        if ($cs.EnableContinuousCrawls -ne $true) {
            $cs.EnableContinuousCrawls = $true
            $cs.Update()
            Write-Host "## ContinuousCrawls is now enabled" -ForegroundColor DarkYellow
            Start-Sleep 10
        }
        else {
            Write-Host "## ContinuousCrawls already enabled" -ForegroundColor DarkYellow
        }
        Write-Host ""
    }

    Write-Host "# Ended Enabling continuous crawl" -ForegroundColor Green
    Write-Host ""
}

function ResetTopology ($servers, $pathIndex) {
    Write-Host "# Starting Reset Topology" -ForegroundColor Green
    Write-Host ""

    $ssa = Get-SPEnterpriseSearchServiceApplication
    $current = Get-SPEnterpriseSearchTopology -SearchApplication $ssa -Active
    $clone = New-SPEnterpriseSearchTopology -SearchApplication $ssa 

    foreach ($server in $servers) {
        $ssi = Get-SPEnterpriseSearchServiceInstance | Where-Object { $_.Server -match $server }

        Write-Host "## Add component on server $($server) " -ForegroundColor Yellow

        Write-Host "### Add Admin Component" -ForegroundColor DarkYellow
        $temp = New-SPEnterpriseSearchAdminComponent -SearchTopology $clone -SearchServiceInstance $ssi

        Write-Host "### Add Crawl Component" -ForegroundColor DarkYellow
        $temp = New-SPEnterpriseSearchCrawlComponent -SearchTopology $clone -SearchServiceInstance $ssi

        Write-Host "### Add Content Processing" -ForegroundColor DarkYellow
        $temp = New-SPEnterpriseSearchContentProcessingComponent -SearchTopology $clone -SearchServiceInstance $ssi

        Write-Host "### Add Analytics Processing" -ForegroundColor DarkYellow
        $temp = New-SPEnterpriseSearchAnalyticsProcessingComponent -SearchTopology $clone -SearchServiceInstance $ssi

        Write-Host "### Add Query Processing" -ForegroundColor DarkYellow
        $temp = New-SPEnterpriseSearchQueryProcessingComponent -SearchTopology $clone -SearchServiceInstance $ssi

        Write-Host "### Add Index Component with the index partion 0 and RootDirectory : $($pathIndex) " -ForegroundColor DarkYellow
        $temp = New-SPEnterpriseSearchIndexComponent -SearchTopology $clone -SearchServiceInstance $ssi -IndexPartition 0 -RootDirectory $pathIndex

        Write-Host ""
    }

    Write-Host "## Apply Topology" -ForegroundColor Yellow
    Set-SPEnterpriseSearchTopology -Identity $clone
    
    Write-Host "## Remove old Topology" -ForegroundColor Yellow
    Remove-SPEnterpriseSearchTopology -Identity $current -Confirm:$false

    Write-Host "# Ended Reset Topology" -ForegroundColor Green
    Write-Host ""
}

function SuspendSearch {
    Write-Host "# Suspend Search" -ForegroundColor Green
    $ssa = Get-SPEnterpriseSearchServiceApplication
    Suspend-SPEnterpriseSearchServiceApplication -Identity $ssa
    Write-Host "# Search is suspended" -ForegroundColor Green
    Write-Host ""
}

function ResumeSearch {
    Write-Host "# Resume Search" -ForegroundColor Green
    $ssa = Get-SPEnterpriseSearchServiceApplication
    Resume-SPEnterpriseSearchServiceApplication -Identity $ssa
    Write-Host "# Search is resumed" -ForegroundColor Green
    Write-Host ""
}

function RemoveIndex ($servers, $drive, $pathIndex) {
    foreach ($server in $servers) { 
        Get-ChildItem -Path "\\$($server)\$($drive)$\$($pathIndex)" -Include * -Recurse | ForEach-Object { Remove-Item -Path $_ -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue } 
    }
}

try {
    $config = Get-Config $Env
    $config

    DisableContinuousCrawl

    ResetIndex

    RemoveIndex -servers $config.Servers -disque $config.ServerDataDrive -pathIndex $config.SearchDataDirectory

    ResetTopology -servers $config.Servers -pathIndex "$($config.ServerDataDrive):\$($config.SearchDataDirectory)"

    EnableContinuousCrawl

    Get-SPEnterpriseSearchTopology -SearchApplication $config.SearchServiceAppName | Where-Object { $_.State -eq "Inactive" } | ForEach-Object { 
        Remove-SPEnterpriseSearchTopology -Identity $_ -Confirm:$false
    } 
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}