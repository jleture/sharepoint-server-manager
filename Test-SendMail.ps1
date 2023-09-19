[CmdletBinding()]
param(
    [string]$Env = "LAB",
    [Parameter(Mandatory = $false)][ValidateNotNullorEmpty()][int]$Port = 25,
    [Parameter(Mandatory = $true)][ValidateNotNullorEmpty()][string]$EmailTo,
    [Parameter(Mandatory = $true)][ValidateNotNullorEmpty()][string]$SiteUrl,
    [Parameter(Mandatory = $false)][ValidateNotNullorEmpty()][string]$Subject = "Email Subject",
    [Parameter(Mandatory = $false)][ValidateNotNullorEmpty()][string]$Body = "Email Body"
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
    Write-Verbose "Port: $Port"
    Write-Verbose "EmailTo: $EmailTo"
    Write-Verbose "SiteUrl: $SiteUrl"
    Write-Verbose "Subject: $Subject"
    Write-Verbose "Body: $Body"

    $config = Get-Config $Env
    $config

    Write-Host "Testing SPUtility..." -NoNewline:$True
    $Web = Get-SPWeb $SiteUrl
    [Microsoft.SharePoint.Utilities.SPUtility]::SendEmail($Web, 0, 0, $EmailTo, "SPUtility - $Subject", $Body) | Out-Null
    Write-Host " [OK]" -ForegroundColor Green

    $SPGlobalAdmin = New-Object Microsoft.SharePoint.Administration.SPGlobalAdmin
    $SMTPServer = $SPGlobalAdmin.OutboundSmtpServer
    $EmailToFrom = $SPGlobalAdmin.MailFromAddress

    Write-Verbose "SMTPServer: $SMTPServer"
    Write-Verbose "EmailToFrom: $EmailToFrom"
    
    $Message = new-object Net.Mail.MailMessage
    $SMTP = new-object Net.Mail.SmtpClient($SMTPServer)
    $SMTP.Port = $Port
    $Message.From = $EmailToFrom
    $Message.To.Add($EmailTo)
    $Message.subject = "MailMessage - $Subject"
    $Message.body = $Body
    
    Write-Host "Testing MailMessage..." -NoNewline:$True
    $SMTP.Send($Message)
    Write-Host " [OK]" -ForegroundColor Green


    Write-Host "Testing Send-MailMessage..." -NoNewline:$True
    Send-MailMessage -To $EmailTo -From $EmailToFrom -Subject "Send-MailMessage - $Subject" -Body $Body -BodyAsHtml -SmtpServer $SmtpServer -Port $Port
    Write-Host " [OK]" -ForegroundColor Green
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}