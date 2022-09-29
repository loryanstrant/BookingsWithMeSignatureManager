param(
    [Parameter(Mandatory=$true)]
    [string]$UPN
)
$ErrorActionPreferance = "Stop"

try {
    $tenantName = "YOURTENANTNAME.onmicrosoft.com"
    $connection = Get-AutomationConnection -Name 'AzureRunAsConnection'
    Connect-ExchangeOnline -CertificateThumbprint $connection.CertificateThumbprint -AppId $connection.ApplicationID -Organization $tenantName
    $userMbx = Get-Mailbox $UPN
    $userExchangeGuid = $userMbx.ExchangeGuid -replace '-'
    Write-Output $userExchangeGuid
    disconnect-ExchangeOnline
}
catch {
    Write-Error "Unable to add retrieve Mailbox GUID for $UPN"
}
