
using module ./FlowmailerAPI.psm1
#Import-Module ./FlowmailerAPI.psm1

. ./config.ps1

$Api = New-FlowmailerAPI -ClientId $client_id -ClientSecret $client_secret -AccountId $account_id # -LoginUrl "https://login.flowmailer.net" -ApiUrl "http://api.flowmailer.net"

$startDate = get-date "01/01/2020"
$endDate = get-date "06/01/2023"

Get-MessageHolds $Api -StartDate $startDate -EndDate $endDate # -Debug -Verbose
  | Where-Object { $_.errorText -like '*Message rejected: Sender domain*' }
  | Where-Object { $_.status -eq 'NEW' }
  | ForEach-Object -Process { Resume-Message $Api $_.messageId }

