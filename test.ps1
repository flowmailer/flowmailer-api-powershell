
using module ./FlowmailerAPI.psm1
#Import-Module ./FlowmailerAPI.psm1

. ./config.ps1

$Api = New-FlowmailerAPI -ClientId $client_id -ClientSecret $client_secret -AccountId $account_id # -LoginUrl "https://login.flowmailer.net" -ApiUrl "http://api.flowmailer.net"

$startDate = get-date "11/16/2020"
$endDate = get-date "05/22/2023"

Get-MessagesByRecipient $Api -Recipient "casper@flowmailer.com" -StartDate $startDate -EndDate $endDate
  | Where-Object { $_.senderAddress -eq 'casper@caspermout.nl' }
  | Where-Object { $_.subject -ne 'test' }
  | ForEach-Object -Process { $_.id }

Get-MessagesBySender $Api -Sender "casper@caspermout.nl" -StartDate $startDate -EndDate $endDate
  | Where-Object { $_.recipientAddress -eq 'casper@flowmailer.com' }
  | Where-Object { $_.subject -ne 'test' }
  | ForEach-Object -Process { $_.id }

Get-Messages $Api -Sender "casper@caspermout.nl" -Recipient "casper@flowmailer.com" -StartDate $startDate -EndDate $endDate # -Debug
  | Where-Object { $_.subject -ne 'test' }
  | ForEach-Object -Process { $_.id }
