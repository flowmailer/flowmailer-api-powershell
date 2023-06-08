
using module ./FlowmailerAPI.psm1

#Import-Module ./FlowmailerAPI.psm1

. ./config.ps1

$api = [FlowmailerAPI]@{
  ClientId = $client_id
  ClientSecret = $client_secret
  AccountId = $account_id

  LoginUrl = "https://login.flowmailer.net"
  ApiUrl = "http://api.flowmailer.net"

#  AccessToken = "MzhmNDQzODE3MDJmZWRhNjk1NzZiM2NmNzkyZmVjNDhmNTNkYzNmYTExODQwZjg0YjE0OGQzODIxZmY1MzE5MAAAMTY4NjIyMzUwMzE4OQBhcGkAADAyMTc2YTg5OWQzMDQ1ZTkwN2Y5ZTg0YmYxNThhYzAxMDNmMjFmZDkAbUgxb2dzY0oA"
}

$startDate = get-date "11/16/2020"
$startDate = get-date "11/16/2021"
$endDate = get-date "05/22/2023"

#Get-MessagesByRecipient $api "casper@flowmailer.com" $startDate $endDate
Get-MessagesBySender $api "casper@caspermout.nl" $startDate $endDate
  | Where-Object { $_.senderAddress -eq 'casper@caspermout.nl' }
  | Where-Object { $_.messageIdHeader -like '*@return.flowmailer.net>' }
#  | ForEach-Object -Process { $_.id }
#  | ForEach-Object -Process { $_.submitted }
#  | ForEach-Object -Process { $_ | ConvertTo-Json }

#  | ConvertTo-Json