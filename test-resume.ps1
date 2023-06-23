
using module ./FlowmailerAPI.psm1
#Import-Module ./FlowmailerAPI.psm1

. $PSScriptRoot\config.ps1

$Api = New-FlowmailerAPI @ApiConfig

$startDate = get-date "01/01/2020"
$endDate = get-date "06/01/2023"

Get-MessageHolds $Api -StartDate $startDate -EndDate $endDate `
  | Where-Object { $_.errorText -like '*Message rejected: Sender domain*' } `
  | Where-Object { $_.status -eq 'NEW' } `
  | ForEach-Object -Process { Resume-Message $Api $_.messageId } `
