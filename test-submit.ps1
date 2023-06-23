
using module .\FlowmailerAPI.psm1
#Import-Module .\FlowmailerAPI.psm1

. $PSScriptRoot\config.ps1

$Api = New-FlowmailerAPI @ApiConfig

Submit-Message $Api `
  -Sender "casper@flowmailer.com" `
  -Recipient "casper@flowmailer.com" `
  -Subject "test 123" `
  -Text "bla" `
