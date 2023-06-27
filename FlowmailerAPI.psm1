
Class FlowmailerAPI
{
    [ValidateNotNullOrEmpty()][string]$ClientId
    [ValidateNotNullOrEmpty()][string]$ClientSecret
    [ValidateNotNullOrEmpty()][string]$AccountId

    [ValidateNotNullOrEmpty()][string]$LoginUrl
    [ValidateNotNullOrEmpty()][string]$ApiUrl
    [ValidateNotNullOrEmpty()][string]$OAuthTokenUrl
    [ValidateNotNullOrEmpty()][string]$Scope

    [string]$AccessToken = ""
}

Function New-FlowmailerAPI {

  [CmdletBinding()]

  Param (
    [ValidateNotNullOrEmpty()][string] $ClientId,
    [ValidateNotNullOrEmpty()][string] $ClientSecret,
    [ValidateNotNullOrEmpty()][string] $AccountId,
    [string] $LoginUrl = "https://login.flowmailer.net",
    [string] $ApiUrl = "http://api.flowmailer.net",
    [string] $OAuthTokenUrl = "",
    [string] $Scope = "api"
  )

  if ($OAuthTokenUrl -eq "") {
    $OAuthTokenUrl = $LoginUrl + "/oauth/token"
  }

  [FlowmailerAPI]@{
    ClientId = $ClientId
    ClientSecret = $ClientSecret
    AccountId = $AccountId

    LoginUrl = $LoginUrl
    ApiUrl = $ApiUrl
    OAuthTokenUrl = $OAuthTokenUrl
    Scope = $Scope
  }
}

Function Get-AccessToken ([FlowmailerAPI]$api) {

  $credentials = @{
      client_id = $api.ClientId
      client_secret = $api.ClientSecret
      grant_type = "client_credentials"
      scope = $api.Scope
  };

  $headers = @{
    "Content-Type" = "application/x-www-form-urlencoded"
    "Accept" = "application/vnd.flowmailer.v1.12+json"
  }

  Write-Verbose $api.OAuthTokenUrl
  Write-Verbose ($credentials | ConvertTo-JSON )
  Write-Verbose ($headers | ConvertTo-JSON )

  $response = Invoke-RestMethod -Uri $api.OAuthTokenUrl -Method Post -Body $credentials -Headers $headers -StatusCodeVariable response_status -SkipHttpErrorCheck

  Write-Verbose ($response_status | ConvertTo-JSON )
  Write-Verbose ($response | ConvertTo-JSON )

  if ($response_status -ne "200") {
    throw "Get-AccessToken " + $response_status + ": " + $response
  }

  return $response.access_token
}

Function Refresh-AccessToken ([FlowmailerAPI]$Api) {
  $Token = Get-AccessToken $Api
  $Api.AccessToken = $Token
}

Function Invoke-FlowmailerAPI ([FlowmailerAPI]$Api, [Microsoft.PowerShell.Commands.WebRequestMethod]$Method, [String]$Url, [System.Collections.Hashtable]$extra_headers = @{}, [System.Collections.Hashtable]$Body = $null, [Int]$tries = 3) {

  if($Api.AccessToken -eq "") {
    Refresh-AccessToken $Api
  }

  # TODO: refresh token als expires_in=60 voorbij is

  $headers = @{
    "Accept" = "application/vnd.flowmailer.v1.12+json;charset=UTF-8"
    "Content-Type" = "application/vnd.flowmailer.v1.12+json;charset=UTF-8"
    "Authorization" = "Bearer " + $api.AccessToken
  }

  ForEach ($Key in $extra_headers.Keys) {
    $headers.$Key = $extra_headers.$Key
  }

  Write-Verbose $Url
#  Write-Host $url
#  Write-Host $tries
#  Write-Host ($headers | ConvertTo-Json)

  $BodyString = $null;
  if ($null -ne $Body) {
    $BodyString = ($Body | ConvertTo-JSON)
  }

  $response = Invoke-RestMethod -Uri ($Api.ApiUrl + "/" + $url) -Method $Method -Headers $headers -Body $BodyString -SkipHeaderValidation -ResponseHeadersVariable response_headers -StatusCodeVariable response_status -SkipHttpErrorCheck

  if ($response_status -eq "401") {
    Write-Verbose ($response_status | ConvertTo-JSON )
    Write-Verbose ($response | ConvertTo-JSON )
    if($tries -lt 0) {
      Throw ("API call failed " + $response_status + ": " + ($response | ConvertTo-JSON ))
    }
    $api.AccessToken = ""
    #Refresh-AccessToken $api
    $tries = $tries - 1
    Write-Verbose ("Retry tries left: " + $tries)
    return Invoke-FlowmailerAPI $api Get $url $extra_headers $Body $tries
  }

  if ($response_status -ne "200" -and $response_status -ne "201" -and $response_status -ne "206") {
    Throw ("HTTP Error " + $response_status + ": " + ($response | ConvertTo-Json))
  }

  return $response, $response_headers
}

# https://flowmailer.com/apidoc/flowmailer-api.html#get_account_id_recipient_recipient_messages
Function Get-MessagesByRecipientPage {

  [CmdletBinding()]

  Param (
    [ValidateNotNullOrEmpty()][FlowmailerAPI] $Api,
    [ValidateNotNullOrEmpty()][string] $Recipient,
    [DateTime] $StartDate,
    [DateTime] $EndDate,
    [String] $Range
  )

  #Write-Debug ("Get Recipient Page: " + $range + " params: " + ($PSBoundParameters | ConvertTo-Json))

  $headers = @{
    "Range" = "items=" + $range
  }

  $matrixParams = ""
  if($startDate) {
    if($endDate) {

      # https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings?view=netframework-4.8
      $format = "yyyy-MM-dd'T'HH:mm:ssK"

      #Write-Host $startDate.toString($format)
      # 2014-12-16T05:50:00Z,2014-12-16T06:00:00Z

      $matrixParams = $matrixParams + ";daterange=" + $startDate.ToUniversalTime().toString($format) + "," + $endDate.ToUniversalTime().toString($format)
      #Write-Host $matrixParams
    }
  }

  $response, $response_headers = Invoke-FlowmailerAPI $Api Get ($Api.AccountId + "/recipient/" + $Recipient + "/messages" + $matrixParams) $headers

  $next_range = if($response_headers['Next-Range']) { $response_headers['Next-Range'].split('=')[1] }

  return $response, $next_range
}

Function Get-MessagesByRecipient {

  [CmdletBinding()]

  Param (
    [ValidateNotNullOrEmpty()][FlowmailerAPI] $Api,
    [ValidateNotNullOrEmpty()][string] $Recipient,
    [DateTime] $StartDate,
    [DateTime] $EndDate
  )

  $range = ":1000"

  while($range) {
    Write-Debug ("Get Recipient Page: " + $range)

    $list, $next_range = Get-MessagesByRecipientPage @PSBoundParameters -Range $range

    foreach ($message in $list) {
      Write-Output $message
    }

    $range = $next_range;
  }
}

# https://flowmailer.com/apidoc/flowmailer-api.html#get_account_id_sender_sender_messages
Function Get-MessagesBySenderPage {

  [CmdletBinding()]

  Param (
    [ValidateNotNullOrEmpty()][FlowmailerAPI] $Api,
    [ValidateNotNullOrEmpty()][string] $Sender,
    [DateTime] $StartDate,
    [DateTime] $EndDate,
    [String] $Range
  )

  #Write-Debug ("Get Sender Page: " + $range + " params: " + ($PSBoundParameters | ConvertTo-Json))

  $headers = @{
    "Range" = "items=" + $range
  }

  $matrixParams = ""
  if($startDate) {
    if($endDate) {

      # https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings?view=netframework-4.8
      $format = "yyyy-MM-dd'T'HH:mm:ssK"

      #Write-Host $startDate.toString($format)
      # 2014-12-16T05:50:00Z,2014-12-16T06:00:00Z

      $matrixParams = $matrixParams + ";daterange=" + $startDate.ToUniversalTime().toString($format) + "," + $endDate.ToUniversalTime().toString($format)
      #Write-Host $matrixParams
    }
  }

  $response, $response_headers = Invoke-FlowmailerAPI $Api Get ($Api.AccountId + "/sender/" + $Sender + "/messages" + $matrixParams) $headers

  $next_range = if($response_headers['Next-Range']) { $response_headers['Next-Range'].split('=')[1] }

  return $response, $next_range
}

# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_cmdletbindingattribute?view=powershell-7.3
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_commonparameters?view=powershell-7.3

Function Get-MessagesBySender {

  [CmdletBinding()]

  Param (
    [ValidateNotNullOrEmpty()][FlowmailerAPI] $Api,
    [ValidateNotNullOrEmpty()][string] $Sender,
    [DateTime] $StartDate,
    [DateTime] $EndDate
  )

  $range = ":1000"

  while($range) {
    Write-Debug ("Get Sender Page: " + $range)

    $list, $next_range = Get-MessagesBySenderPage @PSBoundParameters -Range $range

  #  Write-Host $next_range
  #  Write-Host ($list | ConvertTo-Json -Depth 100)

    foreach ($message in $list) {
      Write-Output $message
    }

    $range = $next_range;
  }
}

Function Get-Messages {

  [CmdletBinding()]

  Param (
    [ValidateNotNullOrEmpty()][FlowmailerAPI] $Api,
    [string] $Sender,
    [string] $Recipient,
    [DateTime] $StartDate,
    [DateTime] $EndDate
  )

  if($Sender -and -not $Recipient) {
    return Get-MessagesBySender @PSBoundParameters
  }
  if($Recipient -and -not $Sender) {
    return Get-MessagesByRecipient @PSBoundParameters
  }

  $range = ":1000"

  $searchPrio = 'Recipient'
  #$searchPrio = 'Sender'

  while($range) {
    Write-Debug ("Get " + $searchPrio + " Page: " + $range)

    #Write-Debug ("" + ($PSBoundParameters | ConvertTo-Json))
    $params = @{} + $PSBoundParameters

    #return

    if($searchPrio -eq 'Sender') {
      $params.Remove('Recipient')
      #Write-Debug ("" + ($params | ConvertTo-Json))

      $list, $next_range = Get-MessagesBySenderPage @params -Range $range
      if(-not $list) {
        break
      }

      $list = $list | Where-Object { $_.recipientAddress -eq $Recipient }

      if(-not $list) {
        $searchPrio = 'Recipient'
      }

    } elseif($searchPrio -eq 'Recipient') {
      $params.Remove('Sender')
      #Write-Debug ("" + ($params | ConvertTo-Json))

      $list, $next_range = Get-MessagesByRecipientPage @params -Range $range
      if(-not $list) {
        break
      }

      $list = $list | Where-Object { $_.senderAddress -eq $Sender }

      if(-not $list) {
        $searchPrio = 'Sender'
      }
    }

    foreach ($message in $list) {
      Write-Output $message
    }

    $range = $next_range;
  }
}

# https://flowmailer.com/apidoc/flowmailer-api.html#get_account_id_message_hold
Function Get-MessageHoldsPage {

  [CmdletBinding()]

  Param (
    [ValidateNotNullOrEmpty()][FlowmailerAPI] $Api,
    [DateTime] $StartDate,
    [DateTime] $EndDate,
    [String] $Range
  )

  $headers = @{
    "Range" = "items=" + $range
  }

  $matrixParams = ""
  if($startDate) {
    if($endDate) {

      # https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings?view=netframework-4.8
      $format = "yyyy-MM-dd'T'HH:mm:ssK"

      #Write-Host $startDate.toString($format)
      # 2014-12-16T05:50:00Z,2014-12-16T06:00:00Z

      $matrixParams = $matrixParams + ";daterange=" + $startDate.ToUniversalTime().toString($format) + "," + $endDate.ToUniversalTime().toString($format)
      #Write-Host $matrixParams
    }
  }

  $response, $response_headers = Invoke-FlowmailerAPI $Api Get ($Api.AccountId + "/message_hold" + $matrixParams) $headers

  $next_range = if($response_headers['Next-Range']) { $response_headers['Next-Range'].split('=')[1] }

  return $response, $next_range
}

Function Get-MessageHolds {

  [CmdletBinding()]

  Param (
    [ValidateNotNullOrEmpty()][FlowmailerAPI] $Api,
    [DateTime] $StartDate,
    [DateTime] $EndDate
  )

  $start = 0
  $pagesize = 100

  while($start -ge 0) {
    Write-Debug ("Get MessageHolds Page: " + $range)

    $list, $next_range = Get-MessageHoldsPage @PSBoundParameters -Range ($start.ToString() + "-" + $pagesize.ToString())

    foreach ($message in $list) {
      Write-Output $message
    }

    if($list.Length -lt $pagesize) {
      break;
    }

    $start = $start + $pagesize;
  }
}

Function Resume-Message {

  [CmdletBinding()]

  Param (
    [ValidateNotNullOrEmpty()][FlowmailerAPI] $Api,
    [string] $MessageId
  )

  $response, $response_headers = Invoke-FlowmailerAPI $Api Post ($Api.AccountId + "/messages/" + $MessageId + "/resume")
}

# https://flowmailer.com/apidoc/flowmailer-api.html#post_account_id_messages_submit
Function Submit-Message {

  [CmdletBinding()]

  Param (
    [ValidateNotNullOrEmpty()][FlowmailerAPI] $Api,
    [string] $Sender,
    [string] $Recipient,
    [string] $Subject,
    [string] $Text
  )

  $SubmitMessage = @{
    messageType = "EMAIL"
    senderAddress = $Sender
    recipientAddress = $Recipient
    subject = $Subject
    text = $Text
  }

  $response, $response_headers = Invoke-FlowmailerAPI $Api Post ($Api.AccountId + "/messages/submit") -Body $SubmitMessage
}

Export-ModuleMember -Function New-FlowmailerAPI,Get-AccessToken,Get-MessagesByRecipient,Get-MessagesBySender,Get-Messages,Get-MessageHolds,Resume-Message,Submit-Message
