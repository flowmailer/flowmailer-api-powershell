
Class FlowmailerAPI
{
    # Optionally, add attributes to prevent invalid values
    [ValidateNotNullOrEmpty()][string]$ClientId
    [ValidateNotNullOrEmpty()][string]$ClientSecret
    [ValidateNotNullOrEmpty()][string]$AccountId

    [ValidateNotNullOrEmpty()][string]$LoginUrl
    [ValidateNotNullOrEmpty()][string]$ApiUrl

    [string]$AccessToken = ""
}

Function New-FlowmailerAPI {

  [CmdletBinding(SupportsPaging)]

  Param (
    [ValidateNotNullOrEmpty()][string] $ClientId,
    [ValidateNotNullOrEmpty()][string] $ClientSecret,
    [ValidateNotNullOrEmpty()][string] $AccountId,
    [string] $LoginUrl = "https://login.flowmailer.net",
    [string] $ApiUrl = "http://api.flowmailer.net"
  )

  [FlowmailerAPI]@{
    ClientId = $ClientId
    ClientSecret = $ClientSecret
    AccountId = $AccountId

    LoginUrl = $LoginUrl
    ApiUrl = $ApiUrl
  }
}

Function Get-AccessToken ([FlowmailerAPI]$api) {

  $credentials = @{
      client_id = $api.ClientId
      client_secret = $api.ClientSecret
      grant_type = "client_credentials"
      scope = "api"
  };

  $headers = @{
    "Content-Type" = "application/x-www-form-urlencoded"
  }

  $response = Invoke-RestMethod -Uri ($api.LoginUrl + "/oauth/token") -Method Post -Body $credentials -Headers $headers

#  Write-Host $response

  return $response.access_token
}

Function Refresh-AccessToken ([FlowmailerAPI]$api) {
  $token = Get-AccessToken $api
  $api.AccessToken = $token
}

Function Invoke-FlowmailerAPI ([FlowmailerAPI]$api, [String]$Url, [System.Collections.Hashtable]$extra_headers = @{}, [Int]$tries = 3) {

  if($api.AccessToken -eq "") {
    Refresh-AccessToken $api
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

#  Write-Host $url
#  Write-Host $tries
#  Write-Host ($headers | ConvertTo-Json)

  $response = Invoke-RestMethod -Uri ($api.ApiUrl + "/" + $url) -Method Get -Headers $headers -SkipHeaderValidation -ResponseHeadersVariable response_headers -StatusCodeVariable response_status -SkipHttpErrorCheck

  if ($response_status -eq "401") {
    if($tries -lt 0) {
      return 1, -1
    }
    $api.AccessToken = ""
    #Refresh-AccessToken $api
    $tries = $tries - 1
    return Invoke-FlowmailerAPI $api $url $extra_headers $tries
  }

  if ($response_status -ne "206") {
    Write-Host $response_status
    Write-Host $response_headers
    Write-Host $response_headers['Next-Range']
    Write-Host ($response | ConvertTo-JSON)
  }

  return $response, $response_headers
}

# https://flowmailer.com/apidoc/flowmailer-api.html#get_account_id_recipient_recipient_messages
Function Get-MessagesByRecipientPage {

  [CmdletBinding(SupportsPaging)]

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

  $response, $response_headers = Invoke-FlowmailerAPI $Api ($Api.AccountId + "/recipient/" + $Recipient + "/messages" + $matrixParams) $headers

  $next_range = if($response_headers['Next-Range']) { $response_headers['Next-Range'].split('=')[1] }

  return $response, $next_range
}

Function Get-MessagesByRecipient {

  [CmdletBinding(SupportsPaging)]

  Param (
    [ValidateNotNullOrEmpty()][FlowmailerAPI] $Api,
    [ValidateNotNullOrEmpty()][string] $Recipient,
    [DateTime] $StartDate,
    [DateTime] $EndDate
  )

  $range = ":10"

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

  [CmdletBinding(SupportsPaging)]

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

  $response, $response_headers = Invoke-FlowmailerAPI $Api ($Api.AccountId + "/sender/" + $Sender + "/messages" + $matrixParams) $headers

  $next_range = if($response_headers['Next-Range']) { $response_headers['Next-Range'].split('=')[1] }

  return $response, $next_range
}

# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_cmdletbindingattribute?view=powershell-7.3
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_commonparameters?view=powershell-7.3

Function Get-MessagesBySender {

  [CmdletBinding(SupportsPaging)]

  Param (
    [ValidateNotNullOrEmpty()][FlowmailerAPI] $Api,
    [ValidateNotNullOrEmpty()][string] $Sender,
    [DateTime] $StartDate,
    [DateTime] $EndDate
  )

  $range = ":10"

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

  [CmdletBinding(SupportsPaging)]

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

  $range = ":10"

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

Export-ModuleMember -Function New-FlowmailerAPI,Get-AccessToken,Get-MessagesByRecipient,Get-MessagesBySender,Get-Messages
