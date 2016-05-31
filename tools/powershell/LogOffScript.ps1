[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

$logoffCommand = "shutdown /l /f /t 00"
$eventsUrl = "http://someSite"

function buildRequestBody ($action) {
  #$timestamp = Get-Date -format u
  $dateStr = Get-Date -format yyyy-MM-dd
  $time = Get-Date -format HH:mm:ss
  $HOSTNAME = $env:computername

  $timestamp = $dateStr + 'T' + $time + '.000Z'

  $ip=get-WmiObject Win32_NetworkAdapterConfiguration|Select ipaddress,IPSubnet,DefaultIPGateway,MACAddress,DNSServerSearchOrder | Where {$_.Ipaddress.length -ge 1}
  $ipAddress = $ip[0].ipaddress[0]
  $mac = $ip[0].MACAddress
  $HOSTLAT = "40.249719"
	$HOSTLON = "-111.649265"
  switch ($action){
    "logoff" {
      $description = "TEC PC user initiated log off"
    }
    "default" {
      $description = "TEC PC logoff dialog cancelled"
    }
  }

  $requestBody = '{"type": "user","timestamp": "' + $timestamp + '","eventTime": "' + $time + '","eventDate": "' + $dateStr + '","device": {"hostname": "' + $HOSTNAME + '", "description": "", "ipAddress": "' + $ipAddress + '", "macAddress": "' + $mac + '"}, "room": { "building": "", "roomNumber": "","coordinates": "' + $HOSTLAT +',' + $HOSTLON +'", "floor": ""},"action": {"actor": "desktopLogoffButton", "description": "' + $description + '"},"session": ""}'
  $converted = ConvertTo-Json $requestBody

  return $requestBody
}

function logEvent ($action) {
  $body = buildRequestBody($action)

  switch ($action){
    "logoff" { $resp = Invoke-RestMethod -Uri $eventsUrl -Body $body -Method Post; Write-Host($resp | Select *); break }
    "cancel" { Invoke-RestMethod -Uri $eventsUrl -Body $body -Method Post; Write-Host($resp | Select *); break }
    default { break }
  }
}

$rv = [Microsoft.VisualBasic.Interaction]::MsgBox('Are you sure you want to log off?','YesNoCancel, Exclamation,MsgBoxSetForeground,SystemModal', 'Accept or Deny')

Switch ($rv) {
  'Yes' {
    logEvent("logoff")
    #iex $logoffCommand
    }
  'No' {
    # Cancelling shutdown
    logEvent("cancel")
    }
  'Cancel' {
    logEvent("cancel")
    }
}
