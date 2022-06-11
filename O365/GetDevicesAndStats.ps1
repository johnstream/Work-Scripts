$CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
 . "$CreateEXOPSSession\CreateExoPSSession.ps1"

$session = Connect-ExOPSSession -UserPrincipalName [ENTER UPN]
 
$csvDevices = "C:\temp\MobileDevices4.csv"
$csvStats = "C:\temp\MobileDeviceStats.csv"

$results = @()
$statistics = @()
$mailboxUsers = get-mailbox -resultsize unlimited
$mobileDevice = @()
$mobileDeviceStats = @()

 
foreach($user in $mailboxUsers)
{
$UPN = $user.UserPrincipalName
$displayName = $user.DisplayName
 
$mobileDevices = Get-MobileDevice -Mailbox $UPN
$mobileDeviceStats = Get-MobileDeviceStatistics -mailbox $UPN
       
      foreach($mobileDevice in $mobileDevices)
      {
          Write-Output "Getting info about a device for $displayName"
          $properties = @{
          Name = $user.name
          UPN = $UPN
          DisplayName = $displayName
          FriendlyName = $mobileDevice.FriendlyName
          ClientType = $mobileDevice.ClientType
          ClientVersion = $mobileDevice.ClientVersion
          DeviceId = $mobileDevice.DeviceId
          DeviceMobileOperator = $mobileDevice.DeviceMobileOperator
          DeviceModel = $mobileDevice.DeviceModel
          DeviceOS = $mobileDevice.DeviceOS
          DeviceTelephoneNumber = $mobileDevice.DeviceTelephoneNumber
          DeviceType = $mobileDevice.DeviceType
          FirstSyncTime = $mobileDevice.FirstSyncTime
          UserDisplayName = $mobileDevice.UserDisplayName
          }
          $results += New-Object psobject -Property $properties
      }
       
      foreach($mobileDevice1 in $mobileDeviceStats)
      {
          Write-Output "Getting stats about a device for $displayName"
          $statproperties = @{
          Name = $user.name
          UPN = $UPN
          DisplayName = $displayName
          DeviceID = $mobileDevice1.DeviceID
          FirstSyncTime = $mobileDevice1.FirstSyncTime
          LastSuccessSync = $mobileDevice1.LastSuccessSync
          DevicePolicyApplied = $mobileDevice1.DevicePolicyApplied
          IsRemoteWipeSupported = $mobileDevice1.IsRemoteWipeSupported
          }
          $statistics += New-Object psobject -Property $statproperties
      }
}
 
$results | Select-Object Name,UPN,FriendlyName,DisplayName,ClientType,ClientVersion,DeviceId,DeviceMobileOperator,DeviceModel,DeviceOS,DeviceTelephoneNumber,DeviceType,FirstSyncTime,UserDisplayName | Export-Csv -notypeinformation -Path $csvDevices

$statistics | Select-Object Name,UPN,DisplayName,DeviceID,FirstSyncTime,LastSuccessSync,DevicePolicyApplied,IsRemoteWipeSupported | Export-Csv -notypeinformation -path $csvStats
 



