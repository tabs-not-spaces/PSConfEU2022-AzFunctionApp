using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)
try {
    #region auth
    Write-Output "Starting the authentication dance.."
    if ($env:MSI_SECRET) { $token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/").Token }
    else {
        
        Disable-AzContextAutosave -Scope Process | Out-Null
        $cred = New-Object System.Management.Automation.PSCredential $env:appId, ($env:secret | ConvertTo-SecureString -AsPlainText -Force)
        Connect-AzAccount -ServicePrincipal -Credential $cred -Tenant $env:tenant | Out-Null
        $token = (Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com').Token
    }
    Write-Output "In the mainframe... let's go!"
    #endregion



    #region params
    $result             = [System.Collections.ArrayList]::new()
    $appName            = "Notepad++"
    $encodedAppName     = [System.Web.HttpUtility]::UrlEncode($appName)
    $groupId            = 'd20a418e-a00f-47a5-a8ed-e12a9d98f83a'
    $baseGraphUri       = 'https://graph.microsoft.com/beta'
    $script:authHeader  = @{Authorization = "Bearer $token" }
    #endregion



    #region functions
    function Invoke-GraphCall {
        [cmdletbinding()]
        param (
            [parameter(Mandatory = $false)]
            [ValidateSet('Get', 'Post', 'Delete')]
            [string]$Method = 'Get',

            [parameter(Mandatory = $false)]
            [hashtable]$Headers = $script:authHeader,

            [parameter(Mandatory = $true)]
            [string]$Uri,

            [parameter(Mandatory = $false)]
            [string]$ContentType = 'Application/Json',

            [parameter(Mandatory = $false)]
            [hashtable]$Body
        )
        try {
            $params = @{
                Method      = $Method
                Headers     = $Headers
                Uri         = $Uri
                ContentType = $ContentType
            }
            if ($Body) {
                $params.Body = $Body | ConvertTo-Json -Depth 20
            }
            $query = Invoke-RestMethod @params
            return $query
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    }
    function Format-Result {
        [cmdletbinding()]
        param (
            [parameter(Mandatory = $true)]
            [string]$DeviceID,

            [parameter(Mandatory = $true)]
            [string]$DeviceName,

            [parameter(Mandatory = $true)]
            [bool]$IsCompliant,

            [parameter(Mandatory = $true)]
            [bool]$IsMember,

            [parameter(Mandatory = $true)]
            [ValidateSet('Added', 'Removed', 'NoActionTaken')]
            [string]$Action
        )
        $result = [PSCustomObject]@{
            DeviceID    = $DeviceID
            DeviceName  = $DeviceName
            IsCompliant = $IsCompliant
            IsMember    = $IsMember
            Action      = $Action
        }
        return $result
    }
    #endregion



    #region Get existing group members
    Write-Output "Getting existing group members.."
    $graphUri = "$baseGraphUri/groups/$groupId/members"
    $groupMembers = Invoke-GraphCall -Uri $graphUri
    Write-Output "Found $($groupMembers.value.count) current members.."
    #endregion



    #region Get devices with notepad++ installed
    Write-Output "Find the detected app object.."
    $detectedAppsBaseUri = "$baseGraphUri/deviceManagement/detectedApps"
    $daFilter = '$filter=(contains(displayName,{0}))' -f "'$([System.Web.HttpUtility]::UrlEncode($encodedAppName))'"
    $daItem = (Invoke-GraphCall -Uri "$($detectedAppsBaseUri)?$daFilter").value
    if ($daItem.deviceCount -gt 0) {
        Write-Output "Find devices with $appName.."
        $detectedDevices = (Invoke-GraphCall -Uri "$detectedAppsBaseUri/$($daItem.id)/managedDevices").value | Select-Object id, deviceName
        Write-Output "Found $($detectedDevices.value.count) devices with $appName installed.."
        
        foreach ($device in $detectedDevices) {
            #region Swap the detected device for Intune + AAD object from Intune object'
            Write-Output "Lets get the Intune / AAD data for device: $($device.id).."
            $intuneDevice = Invoke-GraphCall -Uri "$baseGraphUri/deviceManagement/managedDevices/$($device.id)"
            $aadFilter = '$filter=(deviceId eq {0})' -f "'$($intuneDevice.azureADDeviceId)'"
            $aadDevice = (Invoke-GraphCall -Uri "$baseGraphUri/devices?$aadFilter").value
            $device | Add-Member -MemberType NoteProperty -Name "deviceId" -Value $aadDevice.deviceId
            #endregion



            #region add devices
            if ($groupMembers.value.deviceId -notcontains $aadDevice.deviceId) {
                #region Device not in group and has software
                Write-Output "Adding to group!"
                $graphUri = '{0}/groups/{1}/members/$ref' -f $baseGraphUri, $groupId
                $body = @{"@odata.id" = "$baseGraphUri/directoryObjects/$($aadDevice.id)" }
                Invoke-GraphCall -Uri $graphUri -Method Post -Body $body
                $result.Add($(Format-Result -DeviceId $device.deviceId -DeviceName $device.deviceName -IsCompliant $true -IsMember $true -Action Added)) | Out-Null
                #endregion
            }
            else {
                #region device is compliant and already a member
                Write-Output "Device already a member!"
                $result.Add($(Format-Result -DeviceId $device.deviceId -DeviceName $device.deviceName -IsCompliant $true -IsMember $true -Action NoActionTaken)) | Out-Null
                #endregion
            }
            #endregion
        }



        #region Remove devices
        $devicesToRemove = $groupMembers.value | Where-Object { $_.deviceId -notIn $detectedDevices.deviceId }
        if ($devicesToRemove.count -gt 0) {
            foreach ($dtr in $devicesToRemove) {
                
                
                
                #region Device found in group, but doesnt have software.
                Write-Output "`nDevice $($dtr.id) shouldn't be a member!"
                $graphUri = "$baseGraphUri/groups/$groupId/members/$($dtr.id)/`$ref"
                Invoke-GraphCall -Uri $graphUri -Method Delete
                $result.Add($(Format-Result -DeviceId $dtr.deviceId -DeviceName $dtr.displayName -IsCompliant $false -IsMember $false -Action Removed)) | Out-Null
                #endregion



            }
        
        }
        #endregion
        $statusCode = [HttpStatusCode]::OK
        $body = $result
    }
    #endregion
}
catch {
    Write-Output $_.Exception.Message
    $statusCode = [HttpStatusCode]::BadRequest
    $body = $_.Exception.Message
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = $statusCode
        Body       = $body
    })