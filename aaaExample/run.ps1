using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)
try {
    if ($env:MSI_SECRET) { $token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/").Token }
    else {
        $cred = (New-Object System.Management.Automation.PSCredential $env:appId, ($env:secret | ConvertTo-SecureString -AsPlainText -Force))
        Connect-AzAccount -ServicePrincipal -Credential $cred -Tenant $env:tenant | Out-Null
        $token = (Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com').Token
    }

    $restCall = Invoke-RestMethod -Method Get -uri "https://graph.microsoft.com/beta/devices" -Headers @{Authorization = "Bearer $token"} -ContentType 'Application/Json'
    write-output $($restCall | ConvertTo-Json)
    $resp = $restCall.value | ConvertTo-Json -Depth 100
    $statusCode = [HttpStatusCode]::OK
    $body = $resp
}
catch {
    write-output $_.Exception.Message
    $statusCode = [HttpStatusCode]::BadRequest
    $body = $_.Exception.Message
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = $statusCode
    Body = $body
})


