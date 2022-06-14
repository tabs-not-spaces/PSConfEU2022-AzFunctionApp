using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)
try {
    $token =  (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/").Token
    $restCall = Invoke-RestMethod -Method Get -uri "https://graph.microsoft.com/beta/users" -Headers @{Authorization = "Bearer $token"} -ContentType 'Application/Json'
    write-output $($restCall | ConvertTo-Json)
    $resp = [PSCustomObject]@{
        token = $token
        restCall = $restCall | ConvertTo-Json -Depth 100
    }
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


