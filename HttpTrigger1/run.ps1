using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

Import-Module -Name AzureAD -UseWindowsPowerShell

#user details loaded from App settings
$securedPassword = ConvertTo-SecureString  $Env:password -AsPlainText -Force
$Credential = [System.management.automation.pscredential]::new($Env:myEmail, $SecuredPassword)


# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

Write-Host $Env:myEmail

Connect-AzureAD -Credential $Credential

# Display users
Write-Host "<==== Print MSOnline Users Start ====>"

Get-AzureADUser -Filter "userPrincipalName eq 'xxxxxxxxxxxxxx'"

$TenantId = "xxxxxxxxxxxxxxx"
$ClientId = "xxxxxxxxxxxxxxx"

$TokenRequestParams = @{
    Method = 'POST'
    Uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    ContentType = "application/x-www-form-urlencoded"
    body = @{
       client_id  = $ClientId
       grant_type = "password"
       username = $Env:myEmail
       password = $Env:password
       scope = "user.read openid profile offline_access"
    }
}

$TokenRequest = try{
    Invoke-RestMethod @TokenRequestParams -ErrorAction Stop
}catch{
    $Message = $_.ErrorDetails.Message | ConvertFrom-Json
    if ($Message.error -ne "authorization_pending") {
        throw
    }
}

Write-Output $TokenRequest.access_token

# Interact with query parameters or the body of the request.
$name = $Request.Query.Name
if (-not $name) {
    $name = $Request.Body.Name
}

$body = "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."

if ($name) {
    $body = "Hello, $name. This HTTP triggered function executed successfully."
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})

