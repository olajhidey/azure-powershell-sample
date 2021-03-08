using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)
Import-Module -Name MSOnline -UseWindowsPowerShell
#user details loaded from App settings
$securedPassword = ConvertTo-SecureString  $Env:password -AsPlainText -Force
$Credential = [System.management.automation.pscredential]::new($Env:myEmail, $SecuredPassword)
# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."
# Interact with query parameters or the body of the request.
Connect-Msolservice -Credential $Credential
Write-Host "<==== Print MSOnline Users Start ====>"
Get-MsolUser
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