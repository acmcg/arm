$credential = get-credential
Connect-AzAccount -Credential $credential
$azContext = Get-AzContext
$tenantID  = $azContext.Tenant.TenantId
$subscriptionID = $azContext.Subscription.Id
#get parameters

$ADApplication = get-AzADApplication -DisplayName "ARMaccess"
if(!$ADApplication){
    #create SP and Application
    $ADApplication = New-AzADApplication -DisplayName "ARMaccess" -IdentifierUris https://domain.local
    $ADServicePrincipal = New-AzADServicePrincipal -ApplicationId $ADApplication.ApplicationId -DisplayName "ARMaccess"
    $SecureStringPassword = Read-Host -AsSecureString 
    $appCredential = New-AzADAppCredential -ApplicationId $ADApplication.ApplicationId -Password $SecureStringPassword -EndDate (get-date).AddDays(365)
    New-AzRoleAssignment -ApplicationId $ADApplication.ApplicationId -RoleDefinitionName reader -Scope "/subscriptions/$($azContext.Subscription.Id)"
}

#get token
$apiVersion = "?api-version=2019-03-01"
$uri = "https://management.azure.com/subscriptions/$subscriptionID/$apiVersion"
$body= @{resource="https://management.azure.com/";client_id=$($ADApplication.ApplicationId);grant_type='client_credentials';client_secret = ''}
$authEndPoint = "https://login.microsoftonline.com/$tenantID/oauth2/token"
$token = Invoke-RestMethod -Method Post -Uri $authEndPoint -Body $body
$authHeader = @{Authorization = "Bearer $($token.access_token)"}

#get resource
$apiVersion = "?api-version=2018-11-01"
# start loop with virtualnetworks
$uri = "https://management.azure.com/providers/microsoft.network/$apiVersion"


$virtualNetworks = Invoke-RestMethod -Method Get -Uri $uri -Headers $authHeader
# $virtualNetworks.resourceTypes | Where-Object {$_.locations -match "Australia Central"}

ForEach ($resourceTypes in $virtualNetworks.resourceTypes) {
    
    write-output "$($resourceTypes.resourceType) : $($resourceTypes.apiVersions[0])"
    $apiVersion = "?api-version=$($resourceTypes.apiVersions[0])"
    $uri = "https://management.azure.com/providers/microsoft.network/$($resourceTypes.resourceType)/$apiVersion"
    Invoke-RestMethod -Method Get -Uri $uri -Headers $authHeader


}
