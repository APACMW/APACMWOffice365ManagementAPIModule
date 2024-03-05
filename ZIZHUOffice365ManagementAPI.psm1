<#
 .Synopsis
  PowerShell SDK for Office365ManageAPI

 .Description
  IT Admin can use this PowerShell module to call Office365ManagementAPI. It suppports all operations of Office365ManagementAPI. Also supports Webhook subscriptions and notifications.

 .Example
   # Installl and import this PowerShell Module     
   Set-location 'C:\Program Files\WindowsPowerShell\Modules';
   MD ZIZHUOffice365ManagementAPI
   Copy 2 files(ZIZHUOffice365ManagementAPI.psd1 and ZIZHUOffice365ManagementAPI.psm1) to this folder ZIZHUOffice365ManagementAPI
   Import-Module ZIZHUOffice365ManagementAPI

 .Example
   # Connect to ZIZHUOffice365ManagementAPI module via client secret
$clientID = 'bc4db1db-b705-434a-91ff-145aa94185c8';
$tenantId = 'cff343b2-f0ff-416a-802b-28595997daa2';
$clientSecret = '';
Connect-Office365ManagementAPI -tenantID $tenantId -clientID $clientID -ClientSecret $clientSecret;
   
   # Connect to ZIZHUOffice365ManagementAPI module via client certificate
$clientID = 'bc4db1db-b705-434a-91ff-145aa94185c8';
$tenantId = 'cff343b2-f0ff-416a-802b-28595997daa2';
$thumbprint = '15958E05E3E4C2E563CE9BC346B25A2D70867048';
$clientcertificate= get-item "cert:\localmachine\my\$thumbprint";
Connect-Office365ManagementAPI -tenantID $tenantId -clientID $clientID -clientcertificate $clientcertificate;

   # Connect to ZIZHUOffice365ManagementAPI module via user sign-in
$clientID = '9b0547c4-28b1-466d-a80e-677c6dc42d42';
$tenantId = 'cff343b2-f0ff-416a-802b-28595997daa2';
$redirectUri='https://login.microsoftonline.com/common/oauth2/nativeclient'
$loginHint = 'freeman@vjqg8.onmicrosoft.com';
Connect-Office365ManagementAPI -tenantID $tenantId -clientID $clientID -loginHint $loginHint -redirectUri $redirectUri;

   # List available content and receive audit data
$startTime = "2024-03-02T00:00:00"; 
$endTime = "2024-03-03T00:00:00";
$blobs = Get-AvailableContent -startTime $startTime -endTime $endTime;
Receive-Content -blobs $blobs;

   # Get current subscriptions/Stop subscriptions
Get-CurrentSubscriptions;
Stop-Subscription -contentType AuditSharePoint;
Stop-Subscriptions;

   # Start thesubscriptions. If don't pass $webHookBody, no webhook for the subscription
$webhookEndpoint='https://7104-2404-f801-9000-18-b055-6c1c-9de7-1729.ngrok-free.app/api/O365ManagementAPIHttpFunction';
$authId = 'ZIZHUOffice365ManagementAPINotification20240220';
$expiration= "2024-04-14T00:00:00";
$webHookBody=
@"
{
    "webhook" : {
        "address": "$($webhookEndpoint)",
        "authId": "$($authId)",
        "expiration": "$($expiration)"
    }
}
"@;
Start-Subscription AuditAzureActiveDirectory $webHookBody;
Start-Subscription AuditExchange $webHookBody;
Start-Subscription AuditSharePoint $webHookBody;
Start-Subscription AuditGeneral $webHookBody;
Start-Subscription DLPAll $webHookBody;

  # List the notifications
$startTime = "2024-03-02T00:00:00"; 
$endTime = "2024-03-03T00:00:00";
Get-Notifications -startTime $startTime -endTime $endTime -contentType AuditExchange;

  # Receive the FriendlyNames for DLP Resource
Receive-ResourceFriendlyNames;

  # Clean after usgae
Disconnect-Office365ManagementAPI;
Get-Module ZIZHUOffice365ManagementAPI | Remove-Module;
#>

enum ContentType {
    AuditAzureActiveDirectory    
    AuditExchange
    AuditSharePoint
    AuditGeneral
    DLPAll
};
class Blob {
    [string]$contentUri
    [string]$contentId
    [string]$contentType
    [datetime]$contentCreated
    [datetime]$contentExpiration
    [System.Object[]]$auditRecords
}
class WebHook {
    [string]$authId
    [string]$address
    [string]$expiration
    [string]$status    
}
class Subscription {
    [string]$contentType
    [string]$status
    [WebHook]$webhook     
}
class Notification {
    [string]$contentType
    [string]$contentId
    [string]$contentUri
    [string]$notificationStatus
    [datetime]$contentCreated
    [datetime]$notificationSent    
    [datetime]$contentExpiration    
}
[string]$script:tenantID = $null;
[string]$script:clientId = $null;
[string]$script:clientsecret = $null;
[string]$script:redirectUri = $null;
[string]$script:loginHint = $null;
[X509Certificate]$script:clientcertificate = $null;
$script:AuthResult = $null;
[string]$script:Scope = 'https://manage.office.com/.default';
$script:httpErrorResponse = $null;
$script:maxretries = 10;
$script:sleepSeconds = 1;
function Show-HttpErrorResponse {
    Write-Output $script:httpErrorResponse;
}

function Show-LastErrorDetails {
    param(
        $lastError = $Error[0]
    )
    $lastError | Format-List -Property * -Force;
    $lastError.InvocationInfo | Format-List -Property *;
    $exception = $lastError.Exception;
    for ($depth = 0; $null -ne $exception; $depth++) {
        Write-Host "$depth" * 80 -ForegroundColor Green;                                            
        $exception | Format-List -Property * -Force                 
        $exception = $exception.InnerException                      
    }
}  
function Get-ContentTypeString {
    param(
        [Parameter(Mandatory = $true)][ContentType]$contentTypeData
    )
    [string]$result = $null;
    switch ($contentTypeData) {
        AuditAzureActiveDirectory { $result = "Audit.AzureActiveDirectory"; Break }
        AuditExchange { $result = "Audit.Exchange"; Break }
        AuditSharePoint { $result = "Audit.SharePoint"; Break }
        AuditGeneral { $result = "Audit.General"; Break }
        DLPAll { $result = "DLP.All"; Break }
        Default {
            Write-Error "unknown type: $contentTypeData" -ErrorAction Stop;
        }
    }
    Write-Output $result;
    return;
}
function Get-ContentTypeEnum {
    param(
        [Parameter(Mandatory = $true)][string]$contentTypeString
    )
    switch ($contentTypeString) {
        "Audit.AzureActiveDirectory" { $result = [ContentType]::AuditAzureActiveDirectory; Break }
        "Audit.Exchange" { $result = [ContentType]::AuditExchange; Break }
        "Audit.SharePoint" { $result = [ContentType]::AuditSharePoint; Break }
        "Audit.General" { $result = [ContentType]::AuditGeneral; Break }
        "DLP.All" { $result = [ContentType]::DLPAll; Break }
        Default {
            Write-Error "unknown type: $contentTypeData" -ErrorAction Stop;
        }
    }
    Write-Output $result;
    return;
}
Function Invoke-O365APIHttpRequest {
    param (
        [Parameter(Mandatory = $true)][string]$url,
        [Parameter(Mandatory = $true)][string]$httpVerb,
        [Parameter(Mandatory = $false)][string]$requstBody
    )
    if ($null -eq $script:AuthResult) {
        Write-Error "Not authenticated. Stop." -ErrorAction Stop;
    }
    $httpResponse = $null;
    $script:httpErrorResponse = $null;
    $retryCount = 0;
    do {
        if ($retryCount -gt 1) {
            Start-Sleep -Seconds ($script:sleepSeconds * $retryCount);
        }
        Get-OauthToken;
        $headerParams = @{'Authorization' = "$($script:AuthResult.tokentype) $($script:AuthResult.accesstoken)" };
        try {
            if ($PSBoundParameters.ContainsKey("requstBody")) {
                $httpResponse = Invoke-WebRequest -uri $url -Headers $headerParams -Method $httpVerb -ContentType "application/json" -Body $requstBody;
            }
            else {
                $httpResponse = Invoke-WebRequest -uri $url -Headers $headerParams -Method $httpVerb;
            }
        }
        catch {
            $script:httpErrorResponse = $_.Exception.Response;
        }
        finally {
            $retryCount = $retryCount + 1;
        }
    } until (($null -ne $httpResponse) -or ($retryCount -gt $script:maxretries))

    if (($null -ne $httpResponse) -and ($httpResponse.StatusCode -in (200, 204))) {
        Write-Output $httpResponse;
        return;
    }
    Show-HttpErrorResponse;
    Write-Error "API request fails with error. Stop!" -ErrorAction Stop;
}
function Get-OauthToken {
    $utcNow = (get-date).ToUniversalTime().AddMinutes(1);
    if ($null -ne $script:AuthResult -and ($utcNow -lt $script:AuthResult.ExpiresOn.UtcDateTime)) {
        return;
    }
    if (-not [string]::IsNullOrWhiteSpace($script:redirectUri)) {
        try {
            $script:AuthResult = Get-MsalToken -ClientId $script:clientId -TenantId $script:tenantID -Silent -LoginHint $script:loginHint -RedirectUri $script:redirectUri -Scopes $script:Scope;
        }
        Catch [Microsoft.Identity.Client.MsalUiRequiredException] {
            $script:AuthResult = Get-MsalToken -ClientId $script:clientId -TenantId $script:tenantID -Interactive -LoginHint $script:loginHint -RedirectUri $script:redirectUri -Scopes $script:Scope; ;
        }
        Catch {
            Show-LastErrorDetails;
            Write-Error -Message "Can not get the access token, stop." -ErrorAction Stop;
        }
    }
    else {
        try {
            if (-not [string]::IsNullOrWhiteSpace($script:clientsecret)) {
                $securedclientSecret = ConvertTo-SecureString $script:clientsecret -AsPlainText -Force
                $script:AuthResult = Get-MsalToken -clientID $script:clientId -ClientSecret $securedclientSecret -tenantID $script:tenantID -Scopes $script:Scope;
            }
            elseif ($null -ne $script:clientcertificate) {
                $script:AuthResult = Get-MsalToken -clientID $script:clientId -ClientCertificate $script:clientcertificate -tenantID $script:tenantID -Scopes $script:Scope;
            }        
        }
        catch {
            Show-LastErrorDetails;
            Write-Error -Message "Can not get the access token, stop." -ErrorAction Stop;
        }
    }
}
function Connect-Office365ManagementAPI {
    param (
        [Parameter(Mandatory = $true)][string]$tenantID,
        [Parameter(Mandatory = $true)][String]$clientId,        
        [Parameter(Mandatory = $true, ParameterSetName = "authorizationcode")][String]$redirectUri,
        [Parameter(Mandatory = $true, ParameterSetName = "authorizationcode")][String]$loginHint,    
        [Parameter(Mandatory = $true, ParameterSetName = "clientcredentialsSecret")][String]$clientsecret,
        [Parameter(Mandatory = $true, ParameterSetName = "clientcredentialsCertificate")][X509Certificate]$clientcertificate
    )
    $script:tenantID = $tenantID;
    $script:clientId = $clientId;
    if (-not [string]::IsNullOrWhiteSpace($clientsecret)) {
        $script:clientsecret = $clientsecret;
    }
    elseif ($null -ne $clientcertificate) {
        $script:clientcertificate = $clientcertificate;
    }
    elseif (-not [string]::IsNullOrWhiteSpace($redirectUri)) {
        $script:loginHint = $loginHint;
        $script:redirectUri = $redirectUri;
    }
    else {
        Write-Error "Not implement." -ErrorAction Stop;
    }
    Get-OauthToken;
    if ($null -eq $script:AuthResult) {
        Write-Error "Can not connect to Office365 Management API. Please check your app registration in AAD." -ErrorAction Stop;
    }
    Write-Host "Successfuly Connect to Office365 Management API" -ForegroundColor Green;
}
function Start-Subscription {
    param (
        [Parameter(Mandatory = $true)][ContentType]$contentType,        
        [Parameter(Mandatory = $false)][string]$webhook
    )
    Stop-Subscription -contentType $contentType;
    $contentTypestring = Get-ContentTypeString $contentType;
    $subscriptionUrl = "https://manage.office.com/api/v1.0/$($script:tenantID)/activity/feed/subscriptions/start?contentType=$contentTypestring";
    if ($PSBoundParameters.ContainsKey("webhook")) {
        $httpResponse = Invoke-O365APIHttpRequest -url $subscriptionUrl -httpVerb Post -requstBody $webhook;        
    }
    else {
        $httpResponse = Invoke-O365APIHttpRequest -url $subscriptionUrl -httpVerb Post;        
    }
    Write-Host $httpResponse;
}
function Stop-Subscription {
    param (
        [Parameter(Mandatory = $true)][ContentType]$contentType
    )    
    $contentTypestring = Get-ContentTypeString $contentType;
    $subscriptionUrl = "https://manage.office.com/api/v1.0/$($script:tenantID)/activity/feed/subscriptions/stop?contentType=$contentTypestring";
    $httpResponse = Invoke-O365APIHttpRequest -url $subscriptionUrl -httpVerb Post;
    Write-Host $httpResponse;
}
function Stop-Subscriptions {
    $contentTypes = @(Get-CurrentSubscriptions | Where-Object { $_.status -eq "enabled" } | ForEach-Object { $psitem.contentType });
    $contentTypes | ForEach-Object {
        $contentType = Get-ContentTypeEnum $PSItem;
        Stop-Subscription -contentType $contentType;
    }
}
function Get-CurrentSubscriptions {
    $listSubscriptionURI = "https://manage.office.com/api/v1.0/$($script:tenantID)/activity/feed/subscriptions/list";
    $httpResponse = Invoke-O365APIHttpRequest -url $listSubscriptionURI -httpVerb Get;    
    $convertObjects = $httpResponse.Content | Out-String | ConvertFrom-Json;
    $subscriptions = New-Object Collections.Generic.List[Subscription];
    $convertObjects | ForEach-Object {
        $subscription = New-Object Subscription;
        $subscription.contentType = $psitem.contentType;
        $subscription.status = $psitem.status;
        $subscription.webhook = $psitem.webhook -as [Webhook];
        $subscriptions.Add($subscription);
    }
    Write-Output $subscriptions;
    return;
}
function Get-AvailableContent {
    param (
        [Parameter(Mandatory = $true)][datetime]$startTime,        
        [Parameter(Mandatory = $true)][datetime]$endTime
    )    
    $Subscriptions = Get-CurrentSubscriptions;
    $contentTypes = @($Subscriptions | Where-Object { $_.status -eq "enabled" } | ForEach-Object { $psitem.contentType });
    $availableContent = New-Object Collections.Generic.List[Blob];
    $contentTypes | ForEach-Object {
        $enabledContentType = $psitem;
        # List available content
        $contentUrl = "https://manage.office.com/api/v1.0/$($script:tenantID)/activity/feed/subscriptions/content?contentType=$enabledContentType&startTime=" + $startTime + "&endTime=" + $endTime;
        While (-not [string]::IsNullOrEmpty($contentUrl)) {
            $httpResponse = Invoke-O365APIHttpRequest -url $contentUrl -httpVerb Get; 
            $convertObjs = ConvertFrom-Json $httpResponse.Content;
            $convertObjs | ForEach-Object {
                $blob = New-Object Blob;
                $blob.contentUri = $psitem.contentUri;
                $blob.contentId = $psitem.contentId;
                $blob.contentType = $psitem.contentType;
                $blob.contentCreated = $psitem.contentCreated;
                $blob.contentExpiration = $psitem.contentExpiration;
                $availableContent.Add($blob);
            }
            $contentUrl = $httpResponse.Headers.'NextPageUri';
        }
    }
    Write-Output $availableContent;
    return;
}
function Receive-Content {
    param (
        [Parameter(Mandatory = $true)][System.Object[]]$blobs
    )
    if (($null -eq $blobs) -or ($blobs.Count -eq 0)) {
        Write-Host "No available content, exit!" -ForegroundColor Red;
        return;
    }
    $blobs | ForEach-Object {    
        try {
            $blob = $PSItem -as [Blob];
            if ($null -ne $blob) {
                $httpResponse = Invoke-O365APIHttpRequest -url $blob.contentUri -httpVerb Get; 
                $contents = $httpResponse;
                if ($null -ne $contents) {
                    $auditRecords = $contents.Content | Out-String | ConvertFrom-Json;
                    $blob.auditRecords = $auditRecords;
                }
                else {
                    Write-Error "Can not receive content for $($blob)";
                }
            }            
        }
        catch {
            Show-LastErrorDetails;
        }        
    }
}
function Receive-Notifications {
    Write-Output "The content notifications will be sent to the webhook. Can not implement this method in PowerShell script."
}
function Get-Notifications {
    param (
        [Parameter(Mandatory = $true)][datetime]$startTime,        
        [Parameter(Mandatory = $true)][datetime]$endTime,
        [Parameter(Mandatory = $true)][ContentType]$contentType
    )
    $contentTypestring = Get-ContentTypeString $contentType;
    $listNotificationsUrl = "https://manage.office.com/api/v1.0/$($script:tenantID)/activity/feed/subscriptions/notifications?contentType=$contentTypestring&startTime=" + $startTime + "&endTime=" + $endTime;
    $httpResponse = Invoke-O365APIHttpRequest -url $listNotificationsUrl -httpVerb Get;
    $convertObjs = ConvertFrom-Json $httpResponse.Content;
    $notifications = New-Object Collections.Generic.List[Notification];
    $convertObjs | ForEach-Object {
        $notificationObj = New-Object Notification;
        $notificationObj.contentType = $psitem.contentType;
        $notificationObj.contentId = $psitem.contentId;
        $notificationObj.contentUri = $psitem.contentUri;
        $notificationObj.notificationStatus = $psitem.notificationStatus;
        $notificationObj.contentCreated = $psitem.contentCreated;
        $notificationObj.notificationSent = $psitem.notificationSent;
        $notificationObj.contentExpiration = $psitem.contentExpiration;
        $notifications.Add($notificationObj);
    }
    Write-Output $notifications;        
}
function Receive-ResourceFriendlyNames {
    $url = "https://manage.office.com/api/v1.0/$($script:tenantID)/activity/feed/resources/dlpSensitiveTypes";
    $httpResponse = Invoke-O365APIHttpRequest -url $url -httpVerb Get;
    $friendlyNames = $httpResponse.Content | Out-String | ConvertFrom-Json;
    Write-Output $friendlyNames;
}
function Disconnect-Office365ManagementAPI {
    $script:tenantID = $null;
    $script:clientId = $null;
    $script:clientsecret = $null;
    $script:redirectUri = $null;
    $script:loginHint = $null;
    $script:clientcertificate = $null;
    $script:AuthResult = $null;
    Write-Host "Successfuly Disconnect to Office365 Management API" -ForegroundColor Green;
}
Export-ModuleMember Disconnect-Office365ManagementAPI, Receive-ResourceFriendlyNames, Get-Notifications, Receive-Notifications, Receive-Content, Get-AvailableContent, Get-CurrentSubscriptions, Stop-Subscriptions, Stop-Subscription, Start-Subscription, Connect-Office365ManagementAPI;