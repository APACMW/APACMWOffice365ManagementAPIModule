<#
 .Synopsis
  PowerShell SDK for Office365ManageAPI

 .Description
  IT Admin can use this PowerShell module to call Office365ManagementAPI. It suppports all operations of Office365ManagementAPI. Also supports Webhook subscriptions and notifications.

 .Example
   # Installl and import this PowerShell Module
   https://www.powershellgallery.com/packages/ZIZHUOffice365ManagementAPI/1.0
   https://github.com/APACMW/APACMWOffice365ManagementAPIModule 
   Install-Module -Name ZIZHUOffice365ManagementAPI
   Import-Module -Name ZIZHUOffice365ManagementAPI

 .Example
   # Connect to ZIZHUOffice365ManagementAPI module via client secret
    $clientID = 'bc4db1db-b705-434a-91ff-145aa94185c8';
    $tenantId = 'cff343b2-f0ff-416a-802b-28595997daa2';
    $clientSecret = '';
    Connect-Office365ManagementAPI -tenantID $tenantId -clientID $clientID -ClientSecret $clientSecret;

    $clientID = 'd9499009-1faf-418f-8033-640c29e4a5d7';
    $tenantId = '4ecb5816-21ea-4b5a-b948-fab6471545e1';
    $clientSecret = '';
    Connect-Office365ManagementAPI -tenantID $tenantId -clientID $clientID -ClientSecret $clientSecret -office365SubscriptionPlanType GallatinPlan;

    # Connect to ZIZHUOffice365ManagementAPI module via user sign-in
    $clientID = '0d09e429-1e3f-4050-9fc6-f8bcd3e8c4c5';
    $tenantId = '4ecb5816-21ea-4b5a-b948-fab6471545e1';
    $redirectUri='https://login.microsoftonline.com/common/oauth2/nativeclient'
    $loginHint = 'RiquelTest@jeffreyhe1.partner.onmschina.cn';
    Connect-Office365ManagementAPI -tenantID $tenantId -clientID $clientID -loginHint $loginHint -redirectUri $redirectUri -office365SubscriptionPlanType GallatinPlan;
   
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
    $startTime = "2024-05-14T00:00:00"; 
    $endTime = "2024-05-15T00:00:00";
    $blobs = Get-AvailableContent -startTime $startTime -endTime $endTime;
    Receive-Content -blobs $blobs;

   # Get current subscriptions/Stop subscriptions
    Get-CurrentSubscriptions;
    Stop-Subscription -contentType AuditSharePoint;
    Stop-Subscriptions;

   # Start thesubscriptions. If don't pass $webHookBody, no webhook for the subscription
    $webhookEndpoint='https://5a22-2404-f801-9000-1a-efea-00-23.ngrok-free.app/api/O365ManagementAPIHttpFunction';
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
    $startTime = "2024-04-05T00:00:00"; 
    $endTime = "2024-04-06T00:00:00";
    Get-Notifications -startTime $startTime -endTime $endTime -contentType AuditExchange;

  # Receive the FriendlyNames for DLP Resource
    Receive-ResourceFriendlyNames;

  # Clean after usgae
    Disconnect-Office365ManagementAPI;
    Get-Module ZIZHUOffice365ManagementAPI | Remove-Module;
#>

# Define the tenant environment types
enum Office365SubscriptionPlanType {
    Enterpriseplan    
    GCCGovernmentPlan
    GCCHighGovernmentPlan
    DoDGovernmentPlan
    GallatinPlan
}
# Define the content types
enum ContentType {
    AuditAzureActiveDirectory    
    AuditExchange
    AuditSharePoint
    AuditGeneral
    DLPAll
}
# Define the Blob type as an Azure storage unit to keep the audit data
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
[string]$script:root = $null;
[string]$script:scope = $null; 
$script:httpErrorResponse = $null;
$script:maxretries = 4;
$script:sleepSeconds = 2;
$script:AzureCloudInstance = $null;

function Show-VerboseMessage {    
    param(
        [Parameter(Mandatory = $true)][string]$message
    )    
    Write-Verbose "[$((Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss"))]: $message";
    return;
}
function Show-InformationalMessage {
    param(
        [Parameter(Mandatory = $true)][string]$message,
        [Parameter(Mandatory = $false)][System.ConsoleColor]$consoleColor = [System.ConsoleColor]::Gray
    )
    $defaultConsoleColor = $host.UI.RawUI.ForegroundColor;
    $host.UI.RawUI.ForegroundColor = $consoleColor;
    Write-Information -InformationAction Continue -MessageData "[$((Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss"))]: $message";
    $host.UI.RawUI.ForegroundColor = $defaultConsoleColor;
    return;
}
function Show-HttpErrorResponse {
    param(
        [Parameter(Mandatory = $true)][object]$httpErrorResponse
    )
    $httpError = $httpErrorResponse | Format-List | Out-String;
    Show-InformationalMessage -message $httpError -consoleColor Red;
}
function Show-LastErrorDetails {
    param(
        [Parameter(Mandatory = $false)]$lastError = $Error[0]
    )
    $lastError | Format-List -Property * -Force;
    $lastError.InvocationInfo | Format-List -Property *;
    $exception = $lastError.Exception;
    for ($depth = 0; $null -ne $exception; $depth++) {
        Show-InformationalMessage -message "$depth" * 80 -consoleColor Green;
        $exception | Format-List -Property * -Force;               
        $exception = $exception.InnerException;                
    }
}
function Show-AppPermissions {
    <#
    .SYNOPSIS
    Show the API permissions in the access token
    
    .DESCRIPTION
    Show the API permissions in the access token
    
    .PARAMETER jwtToken
    The accesstoken string
    
    .EXAMPLE
    Show-AppPermissions $accesstoken
    
    .NOTES
    Just show the API permissions. Not enforce to must have the specific permissions
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)][string]$jwtToken
    )
    $decodedToken = Read-JWTtoken -token $jwtToken;
    if ($null -ne $decodedToken -and $null -ne $decodedToken.scp) {
        $permissions = $decodedToken.scp;
    }
    elseif ($null -ne $decodedToken -and $null -ne $decodedToken.roles) {
        $permissions = $decodedToken.roles;
    }
    else {
        $permissions = $null;
    }
    Show-InformationalMessage -message "API permissions in the AccessToke: $($permissions)" -consoleColor Yellow;
}
function Read-JWTtoken {
    <#
    .SYNOPSIS
    Parse the access token/ID token based on https://datatracker.ietf.org/doc/html/rfc7519
    
    .DESCRIPTION
    Parse the access token/ID token based on https://datatracker.ietf.org/doc/html/rfc7519
    
    .PARAMETER token
    The accesstoken/ID token string
    
    .EXAMPLE
    Read-JWTtoken -token $jwtToken
    
    .NOTES
    https://datatracker.ietf.org/doc/html/rfc7519
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)][string]$token
    )
    # Validate Access and ID tokens per RFC 7519
    if (!$token.Contains(".") -or !$token.StartsWith("eyJ")) {
        Show-InformationalMessage -message "Invalid token" -consoleColor Red;
        return;
    }
    # Parse the Header
    $tokenheader = $token.Split(".")[0].Replace('-', '+').Replace('_', '/');
    # Fix padding as needed; keep adding "=" until string length modulus 4 reaches 0
    while ($tokenheader.Length % 4) {
        Show-VerboseMessage -message "Invalid length for a Base-64 char array or string, adding =";
        $tokenheader += "=";
    }
    Show-VerboseMessage -message "Base64 encoded (padded) header:"
    Show-VerboseMessage -message $tokenheader;

    # Convert from Base64 encoded string to PSObject
    Show-VerboseMessage -message "Decoded header:"
    $headers = [System.Text.Encoding]::ASCII.GetString([system.convert]::FromBase64String($tokenheader)) | ConvertFrom-Json | Format-List | Out-String;
    Show-VerboseMessage -message $headers;

    # Payload
    $tokenPayload = $token.Split(".")[1].Replace('-', '+').Replace('_', '/');
    # Fix padding as needed; keep adding "=" until string length modulus 4 reaches 0
    while ($tokenPayload.Length % 4) {
        Show-VerboseMessage -message "Invalid length for a Base-64 char array or string, adding =";
        $tokenPayload += "=";
    }
    Show-VerboseMessage -message "Base64 encoded (padded) payload:";
    Show-VerboseMessage -message $tokenPayload;

    # Convert to Byte array
    $tokenByteArray = [System.Convert]::FromBase64String($tokenPayload);
    # Convert to string array
    $tokenArray = [System.Text.Encoding]::ASCII.GetString($tokenByteArray);
    Show-VerboseMessage -message "Decoded array in JSON format:"
    Show-VerboseMessage -message $tokenArray

    # Convert from JSON to PSObject
    $tokenObj = $tokenArray | ConvertFrom-Json;
    Show-VerboseMessage -message "Decoded Payload:"
    Write-Output $tokenObj;
    return;
}

function Set-RootString {
    <#
    .SYNOPSIS
    Based on the tenant type to specify Office365 management API endpoint
    
    .DESCRIPTION
    Based on the tenant type to specify Office365 management API endpoint
    
    .PARAMETER office365SubscriptionPlanType
    The tenant type (data type enum Office365SubscriptionPlanType)
    #>
    param(
        [Parameter(Mandatory = $true)][Office365SubscriptionPlanType]$office365SubscriptionPlanType
    )
    switch ($office365SubscriptionPlanType) {
        Enterpriseplan {
            $script:root = 'https://manage.office.com';
            $script:AzureCloudInstance = 'AzurePublic';
            Break; 
        }
        GCCHighGovernmentPlan {
            $script:root = 'https://manage.office365.us';
            $script:AzureCloudInstance = 'AzureUsGovernment';
            Break;
        }
        GallatinPlan { 
            $script:root = 'https://manage.office365.cn';
            $script:AzureCloudInstance = 'AzureChina';           
            Break;
        }
        Default {
            Write-Error "unknown/unsupported type: $office365SubscriptionPlanType" -ErrorAction Stop;
        }        
    }
    $script:scope = "$script:root/.default";
    Show-VerboseMessage -message "Root of Office365 Management API endpoint: $($script:root) and scope: $($script:scope)";    
    return;
}
function Get-ContentTypeString {
    <#
    .SYNOPSIS
    From the contentType(enum), to generate the relevant content type script used for Http request
    
    .DESCRIPTION
    From the contentType(enum), to generate the relevant content type script used for Http request
    
    .PARAMETER contentTypeData
    Specify the contentTypeData (enum ContentType)    
    #>
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
    <#
    .SYNOPSIS
    Based on contenttype stirng to get the contentType as enum data type
    
    .DESCRIPTION
    Based on contenttype stirng to get the contentType as enum data type
    
    .PARAMETER contentTypeString
    The content type string    
    #>
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
    <#
    .SYNOPSIS
    Sumbit the Http requests for all Ofice365 Management API operations with retry logic
    
    .DESCRIPTION
    Sumbit the Http requests for all Ofice365 Management API operations with retry logic
    
    .PARAMETER url
    Office365 Management API Request Url
    
    .PARAMETER httpVerb
    Http verb
    
    .PARAMETER requstBody
    The Http request body. Optional  
    #>
    param (
        [Parameter(Mandatory = $true)][string]$url,
        [Parameter(Mandatory = $true)][string]$httpVerb,
        [Parameter(Mandatory = $false)][string]$requstBody
    )
    if ($null -eq $script:AuthResult) {
        Write-Error "Not authenticated. Stop." -ErrorAction Stop;
    }    
    Show-VerboseMessage "Invoke-Webrequest $url $httpVerb";
    if ($PSBoundParameters.ContainsKey('requstBody')) {
        Show-VerboseMessage "request body: $requstBody";        
    }
    # Used for retry logic
    $httpResponse = $null;
    $script:httpErrorResponse = $null;
    $retryCount = 0;

    # If fail, retry 4 times (but stop when response (Unauthorized,BadRequest))
    do {
        if ($retryCount -gt 1) {
            $sleepSeconds = [math]::Pow($script:sleepSeconds, $retryCount);
            Show-VerboseMessage "Retry Invoke-O365APIHttpRequest $($httpVerb) $($url): $retryCount after $($sleepSeconds) seconds";
            Start-Sleep -Seconds $sleepSeconds;
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
            Show-InformationalMessage -message "Http error: $($_.Exception.Response) Body: $($_.ErrorDetails.Message)" -consoleColor Red;
            $script:httpErrorResponse = $_.Exception.Response;
        }
        finally {
            $retryCount = $retryCount + 1;
        }
    } until ((($null -ne $httpResponse) -or ($retryCount -gt $script:maxretries)) -or (($null -ne $script:httpErrorResponse) -and ($script:httpErrorResponse.StatusCode -in @('Unauthorized', 'BadRequest'))))

    # If succeed, then pass the Http response in output stream to caller
    if (($null -ne $httpResponse) -and ($httpResponse.StatusCode -in (200, 204))) {
        Show-VerboseMessage "HTTP Response: $($httpResponse.RawContent)";        
        Write-Output $httpResponse;
        return;
    }

    # If fail, show the error information and stop
    Show-HttpErrorResponse -httpErrorResponse $script:httpErrorResponse;
    Write-Error "API request fails with error. Stop!" -ErrorAction Stop;
}
function Get-OauthToken {
    <#
    .SYNOPSIS
    Use the Msal.ps module to get the access token. Support client credential, Implicit auth flow
    
    .DESCRIPTION
    Use the Msal.ps module to get the access token. Support client credential, Implicit auth flow
    
    .NOTES
    Use the variables from script scope
    #>
    Show-VerboseMessage "Start to invoke Get-OauthToken";
    # If the access token is valid, then use an existing token
    $utcNow = (get-date).ToUniversalTime().AddMinutes(1);
    if ($null -ne $script:AuthResult -and ($utcNow -lt $script:AuthResult.ExpiresOn.UtcDateTime)) {
        Show-VerboseMessage "Current accesstoken is valid before $($script:AuthResult.ExpiresOn.UtcDateTime)";
        return;
    }
    # Implicit auth flow (delegated API permissions). Will try to get the access token silently. If fail, then interactive sign-in
    if (-not [string]::IsNullOrWhiteSpace($script:redirectUri)) {
        try {
            Show-VerboseMessage "Get-MsalToken via user sign-in";
            $script:AuthResult = Get-MsalToken -ClientId $script:clientId -TenantId $script:tenantID -Silent -LoginHint $script:loginHint -RedirectUri $script:redirectUri -Scopes $script:scope -AzureCloudInstance $script:AzureCloudInstance;
        }
        Catch [Microsoft.Identity.Client.MsalUiRequiredException] {
            $script:AuthResult = Get-MsalToken -ClientId $script:clientId -TenantId $script:tenantID -Interactive -LoginHint $script:loginHint -RedirectUri $script:redirectUri -Scopes $script:scope  -AzureCloudInstance $script:AzureCloudInstance;
        }
        Catch {
            Show-LastErrorDetails;
            Write-Error -Message "Can not get the access token, exit." -ErrorAction Stop;
        }
    }
    # Client credential auth flow. Can use the client secret or certificate
    else {
        try {
            if (-not [string]::IsNullOrWhiteSpace($script:clientsecret)) {
                Show-VerboseMessage "Get-MsalToken via client crendential auth flow";
                $securedclientSecret = ConvertTo-SecureString $script:clientsecret -AsPlainText -Force
                $script:AuthResult = Get-MsalToken -clientID $script:clientId -ClientSecret $securedclientSecret -tenantID $script:tenantID -Scopes $script:scope -AzureCloudInstance $script:AzureCloudInstance;
            }
            elseif ($null -ne $script:clientcertificate) {
                $script:AuthResult = Get-MsalToken -clientID $script:clientId -ClientCertificate $script:clientcertificate -tenantID $script:tenantID -Scopes $script:scope -AzureCloudInstance $script:AzureCloudInstance;
            }        
        }
        catch {
            Show-LastErrorDetails;
            Write-Error -Message "Can not get the access token, stop." -ErrorAction Stop;
        }
    }
    Show-VerboseMessage "Succeed to invoke Get-OauthToken";
}
function Connect-Office365ManagementAPI {
    <#
    .SYNOPSIS
    Initilize the script varibles to prepare for calling APIs
    
    .DESCRIPTION
    Initilize the script varibles to prepare for calling APIs
    
    .PARAMETER tenantID
    tenant id
    
    .PARAMETER clientId
    Azure AD application Id
    
    .PARAMETER redirectUri
    The redirectUri used for implicit auth flow
    
    .PARAMETER loginHint
    The loginHint (user's UPN) used for implicit auth flow
    
    .PARAMETER clientsecret
    The clientsecret used for client credential auth flow
    
    .PARAMETER clientcertificate
    The clientcertificate used for client credential auth flow
    
    .PARAMETER office365SubscriptionPlanType
    Tenant type
    
    .EXAMPLE
    Connect-Office365ManagementAPI -tenantID $tenantId -clientID $clientID -ClientSecret $clientSecret;
    
    .NOTES
    Read how to register the app in Azure AD: https://learn.microsoft.com/en-us/office/office-365-management-api/get-started-with-office-365-management-apis
    #>
    param (
        [Parameter(Mandatory = $true)][string]$tenantID,
        [Parameter(Mandatory = $true)][String]$clientId,        
        [Parameter(Mandatory = $true, ParameterSetName = "authorizationcode")][String]$redirectUri,
        [Parameter(Mandatory = $true, ParameterSetName = "authorizationcode")][String]$loginHint,    
        [Parameter(Mandatory = $true, ParameterSetName = "clientcredentialsSecret")][String]$clientsecret,
        [Parameter(Mandatory = $true, ParameterSetName = "clientcredentialsCertificate")][X509Certificate]$clientcertificate,
        [Parameter(Mandatory = $false)][Office365SubscriptionPlanType]$office365SubscriptionPlanType = [Office365SubscriptionPlanType]::Enterpriseplan
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
    Set-RootString $office365SubscriptionPlanType;    
    Get-OauthToken;
    if ($null -eq $script:AuthResult) {
        Write-Error "Can not connect to Office365 Management API. Please check your app registration in AAD." -ErrorAction Stop;
    }    
    Show-AppPermissions $script:AuthResult.accesstoken;
    Show-InformationalMessage -message "Successfuly Connect to Office365 Management API" -consoleColor Green;
}
function Start-Subscription {
    <#
    .SYNOPSIS
    Start Subscription for a content type
    
    .DESCRIPTION
    Start Subscription for a content type
    
    .PARAMETER contentType
    The mandatory content type
    
    .PARAMETER webhook
    The optional webhook 
    
    .EXAMPLE
    Start-Subscription DLPAll $webHookBody;
    
    .NOTES
    reference: https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#start-a-subscription
    #>
    param (
        [Parameter(Mandatory = $true)][ContentType]$contentType,        
        [Parameter(Mandatory = $false)][string]$webhook
    )
    $contentTypestring = Get-ContentTypeString $contentType;
    $Subscriptions = Get-CurrentSubscriptions;
    $subscribedContentType = @($Subscriptions | Where-Object { $PSItem.status -eq "enabled" -and $PSItem.contentType -eq $contentTypeString } );
    if ($subscribedContentType.Count -eq 0) {
        $subscriptionUrl = "$($script:root)/api/v1.0/$($script:tenantID)/activity/feed/subscriptions/start?contentType=$contentTypestring";
        if ($PSBoundParameters.ContainsKey("webhook")) {
            $httpResponse = Invoke-O365APIHttpRequest -url $subscriptionUrl -httpVerb Post -requstBody $webhook;        
        }
        else {
            $httpResponse = Invoke-O365APIHttpRequest -url $subscriptionUrl -httpVerb Post;        
        }
    }
    else {
        Show-InformationalMessage -message "The subscription of $contentType has been started already, and please stop it before start with new parameters" -consoleColor Yellow;
    }
    $httpResponse | Format-List;
}
function Stop-Subscription {
    <#
    .SYNOPSIS
    Stop subscription for a content type
    
    .DESCRIPTION
    Stop subscription for a content type
    
    .PARAMETER contentType
    The mandatory contentType
    
    .EXAMPLE
    Stop-Subscription -contentType AuditSharePoint
    
    .NOTES
    reference: https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#stop-a-subscription
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory = $true)][ContentType]$contentType
    )
    if ($PSCmdlet.ShouldProcess($contentType)) {
        $contentTypestring = Get-ContentTypeString $contentType;
        $subscriptionUrl = "$($script:root)/api/v1.0/$($script:tenantID)/activity/feed/subscriptions/stop?contentType=$contentTypestring";
        $httpResponse = Invoke-O365APIHttpRequest -url $subscriptionUrl -httpVerb Post;
        $httpResponse | Format-List;
    }
    else {
        Show-InformationalMessage -message "The user decide to not stop subscription $contentType" -consoleColor Yellow;
    }
}
function Stop-Subscriptions {
    <#
    .SYNOPSIS
    Stop all subscriptions for this application
    
    .DESCRIPTION
    Stop all subscriptions for this application
    
    .EXAMPLE
    Stop-Subscriptions
    
    .NOTES
    General notes
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param()
    if ($PSCmdlet.ShouldProcess('Current subscriptions')) {
        $contentTypes = @(Get-CurrentSubscriptions | Where-Object { $_.status -eq "enabled" } | ForEach-Object { $psitem.contentType });
        $contentTypes | ForEach-Object {
            $contentType = Get-ContentTypeEnum $PSItem;
            Stop-Subscription -contentType $contentType -Confirm:$false;
        }
    }
}
function Get-CurrentSubscriptions {
    <#
    .SYNOPSIS
    Get current subscriptions for this application
    
    .DESCRIPTION
    This operation returns a collection of the current subscriptions together with the associated webhooks
    
    .EXAMPLE
    Get-CurrentSubscriptions
    
    .NOTES
    reference: https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#list-current-subscriptions
    #>
    [CmdletBinding()]
    param()
    $listSubscriptionURI = "$($script:root)/api/v1.0/$($script:tenantID)/activity/feed/subscriptions/list";
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
    <#
    .SYNOPSIS
    This operation lists the content currently available for retrieval for the specified content type
    
    .DESCRIPTION
    This operation lists the content currently available for retrieval for the specified content type
    
    .PARAMETER startTime
    Datetimes (UTC) indicating the start time of range for the audit content to return
    
    .PARAMETER endTime
    Datetimes (UTC) indicating the end time of range for the audit content to return
    
    .PARAMETER contentType
    The mandatory content type
    
    .EXAMPLE
    $blobs = Get-AvailableContent -startTime $startTime -endTime $endTime
    
    .NOTES
    reference: https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#list-available-content
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][datetime]$startTime,        
        [Parameter(Mandatory = $true)][datetime]$endTime,
        [Parameter(Mandatory = $false)][ContentType]$contentType
    )
    $Subscriptions = Get-CurrentSubscriptions;
    $contentTypes = @($Subscriptions | Where-Object { $_.status -eq "enabled" } | ForEach-Object { $psitem.contentType });
    if ($PSBoundParameters.ContainsKey('contentType')) {
        $contentTypeString = Get-ContentTypeString $contentType;
        $contentTypes = @($contentTypes | Where-Object { $PSItem -eq $contentTypeString });
    }
    if ($contentTypes.Count -eq 0) {
        Write-Warning "The subscription of the specified ContentType or any ContentTypes has not been started yet.";
        return;
    }   
    $availableContent = New-Object Collections.Generic.List[Blob];
    Show-VerboseMessage "Run Get-AvailableContent for the contenttypes $contentTypes";
    $contentTypes | ForEach-Object {
        $enabledContentType = $psitem;
        # List available content        
        $contentUrl = "$($script:root)/api/v1.0/$($script:tenantID)/activity/feed/subscriptions/content?contentType=$enabledContentType&startTime=" + $startTime + "&endTime=" + $endTime;        
        While (-not [string]::IsNullOrEmpty($contentUrl)) {
            Show-VerboseMessage "List available content via the Url $contentUrl";
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
            # Support the paging (https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#pagination)
            if ($null -ne $httpResponse.Headers.'NextPageUri') {
                $nextPageUri = $httpResponse.Headers.'NextPageUri';
                Show-VerboseMessage "NextPageUri is $nextPageUri";
            }
            $contentUrl = $httpResponse.Headers.'NextPageUri';
        }
    }
    Write-Output $availableContent;
    return;
}
function Receive-Content {
    <#
    .SYNOPSIS
    To retrieve a content blob, make a GET request against the corresponding content URI that is included in the list of available content
    
    .DESCRIPTION
    To retrieve a content blob, make a GET request against the corresponding content URI that is included in the list of available content and in the notifications sent to a webhook. The returned content will be a collection of one more actions or events in JSON format
    
    .PARAMETER blobs
    The collection blobs parmater
    
    .EXAMPLE
    Receive-Content -blobs $blobs
    
    .NOTES
    reference:https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#retrieve-content
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][System.Object[]]$blobs
    )
    if (($null -eq $blobs) -or ($blobs.Count -eq 0)) {
        Show-InformationalMessage "No available content, exit!" -consoleColor Yellow;
        return;
    }
    $blobs | ForEach-Object {    
        try {
            $blob = $PSItem -as [Blob];
            if ($null -ne $blob) {
                Show-VerboseMessage "Receive content from the content Url $($blob.contentUri)";
                $httpResponse = Invoke-O365APIHttpRequest -url $blob.contentUri -httpVerb Get; 
                $contents = $httpResponse;
                if ($null -ne $contents) {
                    $auditRecords = $contents.Content | Out-String | ConvertFrom-Json;
                    $blob.auditRecords = $auditRecords;
                }
                else {
                    Write-Error "Can not receive content for $($blob)" -ErrorAction Continue;
                }
            }            
        }
        catch {
            Show-LastErrorDetails;
        }        
    }
}
function Receive-Notifications {
    <#
    .SYNOPSIS
    Not implement. Notifications are sent to the configured webhook for a subscription as new content becomes available
    
    .DESCRIPTION
    Not implement. Notifications are sent to the configured webhook for a subscription as new content becomes available
    
    .NOTES
    reference:https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#receiving-notifications
    #>
    [CmdletBinding()]
    param()
    Show-InformationalMessage -message "The content notifications are sent to the webhook. We can not implement this operation in PowerShell Module" -consoleColor Yellow;
}
function Get-Notifications {
    <#
    .SYNOPSIS
    This operation lists all notification attempts for the specified content type
    
    .DESCRIPTION
    This operation lists all notification attempts for the specified content type
  
    .PARAMETER startTime
    Datetimes (UTC) indicating the start time of range for the notifications to return
    
    .PARAMETER endTime
    Datetimes (UTC) indicating the end time of range for the notifications to return
    
    .PARAMETER contentType
    The mandatory content type

    .EXAMPLE
    $startTime = "2024-04-05T00:00:00"
    $endTime = "2024-04-06T00:00:00"
    Get-Notifications -startTime $startTime -endTime $endTime -contentType AuditExchange
    
    .NOTES
    reference:https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#list-notifications
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][datetime]$startTime,        
        [Parameter(Mandatory = $true)][datetime]$endTime,
        [Parameter(Mandatory = $true)][ContentType]$contentType
    )
    $contentTypestring = Get-ContentTypeString $contentType;
    $listNotificationsUrl = "$($script:root)/api/v1.0/$($script:tenantID)/activity/feed/subscriptions/notifications?contentType=$contentTypestring&startTime=" + $startTime + "&endTime=" + $endTime;
    Show-VerboseMessage -message "List notifications via the Url $listNotificationsUrl";
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
    return;      
}
function Receive-ResourceFriendlyNames {
    <#
    .SYNOPSIS
    This operation retrieves friendly names for objects in the data feed identified by guids. Currently "DlpSensitiveType" is the only supported object
    
    .DESCRIPTION
    This operation retrieves friendly names for objects in the data feed identified by guids. Currently "DlpSensitiveType" is the only supported object
    
    .EXAMPLE
    Receive-ResourceFriendlyNames
    
    .NOTES
    reference:https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#retrieve-resource-friendly-names
    #>
    [CmdletBinding()]
    param()
    $url = "$($script:root)/api/v1.0/$($script:tenantID)/activity/feed/resources/dlpSensitiveTypes";
    Show-VerboseMessage "Receive resource FriendlyNames via the Url $url";
    $httpResponse = Invoke-O365APIHttpRequest -url $url -httpVerb Get;
    $friendlyNames = $httpResponse.Content | Out-String | ConvertFrom-Json;
    Write-Output $friendlyNames;
    return;
}
function Disconnect-Office365ManagementAPI {
    $script:tenantID = $null;
    $script:clientId = $null;
    $script:clientsecret = $null;
    $script:redirectUri = $null;
    $script:loginHint = $null;
    $script:clientcertificate = $null;
    $script:AuthResult = $null;
    $script:root = $null;
    $script:scope = $null; 
    $script:httpErrorResponse = $null;
    Show-InformationalMessage -message "Successfuly Disconnect to Office365 Management API" -consoleColor Green;
}
Export-ModuleMember Disconnect-Office365ManagementAPI, Receive-ResourceFriendlyNames, Get-Notifications, Receive-Notifications, Receive-Content, Get-AvailableContent, Get-CurrentSubscriptions, Stop-Subscriptions, Stop-Subscription, Start-Subscription, Connect-Office365ManagementAPI;