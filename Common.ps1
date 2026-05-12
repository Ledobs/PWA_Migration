#Version: 16.0.7715.1200
$ErrorActionPreference = "Stop"  # http://technet.microsoft.com/en-us/library/dd347731.aspx
Set-StrictMode -Version "Latest" # http://technet.microsoft.com/en-us/library/dd347614.aspx

[Reflection.Assembly]::LoadFrom("$($PSScriptRoot)\Microsoft.Identity.Client.dll") | Out-Null

Add-Type -Path "$($PSScriptRoot)\UserData.cs" -ReferencedAssemblies System.Data,System.ServiceModel,System.Xml,System.Runtime.Serialization
Add-Type -Path "$($PSScriptRoot)\UserDataRequestBehavior.cs" -ReferencedAssemblies System.Data,System.ServiceModel
Add-Type -Path "$($PSScriptRoot)\Utilities.cs" -ReferencedAssemblies System.Data,System.ServiceModel,System.Xml

function Get-AADAuthResult ([Uri] $Uri, [string] $Region) 
{
    $resource = $Uri.GetLeftPart([System.UriPartial]::Authority);

    $authUrl = "https://login.microsoftonline.com/common"
    switch ($Region)
    {
        "ITAR"
        {
            $authUrl = "https://login.microsoftonline.us/common"
        }

        "Germany"
        {
            $authUrl = "https://login.microsoftonline.de/common"
        }

        "China"
        {
            $authUrl = "https://login.chinacloudapi.cn/common"
        }
    }

    if ($Uri.Host -like "*.spoppe.com" -or $Uri.Host -like "*.spgrid.com") # for eDog or Sandbox
    {
        $authUrl = "https://login.windows-ppe.net/common"
    }

    $clientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"

    $pcaConfig = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($clientId).WithRedirectUri($redirectUri).WithAuthority($authUrl);
    $scopes = New-Object System.Collections.Generic.List[string]
    $scope = $resource.ToString() + "/.default"
    $scopes.Add($scope)
    $authenticationResult = $pcaConfig.Build().AcquireTokenInteractive($scopes).WithPrompt([Microsoft.Identity.Client.Prompt]::ForceLogin).ExecuteAsync().Result;

    return $authenticationResult
}

function GetAuthCookie([Uri] $Uri, [string] $Region)
{
    $authResult = Get-AADAuthResult -Uri $Uri -Region $Region

    if ($authResult -ne $null)
    {
        $accessHeader = $authResult.CreateAuthorizationHeader()

        $OAuthUri = New-Object Uri -ArgumentList @($Uri, "_api/SP.OAuth.NativeClient/Authenticate")
        $Body = [System.String]::Empty
        $headers = @{ "X-IDCRL_ACCEPTED" = "t"; "Authorization" = $accessHeader; }

        $Response = Invoke-WebRequest -Headers $headers -Method "POST" -Body $Body -Uri $OAuthUri
        return $Response.Headers["Set-Cookie"]
    }
}

function GetXRequestDigest([Uri] $Uri, [Boolean] $OnPrem, [string] $AuthCookie)
{
    $ContextInfoUri = New-Object Uri -ArgumentList @($Uri, "_api/contextinfo")
    $Method = "POST"

    $Props = @{
        Uri   = $ContextInfoUri
        Method = $Method
    }

    if ($OnPrem)
    {
        $Props.UseDefaultCredentials = $true
    }
    else
    {
        $WebSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $WebSession.Cookies.SetCookies($Uri, $AuthCookie)
        $Props.WebSession = $WebSession
    }

    $ret = Invoke-RestMethod @Props
    return $ret.GetContextWebInformation.FormDigestValue
}

function GetPSIProxy([string] $Url, [string] $Region, [Boolean] $OnPrem)
{
    if (-not $Url.EndsWith("/"))
    {
        $Url = $Url + "/"
    }

    $Uri = New-Object Uri $Url
    $PSIUri = New-Object Uri -ArgumentList @($Uri, "_vti_bin/PSI/ProjectServer.svc")

    $authCookie = $null
    if (-not $OnPrem)
    {
        $authCookie = GetAuthCookie -Uri $Uri -Region $Region
    }

    $requestDigest = GetXRequestDigest -Uri $Uri -OnPrem $OnPrem -AuthCookie $authCookie

    $requestBehavior = New-Object UserDataRequestBehavior
    $requestBehavior.AuthCookie = $authCookie
    $requestBehavior.RequestDigest = $requestDigest
    $requestBehavior.RequestUri = $Uri


    if ($Uri.Scheme -eq "http")
    {
        $binding = New-Object System.ServiceModel.BasicHttpBinding 
    }
    elseif($Uri.Scheme -eq "https")
    {
        $binding = New-Object System.ServiceModel.BasicHttpsBinding
    }
    else
    {
        Write-Error "Unknown scheme from the URL $($Uri)"
    }

    $binding.MaxBufferSize = 4194304
    $binding.MaxReceivedMessageSize = 1073741824
    $binding.TransferMode = "StreamedResponse"
    $binding.ReaderQuotas.MaxStringContentLength = 1073741824
    $tenMinsTimeSpan = [System.TimeSpan]::FromMinutes(10)
    $binding.CloseTimeout = $tenMinsTimeSpan
    $binding.OpenTimeout = $tenMinsTimeSpan
    $binding.ReceiveTimeout = $tenMinsTimeSpan
    $binding.SendTimeout = $tenMinsTimeSpan

    if ($OnPrem)
    {
        $binding.Security.Mode = [System.ServiceModel.BasicHttpSecurityMode]::TransportCredentialOnly
        $binding.Security.Transport.ClientCredentialType = [System.ServiceModel.HttpClientCredentialType]::Ntlm
    }

    $endpoint = New-Object System.ServiceModel.EndpointAddress($PSIUri)
    $proxy = New-Object UserDataSoapClient($binding, $endpoint)
    $proxy.ChannelFactory.Endpoint.Behaviors.Add($requestBehavior)

    if ($OnPrem)
    {
        $proxy.ClientCredentials.Windows.AllowedImpersonationLevel = [System.Security.Principal.TokenImpersonationLevel]::Impersonation
        $proxy.ChannelFactory.Credentials.Windows.ClientCredential = [System.Net.CredentialCache]::DefaultNetworkCredentials
    }

    return $proxy
}

function GetEncodedClaim([string] $LoginName, [bool] $OnPrem)
{
    # TODO: We should get rid of $OnPrem switch before publishing the script to public - this is for making test easier from dev box.
    if ($OnPrem)
    {
        return "i:0#.w|$($LoginName)"
    }
    else
    {
        return "i:0#.f|membership|$($LoginName)"
    }
}

function IsGeneralSecurityAccessDenied($Exception)
{
    # two cases
    # 1. do not have access to the site at all (sharepoint level) : getting proxy will fail because getting auth cookie will fail.
    # 2. do not have enouth permission to call PSI/OData - GeneralSecurityAccessDenied : 20010

    [bool] $isGeneralSecurityAccessDenied = $false

    try
    {
        $mf = $Exception.InnerException.CreateMessageFault()
        $detail = [MessageFaultHelper]::GetMessageDetail($mf)
        $errInfos = $detail.SelectNodes("//errinfo")
        if ($errInfos[0].general.class.error.name -eq "GeneralSecurityAccessDenied")
        {
            $isGeneralSecurityAccessDenied = $true
        }
    }
    catch
    {
    }

    return $isGeneralSecurityAccessDenied
}

function IsThrottled($Exception)
{
    [System.Net.WebException] $we = $null
    if ($Exception -is [System.Net.WebException])
    {
        $we = $Exception -as [System.Net.WebException]
    }
    elseif ($Exception.InnerException -ne $null -and $Exception.InnerException -is [System.Net.WebException])
    {
        $we = $Exception.InnerException -as [System.Net.WebException]
    }

    # Invoke-RestMethod throws WebException with 503 or 429 StatusCode when throttled
    if ($we -ne $null)
    {
        if ($we.Response.StatusCode -eq [System.Net.HttpStatusCode]::ServiceUnavailable -or $we.Response.StatusCode -eq 429)
        {
            return $true
        }
    }

    # PSI proxy throws ServerTooBusyException when throttled.
    if ($Exception -ne $null -and
        $Exception.InnerException -ne $null -and
        $Exception.InnerException -is [System.ServiceModel.ServerTooBusyException])
    {
        return $true
    }
    
    return $false
}

function ShowExceptionDetail ($Exception)
{
    try
    {
        $mf = $Exception.InnerException.CreateMessageFault()
        $detail = [MessageFaultHelper]::GetMessageDetail($mf)
        $errInfos = $detail.SelectNodes("//errinfo")
        foreach ($errInfo in $errInfos)
        {
            Write-Host -ForegroundColor Red "$($errInfo.general.class.name) - $($errInfo.general.class.error.OuterXml)"
        }
    }
    catch
    {
        Write-Host -ForegroundColor Red $Exception.Message
    }
}

# retry when throttled
function RetryOnThrottled ($ScriptBlock)
{
    $sleepSeconds = 300
    $retryMax = 5
    $retryCount = 0

    do
    {
        $retryCount++

        try
        {
            Write-Host -ForegroundColor Red "Request throttled, will retry again in $($sleepSeconds) seconds."
            Start-Sleep -Seconds $sleepSeconds
            Invoke-Command -ScriptBlock $ScriptBlock
            break
        }
        catch
        {
            if (IsThrottled($_.Exception) -and $retryCount -le $retryMax)
            {
                continue
            }

            # stop retrying on non-throttled exception
            ShowExceptionDetail($_.Exception)
            break
        }
    } while ($retryCount -le $retryMax)
}

# try-catch exception and based on the exception, do stop, retry or continue.
function ContinueOnError ($ScriptBlock)
{
    try
    {
        Invoke-Command -ScriptBlock $ScriptBlock
    }
    catch
    {
        # stop on access denied
        if ((IsGeneralSecurityAccessDenied -Exception $_.Exception))
        {
            # clear stored credentials - script will ask credential again.
            ShowExceptionDetail($_.Exception)
            throw "You don't have enough permission on target site. Please try again with different account or after granting permission."
        }
        # retry when throttled
        elseif (IsThrottled($_.Exception))
        {
            RetryOnThrottled($ScriptBlock)
        }
        # continue on other unknown errors
        else
        {
            ShowExceptionDetail($_.Exception)
        }
    }
}

# Formats JSON in a nicer format than the built-in ConvertTo-Json does.
function Get-Json($data) {
    return $data | ConvertFrom-Json | ConvertTo-Json -Depth 99
}

function CallOData([UserDataSoapClient] $Proxy, [string] $BaseUrl, [string] $Request, [bool] $OnPrem)
{    
    if (-not $BaseUrl.EndsWith("/"))
    {
        $BaseUrl = $BaseUrl + "/"
    }

    # get baseuri (for authentication), odata base uri and actual request uri
    $baseUri = New-Object Uri $BaseUrl
    $oDataBaseUri = New-Object Uri -ArgumentList @($baseUri, "_api/projectdata/")
    $oDataRequestUri = New-Object Uri -ArgumentList @($oDataBaseUri, $Request)

    $headers = @{
        Accept = "application/json"
    }

    $parameters = @{
        Uri = $oDataRequestUri.AbsoluteUri
        Method = "GET"
        Headers = $headers
        UseBasicParsing = $true # for backcompat only
    }

    if (-not $OnPrem)
    {
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

        # get auth cookie from the behavior attached to the proxy
        foreach ($behavior in $Proxy.ChannelFactory.Endpoint.Behaviors)
        {
            $userDataBehavior = $behavior -as [UserDataRequestBehavior]
            if ($userDataBehavior -ne $null)
            {
                $session.Cookies.SetCookies($baseUri, $behavior.AuthCookie)
                break
            }
        }

        $parameters.WebSession = $session
    }
    else
    {
        $parameters.UseDefaultCredentials = $true
    }

    return Invoke-RestMethod @parameters
}

# Save all odata pages
# ex> SaveODataAllPages -BaseUrl "http://projectserver:81/native" -Request "TaskTimephasedDataSet?`$filter=ProjectId eq guid'd7a85317-e033-e811-8518-bc8385e6294b'" -OutputDirectory C:\ -OutputFileName "timephased.json" -OnPrem $true
function SaveODataAllPages([UserDataSoapClient] $Proxy, [string] $BaseUrl, [string] $Request, [string] $OutputDirectory, [string] $OutputFileName, [bool] $OnPrem)
{
    if (-not $OutputDirectory.EndsWith("\"))
    {
        $OutputDirectory = $OutputDirectory + "\" 
    }

    $fullFileName = "$($OutputDirectory)$($OutputFileName)"

    "{ ""all"" : [ " > "$fullfileName"

    [int] $pageCount = 1
    $ret = CallOData -Proxy $Proxy -BaseUrl $BaseUrl -Request $Request -OnPrem $OnPrem
    do
    {
        if ($pageCount -gt 1)
        {
            ", " >> "$fullFileName"
        }
        $ret | ConvertTo-Json >> "$fullFileName"
        if (Get-Member -InputObject $ret -Name 'odata.nextLink')
        {
            # Write-Host "Getting next page $($ret.'odata.nextLink')"
            $pageCount++
            $ret = CallOData -Proxy $Proxy -BaseUrl $BaseUrl -Request $ret.'odata.nextLink' -OnPrem $OnPrem
        }
        else
        {
            break
        }
    }
    while ($ret -ne $null)

    "] }" >> "$fullFileName"
}
