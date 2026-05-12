#requires -version 5.1
<#
.SYNOPSIS
Targets existing Project Online enterprise resources by UID to inactivate them,
append/apply an archive name suffix, and optionally attempt an unsupported user
account link dissociation.

.IMPORTANT
The direct dissociation of an Enterprise Resource "User Logon Account" through
CSOM is not confirmed by an official Microsoft source. This script blocks that
operation unless -EnableExperimentalUserDissociation is explicitly provided.

The script never deletes or recreates resources.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Custom', 'Idexia', 'SQI')]
    [string] $TenantPreset = 'Custom',

    [Parameter(Mandatory = $false)]
    [string] $PwaUrl,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Custom', 'Default', 'TestIdexia', 'TestSQI')]
    [string] $InputPreset = 'Custom',

    [Parameter(Mandatory = $false)]
    [string] $InputXlsxPath,

    [Parameter(Mandatory = $false)]
    [string] $Suffix = ' (archive)',

    [Parameter(Mandatory = $false)]
    [string] $LogFolder = (Join-Path $PSScriptRoot 'logs'),

    [Parameter(Mandatory = $false)]
    [string] $SharePointClientRuntimeDllPath = 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll',

    [Parameter(Mandatory = $false)]
    [string] $SharePointClientDllPath = 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll',

    [Parameter(Mandatory = $false)]
    [string] $ProjectServerClientDllPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Default', 'ITAR', 'Germany', 'China')]
    [string] $Region = 'Default',

    [Parameter(Mandatory = $false)]
    [int] $MaxRows = 0,

    [Parameter(Mandatory = $false)]
    [guid[]] $TargetResourceUid,

    [Parameter(Mandatory = $false)]
    [string] $SourceNameContains,

    [Parameter(Mandatory = $false)]
    [switch] $AllowFullWorkbook,

    [Parameter(Mandatory = $false)]
    [switch] $ValidateInputOnly,

    [Parameter(Mandatory = $false)]
    [switch] $EnableExperimentalUserDissociation,

    [Parameter(Mandatory = $false)]
    [switch] $ForceCheckInBeforeUpdate,

    [Parameter(Mandatory = $false)]
    [switch] $ForceCheckInAfterUpdate,

    [Parameter(Mandatory = $false)]
    [bool] $RetryOnConflict = $true,

    [Parameter(Mandatory = $false)]
    [switch] $StopOnFirstError
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($TenantPreset -eq 'Idexia' -and [string]::IsNullOrWhiteSpace($PwaUrl)) {
    $PwaUrl = 'https://idexia365.sharepoint.com/sites/pwa/'
}
elseif ($TenantPreset -eq 'SQI' -and [string]::IsNullOrWhiteSpace($PwaUrl)) {
    $PwaUrl = 'https://sqi365.sharepoint.com/sites/pwa/'
}

if ($InputPreset -eq 'Default' -and [string]::IsNullOrWhiteSpace($InputXlsxPath)) {
    $InputXlsxPath = Join-Path $PSScriptRoot 'CorrectionResourceMapping.xlsx'
}
elseif ($InputPreset -eq 'TestIdexia' -and [string]::IsNullOrWhiteSpace($InputXlsxPath)) {
    $InputXlsxPath = Join-Path $PSScriptRoot 'CorrectionResourceMappingTest.xlsx'
}
elseif ($InputPreset -eq 'TestSQI' -and [string]::IsNullOrWhiteSpace($InputXlsxPath)) {
    $InputXlsxPath = Join-Path $PSScriptRoot 'CorrectionResourceMappingTestSQI.xlsx'
}
elseif ([string]::IsNullOrWhiteSpace($InputXlsxPath)) {
    $InputXlsxPath = Join-Path $PSScriptRoot 'CorrectionResourceMapping.xlsx'
}

if ([string]::IsNullOrWhiteSpace($PwaUrl)) {
    throw "PwaUrl is required unless -TenantPreset Idexia or -TenantPreset SQI is used."
}

function Resolve-RequiredFile {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $Purpose
    )

    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($null -eq $resolved) {
        throw "Required file not found for ${Purpose}: $Path"
    }

    return $resolved.ProviderPath
}

function Find-ProjectServerClientDll {
    param([string] $ExplicitPath)

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        return (Resolve-RequiredFile -Path $ExplicitPath -Purpose 'Project Server CSOM')
    }

    $candidates = @(
        'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.ProjectServer.Client.dll',
        'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.ProjectServer.Client.Portable.dll',
        'C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.ProjectServer.Client.dll'
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    throw "Microsoft.ProjectServer.Client.dll was not found. Install/copy the Project Server CSOM assemblies and pass -ProjectServerClientDllPath."
}

function Import-CsomAssemblies {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SharePointClientRuntimeDllPath,

        [Parameter(Mandatory = $true)]
        [string] $SharePointClientDllPath,

        [Parameter(Mandatory = $true)]
        [string] $ProjectServerClientDllPath
    )

    Add-Type -Path (Resolve-RequiredFile -Path $SharePointClientRuntimeDllPath -Purpose 'SharePoint CSOM runtime')
    Add-Type -Path (Resolve-RequiredFile -Path $SharePointClientDllPath -Purpose 'SharePoint CSOM')
    Add-Type -Path (Resolve-RequiredFile -Path $ProjectServerClientDllPath -Purpose 'Project Server CSOM')
}

function Find-CsomDllFolder {
    $nugetRoot = Join-Path $env:USERPROFILE '.nuget\packages\microsoft.sharepointonline.csom'
    if (-not (Test-Path -LiteralPath $nugetRoot)) {
        return $null
    }

    $latest = Get-ChildItem -LiteralPath $nugetRoot -Directory | Sort-Object Name -Descending | Select-Object -First 1
    if ($null -eq $latest) {
        return $null
    }

    $net45 = Join-Path $latest.FullName 'lib\net45'
    if (Test-Path -LiteralPath $net45) {
        return $net45
    }

    return $null
}

function Resolve-CsomDllPath {
    param(
        [string] $ExplicitPath,
        [Parameter(Mandatory = $true)]
        [string] $DllName,
        [Parameter(Mandatory = $true)]
        [string] $Purpose
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        return (Resolve-RequiredFile -Path $ExplicitPath -Purpose $Purpose)
    }

    $folder = Find-CsomDllFolder
    if ($null -ne $folder) {
        $candidate = Join-Path $folder $DllName
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return (Resolve-RequiredFile -Path $DllName -Purpose $Purpose)
}

function Read-ZipEntryText {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive] $Zip,

        [Parameter(Mandatory = $true)]
        [string] $EntryName
    )

    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) {
        throw "XLSX entry not found: $EntryName"
    }

    $reader = New-Object System.IO.StreamReader($entry.Open())
    try {
        return $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }
}

function Convert-ExcelColumnRefToIndex {
    param([Parameter(Mandatory = $true)][string] $CellRef)

    $letters = ([regex]::Match($CellRef, '^[A-Z]+')).Value
    $index = 0
    foreach ($char in $letters.ToCharArray()) {
        $index = ($index * 26) + ([int][char]$char - [int][char]'A' + 1)
    }

    return $index - 1
}

function Get-XlsxCellValue {
    param(
        [Parameter(Mandatory = $true)] $Cell,
        [Parameter(Mandatory = $true)] [string[]] $SharedStrings
    )

    $cellType = $Cell.GetAttribute('t')
    $valueNode = $Cell.SelectSingleNode('*[local-name()="v"]')
    $cellValue = if ($null -ne $valueNode) { $valueNode.InnerText } else { '' }

    if ($cellType -eq 's' -and -not [string]::IsNullOrWhiteSpace($cellValue)) {
        return $SharedStrings[[int]$cellValue]
    }

    if ($cellType -eq 'inlineStr') {
        $inlineString = $Cell.SelectSingleNode('*[local-name()="is"]')
        if ($null -ne $inlineString) {
            return ($inlineString.ChildNodes | ForEach-Object { $_.InnerText }) -join ''
        }

        return ''
    }

    return $cellValue
}

function Import-ResourceMappingXlsx {
    param([Parameter(Mandatory = $true)][string] $Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $xlsxPath = Resolve-RequiredFile -Path $Path -Purpose 'resource mapping input'
    $tempXlsxPath = Join-Path $env:TEMP ("ProjectOnlineResourceMapping-{0}.xlsx" -f [guid]::NewGuid())
    Copy-Item -LiteralPath $xlsxPath -Destination $tempXlsxPath -Force -WhatIf:$false
    $zip = [System.IO.Compression.ZipFile]::OpenRead($tempXlsxPath)
    try {
        [xml] $sharedXml = Read-ZipEntryText -Zip $zip -EntryName 'xl/sharedStrings.xml'
        $sharedStrings = New-Object System.Collections.Generic.List[string]
        foreach ($si in $sharedXml.sst.si) {
            $parts = New-Object System.Collections.Generic.List[string]
            foreach ($node in $si.ChildNodes) {
                if ($node.Name -eq 't') { [void] $parts.Add($node.InnerText) }
                if ($node.Name -eq 'r') { [void] $parts.Add($node.t.InnerText) }
            }
            [void] $sharedStrings.Add(($parts -join ''))
        }

        [xml] $sheet = Read-ZipEntryText -Zip $zip -EntryName 'xl/worksheets/sheet1.xml'
        $rows = @()
        foreach ($row in $sheet.worksheet.sheetData.row) {
            $values = @{}
            foreach ($cell in $row.c) {
                $idx = Convert-ExcelColumnRefToIndex -CellRef $cell.GetAttribute('r')
                $values[$idx] = Get-XlsxCellValue -Cell $cell -SharedStrings $sharedStrings.ToArray()
            }

            if ($values.Count -gt 0) {
                $max = ($values.Keys | Measure-Object -Maximum).Maximum
                $rowValues = for ($i = 0; $i -le $max; $i++) {
                    if ($values.ContainsKey($i)) { $values[$i] } else { '' }
                }
                $rows += ,$rowValues
            }
        }

        if ($rows.Count -lt 2) {
            throw "Input workbook does not contain any data rows."
        }

        $headers = $rows[0]
        $requiredHeaders = @('New Name', 'Source Name', 'UID')
        foreach ($required in $requiredHeaders) {
            if ($headers -notcontains $required) {
                throw "Input workbook is missing required column: $required"
            }
        }

        $objects = New-Object System.Collections.Generic.List[object]
        for ($r = 1; $r -lt $rows.Count; $r++) {
            $obj = [ordered]@{}
            for ($c = 0; $c -lt $headers.Count; $c++) {
                if (-not [string]::IsNullOrWhiteSpace($headers[$c])) {
                    $obj[$headers[$c]] = if ($c -lt $rows[$r].Count) { $rows[$r][$c] } else { '' }
                }
            }
            [void] $objects.Add([pscustomobject]$obj)
        }

        return $objects
    }
    finally {
        $zip.Dispose()
        Remove-Item -LiteralPath $tempXlsxPath -Force -ErrorAction SilentlyContinue -WhatIf:$false
    }
}

function Connect-ProjectOnline {
    param(
        [Parameter(Mandatory = $true)]
        [string] $PwaUrl,

        [Parameter(Mandatory = $true)]
        [string] $Region
    )

    $uri = New-Object Uri($PwaUrl)
    $script:ProjectOnlineAuthCookie = GetAuthCookie -Uri $uri -Region $Region
    if ([string]::IsNullOrWhiteSpace($script:ProjectOnlineAuthCookie)) {
        throw "Could not acquire Project Online auth cookie."
    }

    $context = New-Object Microsoft.ProjectServer.Client.ProjectContext($PwaUrl)
    $context.add_ExecutingWebRequest({
        param($sender, $eventArgs)
        $eventArgs.WebRequestExecutor.RequestHeaders['Cookie'] = $script:ProjectOnlineAuthCookie
    })

    return $context
}

function New-ProjectOnlineWebSession {
    param([Parameter(Mandatory = $true)][Uri] $PwaUri)

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $session.Cookies.SetCookies($PwaUri, $script:ProjectOnlineAuthCookie)
    return $session
}

function Force-EnterpriseResourceCheckIn {
    param(
        [Parameter(Mandatory = $true)]
        [string] $PwaUrl,

        [Parameter(Mandatory = $true)]
        [guid] $ResourceUid,

        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.WebRequestSession] $WebSession,

        [Parameter(Mandatory = $true)]
        [string] $RequestDigest
    )

    $url = "$($PwaUrl.TrimEnd('/'))/_api/ProjectServer/EnterpriseResources('$($ResourceUid.ToString())')/ForceCheckIn"
    Invoke-RestMethod -Uri $url -Method POST -UseBasicParsing -WebSession $WebSession -Headers @{
        'X-RequestDigest' = $RequestDigest
        'Content-Length' = '0'
    } | Out-Null
}

function Test-IsHttpConflict {
    param([Parameter(Mandatory = $true)] $Exception)

    $current = $Exception
    while ($null -ne $current) {
        if ($current -is [System.Net.WebException] -and
            $null -ne $current.Response -and
            [int]$current.Response.StatusCode -eq 409) {
            return $true
        }
        $current = $current.InnerException
    }

    return ($Exception.Message -like '*409*' -or $Exception.Message -like '*Conflit*' -or $Exception.Message -like '*Conflict*')
}

function Invoke-ResourceForceCheckInIfNeeded {
    param(
        [Parameter(Mandatory = $true)]
        [string] $PwaUrl,

        [Parameter(Mandatory = $true)]
        [guid] $ResourceUid,

        [Parameter(Mandatory = $true)]
        [bool] $IsCheckedOut,

        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.WebRequestSession] $WebSession,

        [Parameter(Mandatory = $true)]
        [string] $RequestDigest,

        [Parameter(Mandatory = $true)]
        [string] $Stage
    )

    if (-not $IsCheckedOut) {
        return "$Stage`:SkippedNotCheckedOut"
    }

    try {
        Force-EnterpriseResourceCheckIn -PwaUrl $PwaUrl -ResourceUid $ResourceUid -WebSession $WebSession -RequestDigest $RequestDigest
        return "$Stage`:Succeeded"
    }
    catch {
        if (Test-IsHttpConflict -Exception $_.Exception) {
            return "$Stage`:Conflict"
        }

        throw
    }
}

function Get-EnterpriseResourceByUid {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.ProjectServer.Client.ProjectContext] $Context,

        [Parameter(Mandatory = $true)]
        [guid] $ResourceUid
    )

    $resource = $Context.EnterpriseResources.GetByGuid($ResourceUid)
    $Context.Load($resource)
    $Context.ExecuteQuery()
    return $resource
}

function Get-ResourceSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        $Resource
    )

    $linkedUser = $null
    try {
        if ($Resource.IsObjectPropertyInstantiated('User') -and $null -ne $Resource.User) {
            $linkedUser = $Resource.User.LoginName
        }
    }
    catch {
        $linkedUser = "Unreadable: $($_.Exception.Message)"
    }

    [pscustomobject]@{
        ResourceUid = $Resource.Id
        Name = $Resource.Name
        IsActive = $Resource.IsActive
        IsCheckedOut = $Resource.IsCheckedOut
        ResourceType = $Resource.ResourceType
        LinkedUser = $linkedUser
    }
}

function Get-TargetName {
    param(
        [Parameter(Mandatory = $true)]
        $Row,

        [Parameter(Mandatory = $true)]
        [string] $Suffix
    )

    if (-not [string]::IsNullOrWhiteSpace($Row.'New Name')) {
        return $Row.'New Name'.Trim()
    }

    return ($Row.'Source Name'.Trim() + $Suffix)
}

function Set-ResourceInactive {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.ProjectServer.Client.ProjectContext] $Context,

        [Parameter(Mandatory = $true)]
        $Resource
    )

    if ($Resource.IsActive -eq $false) {
        return 'AlreadyInactive'
    }

    $Resource.IsActive = $false
    $Context.EnterpriseResources.Update()
    $Context.ExecuteQuery()
    return 'Inactivated'
}

function Set-ResourceInactiveAndDisplayName {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.ProjectServer.Client.ProjectContext] $Context,

        [Parameter(Mandatory = $true)]
        $Resource,

        [Parameter(Mandatory = $true)]
        [string] $TargetName
    )

    $actions = New-Object System.Collections.Generic.List[string]

    if ($Resource.IsActive -ne $false) {
        $Resource.IsActive = $false
        [void] $actions.Add('Inactivated')
    }
    else {
        [void] $actions.Add('AlreadyInactive')
    }

    if ($Resource.Name -ne $TargetName) {
        $Resource.Name = $TargetName
        [void] $actions.Add('Renamed')
    }
    else {
        [void] $actions.Add('NameAlreadySet')
    }

    $Context.EnterpriseResources.Update()
    $Context.ExecuteQuery()
    return ($actions -join ';')
}

function Clear-ResourceUserAccountLink {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.ProjectServer.Client.ProjectContext] $Context,

        [Parameter(Mandatory = $true)]
        $Resource,

        [Parameter(Mandatory = $true)]
        [bool] $Enabled
    )

    if (-not $Enabled) {
        return 'Skipped: dissociation not enabled. A verifier dans l''environnement / non confirme par source officielle.'
    }

    $property = $Resource.GetType().GetProperty('User')
    if ($null -eq $property -or -not $property.CanWrite) {
        throw 'Cannot dissociate user link: EnterpriseResource.User is not writable in this CSOM assembly. A verifier dans l''environnement / non confirme par source officielle.'
    }

    $Resource.User = $null
    $Context.EnterpriseResources.Update()
    $Context.ExecuteQuery()
    return 'ExperimentalUserLinkCleared'
}

function New-LogRow {
    param(
        [string] $Status,
        [object] $Row,
        [object] $Before,
        [object] $After,
        [string] $TargetName,
        [string] $DissociationResult,
        [string] $InactiveResult,
        [string] $RenameResult,
        [string] $CheckInBeforeResult,
        [string] $CheckInAfterResult,
        [string] $RetryResult,
        [string] $ErrorMessage
    )

    [pscustomobject]@{
        Timestamp = (Get-Date).ToString('s')
        Status = $Status
        ResourceUid = $Row.UID
        SourceName = $Row.'Source Name'
        InputNewName = $Row.'New Name'
        TargetName = $TargetName
        BeforeName = if ($Before) { $Before.Name } else { $null }
        AfterName = if ($After) { $After.Name } else { $null }
        BeforeIsActive = if ($Before) { $Before.IsActive } else { $null }
        AfterIsActive = if ($After) { $After.IsActive } else { $null }
        BeforeLinkedUser = if ($Before) { $Before.LinkedUser } else { $null }
        AfterLinkedUser = if ($After) { $After.LinkedUser } else { $null }
        DissociationResult = $DissociationResult
        InactiveResult = $InactiveResult
        RenameResult = $RenameResult
        CheckInBeforeResult = $CheckInBeforeResult
        CheckInAfterResult = $CheckInAfterResult
        RetryResult = $RetryResult
        ErrorMessage = $ErrorMessage
    }
}

$inputPath = Resolve-RequiredFile -Path $InputXlsxPath -Purpose 'resource mapping input'
$rows = Import-ResourceMappingXlsx -Path $inputPath

if (-not $AllowFullWorkbook -and $null -eq $TargetResourceUid -and [string]::IsNullOrWhiteSpace($SourceNameContains)) {
    throw "Refusing to process the full workbook without an explicit filter. Use -TargetResourceUid, -SourceNameContains, or -AllowFullWorkbook."
}

if ($null -ne $TargetResourceUid -and $TargetResourceUid.Count -gt 0) {
    $uidSet = @{}
    foreach ($uid in $TargetResourceUid) {
        $uidSet[$uid.ToString().ToLowerInvariant()] = $true
    }

    $rows = $rows | Where-Object {
        $candidate = [guid]::Empty
        [guid]::TryParse([string]$_.UID, [ref]$candidate) -and $uidSet.ContainsKey($candidate.ToString().ToLowerInvariant())
    }
}

if (-not [string]::IsNullOrWhiteSpace($SourceNameContains)) {
    $rows = $rows | Where-Object {
        $_.'Source Name' -like "*$SourceNameContains*" -or $_.'New Name' -like "*$SourceNameContains*"
    }
}

if ($MaxRows -gt 0) {
    $rows = $rows | Select-Object -First $MaxRows
}

if ($ValidateInputOnly) {
    $rows | Select-Object 'New Name','Source Name','UID','UPN','Source Account','Target Account','Action' | Format-Table -AutoSize
    Write-Host "Input validation only. Selected rows: $(@($rows).Count)"
    return
}

if (@($rows).Count -eq 0) {
    throw "No input row matched the requested filter."
}

$projectDll = Resolve-CsomDllPath -ExplicitPath $ProjectServerClientDllPath -DllName 'Microsoft.ProjectServer.Client.dll' -Purpose 'Project Server CSOM'
$spRuntimeDll = Resolve-CsomDllPath -ExplicitPath $SharePointClientRuntimeDllPath -DllName 'Microsoft.SharePoint.Client.Runtime.dll' -Purpose 'SharePoint CSOM runtime'
$spClientDll = Resolve-CsomDllPath -ExplicitPath $SharePointClientDllPath -DllName 'Microsoft.SharePoint.Client.dll' -Purpose 'SharePoint CSOM'
Import-CsomAssemblies -SharePointClientRuntimeDllPath $spRuntimeDll -SharePointClientDllPath $spClientDll -ProjectServerClientDllPath $projectDll

. "$PSScriptRoot\Common.ps1"

if (-not (Test-Path -LiteralPath $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory -WhatIf:$false | Out-Null
}

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$logPath = Join-Path $LogFolder "orphan-resource-cleanup-$timestamp.csv"

$context = Connect-ProjectOnline -PwaUrl $PwaUrl -Region $Region
$pwaUri = New-Object Uri($PwaUrl)
$webSession = New-ProjectOnlineWebSession -PwaUri $pwaUri
$requestDigest = if ($ForceCheckInBeforeUpdate -or $ForceCheckInAfterUpdate) {
    GetXRequestDigest -Uri $pwaUri -OnPrem $false -AuthCookie $script:ProjectOnlineAuthCookie
}
else {
    ''
}
$log = New-Object System.Collections.Generic.List[object]

foreach ($row in $rows) {
    $before = $null
    $after = $null
    $targetName = $null
    $dissociationResult = $null
    $inactiveResult = $null
    $renameResult = $null
    $checkInBeforeResult = $null
    $checkInAfterResult = $null
    $retryResult = $null

    try {
        [guid] $resourceUid = [guid]::Empty
        if (-not [guid]::TryParse([string]$row.UID, [ref]$resourceUid) -or $resourceUid -eq [guid]::Empty) {
            throw "Invalid or empty UID: $($row.UID)"
        }

        $targetName = Get-TargetName -Row $row -Suffix $Suffix
        if ([string]::IsNullOrWhiteSpace($targetName)) {
            throw "Target resource name is empty for UID $resourceUid"
        }

        $resource = Get-EnterpriseResourceByUid -Context $context -ResourceUid $resourceUid
        $before = Get-ResourceSnapshot -Resource $resource

        $operation = "set inactive and rename to '$targetName'"
        if ($EnableExperimentalUserDissociation) {
            $operation = "experimentally dissociate orphan account, set inactive, rename to '$targetName'"
        }
        if ($PSCmdlet.ShouldProcess("$resourceUid / $($resource.Name)", $operation)) {
            if ($ForceCheckInBeforeUpdate) {
                $checkInBeforeResult = Invoke-ResourceForceCheckInIfNeeded -PwaUrl $PwaUrl -ResourceUid $resourceUid -IsCheckedOut ([bool]$before.IsCheckedOut) -WebSession $webSession -RequestDigest $requestDigest -Stage 'ForceCheckInBeforeUpdate'
            }
            else {
                $checkInBeforeResult = 'Skipped'
            }

            $resource = Get-EnterpriseResourceByUid -Context $context -ResourceUid $resourceUid
            $dissociationResult = Clear-ResourceUserAccountLink -Context $context -Resource $resource -Enabled ([bool]$EnableExperimentalUserDissociation)

            $resource = Get-EnterpriseResourceByUid -Context $context -ResourceUid $resourceUid
            try {
                $combinedResult = Set-ResourceInactiveAndDisplayName -Context $context -Resource $resource -TargetName $targetName
                $inactiveResult = $combinedResult
                $renameResult = $combinedResult
                $retryResult = 'NotNeeded'
            }
            catch {
                if ($RetryOnConflict -and (Test-IsHttpConflict -Exception $_.Exception)) {
                    $retryResult = 'ConflictDetected;ForceCheckInAndRetry'
                    Force-EnterpriseResourceCheckIn -PwaUrl $PwaUrl -ResourceUid $resourceUid -WebSession $webSession -RequestDigest (GetXRequestDigest -Uri $pwaUri -OnPrem $false -AuthCookie $script:ProjectOnlineAuthCookie)
                    $resource = Get-EnterpriseResourceByUid -Context $context -ResourceUid $resourceUid
                    $combinedResult = Set-ResourceInactiveAndDisplayName -Context $context -Resource $resource -TargetName $targetName
                    $inactiveResult = $combinedResult
                    $renameResult = $combinedResult
                    $retryResult = "$retryResult;RetrySucceeded"
                }
                else {
                    throw
                }
            }

            if ($ForceCheckInAfterUpdate) {
                $resourceAfterUpdate = Get-EnterpriseResourceByUid -Context $context -ResourceUid $resourceUid
                $snapshotAfterUpdate = Get-ResourceSnapshot -Resource $resourceAfterUpdate
                $checkInAfterResult = Invoke-ResourceForceCheckInIfNeeded -PwaUrl $PwaUrl -ResourceUid $resourceUid -IsCheckedOut ([bool]$snapshotAfterUpdate.IsCheckedOut) -WebSession $webSession -RequestDigest $requestDigest -Stage 'ForceCheckInAfterUpdate'
            }
            else {
                $checkInAfterResult = 'Skipped'
            }

            $resource = Get-EnterpriseResourceByUid -Context $context -ResourceUid $resourceUid
            $after = Get-ResourceSnapshot -Resource $resource
        }
        else {
            $dissociationResult = 'WhatIf'
            $inactiveResult = 'WhatIf'
            $renameResult = 'WhatIf'
            $checkInBeforeResult = 'WhatIf'
            $checkInAfterResult = 'WhatIf'
            $retryResult = 'WhatIf'
            $after = $before
        }

        [void] $log.Add((New-LogRow -Status 'Success' -Row $row -Before $before -After $after -TargetName $targetName -DissociationResult $dissociationResult -InactiveResult $inactiveResult -RenameResult $renameResult -CheckInBeforeResult $checkInBeforeResult -CheckInAfterResult $checkInAfterResult -RetryResult $retryResult -ErrorMessage ''))
    }
    catch {
        [void] $log.Add((New-LogRow -Status 'Failed' -Row $row -Before $before -After $after -TargetName $targetName -DissociationResult $dissociationResult -InactiveResult $inactiveResult -RenameResult $renameResult -CheckInBeforeResult $checkInBeforeResult -CheckInAfterResult $checkInAfterResult -RetryResult $retryResult -ErrorMessage $_.Exception.Message))
        Write-Host "ERROR for resource $($row.UID): $($_.Exception.Message)" -ForegroundColor Red
        if ($StopOnFirstError) {
            break
        }
    }
}

$log | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
Write-Host "Processing log written to: $logPath"
