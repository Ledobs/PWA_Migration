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
    [Parameter(Mandatory = $true)]
    [string] $PwaUrl,

    [Parameter(Mandatory = $false)]
    [string] $InputXlsxPath = (Join-Path $PSScriptRoot 'CorrectionResourceMapping.xlsx'),

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
    [switch] $StopOnFirstError
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

    if ($Cell.t -eq 's' -and -not [string]::IsNullOrWhiteSpace([string]$Cell.v)) {
        return $SharedStrings[[int]$Cell.v]
    }

    if ($Cell.t -eq 'inlineStr') {
        return ($Cell.is.t | ForEach-Object { $_.'#text' }) -join ''
    }

    if ($null -ne $Cell.v) {
        return [string]$Cell.v
    }

    return ''
}

function Import-ResourceMappingXlsx {
    param([Parameter(Mandatory = $true)][string] $Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $xlsxPath = Resolve-RequiredFile -Path $Path -Purpose 'resource mapping input'
    $tempXlsxPath = Join-Path $env:TEMP ("ProjectOnlineResourceMapping-{0}.xlsx" -f [guid]::NewGuid())
    Copy-Item -LiteralPath $xlsxPath -Destination $tempXlsxPath -Force
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
                $idx = Convert-ExcelColumnRefToIndex -CellRef $cell.r
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
        Remove-Item -LiteralPath $tempXlsxPath -Force -ErrorAction SilentlyContinue
    }
}

function Connect-ProjectOnline {
    param(
        [Parameter(Mandatory = $true)]
        [string] $PwaUrl,

        [Parameter(Mandatory = $true)]
        [string] $Region
    )

    . "$PSScriptRoot\Common.ps1"

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

function Set-ResourceDisplayName {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.ProjectServer.Client.ProjectContext] $Context,

        [Parameter(Mandatory = $true)]
        $Resource,

        [Parameter(Mandatory = $true)]
        [string] $TargetName
    )

    if ($Resource.Name -eq $TargetName) {
        return 'NameAlreadySet'
    }

    $Resource.Name = $TargetName
    $Context.EnterpriseResources.Update()
    $Context.ExecuteQuery()
    return 'Renamed'
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
        ErrorMessage = $ErrorMessage
    }
}

$inputPath = Resolve-RequiredFile -Path $InputXlsxPath -Purpose 'resource mapping input'
$projectDll = Find-ProjectServerClientDll -ExplicitPath $ProjectServerClientDllPath
Import-CsomAssemblies -SharePointClientRuntimeDllPath $SharePointClientRuntimeDllPath -SharePointClientDllPath $SharePointClientDllPath -ProjectServerClientDllPath $projectDll

if (-not (Test-Path -LiteralPath $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory | Out-Null
}

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$logPath = Join-Path $LogFolder "orphan-resource-cleanup-$timestamp.csv"
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

$context = Connect-ProjectOnline -PwaUrl $PwaUrl -Region $Region
$log = New-Object System.Collections.Generic.List[object]

foreach ($row in $rows) {
    $before = $null
    $after = $null
    $targetName = $null
    $dissociationResult = $null
    $inactiveResult = $null
    $renameResult = $null

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

        if ($resource.IsCheckedOut) {
            throw "Resource is checked out and cannot be safely modified: $resourceUid"
        }

        $operation = "dissociate orphan account, set inactive, rename to '$targetName'"
        if ($PSCmdlet.ShouldProcess("$resourceUid / $($resource.Name)", $operation)) {
            $dissociationResult = Clear-ResourceUserAccountLink -Context $context -Resource $resource -Enabled ([bool]$EnableExperimentalUserDissociation)

            $resource = Get-EnterpriseResourceByUid -Context $context -ResourceUid $resourceUid
            $inactiveResult = Set-ResourceInactive -Context $context -Resource $resource

            $resource = Get-EnterpriseResourceByUid -Context $context -ResourceUid $resourceUid
            $renameResult = Set-ResourceDisplayName -Context $context -Resource $resource -TargetName $targetName

            $resource = Get-EnterpriseResourceByUid -Context $context -ResourceUid $resourceUid
            $after = Get-ResourceSnapshot -Resource $resource
        }
        else {
            $dissociationResult = 'WhatIf'
            $inactiveResult = 'WhatIf'
            $renameResult = 'WhatIf'
            $after = $before
        }

        [void] $log.Add((New-LogRow -Status 'Success' -Row $row -Before $before -After $after -TargetName $targetName -DissociationResult $dissociationResult -InactiveResult $inactiveResult -RenameResult $renameResult -ErrorMessage ''))
    }
    catch {
        [void] $log.Add((New-LogRow -Status 'Failed' -Row $row -Before $before -After $after -TargetName $targetName -DissociationResult $dissociationResult -InactiveResult $inactiveResult -RenameResult $renameResult -ErrorMessage $_.Exception.Message))
        Write-Error $_.Exception.Message
        if ($StopOnFirstError) {
            break
        }
    }
}

$log | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
Write-Host "Processing log written to: $logPath"
