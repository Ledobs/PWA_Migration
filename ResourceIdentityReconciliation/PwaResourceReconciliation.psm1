#requires -version 5.1
<#
====================================================================================================
 Nom du fichier : PwaResourceReconciliation.psm1
 Auteur         : François Breton
 Date           : 2026-08-06
 Révision       : 1.0.1
 Objet          : Fonctions de lecture, normalisation, réconciliation et production du rapport Excel.
 Mode           : INVENTORY seulement - aucune écriture dans Project Online, PSSE ou SQL Server.
====================================================================================================
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function ConvertTo-CanonicalGuid {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object] $Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    $guid = [guid]::Empty
    if (-not [guid]::TryParse(([string]$Value).Trim(), [ref]$guid)) {
        throw "La valeur '$Value' n'est pas un GUID valide."
    }

    return $guid.ToString('D').ToLowerInvariant()
}

function ConvertTo-OuiNon {
    param([bool] $Value)
    if ($Value) { return 'Oui' }
    return 'Non'
}

function Get-FirstPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object] $InputObject,
        [Parameter(Mandatory = $true)][string[]] $Names,
        [object] $DefaultValue = $null
    )

    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    return $DefaultValue
}

function Resolve-PropertyName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object] $Sample,
        [Parameter(Mandatory = $true)][string[]] $Candidates,
        [Parameter(Mandatory = $true)][string] $Purpose,
        [switch] $Optional
    )

    foreach ($candidate in $Candidates) {
        if ($null -ne $Sample.PSObject.Properties[$candidate]) {
            return $candidate
        }
    }

    if ($Optional) { return $null }
    throw "Aucune propriété compatible trouvée pour '$Purpose'. Candidats : $($Candidates -join ', ')."
}

function Write-ReconciliationLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $LogPath,
        [Parameter(Mandatory = $true)][ValidateSet('INFO','AVERTISSEMENT','ERREUR')][string] $Level,
        [Parameter(Mandatory = $true)][string] $Message
    )

    $line = '{0} [{1}] {2}' -f (Get-Date).ToString('s'), $Level, $Message
    $line | Add-Content -LiteralPath $LogPath -Encoding UTF8
    Write-Host $line
}

function Resolve-RequiredFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Path,
        [Parameter(Mandatory = $true)][string] $Purpose
    )

    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($null -eq $resolved) {
        throw "Fichier requis introuvable pour '$Purpose' : $Path"
    }
    return $resolved.ProviderPath
}

function Import-ProjectCsomAssemblies {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable] $Csom)

    $runtime = Resolve-RequiredFile -Path $Csom.SharePointClientRuntimeDllPath -Purpose 'SharePoint CSOM Runtime'
    $client = Resolve-RequiredFile -Path $Csom.SharePointClientDllPath -Purpose 'SharePoint CSOM'
    $project = Resolve-RequiredFile -Path $Csom.ProjectServerClientDllPath -Purpose 'Project Server CSOM'

    foreach ($assembly in @($runtime, $client, $project)) {
        if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies().Location -contains $assembly)) {
            Add-Type -Path $assembly
        }
    }
}

function New-ProjectOnlineConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $PwaUrl,
        [Parameter(Mandatory = $true)][string] $CommonScriptPath,
        [ValidateSet('Default','ITAR','Germany','China')][string] $Region = 'Default'
    )

    . (Resolve-RequiredFile -Path $CommonScriptPath -Purpose 'authentification Project Online (Common.ps1)')
    if (-not (Get-Command GetAuthCookie -ErrorAction SilentlyContinue)) {
        throw "La fonction GetAuthCookie n'a pas été chargée depuis Common.ps1."
    }

    $pwaUri = New-Object System.Uri($PwaUrl)
    $authCookie = GetAuthCookie -Uri $pwaUri -Region $Region
    if ([string]::IsNullOrWhiteSpace($authCookie)) {
        throw 'Impossible d’obtenir le cookie d’authentification Project Online.'
    }

    $webSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $webSession.Cookies.SetCookies($pwaUri, $authCookie)

    $context = New-Object Microsoft.ProjectServer.Client.ProjectContext($PwaUrl)
    $cookieForHandler = $authCookie
    $requestHandler = {
        param($sender, $eventArgs)
        $eventArgs.WebRequestExecutor.RequestHeaders['Cookie'] = $cookieForHandler
    }.GetNewClosure()
    $context.add_ExecutingWebRequest($requestHandler)

    return [pscustomobject]@{
        WebSession = $webSession
        ProjectContext = $context
        AuthenticationType = 'Microsoft 365 / cookie interactif'
    }
}

function New-PsseProjectContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $PwaUrl,
        [System.Management.Automation.PSCredential] $Credential
    )

    $context = New-Object Microsoft.ProjectServer.Client.ProjectContext($PwaUrl)
    if ($null -ne $Credential) {
        $context.Credentials = $Credential.GetNetworkCredential()
    }
    else {
        $context.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    }
    return $context
}

function Get-ODataItemsFromResponse {
    param([Parameter(Mandatory = $true)][object] $Response)

    if ($null -ne $Response.PSObject.Properties['value']) {
        return @($Response.value)
    }
    if ($null -ne $Response.PSObject.Properties['d']) {
        if ($null -ne $Response.d.PSObject.Properties['results']) {
            return @($Response.d.results)
        }
        return @($Response.d)
    }
    return @($Response)
}

function Get-ODataNextLink {
    param([Parameter(Mandatory = $true)][object] $Response)

    foreach ($name in @('@odata.nextLink','odata.nextLink')) {
        if ($null -ne $Response.PSObject.Properties[$name]) {
            return [string]$Response.PSObject.Properties[$name].Value
        }
    }
    if ($null -ne $Response.PSObject.Properties['d'] -and
        $null -ne $Response.d.PSObject.Properties['__next']) {
        return [string]$Response.d.__next
    }
    return $null
}

function Invoke-ProjectDataQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $ProjectDataUrl,
        [Parameter(Mandatory = $true)][string] $EntitySet,
        [Parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession] $WebSession,
        [string] $Query = ''
    )

    $base = $ProjectDataUrl.TrimEnd('/')
    $nextUrl = "$base/$EntitySet"
    if (-not [string]::IsNullOrWhiteSpace($Query)) {
        $nextUrl = "$nextUrl`?$Query"
    }

    $items = New-Object System.Collections.Generic.List[object]
    do {
        $response = Invoke-RestMethod -Uri $nextUrl -Method Get -UseBasicParsing -WebSession $WebSession -Headers @{
            Accept = 'application/json;odata=verbose'
        }
        foreach ($item in @(Get-ODataItemsFromResponse -Response $response)) {
            [void]$items.Add($item)
        }
        $nextUrl = Get-ODataNextLink -Response $response
    } while (-not [string]::IsNullOrWhiteSpace($nextUrl))

    return $items.ToArray()
}

function Get-ProjectOnlineReportingInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $ProjectDataUrl,
        [Parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession] $WebSession,
        [Parameter(Mandatory = $true)][string[]] $ResourceUids
    )

    # Les requêtes d’échantillon servent à découvrir les noms réels des propriétés du tenant.
    $resourceSample = @(Invoke-ProjectDataQuery -ProjectDataUrl $ProjectDataUrl -EntitySet 'Resources' -WebSession $WebSession -Query '$top=1') | Select-Object -First 1
    $projectSample = @(Invoke-ProjectDataQuery -ProjectDataUrl $ProjectDataUrl -EntitySet 'Projects' -WebSession $WebSession -Query '$top=1') | Select-Object -First 1
    $assignmentSample = @(Invoke-ProjectDataQuery -ProjectDataUrl $ProjectDataUrl -EntitySet 'Assignments' -WebSession $WebSession -Query '$top=1') | Select-Object -First 1

    if ($null -eq $resourceSample -or $null -eq $projectSample) {
        throw 'Les entités Resources ou Projects de ProjectData ne retournent aucun échantillon.'
    }

    $resourceUidField = Resolve-PropertyName -Sample $resourceSample -Candidates @('ResourceUID','ResourceUid','ResourceId') -Purpose 'ResourceUID'
    $projectUidField = Resolve-PropertyName -Sample $projectSample -Candidates @('ProjectUID','ProjectUid','ProjectId') -Purpose 'ProjectUID'
    $ownerUidField = Resolve-PropertyName -Sample $projectSample -Candidates @('ProjectOwnerResourceUID','ProjectOwnerResourceUid','ProjectOwnerId') -Purpose 'Project Owner UID'

    $resourceFilterParts = foreach ($uid in $ResourceUids) { "$resourceUidField eq guid'$uid'" }
    $resourceQuery = '$filter=' + [uri]::EscapeDataString(($resourceFilterParts -join ' or '))
    $rawResources = @(Invoke-ProjectDataQuery -ProjectDataUrl $ProjectDataUrl -EntitySet 'Resources' -WebSession $WebSession -Query $resourceQuery)
    $rawProjects = @(Invoke-ProjectDataQuery -ProjectDataUrl $ProjectDataUrl -EntitySet 'Projects' -WebSession $WebSession)

    $rawAssignments = @()
    $assignmentResourceUidField = $null
    if ($null -ne $assignmentSample) {
        $assignmentResourceUidField = Resolve-PropertyName -Sample $assignmentSample -Candidates @('ResourceUID','ResourceUid','ResourceId') -Purpose 'Assignment ResourceUID'
        $assignmentFilterParts = foreach ($uid in $ResourceUids) { "$assignmentResourceUidField eq guid'$uid'" }
        $assignmentQuery = '$filter=' + [uri]::EscapeDataString(($assignmentFilterParts -join ' or '))
        $rawAssignments = @(Invoke-ProjectDataQuery -ProjectDataUrl $ProjectDataUrl -EntitySet 'Assignments' -WebSession $WebSession -Query $assignmentQuery)
    }

    $resources = foreach ($item in $rawResources) {
        [pscustomobject]@{
            ResourceUID = ConvertTo-CanonicalGuid (Get-FirstPropertyValue $item @($resourceUidField))
            ResourceName = [string](Get-FirstPropertyValue $item @('ResourceName','Name'))
            ResourceAccount = [string](Get-FirstPropertyValue $item @('ResourceAccount','ResourceNTAccount','UserAccount'))
            UserClaimsAccount = [string](Get-FirstPropertyValue $item @('UserClaimsAccount','ResourceClaimsAccount','ClaimsAccount'))
            ResourceEmailAddress = [string](Get-FirstPropertyValue $item @('ResourceEmailAddress','Email'))
            ResourceIsActive = Get-FirstPropertyValue $item @('ResourceIsActive','IsActive')
        }
    }

    $projects = foreach ($item in $rawProjects) {
        [pscustomobject]@{
            ProjectUID = ConvertTo-CanonicalGuid (Get-FirstPropertyValue $item @($projectUidField))
            ProjectName = [string](Get-FirstPropertyValue $item @('ProjectName','Name'))
            ProjectOwnerName = [string](Get-FirstPropertyValue $item @('ProjectOwnerName','OwnerName'))
            ProjectOwnerResourceUID = if ($null -ne $ownerUidField) { ConvertTo-CanonicalGuid (Get-FirstPropertyValue $item @($ownerUidField)) } else { $null }
        }
    }

    $assignments = foreach ($item in $rawAssignments) {
        [pscustomobject]@{
            ProjectUID = ConvertTo-CanonicalGuid (Get-FirstPropertyValue $item @('ProjectUID','ProjectUid','ProjectId'))
            ProjectName = [string](Get-FirstPropertyValue $item @('ProjectName'))
            TaskUID = ConvertTo-CanonicalGuid (Get-FirstPropertyValue $item @('TaskUID','TaskUid','TaskId'))
            TaskName = [string](Get-FirstPropertyValue $item @('TaskName'))
            AssignmentUID = ConvertTo-CanonicalGuid (Get-FirstPropertyValue $item @('AssignmentUID','AssignmentUid','AssignmentId'))
            ResourceUID = ConvertTo-CanonicalGuid (Get-FirstPropertyValue $item @($assignmentResourceUidField))
        }
    }

    return [pscustomobject]@{
        Resources = @($resources)
        Projects = @($projects)
        Assignments = @($assignments)
        FieldMap = [ordered]@{
            ResourceUID = $resourceUidField
            ProjectUID = $projectUidField
            ProjectOwnerResourceUID = $ownerUidField
            AssignmentResourceUID = $assignmentResourceUidField
        }
    }
}

function Get-PublishedProjectTeamInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object] $Context,
        [Parameter(Mandatory = $true)][string[]] $ResourceUids,
        [Parameter(Mandatory = $true)][string] $Environment,
        [Parameter(Mandatory = $true)][string] $LogPath
    )

    $wanted = @{}
    foreach ($uid in $ResourceUids) { $wanted[(ConvertTo-CanonicalGuid $uid)] = $true }

    $projects = $Context.Projects
    $Context.Load($projects)
    $Context.ExecuteQuery()
    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($project in $projects) {
        try {
            $team = $project.ProjectResources
            $Context.Load($team)
            $Context.ExecuteQuery()
            foreach ($projectResource in $team) {
                $uidValue = Get-FirstPropertyValue $projectResource @('Id','ResourceUid','ResourceUID')
                if ($null -eq $uidValue) { continue }
                $uid = ConvertTo-CanonicalGuid $uidValue
                if ($wanted.ContainsKey($uid)) {
                    [void]$rows.Add([pscustomobject]@{
                        ProjectUID = ConvertTo-CanonicalGuid $project.Id
                        ProjectName = [string]$project.Name
                        ResourceUID = $uid
                        Environment = $Environment
                    })
                }
            }
        }
        catch {
            Write-ReconciliationLog -LogPath $LogPath -Level AVERTISSEMENT -Message "Project Team non lisible pour $($project.Id) - $($project.Name) : $($_.Exception.Message)"
        }
    }

    return $rows.ToArray()
}

function New-ReadOnlySqlConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable] $Sql,
        [System.Management.Automation.PSCredential] $Credential
    )

    $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $builder['Data Source'] = $Sql.Server
    $builder['Initial Catalog'] = $Sql.Database
    $builder['Application Name'] = 'PWA Resource Reconciliation Inventory'
    $builder['ApplicationIntent'] = 'ReadOnly'
    $builder['Encrypt'] = [bool]$Sql.Encrypt
    $builder['TrustServerCertificate'] = [bool]$Sql.TrustServerCertificate
    $builder['Connect Timeout'] = if ($Sql.ContainsKey('ConnectTimeout')) { [int]$Sql.ConnectTimeout } else { 30 }

    if ($null -eq $Credential) {
        $builder['Integrated Security'] = $true
        return New-Object System.Data.SqlClient.SqlConnection($builder.ConnectionString)
    }

    $builder['Integrated Security'] = $false
    $secureCredential = New-Object System.Data.SqlClient.SqlCredential($Credential.UserName, $Credential.Password)
    return New-Object System.Data.SqlClient.SqlConnection($builder.ConnectionString, $secureCredential)
}

function Invoke-ReadOnlySqlDataTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Data.SqlClient.SqlConnection] $Connection,
        [Parameter(Mandatory = $true)][string] $Query,
        [hashtable] $Parameters = @{}
    )

    if ($Query.TrimStart() -notmatch '^(?is)SELECT\b' -or
        $Query -match '(?is)\b(INSERT|UPDATE|DELETE|MERGE|DROP|ALTER|CREATE|TRUNCATE|EXEC|EXECUTE|GRANT|REVOKE)\b') {
        throw 'Seules les requêtes SELECT sans instruction de modification sont permises en mode Inventory.'
    }

    $command = $Connection.CreateCommand()
    $command.CommandText = $Query
    $command.CommandTimeout = 120
    foreach ($entry in $Parameters.GetEnumerator()) {
        $parameter = $command.Parameters.Add("@$($entry.Key)", [System.Data.SqlDbType]::UniqueIdentifier)
        $parameter.Value = [guid]$entry.Value
    }

    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
    $table = New-Object System.Data.DataTable
    [void]$adapter.Fill($table)
    Write-Output -NoEnumerate $table
}

function Get-SqlViewColumns {
    param(
        [Parameter(Mandatory = $true)][System.Data.SqlClient.SqlConnection] $Connection,
        [Parameter(Mandatory = $true)][string] $Schema,
        [Parameter(Mandatory = $true)][string] $View
    )

    $command = $Connection.CreateCommand()
    $command.CommandText = @'
SELECT c.name
FROM sys.columns AS c
INNER JOIN sys.views AS v ON v.object_id = c.object_id
INNER JOIN sys.schemas AS s ON s.schema_id = v.schema_id
WHERE s.name = @SchemaName AND v.name = @ViewName
ORDER BY c.column_id;
'@
    [void]$command.Parameters.Add('@SchemaName',[System.Data.SqlDbType]::NVarChar,128)
    [void]$command.Parameters.Add('@ViewName',[System.Data.SqlDbType]::NVarChar,128)
    $command.Parameters['@SchemaName'].Value = $Schema
    $command.Parameters['@ViewName'].Value = $View
    $reader = $command.ExecuteReader()
    $columns = New-Object System.Collections.Generic.List[string]
    try {
        while ($reader.Read()) { [void]$columns.Add($reader.GetString(0)) }
    }
    finally { $reader.Dispose() }
    if ($columns.Count -eq 0) { throw "Vue SQL introuvable ou inaccessible : [$Schema].[$View]." }
    return $columns.ToArray()
}

function Resolve-SqlColumn {
    param(
        [Parameter(Mandatory = $true)][string[]] $Available,
        [Parameter(Mandatory = $true)][string[]] $Candidates,
        [Parameter(Mandatory = $true)][string] $Purpose,
        [switch] $Optional
    )
    foreach ($candidate in $Candidates) {
        $match = $Available | Where-Object { $_ -ieq $candidate } | Select-Object -First 1
        if ($null -ne $match) { return $match }
    }
    if ($Optional) { return $null }
    throw "Colonne SQL compatible introuvable pour '$Purpose'. Candidats : $($Candidates -join ', ')."
}

function Get-SqlSelectExpression {
    param([string] $Column, [string] $Alias)
    if ([string]::IsNullOrWhiteSpace($Column)) { return "CAST(NULL AS nvarchar(4000)) AS [$Alias]" }
    return "[$($Column.Replace(']',']]'))] AS [$Alias]"
}

function Convert-DataRowsToObjects {
    param([Parameter(Mandatory = $true)][System.Data.DataTable] $Table)
    foreach ($row in $Table.Rows) {
        $record = [ordered]@{}
        foreach ($column in $Table.Columns) {
            $value = $row[$column.ColumnName]
            $record[$column.ColumnName] = if ($value -is [System.DBNull]) { $null } else { $value }
        }
        [pscustomobject]$record
    }
}

function Get-PsseReportingInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Data.SqlClient.SqlConnection] $Connection,
        [Parameter(Mandatory = $true)][hashtable] $Views,
        [Parameter(Mandatory = $true)][string[]] $ResourceUids
    )

    $schema = if ($Views.ContainsKey('Schema')) { $Views.Schema } else { 'dbo' }
    foreach ($identifier in @($schema, $Views.ResourceView, $Views.ProjectView, $Views.AssignmentView)) {
        if ([string]$identifier -notmatch '^[A-Za-z0-9_]+$') {
            throw "Identifiant SQL non sécuritaire dans la configuration : '$identifier'."
        }
    }
    $resourceColumns = Get-SqlViewColumns $Connection $schema $Views.ResourceView
    $projectColumns = Get-SqlViewColumns $Connection $schema $Views.ProjectView
    $assignmentColumns = Get-SqlViewColumns $Connection $schema $Views.AssignmentView

    $rUid = Resolve-SqlColumn $resourceColumns @('ResourceUID','ResourceUid','ResourceId') 'ResourceUID'
    $rName = Resolve-SqlColumn $resourceColumns @('ResourceName','Name') 'ResourceName'
    $rAccount = Resolve-SqlColumn $resourceColumns @('ResourceNTAccount','ResourceAccount','UserAccount') 'Resource account' -Optional
    $rClaims = Resolve-SqlColumn $resourceColumns @('UserClaimsAccount','ResourceClaimsAccount','ClaimsAccount') 'Claims account' -Optional
    $rEmail = Resolve-SqlColumn $resourceColumns @('ResourceEmailAddress','Email') 'Resource email' -Optional
    $rActive = Resolve-SqlColumn $resourceColumns @('ResourceIsActive','IsActive') 'Resource active' -Optional

    $pUid = Resolve-SqlColumn $projectColumns @('ProjectUID','ProjectUid','ProjectId') 'ProjectUID'
    $pName = Resolve-SqlColumn $projectColumns @('ProjectName','Name') 'ProjectName'
    $pOwnerUid = Resolve-SqlColumn $projectColumns @('ProjectOwnerResourceUID','ProjectOwnerResourceUid','ProjectOwnerId') 'Project Owner UID'
    $pOwnerName = Resolve-SqlColumn $projectColumns @('ProjectOwnerName','OwnerName') 'Project Owner Name' -Optional

    $aProjectUid = Resolve-SqlColumn $assignmentColumns @('ProjectUID','ProjectUid','ProjectId') 'Assignment ProjectUID'
    $aTaskUid = Resolve-SqlColumn $assignmentColumns @('TaskUID','TaskUid','TaskId') 'TaskUID'
    $aAssignmentUid = Resolve-SqlColumn $assignmentColumns @('AssignmentUID','AssignmentUid','AssignmentId') 'AssignmentUID'
    $aResourceUid = Resolve-SqlColumn $assignmentColumns @('ResourceUID','ResourceUid','ResourceId') 'Assignment ResourceUID'
    $aProjectName = Resolve-SqlColumn $assignmentColumns @('ProjectName') 'Assignment ProjectName' -Optional
    $aTaskName = Resolve-SqlColumn $assignmentColumns @('TaskName') 'TaskName' -Optional

    $parameters = @{ UserUid = $ResourceUids[0]; MigrationUid = $ResourceUids[1] }
    $resourceSql = @"
SELECT $(Get-SqlSelectExpression $rUid 'ResourceUID'),
       $(Get-SqlSelectExpression $rName 'ResourceName'),
       $(Get-SqlSelectExpression $rAccount 'ResourceAccount'),
       $(Get-SqlSelectExpression $rClaims 'UserClaimsAccount'),
       $(Get-SqlSelectExpression $rEmail 'ResourceEmailAddress'),
       $(Get-SqlSelectExpression $rActive 'ResourceIsActive')
FROM [$schema].[$($Views.ResourceView)]
WHERE [$rUid] IN (@UserUid, @MigrationUid);
"@
    $projectSql = @"
SELECT $(Get-SqlSelectExpression $pUid 'ProjectUID'),
       $(Get-SqlSelectExpression $pName 'ProjectName'),
       $(Get-SqlSelectExpression $pOwnerUid 'ProjectOwnerResourceUID'),
       $(Get-SqlSelectExpression $pOwnerName 'ProjectOwnerName')
FROM [$schema].[$($Views.ProjectView)];
"@
    $assignmentSql = @"
SELECT $(Get-SqlSelectExpression $aProjectUid 'ProjectUID'),
       $(Get-SqlSelectExpression $aProjectName 'ProjectName'),
       $(Get-SqlSelectExpression $aTaskUid 'TaskUID'),
       $(Get-SqlSelectExpression $aTaskName 'TaskName'),
       $(Get-SqlSelectExpression $aAssignmentUid 'AssignmentUID'),
       $(Get-SqlSelectExpression $aResourceUid 'ResourceUID')
FROM [$schema].[$($Views.AssignmentView)]
WHERE [$aResourceUid] IN (@UserUid, @MigrationUid);
"@

    $resourceRows = Convert-DataRowsToObjects (Invoke-ReadOnlySqlDataTable $Connection $resourceSql $parameters)
    $projectRows = Convert-DataRowsToObjects (Invoke-ReadOnlySqlDataTable $Connection $projectSql)
    $assignmentRows = Convert-DataRowsToObjects (Invoke-ReadOnlySqlDataTable $Connection $assignmentSql $parameters)

    $resources = foreach ($row in $resourceRows) {
        [pscustomobject]@{
            ResourceUID = ConvertTo-CanonicalGuid $row.ResourceUID
            ResourceName = [string]$row.ResourceName
            ResourceAccount = [string]$row.ResourceAccount
            UserClaimsAccount = [string]$row.UserClaimsAccount
            ResourceEmailAddress = [string]$row.ResourceEmailAddress
            ResourceIsActive = $row.ResourceIsActive
        }
    }
    $projects = foreach ($row in $projectRows) {
        [pscustomobject]@{
            ProjectUID = ConvertTo-CanonicalGuid $row.ProjectUID
            ProjectName = [string]$row.ProjectName
            ProjectOwnerName = [string]$row.ProjectOwnerName
            ProjectOwnerResourceUID = ConvertTo-CanonicalGuid $row.ProjectOwnerResourceUID
        }
    }
    $assignments = foreach ($row in $assignmentRows) {
        [pscustomobject]@{
            ProjectUID = ConvertTo-CanonicalGuid $row.ProjectUID
            ProjectName = [string]$row.ProjectName
            TaskUID = ConvertTo-CanonicalGuid $row.TaskUID
            TaskName = [string]$row.TaskName
            AssignmentUID = ConvertTo-CanonicalGuid $row.AssignmentUID
            ResourceUID = ConvertTo-CanonicalGuid $row.ResourceUID
        }
    }

    return [pscustomobject]@{
        Resources = @($resources)
        Projects = @($projects)
        Assignments = @($assignments)
        FieldMap = [ordered]@{
            ResourceUID = $rUid
            ProjectUID = $pUid
            ProjectOwnerResourceUID = $pOwnerUid
            AssignmentResourceUID = $aResourceUid
        }
    }
}

function New-IdentityRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]] $SourceResources,
        [Parameter(Mandatory = $true)][object[]] $TargetResources,
        [Parameter(Mandatory = $true)][hashtable] $Identity
    )

    $definitions = @(
        @{ Classification = 'UTILISATEUR'; UID = ConvertTo-CanonicalGuid $Identity.UserResourceUID },
        @{ Classification = 'RESSOURCE_MIGRATION'; UID = ConvertTo-CanonicalGuid $Identity.MigrationResourceUID }
    )

    foreach ($definition in $definitions) {
        $source = $SourceResources | Where-Object { $_.ResourceUID -eq $definition.UID } | Select-Object -First 1
        $target = $TargetResources | Where-Object { $_.ResourceUID -eq $definition.UID } | Select-Object -First 1
        [pscustomobject][ordered]@{
            Classification = $definition.Classification
            ResourceUID = $definition.UID
            'Nom source' = if ($source) { $source.ResourceName } else { $null }
            'Compte source' = if ($source) { $source.ResourceAccount } else { $null }
            'Claims source' = if ($source) { $source.UserClaimsAccount } else { $null }
            'Courriel source' = if ($source) { $source.ResourceEmailAddress } else { $null }
            'Nom cible' = if ($target) { $target.ResourceName } else { $null }
            'Compte cible' = if ($target) { $target.ResourceAccount } else { $null }
            'Claims cible' = if ($target) { $target.UserClaimsAccount } else { $null }
            'Ressource active' = 'Source:{0}; Cible:{1}' -f (if ($source) { $source.ResourceIsActive } else { 'Introuvable' }), (if ($target) { $target.ResourceIsActive } else { 'Introuvable' })
            Commentaire = if (-not $source -or -not $target) { 'Ressource absente de l’un des inventaires.' } else { 'ResourceUID préservé et utilisé comme clé technique.' }
        }
    }
}

function New-EnvironmentProjectRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Source','Cible')][string] $Side,
        [Parameter(Mandatory = $true)][object] $Inventory,
        [Parameter(Mandatory = $true)][object[]] $Teams,
        [Parameter(Mandatory = $true)][string] $UserResourceUID,
        [Parameter(Mandatory = $true)][string] $MigrationResourceUID,
        [Parameter(Mandatory = $true)][string] $Environment
    )

    $userUid = ConvertTo-CanonicalGuid $UserResourceUID
    $migrationUid = ConvertTo-CanonicalGuid $MigrationResourceUID
    $projectUids = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($project in $Inventory.Projects) {
        if ($project.ProjectOwnerResourceUID -eq $userUid -or $project.ProjectOwnerResourceUID -eq $migrationUid) {
            [void]$projectUids.Add($project.ProjectUID)
        }
    }
    foreach ($assignment in $Inventory.Assignments) { [void]$projectUids.Add($assignment.ProjectUID) }
    foreach ($team in $Teams) { [void]$projectUids.Add($team.ProjectUID) }

    foreach ($projectUid in ($projectUids | Sort-Object)) {
        $project = $Inventory.Projects | Where-Object { $_.ProjectUID -eq $projectUid } | Select-Object -First 1
        $projectAssignments = @($Inventory.Assignments | Where-Object { $_.ProjectUID -eq $projectUid })
        $projectTeam = @($Teams | Where-Object { $_.ProjectUID -eq $projectUid })
        $ownerUser = $null -ne $project -and $project.ProjectOwnerResourceUID -eq $userUid
        $ownerMigration = $null -ne $project -and $project.ProjectOwnerResourceUID -eq $migrationUid
        $teamUser = @($projectTeam | Where-Object { $_.ResourceUID -eq $userUid }).Count -gt 0
        $teamMigration = @($projectTeam | Where-Object { $_.ResourceUID -eq $migrationUid }).Count -gt 0
        $assignmentUserCount = @($projectAssignments | Where-Object { $_.ResourceUID -eq $userUid }).Count
        $assignmentMigrationCount = @($projectAssignments | Where-Object { $_.ResourceUID -eq $migrationUid }).Count
        $name = if ($project) { $project.ProjectName } elseif ($projectAssignments.Count -gt 0) { $projectAssignments[0].ProjectName } elseif ($projectTeam.Count -gt 0) { $projectTeam[0].ProjectName } else { $null }

        $row = [ordered]@{ ProjectUID = $projectUid; ProjectName = $name }
        $row["${Side}_Owner_User"] = ConvertTo-OuiNon $ownerUser
        $row["${Side}_Owner_Migration"] = ConvertTo-OuiNon $ownerMigration
        $row["${Side}_Team_User"] = ConvertTo-OuiNon $teamUser
        $row["${Side}_Team_Migration"] = ConvertTo-OuiNon $teamMigration
        $row["${Side}_Assignment_User"] = ConvertTo-OuiNon ($assignmentUserCount -gt 0)
        $row["${Side}_Assignment_Migration"] = ConvertTo-OuiNon ($assignmentMigrationCount -gt 0)
        $row['Nombre_Assignments_User'] = $assignmentUserCount
        $row['Nombre_Assignments_Migration'] = $assignmentMigrationCount
        $row['Accès_Attendu_User'] = ConvertTo-OuiNon ($ownerUser -or $teamUser -or $assignmentUserCount -gt 0)
        $row['Environnement'] = $Environment
        $row['Commentaire'] = 'Accès attendu basé sur Owner, Project Team ou Assignment; ce n’est pas une permission PWA effective.'
        [pscustomobject]$row
    }
}

function Get-RelationList {
    param([object] $Row, [string] $Side, [string] $IdentitySuffix)
    if ($null -eq $Row) { return '' }
    $relations = New-Object System.Collections.Generic.List[string]
    if ($Row."${Side}_Owner_${IdentitySuffix}" -eq 'Oui') { [void]$relations.Add('OWNER') }
    if ($Row."${Side}_Team_${IdentitySuffix}" -eq 'Oui') { [void]$relations.Add('PROJECT_TEAM') }
    if ($Row."${Side}_Assignment_${IdentitySuffix}" -eq 'Oui') { [void]$relations.Add('ASSIGNMENT') }
    return ($relations -join '|')
}

function Get-RelationDiagnostic {
    param(
        [bool] $SourceUser,
        [bool] $SourceMigration,
        [bool] $TargetUser,
        [bool] $TargetMigration
    )

    if ($SourceMigration) { return 'ANOMALIE_SOURCE' }
    if ($SourceUser -and -not $TargetUser -and $TargetMigration) { return 'À_CORRIGER_UID' }
    if ($SourceUser -and $TargetUser -and $TargetMigration) { return 'DOUBLON_CIBLE' }
    if ($SourceUser -and -not $TargetUser -and -not $TargetMigration) { return 'RELATION_MANQUANTE_CIBLE' }
    if (-not $SourceUser -and -not $SourceMigration -and ($TargetUser -or $TargetMigration)) { return 'RELATION_SUPPLÉMENTAIRE_CIBLE' }
    if ($SourceUser -eq $TargetUser -and $SourceMigration -eq $TargetMigration) { return 'CONFORME' }
    return 'À_INVESTIGUER'
}

function Get-AggregateDiagnostic {
    param([string[]] $Diagnostics)
    $precedence = @('ANOMALIE_SOURCE','À_CORRIGER_UID','DOUBLON_CIBLE','RELATION_MANQUANTE_CIBLE','RELATION_SUPPLÉMENTAIRE_CIBLE','À_INVESTIGUER','CONFORME')
    foreach ($value in $precedence) {
        if ($Diagnostics -contains $value) { return $value }
    }
    return 'À_INVESTIGUER'
}

function New-ReconciliationRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]] $SourceRows,
        [Parameter(Mandatory = $true)][object[]] $TargetRows,
        [Parameter(Mandatory = $true)][object[]] $SourceProjects,
        [Parameter(Mandatory = $true)][object[]] $TargetProjects
    )

    $projectUids = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($row in $SourceRows) { [void]$projectUids.Add($row.ProjectUID) }
    foreach ($row in $TargetRows) { [void]$projectUids.Add($row.ProjectUID) }

    foreach ($projectUid in ($projectUids | Sort-Object)) {
        $source = $SourceRows | Where-Object { $_.ProjectUID -eq $projectUid } | Select-Object -First 1
        $target = $TargetRows | Where-Object { $_.ProjectUID -eq $projectUid } | Select-Object -First 1
        $sourceProject = $SourceProjects | Where-Object { $_.ProjectUID -eq $projectUid } | Select-Object -First 1
        $targetProject = $TargetProjects | Where-Object { $_.ProjectUID -eq $projectUid } | Select-Object -First 1

        $diagnostics = New-Object System.Collections.Generic.List[string]
        $detail = New-Object System.Collections.Generic.List[string]
        foreach ($relation in @(
            @{ Name='OWNER'; Property='Owner' },
            @{ Name='PROJECT_TEAM'; Property='Team' },
            @{ Name='ASSIGNMENT'; Property='Assignment' }
        )) {
            $su = $null -ne $source -and $source."Source_$($relation.Property)_User" -eq 'Oui'
            $sm = $null -ne $source -and $source."Source_$($relation.Property)_Migration" -eq 'Oui'
            $tu = $null -ne $target -and $target."Cible_$($relation.Property)_User" -eq 'Oui'
            $tm = $null -ne $target -and $target."Cible_$($relation.Property)_Migration" -eq 'Oui'
            $diagnostic = Get-RelationDiagnostic $su $sm $tu $tm
            [void]$diagnostics.Add($diagnostic)
            [void]$detail.Add("$($relation.Name):$diagnostic")
        }
        $aggregate = Get-AggregateDiagnostic $diagnostics.ToArray()

        $suggestions = New-Object System.Collections.Generic.List[string]
        if ($detail -match 'OWNER:(À_CORRIGER_UID|DOUBLON_CIBLE|RELATION_MANQUANTE_CIBLE)') { [void]$suggestions.Add('1- Valider/corriger le Project Owner.') }
        if ($detail -match 'PROJECT_TEAM:(À_CORRIGER_UID|DOUBLON_CIBLE|RELATION_MANQUANTE_CIBLE)') { [void]$suggestions.Add('2- Ajouter la bonne ressource au Project Team, publier, valider, puis évaluer le retrait du doublon.') }
        if ($detail -match 'ASSIGNMENT:(À_CORRIGER_UID|DOUBLON_CIBLE|RELATION_MANQUANTE_CIBLE)') { [void]$suggestions.Add('4- Analyser les assignments; aucune correction automatique dans cette révision.') }
        if ($aggregate -eq 'ANOMALIE_SOURCE') { [void]$suggestions.Add('Investigation pilote requise; ne pas corriger sur le seul nom « migration ».') }
        if ($aggregate -eq 'RELATION_SUPPLÉMENTAIRE_CIBLE') { [void]$suggestions.Add('Valider si la relation cible a été ajoutée volontairement après la migration.') }

        [pscustomobject][ordered]@{
            ProjectUID = $projectUid
            ProjectName = if ($source) { $source.ProjectName } elseif ($target) { $target.ProjectName } else { $null }
            ProjectName_Cible = if ($target) { $target.ProjectName } else { $null }
            'ProjectUID_Préservé' = ConvertTo-OuiNon ($null -ne $sourceProject -and $null -ne $targetProject)
            Relation_Source_User = Get-RelationList $source 'Source' 'User'
            Relation_Source_Migration = Get-RelationList $source 'Source' 'Migration'
            Relation_Cible_User = Get-RelationList $target 'Cible' 'User'
            Relation_Cible_Migration = Get-RelationList $target 'Cible' 'Migration'
            Owner_Source_UID = if ($sourceProject) { $sourceProject.ProjectOwnerResourceUID } else { $null }
            Owner_Cible_UID = if ($targetProject) { $targetProject.ProjectOwnerResourceUID } else { $null }
            Team_Source_User = if ($source) { $source.Source_Team_User } else { 'Non' }
            Team_Source_Migration = if ($source) { $source.Source_Team_Migration } else { 'Non' }
            Team_Cible_User = if ($target) { $target.Cible_Team_User } else { 'Non' }
            Team_Cible_Migration = if ($target) { $target.Cible_Team_Migration } else { 'Non' }
            Assignment_Source_User = if ($source) { $source.Source_Assignment_User } else { 'Non' }
            Assignment_Source_Migration = if ($source) { $source.Source_Assignment_Migration } else { 'Non' }
            Assignment_Cible_User = if ($target) { $target.Cible_Assignment_User } else { 'Non' }
            Assignment_Cible_Migration = if ($target) { $target.Cible_Assignment_Migration } else { 'Non' }
            Diagnostic = $aggregate
            'Diagnostic_Détail' = $detail -join '; '
            'Correction_Suggérée' = $suggestions -join ' '
            Correction_Requise = if ($aggregate -eq 'CONFORME') { 'Non' } elseif ($aggregate -eq 'ANOMALIE_SOURCE' -or $aggregate -eq 'À_INVESTIGUER') { 'À valider' } else { 'Oui' }
            'Correction_Validée' = ''
            'Correction_Effectuée' = ''
            Validation_Post_Correction = ''
            Commentaire_Pilote = ''
        }
    }
}

function New-AssignmentDetailRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]] $SourceAssignments,
        [Parameter(Mandatory = $true)][object[]] $TargetAssignments,
        [Parameter(Mandatory = $true)][string] $UserResourceUID,
        [Parameter(Mandatory = $true)][string] $MigrationResourceUID,
        [Parameter(Mandatory = $true)][string] $TargetEnvironment
    )

    $userUid = ConvertTo-CanonicalGuid $UserResourceUID
    $migrationUid = ConvertTo-CanonicalGuid $MigrationResourceUID
    foreach ($set in @(
        @{ Environment='SOURCE_PROJECT_ONLINE'; Rows=$SourceAssignments },
        @{ Environment=$TargetEnvironment; Rows=$TargetAssignments }
    )) {
        foreach ($row in $set.Rows) {
            [pscustomobject][ordered]@{
                Environnement = $set.Environment
                ProjectUID = $row.ProjectUID
                ProjectName = $row.ProjectName
                TaskUID = $row.TaskUID
                TaskName = $row.TaskName
                AssignmentUID = $row.AssignmentUID
                ResourceUID = $row.ResourceUID
                Classification = if ($row.ResourceUID -eq $userUid) { 'UTILISATEUR' } elseif ($row.ResourceUID -eq $migrationUid) { 'RESSOURCE_MIGRATION' } else { 'AUTRE' }
                Action = if ($row.ResourceUID -eq $migrationUid) { 'À analyser; aucune correction automatique.' } else { 'Aucune action déduite.' }
            }
        }
    }
}

function Export-ReconciliationWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Path,
        [Parameter(Mandatory = $true)][object[]] $IdentityRows,
        [Parameter(Mandatory = $true)][object[]] $SourceRows,
        [Parameter(Mandatory = $true)][object[]] $TargetRows,
        [Parameter(Mandatory = $true)][object[]] $ReconciliationRows,
        [Parameter(Mandatory = $true)][object[]] $AssignmentRows,
        [Parameter(Mandatory = $true)][object[]] $ControlRows
    )

    $module = Get-Module -ListAvailable -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
    if ($null -eq $module) {
        throw 'Module ImportExcel absent. Installer avec : Install-Module ImportExcel -Scope CurrentUser -MinimumVersion 7.8.10'
    }
    Import-Module ImportExcel -MinimumVersion 7.8.10 -ErrorAction Stop

    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent)) { [void](New-Item -ItemType Directory -Path $parent -Force) }
    if (Test-Path -LiteralPath $Path) { throw "Le rapport existe déjà : $Path" }

    $exports = @(
        @{ Name='01_Identités'; Table='Identites'; Rows=$IdentityRows },
        @{ Name='02_Source_ProjectOnline'; Table='SourceProjectOnline'; Rows=$SourceRows },
        @{ Name='03_Cible_PSSE'; Table='CiblePSSE'; Rows=$TargetRows },
        @{ Name='04_Réconciliation'; Table='Reconciliation'; Rows=$ReconciliationRows },
        @{ Name='05_Assignments_Détail'; Table='AssignmentsDetail'; Rows=$AssignmentRows },
        @{ Name='06_Contrôle_Exécution'; Table='ControleExecution'; Rows=$ControlRows }
    )

    $first = $true
    foreach ($export in $exports) {
        $rows = @($export.Rows)
        if ($rows.Count -eq 0) { $rows = @([pscustomobject]@{ Information = 'Aucune donnée.' }) }
        $parameters = @{
            Path = $Path
            WorksheetName = $export.Name
            TableName = $export.Table
            AutoSize = $true
            AutoFilter = $true
            FreezeTopRow = $true
            BoldTopRow = $true
            ClearSheet = $true
        }
        if (-not $first) { $parameters['Append'] = $true }
        $rows | Export-Excel @parameters
        $first = $false
    }

    $package = Open-ExcelPackage -Path $Path
    try {
        $sheet = $package.Workbook.Worksheets['04_Réconciliation']
        if ($null -ne $sheet -and $sheet.Dimension.Rows -gt 1) {
            $diagnosticColumn = $null
            for ($column = 1; $column -le $sheet.Dimension.Columns; $column++) {
                if ($sheet.Cells[1,$column].Text -eq 'Diagnostic') { $diagnosticColumn = $column; break }
            }
            if ($null -ne $diagnosticColumn) {
                for ($row = 2; $row -le $sheet.Dimension.Rows; $row++) {
                    $cell = $sheet.Cells[$row,$diagnosticColumn]
                    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    switch ($cell.Text) {
                        'CONFORME' { $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen) }
                        'À_CORRIGER_UID' { $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightCoral) }
                        'DOUBLON_CIBLE' { $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Orange) }
                        default { $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Khaki) }
                    }
                }
            }
        }
        Close-ExcelPackage $package
    }
    catch {
        if ($null -ne $package) { Close-ExcelPackage $package -NoSave }
        throw
    }
}

Export-ModuleMember -Function @(
    'ConvertTo-CanonicalGuid',
    'Write-ReconciliationLog',
    'Import-ProjectCsomAssemblies',
    'New-ProjectOnlineConnection',
    'New-PsseProjectContext',
    'Get-ProjectOnlineReportingInventory',
    'Get-PublishedProjectTeamInventory',
    'New-ReadOnlySqlConnection',
    'Get-PsseReportingInventory',
    'New-IdentityRows',
    'New-EnvironmentProjectRows',
    'New-ReconciliationRows',
    'New-AssignmentDetailRows',
    'Export-ReconciliationWorkbook'
)
