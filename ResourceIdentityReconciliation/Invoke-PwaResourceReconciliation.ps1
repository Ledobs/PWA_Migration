#requires -version 5.1
<#
====================================================================================================
 Nom du fichier : Invoke-PwaResourceReconciliation.ps1
 Auteur         : François Breton
 Date           : 2026-08-06
 Révision       : 1.0.1
 Objet          : Inventorier et comparer les relations de deux ResourceUID entre Project Online et PSSE.
 Mode           : INVENTORY seulement - aucune écriture dans Project Online, PSSE ou SQL Server.
====================================================================================================

.SYNOPSIS
Produit un rapport Excel de réconciliation pour les relations OWNER, PROJECT_TEAM et ASSIGNMENT.

.DESCRIPTION
Le script lit Project Online par ProjectData et CSOM, puis PSSE par les vues SQL de reporting et
CSOM. ResourceUID est toujours la clé de rapprochement des ressources. ProjectUID est la clé des
projets lorsque sa préservation est confirmée dans les deux inventaires.

.EXAMPLE
.\Invoke-PwaResourceReconciliation.ps1 `
  -ConfigPath .\config\ResourceReconciliation.UNIT.psd1
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $ConfigPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Inventory','Validate','Fix')]
    [string] $Mode = 'Inventory',

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential] $TargetWindowsCredential,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential] $TargetSqlCredential
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($Mode -ne 'Inventory') {
    throw "Le mode '$Mode' est réservé à une révision future. Cette version n’autorise que Inventory."
}

$resolvedConfig = (Resolve-Path -LiteralPath $ConfigPath -ErrorAction Stop).ProviderPath
$config = Import-PowerShellDataFile -LiteralPath $resolvedConfig
$modulePath = Join-Path $PSScriptRoot 'PwaResourceReconciliation.psd1'
Import-Module $modulePath -Force

foreach ($requiredSection in @('Source','Target','Identity','Csom','Sql','Output')) {
    if (-not $config.ContainsKey($requiredSection)) {
        throw "Section de configuration obligatoire absente : $requiredSection"
    }
}

$userUid = ConvertTo-CanonicalGuid $config.Identity.UserResourceUID
$migrationUid = ConvertTo-CanonicalGuid $config.Identity.MigrationResourceUID
if ($userUid -eq $migrationUid) {
    throw 'UserResourceUID et MigrationResourceUID doivent être différents.'
}

function Resolve-ConfigurationPath {
    param([Parameter(Mandatory = $true)][string] $Value)
    if ([System.IO.Path]::IsPathRooted($Value)) { return $Value }
    return [System.IO.Path]::GetFullPath((Join-Path (Split-Path -Parent $resolvedConfig) $Value))
}

$config.Csom.CommonScriptPath = Resolve-ConfigurationPath $config.Csom.CommonScriptPath
$config.Csom.SharePointClientRuntimeDllPath = Resolve-ConfigurationPath $config.Csom.SharePointClientRuntimeDllPath
$config.Csom.SharePointClientDllPath = Resolve-ConfigurationPath $config.Csom.SharePointClientDllPath
$config.Csom.ProjectServerClientDllPath = Resolve-ConfigurationPath $config.Csom.ProjectServerClientDllPath
$outputFolder = Resolve-ConfigurationPath $config.Output.OutputPath

if (-not (Test-Path -LiteralPath $outputFolder)) {
    [void](New-Item -ItemType Directory -Path $outputFolder -Force)
}
$runId = Get-Date -Format 'yyyyMMdd-HHmmss'
$logFolder = Join-Path $outputFolder 'logs'
if (-not (Test-Path -LiteralPath $logFolder)) {
    [void](New-Item -ItemType Directory -Path $logFolder -Force)
}
$logPath = Join-Path $logFolder "PwaResourceReconciliation-$runId.log"
$reportPath = Join-Path $outputFolder "PwaResourceReconciliation-$($config.Target.Environment)-$runId.xlsx"

$sourceConnection = $null
$targetContext = $null
$sqlConnection = $null
$startedAt = Get-Date

try {
    Write-ReconciliationLog $logPath INFO "Début de l’inventaire. Révision 1.0.1. Mode lecture seule."
    Write-ReconciliationLog $logPath INFO "ResourceUID utilisateur : $userUid"
    Write-ReconciliationLog $logPath INFO "ResourceUID migration   : $migrationUid"

    Import-ProjectCsomAssemblies -Csom $config.Csom

    Write-ReconciliationLog $logPath INFO "Connexion à Project Online : $($config.Source.PwaUrl)"
    $sourceConnection = New-ProjectOnlineConnection `
        -PwaUrl $config.Source.PwaUrl `
        -CommonScriptPath $config.Csom.CommonScriptPath `
        -Region $config.Source.Region

    Write-ReconciliationLog $logPath INFO 'Lecture de Resources, Projects et Assignments dans ProjectData.'
    $sourceInventory = Get-ProjectOnlineReportingInventory `
        -ProjectDataUrl $config.Source.ProjectDataUrl `
        -WebSession $sourceConnection.WebSession `
        -ResourceUids @($userUid, $migrationUid)

    Write-ReconciliationLog $logPath INFO 'Lecture des véritables Project Teams dans Project Online par CSOM.'
    $sourceTeams = @(Get-PublishedProjectTeamInventory `
        -Context $sourceConnection.ProjectContext `
        -ResourceUids @($userUid, $migrationUid) `
        -Environment 'SOURCE_PROJECT_ONLINE' `
        -LogPath $logPath)

    Write-ReconciliationLog $logPath INFO "Connexion SQL en lecture seule : $($config.Sql.Server) / $($config.Sql.Database)"
    $sqlConnection = New-ReadOnlySqlConnection -Sql $config.Sql -Credential $TargetSqlCredential
    $sqlConnection.Open()
    $targetInventory = Get-PsseReportingInventory `
        -Connection $sqlConnection `
        -Views $config.Sql.Views `
        -ResourceUids @($userUid, $migrationUid)

    Write-ReconciliationLog $logPath INFO "Lecture des Project Teams dans PSSE $($config.Target.Environment) par CSOM."
    $targetContext = New-PsseProjectContext -PwaUrl $config.Target.PwaUrl -Credential $TargetWindowsCredential
    $targetTeams = @(Get-PublishedProjectTeamInventory `
        -Context $targetContext `
        -ResourceUids @($userUid, $migrationUid) `
        -Environment $config.Target.Environment `
        -LogPath $logPath)

    $identityRows = @(New-IdentityRows `
        -SourceResources $sourceInventory.Resources `
        -TargetResources $targetInventory.Resources `
        -Identity $config.Identity)

    $sourceRows = @(New-EnvironmentProjectRows `
        -Side Source `
        -Inventory $sourceInventory `
        -Teams $sourceTeams `
        -UserResourceUID $userUid `
        -MigrationResourceUID $migrationUid `
        -Environment 'SOURCE_PROJECT_ONLINE')

    $targetRows = @(New-EnvironmentProjectRows `
        -Side Cible `
        -Inventory $targetInventory `
        -Teams $targetTeams `
        -UserResourceUID $userUid `
        -MigrationResourceUID $migrationUid `
        -Environment $config.Target.Environment)

    $reconciliationRows = @(New-ReconciliationRows `
        -SourceRows $sourceRows `
        -TargetRows $targetRows `
        -SourceProjects $sourceInventory.Projects `
        -TargetProjects $targetInventory.Projects)

    $assignmentRows = @(New-AssignmentDetailRows `
        -SourceAssignments $sourceInventory.Assignments `
        -TargetAssignments $targetInventory.Assignments `
        -UserResourceUID $userUid `
        -MigrationResourceUID $migrationUid `
        -TargetEnvironment $config.Target.Environment)

    $sourceRelevantUids = @($sourceRows | Select-Object -ExpandProperty ProjectUID -Unique)
    $targetProjectUids = @($targetInventory.Projects | Select-Object -ExpandProperty ProjectUID -Unique)
    $preservedCount = @($sourceRelevantUids | Where-Object { $targetProjectUids -contains $_ }).Count
    $preservationRate = if ($sourceRelevantUids.Count -eq 0) { 'S.O.' } else { '{0:P1}' -f ($preservedCount / [double]$sourceRelevantUids.Count) }

    $controlRows = @(
        [pscustomobject]@{ 'Contrôle'='Mode'; Valeur='Inventory'; 'Résultat'='Lecture seule' },
        [pscustomobject]@{ 'Contrôle'='Début'; Valeur=$startedAt.ToString('s'); 'Résultat'='Information' },
        [pscustomobject]@{ 'Contrôle'='Source PWA'; Valeur=$config.Source.PwaUrl; 'Résultat'='Lu' },
        [pscustomobject]@{ 'Contrôle'='Cible PWA'; Valeur=$config.Target.PwaUrl; 'Résultat'="Lu - $($config.Target.Environment)" },
        [pscustomobject]@{ 'Contrôle'='ResourceUID'; Valeur="$userUid / $migrationUid"; 'Résultat'='Clés techniques des ressources' },
        [pscustomobject]@{ 'Contrôle'='ProjectUID préservés'; Valeur="$preservedCount / $($sourceRelevantUids.Count) ($preservationRate)"; 'Résultat'='Comparer les UID communs; aucun rapprochement par nom' },
        [pscustomobject]@{ 'Contrôle'='Champs ProjectData'; Valeur=($sourceInventory.FieldMap | ConvertTo-Json -Compress); 'Résultat'='Résolus dynamiquement' },
        [pscustomobject]@{ 'Contrôle'='Champs SQL'; Valeur=($targetInventory.FieldMap | ConvertTo-Json -Compress); 'Résultat'='Résolus dynamiquement' },
        [pscustomobject]@{ 'Contrôle'='Projets source pertinents'; Valeur=$sourceRows.Count; 'Résultat'='Owner, Team ou Assignment pour un UID suivi' },
        [pscustomobject]@{ 'Contrôle'='Projets cible pertinents'; Valeur=$targetRows.Count; 'Résultat'='Owner, Team ou Assignment pour un UID suivi' },
        [pscustomobject]@{ 'Contrôle'='Permissions PWA'; Valeur='Non calculées'; 'Résultat'='Accès_Attendu seulement; groupes/catégories/RBS hors périmètre v1' },
        [pscustomobject]@{ 'Contrôle'='Corrections'; Valeur='Aucune'; 'Résultat'='Validate et Fix non implémentés' }
    )

    Write-ReconciliationLog $logPath INFO "Production du rapport Excel : $reportPath"
    Export-ReconciliationWorkbook `
        -Path $reportPath `
        -IdentityRows $identityRows `
        -SourceRows $sourceRows `
        -TargetRows $targetRows `
        -ReconciliationRows $reconciliationRows `
        -AssignmentRows $assignmentRows `
        -ControlRows $controlRows

    Write-ReconciliationLog $logPath INFO "Inventaire terminé. Projets réconciliés : $($reconciliationRows.Count)."
    Write-ReconciliationLog $logPath INFO "Rapport : $reportPath"
}
catch {
    Write-ReconciliationLog $logPath ERREUR $_.Exception.Message
    throw
}
finally {
    if ($null -ne $sqlConnection) { $sqlConnection.Dispose() }
    if ($null -ne $targetContext) { $targetContext.Dispose() }
    if ($null -ne $sourceConnection -and $null -ne $sourceConnection.ProjectContext) { $sourceConnection.ProjectContext.Dispose() }
}

[pscustomobject]@{
    Mode = $Mode
    ReportPath = $reportPath
    LogPath = $logPath
    TargetEnvironment = $config.Target.Environment
}
