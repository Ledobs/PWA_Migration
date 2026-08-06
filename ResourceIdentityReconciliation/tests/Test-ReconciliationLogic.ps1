#requires -version 5.1
<#
====================================================================================================
 Nom du fichier : Test-ReconciliationLogic.ps1
 Auteur         : François Breton
 Date           : 2026-08-06
 Révision       : 1.0.1
 Objet          : Tests autonomes des diagnostics, sans connexion à Project Online, PSSE ou SQL.
====================================================================================================
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'PwaResourceReconciliation.psd1') -Force

$projectUid = '11111111-1111-1111-1111-111111111111'
$userUid = '22222222-2222-2222-2222-222222222222'
$migrationUid = '33333333-3333-3333-3333-333333333333'

function New-TestEnvironmentRow {
    param(
        [ValidateSet('Source','Cible')][string] $Side,
        [bool] $User,
        [bool] $Migration
    )
    $row = [ordered]@{ ProjectUID=$projectUid; ProjectName='Projet test' }
    foreach ($relation in @('Owner','Team','Assignment')) {
        $row["${Side}_${relation}_User"] = if ($User) { 'Oui' } else { 'Non' }
        $row["${Side}_${relation}_Migration"] = if ($Migration) { 'Oui' } else { 'Non' }
    }
    [pscustomobject]$row
}

function Assert-Diagnostic {
    param(
        [string] $Name,
        [bool] $SourceUser,
        [bool] $SourceMigration,
        [bool] $TargetUser,
        [bool] $TargetMigration,
        [string] $Expected
    )
    $source = New-TestEnvironmentRow Source $SourceUser $SourceMigration
    $target = New-TestEnvironmentRow Cible $TargetUser $TargetMigration
    $projects = @([pscustomobject]@{
        ProjectUID=$projectUid
        ProjectName='Projet test'
        ProjectOwnerResourceUID=if ($SourceUser) { $userUid } elseif ($SourceMigration) { $migrationUid } else { $null }
    })
    $result = @(New-ReconciliationRows -SourceRows @($source) -TargetRows @($target) -SourceProjects $projects -TargetProjects $projects)
    if ($result.Count -ne 1 -or $result[0].Diagnostic -ne $Expected) {
        throw "ÉCHEC $Name : attendu '$Expected', obtenu '$($result[0].Diagnostic)'."
    }
    Write-Host "RÉUSSI $Name : $Expected"
}

Assert-Diagnostic 'Situation conforme' $true $false $true $false 'CONFORME'
Assert-Diagnostic 'Mauvais UID cible' $true $false $false $true 'À_CORRIGER_UID'
Assert-Diagnostic 'Doublon cible' $true $false $true $true 'DOUBLON_CIBLE'
Assert-Diagnostic 'Anomalie source' $false $true $false $true 'ANOMALIE_SOURCE'
Assert-Diagnostic 'Relation manquante cible' $true $false $false $false 'RELATION_MANQUANTE_CIBLE'
Assert-Diagnostic 'Relation supplémentaire cible' $false $false $true $false 'RELATION_SUPPLÉMENTAIRE_CIBLE'

Write-Host 'Tous les diagnostics de base sont conformes.'
