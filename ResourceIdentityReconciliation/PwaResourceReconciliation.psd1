@{
    #================================================================================================
    # Nom du fichier : PwaResourceReconciliation.psd1
    # Auteur         : François Breton
    # Date           : 2026-08-06
    # Révision       : 1.0.1
    # Objet          : Manifeste du module de réconciliation des identités de ressources PWA.
    #================================================================================================

    RootModule = 'PwaResourceReconciliation.psm1'
    ModuleVersion = '1.0.1'
    GUID = '4d62bcbf-f071-4e09-9f4a-8d8911b8bb6a'
    Author = 'François Breton'
    CompanyName = 'Dataligence.ca'
    Copyright = '(c) 2026 François Breton. Tous droits réservés.'
    Description = 'Inventaire en lecture seule des relations de ressources entre Project Online et Project Server SE.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
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
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('ProjectOnline','ProjectServer','PSSE','Migration','Inventory','PowerShell')
        }
    }
}
