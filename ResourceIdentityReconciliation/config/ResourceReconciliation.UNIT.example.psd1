@{
    #================================================================================================
    # Nom du fichier : ResourceReconciliation.UNIT.example.psd1
    # Auteur         : François Breton
    # Date           : 2026-08-06
    # Révision       : 1.0.1
    # Objet          : Configuration générique d’exemple pour un inventaire dans PSSE UNIT.
    #================================================================================================

    # Copier ce fichier sous ResourceReconciliation.UNIT.psd1, puis remplacer toutes les valeurs
    # d’exemple. Le fichier sans le suffixe .example est exclu du dépôt par .gitignore.
    Source = @{
        PwaUrl = 'https://contoso.sharepoint.com/sites/pwa/'
        ProjectDataUrl = 'https://contoso.sharepoint.com/sites/pwa/_api/ProjectData/'
        Region = 'Default'
    }

    Target = @{
        # Utiliser UNIT jusqu’à ce que le rapport soit validé par les pilotes.
        PwaUrl = 'https://pwa-unit.contoso.local/'
        Environment = 'UNIT'
    }

    Identity = @{
        DisplayName = 'Example; Alice'
        UserResourceUID = '11111111-1111-1111-1111-111111111111'
        MigrationResourceUID = '22222222-2222-2222-2222-222222222222'
        SourceUserUPN = 'alice.example@contoso.com'
        TargetWindowsAccount = 'CONTOSO\aexample'
    }

    Csom = @{
        # Le chemin de Common.ps1 est résolu relativement à ce fichier de configuration.
        CommonScriptPath = '..\..\Common.ps1'
        SharePointClientRuntimeDllPath = 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
        SharePointClientDllPath = 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
        ProjectServerClientDllPath = 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.ProjectServer.Client.dll'
    }

    Sql = @{
        Server = 'SQL-SERVER-UNIT'
        Database = 'ProjectWebApp_UNIT'
        Encrypt = $true
        TrustServerCertificate = $true
        ConnectTimeout = 30
        Views = @{
            Schema = 'dbo'
            ResourceView = 'MSP_EpmResource_UserView'
            ProjectView = 'MSP_EpmProject_UserView'
            AssignmentView = 'MSP_EpmAssignment_UserView'
        }
    }

    Output = @{
        OutputPath = 'C:\PWA\Reconciliation'
    }
}
