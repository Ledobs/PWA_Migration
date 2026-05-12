# Orphan Enterprise Resource Cleanup - Package client

Ce dossier contient uniquement les fichiers necessaires pour executer le script de nettoyage des ressources Project Online.

## Pre-requis

- Windows PowerShell 5.1
- Acces administrateur/Project Online suffisant sur le PWA cible
- Un fichier Excel de mapping fourni par le client

## Fichier Excel attendu

Le fichier doit contenir au minimum les colonnes suivantes :

- UID
- Source Name
- New Name

Les lignes vides ou avec UID invalide sont ignorees.

Placez le fichier Excel dans le dossier `input`.

## Validation locale

```powershell
.\Invoke-OrphanEnterpriseResourceCleanup.ps1 `
  -PwaUrl "https://tenant.sharepoint.com/sites/pwa/" `
  -InputXlsxPath ".\input\CorrectionResourceMapping.xlsx" `
  -SharePointClientRuntimeDllPath ".\lib\Microsoft.SharePoint.Client.Runtime.dll" `
  -SharePointClientDllPath ".\lib\Microsoft.SharePoint.Client.dll" `
  -ProjectServerClientDllPath ".\lib\Microsoft.ProjectServer.Client.dll" `
  -ValidateInputOnly
```

## Simulation

```powershell
.\Invoke-OrphanEnterpriseResourceCleanup.ps1 `
  -PwaUrl "https://tenant.sharepoint.com/sites/pwa/" `
  -InputXlsxPath ".\input\CorrectionResourceMapping.xlsx" `
  -SharePointClientRuntimeDllPath ".\lib\Microsoft.SharePoint.Client.Runtime.dll" `
  -SharePointClientDllPath ".\lib\Microsoft.SharePoint.Client.dll" `
  -ProjectServerClientDllPath ".\lib\Microsoft.ProjectServer.Client.dll" `
  -WhatIf
```

## Execution reelle

```powershell
.\Invoke-OrphanEnterpriseResourceCleanup.ps1 `
  -PwaUrl "https://tenant.sharepoint.com/sites/pwa/" `
  -InputXlsxPath ".\input\CorrectionResourceMapping.xlsx" `
  -SharePointClientRuntimeDllPath ".\lib\Microsoft.SharePoint.Client.Runtime.dll" `
  -SharePointClientDllPath ".\lib\Microsoft.SharePoint.Client.dll" `
  -ProjectServerClientDllPath ".\lib\Microsoft.ProjectServer.Client.dll" `
  -ForceCheckInBeforeUpdate `
  -ForceCheckInAfterUpdate `
  -StopOnFirstError
```

## Notes importantes

- Le script ne supprime aucune ressource.
- Le script ne recree aucune ressource.
- Le script traite les ressources par UID.
- La dissociation directe du User Logon Account par CSOM n'est pas confirmee officiellement et reste desactivee par defaut.
- Les logs sont ecrits dans `logs`.
