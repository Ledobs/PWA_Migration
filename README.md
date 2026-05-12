# PWA Migration

Scripts PowerShell pour analyser et traiter prudemment des ressources d'entreprise Project Online avant migration.

## Contenu principal

- `Invoke-OrphanEnterpriseResourceCleanup.ps1` : script pilote ciblant une ressource existante par UID pour la rendre inactive et appliquer un nom d'archive.
- `Common.ps1` : fonctions communes d'authentification et de connexion Project Online issues du package Microsoft.
- `ExportProjectUserContent.ps1` et `Invoke-RedactProjectUser.ps1` : scripts Microsoft de reference pour export/redaction Project Online.

## Donnees exclues du depot

Les fichiers Excel, journaux et exports sont exclus par `.gitignore`, car ils peuvent contenir des noms, UPN, comptes ou donnees de migration.

## Usage pilote

Executer avec Windows PowerShell 5.1.

```powershell
.\Invoke-OrphanEnterpriseResourceCleanup.ps1 `
  -PwaUrl "https://idexia365.sharepoint.com/sites/pwa/" `
  -InputXlsxPath ".\CorrectionResourceMappingTest.xlsx" `
  -TargetResourceUid "08aa95c4-80eb-f011-9aad-00155dc06909" `
  -ProjectServerClientDllPath "C:\Users\FrançoisBreton\.nuget\packages\microsoft.sharepointonline.csom\16.1.26615.12013\lib\net45\Microsoft.ProjectServer.Client.dll" `
  -SharePointClientRuntimeDllPath "C:\Users\FrançoisBreton\.nuget\packages\microsoft.sharepointonline.csom\16.1.26615.12013\lib\net45\Microsoft.SharePoint.Client.Runtime.dll" `
  -SharePointClientDllPath "C:\Users\FrançoisBreton\.nuget\packages\microsoft.sharepointonline.csom\16.1.26615.12013\lib\net45\Microsoft.SharePoint.Client.dll" `
  -WhatIf
```

## Limite importante

La dissociation directe du `User Logon Account` par CSOM n'est pas confirmee par source Microsoft officielle dans ce projet. Le script la bloque par defaut.
