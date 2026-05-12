# PWA Migration

Scripts PowerShell pour traiter prudemment des ressources d'entreprise Project Online avant migration.

## Contenu principal

- `Invoke-OrphanEnterpriseResourceCleanup.ps1` : lit un fichier Excel de mapping et traite les ressources existantes par `UID`.
- `Common.ps1` : fonctions communes d'authentification et de connexion Project Online issues du package Microsoft.
- `ExportProjectUserContent.ps1` et `Invoke-RedactProjectUser.ps1` : scripts Microsoft de reference pour export/redaction Project Online.

## Donnees exclues du depot

Les fichiers Excel, journaux et exports sont exclus par `.gitignore`, car ils peuvent contenir des noms, UPN, comptes ou donnees de migration.

## Format Excel attendu

Le fichier Excel est la liste de traitement. Les lignes valides doivent contenir au minimum :

- `UID`
- `Source Name`
- `New Name`

Les lignes vides ou avec `UID` invalide sont ignorees.

## Usage

Executer avec Windows PowerShell 5.1.

Validation locale du fichier :

```powershell
.\Invoke-OrphanEnterpriseResourceCleanup.ps1 `
  -PwaUrl "https://sqi365.sharepoint.com/sites/pwa/" `
  -InputXlsxPath ".\CorrectionResourceMappingTestSQI.xlsx" `
  -ValidateInputOnly
```

Simulation :

```powershell
.\Invoke-OrphanEnterpriseResourceCleanup.ps1 `
  -PwaUrl "https://sqi365.sharepoint.com/sites/pwa/" `
  -InputXlsxPath ".\CorrectionResourceMappingTestSQI.xlsx" `
  -WhatIf
```

Execution :

```powershell
.\Invoke-OrphanEnterpriseResourceCleanup.ps1 `
  -PwaUrl "https://sqi365.sharepoint.com/sites/pwa/" `
  -InputXlsxPath ".\CorrectionResourceMappingTestSQI.xlsx" `
  -ForceCheckInBeforeUpdate `
  -ForceCheckInAfterUpdate `
  -StopOnFirstError
```

## Limite importante

La dissociation directe du `User Logon Account` par CSOM n'est pas confirmee par source Microsoft officielle dans ce projet. Le script la bloque par defaut.
