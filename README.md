# PWA Migration

Scripts PowerShell pour traiter prudemment des ressources d'entreprise Project Online avant migration.

## Contenu principal

- `ResourceIdentityReconciliation/` : inventaire en lecture seule et rapport Excel de comparaison des
  relations `OWNER`, `PROJECT_TEAM` et `ASSIGNMENT` entre Project Online et Project Server SE.
- `Invoke-OrphanEnterpriseResourceCleanup.ps1` : lit un fichier Excel de mapping et traite les ressources existantes par `UID`.
- `Common.ps1` : fonctions communes d'authentification et de connexion Project Online issues du package Microsoft.
- `ExportProjectUserContent.ps1` et `Invoke-RedactProjectUser.ps1` : scripts Microsoft de reference pour export/redaction Project Online.

## Donnees exclues du depot

Les fichiers Excel, journaux, exports et configurations locales sont exclus par `.gitignore`, car ils
peuvent contenir des noms, UPN, comptes, URL ou donnees de migration.

## Réconciliation des identités

Le mode Inventory est documenté dans
[`ResourceIdentityReconciliation/README.md`](ResourceIdentityReconciliation/README.md). Son exemple
de configuration est générique; les paramètres propres au client demeurent dans un fichier local
ignoré par Git.

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
  -PwaUrl "https://tenant.sharepoint.com/sites/pwa/" `
  -InputXlsxPath ".\CorrectionResourceMappingTest.xlsx" `
  -ValidateInputOnly
```

Simulation :

```powershell
.\Invoke-OrphanEnterpriseResourceCleanup.ps1 `
  -PwaUrl "https://tenant.sharepoint.com/sites/pwa/" `
  -InputXlsxPath ".\CorrectionResourceMappingTest.xlsx" `
  -WhatIf
```

Execution :

```powershell
.\Invoke-OrphanEnterpriseResourceCleanup.ps1 `
  -PwaUrl "https://tenant.sharepoint.com/sites/pwa/" `
  -InputXlsxPath ".\CorrectionResourceMappingTest.xlsx" `
  -ForceCheckInBeforeUpdate `
  -ForceCheckInAfterUpdate `
  -StopOnFirstError
```

## Limite importante

La dissociation directe du `User Logon Account` par CSOM n'est pas confirmee par source Microsoft officielle dans ce projet. Le script la bloque par defaut.
