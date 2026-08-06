# Réconciliation des identités de ressources PWA

| Contrôle documentaire | Valeur |
|---|---|
| Auteur | François Breton |
| Date | 2026-08-06 |
| Révision | 1.0.1 |
| Mode disponible | `Inventory` seulement |

Cette solution inventorie les relations de deux ressources d’entreprise avant et après une migration
Project Online vers Project Server Subscription Edition (PSSE). Elle compare les relations `OWNER`,
`PROJECT_TEAM` et `ASSIGNMENT`, puis produit un rapport Excel destiné aux pilotes.

La version 1.0.1 n’effectue aucune correction. Elle ne contient aucune requête SQL d’écriture et bloque
explicitement les modes `Validate` et `Fix`.

## Principes de conception

- `ResourceUID` est la clé technique des ressources. Le nom, le UPN, le compte Windows et le Claims
  Account servent uniquement à expliquer l’identité.
- `ProjectUID` est la clé des projets lorsque sa présence est confirmée dans les deux environnements.
- Le vrai `Project Team` provient de `PublishedProject.ProjectResources` par CSOM. Une absence
  d’affectation ne signifie donc pas une absence du Team.
- Project Online est lu par `ProjectData` pour le reporting et par CSOM pour le Team.
- PSSE est lu par les vues SQL de reporting et par CSOM pour le Team.
- `Accès_Attendu_User` n’est pas une permission PWA effective. Le calcul des groupes, catégories,
  règles dynamiques et RBS demeure hors périmètre de la révision 1.0.0.

Microsoft documente `ProjectContext` comme point d’entrée CSOM, la propriété
[`PublishedProject.ProjectResources`](https://learn.microsoft.com/en-us/dotnet/api/microsoft.projectserver.client.publishedproject.projectresources?view=sharepoint-csom)
pour les ressources d’un projet et l’utilisation des tables/vues de reporting pour les lectures
on-premises. Les objets `draft`, `pub` et `ver` ne doivent pas servir au reporting ou aux modifications :
[`What's new for IT pros in Project Server 2013`](https://learn.microsoft.com/en-us/project/what-s-new-for-it-pros-in-project-server-2013).

## Contenu

| Fichier | Rôle |
|---|---|
| `Invoke-PwaResourceReconciliation.ps1` | Orchestration du mode Inventory |
| `PwaResourceReconciliation.psd1` / `.psm1` | Manifeste, connecteurs de lecture, normalisation, diagnostics et export Excel |
| `config/ResourceReconciliation.UNIT.example.psd1` | Exemple générique sans donnée client |
| `docs/Modele-de-donnees-et-diagnostics.md` | Modèle normalisé et règles de diagnostic |
| `docs/Validate-et-Fix.md` | Design des futures phases, sans code de correction |
| `tests/Test-ReconciliationLogic.ps1` | Tests unitaires autonomes de la logique de diagnostic |

## Prérequis

- Windows PowerShell 5.1 sur un poste ou serveur pouvant joindre simultanément :
  - Project Online;
  - le PWA PSSE UNIT;
  - SQL Server de l’instance PSSE UNIT.
- Un compte ayant les droits de lecture nécessaires dans les deux PWA et les vues SQL de reporting.
- Les DLL CSOM SharePoint et Project Server. Les chemins par défaut de la configuration pointent vers
  le dossier `ISAPI` d’une installation SharePoint/Project Server.
- Le fichier `Common.ps1` déjà présent à la racine du dépôt pour l’authentification Project Online.
- Le module PowerShell `ImportExcel` 7.8.10 ou plus récent.

Installation d’ImportExcel :

```powershell
Set-PSRepository PSGallery -InstallationPolicy Trusted
Install-Module ImportExcel -Scope CurrentUser -MinimumVersion 7.8.10
```

ImportExcel est la seule dépendance ajoutée. Il permet de produire un vrai classeur `.xlsx` sans
installer Microsoft Excel ni utiliser COM Automation.

## Préparer la configuration UNIT

1. Copier le fichier d’exemple sans le suffixe `.example`. Le fichier local obtenu est ignoré par
   Git afin de ne pas publier les URL, comptes ou UID du client :

   ```powershell
   Copy-Item `
     .\config\ResourceReconciliation.UNIT.example.psd1 `
     .\config\ResourceReconciliation.UNIT.psd1
   ```

2. Compléter `Sql.Server` et `Sql.Database` avec le listener/serveur SQL et la base ProjectService de
   l’instance UNIT.
3. Confirmer les trois noms de vues dans `Sql.Views`. Les valeurs proposées sont :
   `MSP_EpmResource_UserView`, `MSP_EpmProject_UserView` et `MSP_EpmAssignment_UserView`.
4. Confirmer les chemins des DLL CSOM.
5. Inscrire l’URL du PWA UNIT dans `Target.PwaUrl` et conserver `Environment = UNIT`.

Inscrire les valeurs propres à la ressource dans `Identity` : les deux UID, le UPN, le compte Windows
et le nom d’affichage. Pour analyser une autre ressource, copier la configuration locale et remplacer
ces valeurs. Aucun changement de code n’est requis.

## Exécuter l’inventaire

Depuis le dossier `ResourceIdentityReconciliation` :

```powershell
.\Invoke-PwaResourceReconciliation.ps1 `
  -ConfigPath .\config\ResourceReconciliation.UNIT.psd1 `
  -Verbose
```

`ConfigPath` est obligatoire. Si le script est lancé sans ce paramètre, PowerShell affiche
`Fournissez des valeurs pour les paramètres suivants : ConfigPath:`. On peut alors saisir le chemin,
mais il est préférable d’annuler avec `Ctrl+C`, de compléter la configuration, puis de relancer la
commande ci-dessus.

Par défaut :

- Project Online utilise le mécanisme interactif de `Common.ps1`;
- PSSE utilise les credentials Windows courants;
- SQL Server utilise l’authentification Windows intégrée.

Un credential Windows différent peut être fourni à CSOM, et un credential SQL explicite peut être
fourni sans l’inscrire dans la configuration :

```powershell
$psseCredential = Get-Credential 'CONTOSO\service-account'
$sqlCredential = Get-Credential 'compte_sql_lecture'

.\Invoke-PwaResourceReconciliation.ps1 `
  -ConfigPath .\config\ResourceReconciliation.UNIT.psd1 `
  -TargetWindowsCredential $psseCredential `
  -TargetSqlCredential $sqlCredential
```

Les credentials ne sont jamais journalisés ni enregistrés dans le classeur.

## Rapport produit

Le nom inclut l’environnement et un timestamp afin de ne pas écraser un rapport existant :

`PwaResourceReconciliation-UNIT-AAAAMMJJ-HHMMSS.xlsx`

| Feuille | Utilité |
|---|---|
| `01_Identités` | Identités source/cible des deux ResourceUID |
| `02_Source_ProjectOnline` | Relations avant migration |
| `03_Cible_PSSE` | Relations dans UNIT ou PROD |
| `04_Réconciliation` | Diagnostic principal, correction suggérée et colonnes de pilotage |
| `05_Assignments_Détail` | Assignments exacts utilisant l’un des deux ResourceUID |
| `06_Contrôle_Exécution` | Sources, champs résolus, volumes et taux de ProjectUID préservés |

Les logs sont écrits dans le sous-dossier `logs`. Le dépôt exclut déjà les fichiers `.xlsx` et les
logs afin d’éviter d’y publier des identités ou des données de projets.

## Procédure de test dans UNIT

1. Exécuter `tests/Test-ReconciliationLogic.ps1` pour valider les six diagnostics de base.
2. Exécuter l’inventaire avec la configuration UNIT.
3. Dans `01_Identités`, confirmer les deux UID et les comptes inscrits dans la configuration locale.
4. Dans `06_Contrôle_Exécution`, valider le taux de `ProjectUID` préservés. Aucun rapprochement par
   `ProjectName` n’est substitué à un UID absent.
5. Filtrer `04_Réconciliation` sur les diagnostics autres que `CONFORME`.
6. Vérifier manuellement quelques projets dans PWA UNIT : Owner, Build Team et assignments publiés.
7. Conserver le rapport comme preuve avant toute correction manuelle.
8. Après correction manuelle, relancer Inventory et comparer les rapports. Le mode `Validate` formel
   viendra dans la prochaine phase.

## Lire les anomalies pour une ressource

- Mauvais Owner : `OWNER:À_CORRIGER_UID` et `Owner_Cible_UID` égal au MigrationResourceUID.
- Mauvais membre du Team : `PROJECT_TEAM:À_CORRIGER_UID`.
- Mauvaise ressource dans les assignments : `ASSIGNMENT:À_CORRIGER_UID`; les lignes exactes sont dans
  `05_Assignments_Détail`.
- Doublon : relation présente pour les deux UID dans la cible, diagnostic `DOUBLON_CIBLE`.
- Anomalie déjà présente dans Online : une relation source utilise le MigrationResourceUID,
  diagnostic `ANOMALIE_SOURCE`. Aucun remplacement automatique ne doit être déduit du suffixe du nom.

## Limites connues

- La première exécution doit confirmer les noms de colonnes réellement exposés par ProjectData et les
  vues SQL. Le script accepte les variantes `UID`, `Uid` et `Id` courantes et inscrit le mapping retenu
  dans la feuille de contrôle.
- Le Project Team est lu projet par projet. Le traitement peut être long sur une instance importante.
- Les permissions effectives PWA ne sont pas calculées.
- Aucune correction d’assignments n’est prévue dans cette révision.
