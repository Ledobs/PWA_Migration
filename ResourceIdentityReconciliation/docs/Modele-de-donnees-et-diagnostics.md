# Modèle de données et diagnostics

| Contrôle documentaire | Valeur |
|---|---|
| Auteur | François Breton |
| Date | 2026-08-06 |
| Révision | 1.0.1 |

## Modèle normalisé

Les connecteurs traduisent les données source dans quatre objets communs.

| Objet | Clé | Attributs principaux |
|---|---|---|
| Resource | `ResourceUID` | nom, compte, Claims Account, courriel, actif |
| Project | `ProjectUID` | nom, OwnerResourceUID, OwnerName |
| ProjectTeam | `ProjectUID + ResourceUID` | environnement |
| Assignment | `AssignmentUID` | ProjectUID, TaskUID, ResourceUID, noms explicatifs |

Les noms sont des attributs explicatifs. Ils ne remplacent jamais les UID.

## Résolution des champs

ProjectData et les vues SQL n’exposent pas toujours la même graphie. Le connecteur résout les alias
suivants, puis les normalise :

| Champ normalisé | Alias acceptés |
|---|---|
| ResourceUID | `ResourceUID`, `ResourceUid`, `ResourceId` |
| ProjectUID | `ProjectUID`, `ProjectUid`, `ProjectId` |
| AssignmentUID | `AssignmentUID`, `AssignmentUid`, `AssignmentId` |
| TaskUID | `TaskUID`, `TaskUid`, `TaskId` |
| ProjectOwnerResourceUID | `ProjectOwnerResourceUID`, `ProjectOwnerResourceUid`, `ProjectOwnerId` |

Le mapping réellement retenu est inscrit dans `06_Contrôle_Exécution`.

## Matrice de diagnostic par relation

La même fonction est appliquée séparément à `OWNER`, `PROJECT_TEAM` et `ASSIGNMENT`.

| Source User | Source Migration | Cible User | Cible Migration | Diagnostic |
|---|---|---|---|---|
| Oui | Non | Oui | Non | `CONFORME` |
| Oui | Non | Non | Oui | `À_CORRIGER_UID` |
| Oui | Non | Oui | Oui | `DOUBLON_CIBLE` |
| Oui | Non | Non | Non | `RELATION_MANQUANTE_CIBLE` |
| Non | Non | Oui ou Non | Oui ou Non, avec au moins une présence cible | `RELATION_SUPPLÉMENTAIRE_CIBLE` |
| Peu importe | Oui | Peu importe | Peu importe | `ANOMALIE_SOURCE` |
| Autre combinaison | | | | `À_INVESTIGUER` |

Si plusieurs relations d’un même projet ont des diagnostics différents, le diagnostic global suit
un ordre prudent : anomalie source, mauvais UID, doublon, relation manquante, relation supplémentaire,
à investiguer, conforme. `Diagnostic_Détail` conserve toujours le résultat de chaque relation.

## ProjectUID préservé

Le script mesure combien de ProjectUID pertinents de la source existent aussi dans la liste des
projets cible. Il ne tente aucun rapprochement par nom lorsqu’un UID n’est pas retrouvé. Une valeur
inférieure à 100 % doit être investiguée avant de se servir du rapport pour automatiser une action.

## Accès attendu et permission effective

`Accès_Attendu_User = Oui` lorsqu’au moins une des relations suivantes existe : Owner, Project Team ou
Assignment. Cette valeur sert au pilotage, mais ne tient pas compte des groupes, catégories, règles
dynamiques, du RBS ou d’un modèle SharePoint Permission Mode.

Microsoft précise que le CSOM ne couvre pas l’administration des groupes et catégories de sécurité :
[`What the CSOM does and does not do`](https://learn.microsoft.com/en-us/office/client-developer/project/what-the-csom-does-and-does-not-do).
