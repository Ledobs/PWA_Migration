# Design des futurs modes Validate et Fix

| Contrôle documentaire | Valeur |
|---|---|
| Auteur | François Breton |
| Date | 2026-08-06 |
| Révision | 1.0.1 |

Ces modes ne sont pas codés dans la révision 1.0.1. L’orchestrateur reconnaît leurs noms, mais refuse
de les exécuter afin d’empêcher une écriture accidentelle.

## Validate

Le mode `Validate` réutilisera exactement les connecteurs et le modèle normalisé d’Inventory.

Entrées proposées :

- le fichier de configuration;
- un rapport Inventory de référence;
- facultativement un fichier de corrections approuvées.

Traitement proposé :

1. relire Project Online et PSSE;
2. recalculer les trois relations;
3. comparer le nouvel état au résultat attendu;
4. produire `CONFORME` ou `NON_CONFORME` par action;
5. conserver l’état avant/après, l’heure et l’identité de l’opérateur.

## Fix

Le mode `Fix` utilisera uniquement les API supportées Project Server/CSOM. Il n’effectuera aucune
écriture SQL.

Fichier de contrôle proposé :

| Exécuter | Action | ProjectUID | AncienResourceUID | NouveauResourceUID | Commentaire |
|---|---|---|---|---|---|
| Oui/Non | `SET_OWNER` ou `ADD_TEAM_MEMBER` | GUID | GUID | GUID | Approbation pilote |

Garde-fous proposés :

- `Exécuter` doit être exactement `Oui`;
- environnement `UNIT` obligatoire par défaut;
- `SupportsShouldProcess`, `-WhatIf` et confirmation à impact élevé;
- rapport Inventory récent et taux ProjectUID préservés validé;
- snapshot avant correction;
- ajout de la bonne ressource au Team, publication et validation avant tout retrait du doublon;
- journal immuable des actions et QueueJobUID;
- relance automatique du mode Validate après publication.

## Ordre des corrections

1. Project Owner;
2. Project Team;
3. validation de l’accès utilisateur;
4. assignments.

La correction des assignments doit demeurer une stratégie distincte. Work, Actual Work, dates,
Status Manager, timesheets et historique peuvent être touchés. Aucun code de correction d’assignment
ne doit être ajouté sans une approbation et une procédure de retour arrière propres à ce scénario.
