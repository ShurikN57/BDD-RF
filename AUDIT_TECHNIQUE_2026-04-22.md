# Audit technique VBA — BDD-RF

Date: 2026-04-22
Périmètre: export `src/*.txt` (modules VBA)

## Synthèse

- **Statut global**: code globalement structuré (Option Explicit présent), mais avec **2 points prioritaires** à corriger.
- **Priorité haute**:
  1. Gestion d'état `Application.EnableEvents` incohérente dans `zRFCollage`.
  2. Chemin dépôt local codé en dur dans `zRFGitSync`.

## Points forts observés

1. `Option Explicit` est utilisé dans tous les modules exportés.
2. Le module d'import/export `zRFGitSync` est découpé en fonctions helpers lisibles.
3. Plusieurs procédures restaurent correctement `ScreenUpdating/EnableEvents` via variables de sauvegarde.

## Constats détaillés

### 1) État des événements Excel potentiellement corrompu (Priorité: Haute)

Dans `CollerValeursRecherche`, l'état initial des événements est mémorisé (`prevEnableEvents`) mais la sortie de procédure force ensuite `Application.EnableEvents = True` (chemins normal + erreur), au lieu de restaurer l'état d'origine.

Conséquences possibles:
- Réactivation involontaire des événements si l'appelant les avait volontairement désactivés.
- Effets de bord dans des macros chaînées.

Indicateurs:
- Affectation redondante `Application.EnableEvents = prevEnableEvents` en double.
- Sorties `Fin`/`FinAvecErreur` qui forcent `True`.

### 2) Chemin local absolu en dur (Priorité: Haute)

`DOSSIER_REPO` pointe vers un chemin utilisateur spécifique (`C:\Users\...\Desktop\BDD-RF-GitHub`).

Conséquences possibles:
- Non-portabilité du classeur entre postes.
- Échecs silencieux d'export/import hors machine d'origine.

Recommandation:
- Basculer vers un chemin configurable (feuille de config, variable d'environnement, ou détection relative au classeur).

### 3) Usage fréquent de `On Error Resume Next` sans traçage (Priorité: Moyenne)

Le pattern est utilisé sur des opérations fichiers (`Kill`, suppressions, etc.) sans journalisation des erreurs attendues/non attendues.

Conséquences possibles:
- Diagnostic plus difficile en cas d'incident réel.

Recommandation:
- Encadrer les sections `On Error Resume Next` par un contrôle explicite de `Err.Number` quand une suppression échoue de façon inattendue.

## Plan d'action recommandé

1. **Corriger `zRFCollage`**: restaurer systématiquement `Application.EnableEvents = prevEnableEvents` sur toutes les sorties.
2. **Externaliser `DOSSIER_REPO`** dans `zRFGitSync`.
3. **Ajouter une journalisation minimale** des erreurs d'E/S ignorées.
4. (Optionnel) Ajouter une macro de self-check qui valide les prérequis (accès VBProject, chemins, droits d'écriture).

## Contrôles réalisés pour cet audit

- Recherche ciblée des patterns à risque (`On Error Resume Next`, `Kill`, `Workbook_Open`, toggles d'état Application).
- Revue manuelle des modules `zRFCollage` et `zRFGitSync`.
- Vérification automatisée de la présence de `Option Explicit` sur tous les fichiers `src/*.txt`.
