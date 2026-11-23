# Boîte à outils Excel

Classeur d'exemples qui fonctionnent, construit à partir de `referentiel-excel.md` (une feuille par thème, charte visuelle du cabinet).

- `Boite_a_outils_Excel.xlsx` : le classeur à ouvrir dans Excel.
- `construire.py` : le script qui le génère (`python3 construire.py`).
- `referentiel-excel.md` : le référentiel (cahier des charges).

## Contenu
| Feuille | Fiches |
|---|---|
| Accueil | Boutons vers chaque fiche, mode d'emploi, journal des ajouts |
| Saisie | 1 · Liste déroulante dynamique |
| Dates | 2 · Période manquante · 3 · Jour ouvré suivant |
| Navigation | 4 · Bouton vers une autre feuille |
| Mise en forme | 5 · Couleur selon un statut |
| Calcul | 6 · Compter selon plusieurs critères |
| Recherche | 7 · INDEX / EQUIV |
| Astuces | 8 · Raccourcis |
| Clients démo | Toutes les fiches appliquées à un même tableau de 10 clients fictifs |
| Paramètres | Listes `L_…`, date de référence `E_Ref`, seuil `E_Urgent`, fériés 2026-2030 `E_Feries` |

Jaune pâle = à modifier pour essayer · blanc = calculé · bleu = titres et libellés.

## Ajouter une fiche
1. Ajouter la fiche dans `referentiel-excel.md`.
2. Ajouter un bloc dans `construire.py` (copier le bloc d'une fiche existante : `bloc`, `ligne`, `joker`, puis la zone d'essai), et un bouton dans `GROUPES` (Accueil).
3. Relancer `python3 construire.py`.
