# Référentiel Excel – techniques et formules

Mémo des techniques Excel apprises au fil des outils du cabinet. Une fiche = une technique.

**Pour ajouter une fiche :** dans une conversation du projet, demander « ajoute ça au référentiel Excel ». La fiche est rédigée sur le même modèle et ajoutée ici.

**Conventions**
- Les formules sont en version française (Excel FR, séparateur `;`), avec la version anglaise en dessous.
- « Mes fichiers » indique où la technique est utilisée.
- Les fiches sont numérotées et classées par thème.

---

## Sommaire

| N° | Fiche | Thème |
|---|---|---|
| 1 | Liste déroulante dynamique | Saisie |
| 2 | Période manquante par rapport au mois précédent | Dates |
| 3 | Reporter une date au jour ouvré suivant | Dates |
| 4 | Bouton pour aller sur une autre feuille | Navigation |
| 5 | Couleur selon un statut (mise en forme conditionnelle) | Mise en forme |
| 6 | Compter selon plusieurs critères | Calcul |
| 7 | Rechercher une valeur dans un tableau | Recherche |
| 8 | Raccourcis et réglages utiles | Astuces |

---

## Thème : Saisie

### 1. Liste déroulante dynamique

**Objectif :** proposer une liste de choix dans une colonne. La liste s'allonge toute seule quand on ajoute une valeur.

**Étapes**
1. **La liste :** dans la feuille Paramètres, utiliser la première colonne libre (AC, puis AD…). Mettre le titre en ligne 2 et les valeurs à partir de la ligne 3, une par cellule, sans ligne vide.
2. **Le nom :** Formules > Gestionnaire de noms > Nouveau.
   - Nom : `L_MaListe` (préfixe `L_`, sans espace ni accent).
   - Fait référence à :
     ```
     =DECALER(Paramètres!$AC$3;0;0;MAX(1;NBVAL(Paramètres!$AC$3:$AC$42));1)
     ```
     EN : `=OFFSET(Paramètres!$AC$3,0,0,MAX(1,COUNTA(Paramètres!$AC$3:$AC$42)),1)`
3. **Le menu :** sélectionner les cellules cibles, puis Données > Validation des données > Autoriser : **Liste** > Source : `=L_MaListe`.
4. Facultatif : un message d'aide dans l'onglet « Message de saisie ».

**Comment ça marche :** `DECALER` part de la 1re valeur (AC3) et prend autant de lignes qu'il y a de valeurs remplies (`NBVAL`), jusqu'à 40. Le `MAX(1;…)` évite une erreur quand la liste est vide.

**Pièges**
- Une cellule vide au milieu de la liste coupe la fin de la liste.
- Une liste plus longue que 40 valeurs : remplacer `$AC$42` par une ligne plus basse.

**Mes fichiers :** Tableau de Bord 2026, avec `L_Clients`, `L_Paie`, `L_Logiciels`, `L_OuiNon`, etc. Les listes sont dans Paramètres, colonnes S à AB.

---

## Thème : Dates

### 2. Période manquante par rapport au mois précédent

**Objectif :** à partir du dernier mois saisi (colonne C), ne rien afficher si on est à jour. Sinon, afficher le ou les mois manquants jusqu'au mois précédent.

```
=SI(C3="";"";SI(ANNEE(C3)*12+MOIS(C3)>=ANNEE(AUJOURDHUI())*12+MOIS(AUJOURDHUI())-1;"";SI(ANNEE(C3)*12+MOIS(C3)=ANNEE(AUJOURDHUI())*12+MOIS(AUJOURDHUI())-2;TEXTE(MOIS.DECALER(C3;1);"mmmm aaaa");TEXTE(MOIS.DECALER(C3;1);"mmmm aaaa")&" à "&TEXTE(FIN.MOIS(AUJOURDHUI();-1);"mmmm aaaa"))))
```
EN : `=IF(C3="","",IF(YEAR(C3)*12+MONTH(C3)>=YEAR(TODAY())*12+MONTH(TODAY())-1,"",IF(YEAR(C3)*12+MONTH(C3)=YEAR(TODAY())*12+MONTH(TODAY())-2,TEXT(EDATE(C3,1),"mmmm yyyy"),TEXT(EDATE(C3,1),"mmmm yyyy")&" à "&TEXT(EOMONTH(TODAY(),-1),"mmmm yyyy"))))`

**Exemple au 20/11/2026**

| Dernier mois saisi | Résultat |
|---|---|
| octobre 2026 | vide |
| septembre 2026 | octobre 2026 |
| juillet 2026 | août 2026 à octobre 2026 |

**L'astuce à retenir :** `ANNEE(date)*12+MOIS(date)` transforme une date en « numéro de mois ». Comparer deux numéros de mois revient à comparer des mois sans s'occuper des jours. On compte aussi facilement l'écart en mois en soustrayant les deux numéros.

**Fonctions utiles**
- `MOIS.DECALER(date;n)` : même jour, n mois plus tard.
- `FIN.MOIS(date;-1)` : dernier jour du mois précédent.
- `TEXTE(date;"mmmm aaaa")` : affiche « octobre 2026 ».

**Mes fichiers :** Tableau de Bord 2026, feuille Salaires, colonne D.

### 3. Reporter une date au jour ouvré suivant

**Objectif :** une échéance qui tombe un samedi, un dimanche ou un jour férié est reportée au jour ouvré suivant.

```
=SERIE.JOUR.OUVRE(A2-1;1;Feries)
```
EN : `=WORKDAY(A2-1,1,Feries)`

**Comment ça marche :** on recule d'un jour, puis on avance d'un jour ouvré. Si A2 est déjà un jour ouvré, on retombe dessus ; sinon, on obtient le jour ouvré suivant.

**Prérequis :** une liste des jours fériés, nommée `Feries` (même méthode que la fiche 1 pour créer le nom).

**Mes fichiers :** Tableau de Bord 2026. La feuille Échéances s'en sert pour toutes les dates limites, et la liste des fériés 2026-2030 est dans « Réf. échéances » (nom `E_Feries`).

---

## Thème : Navigation

### 4. Bouton pour aller sur une autre feuille

**Objectif :** une cellule cliquable qui emmène sur une autre feuille, par exemple un bouton « ◄ Accueil ».

```
=LIEN_HYPERTEXTE("#'Accueil'!A1";"◄ Accueil")
```
EN : `=HYPERLINK("#'Accueil'!A1","◄ Accueil")`

**Règles**
- Le `#` indique « dans ce classeur ».
- Le nom de la feuille va entre apostrophes s'il contient un espace : `"#'Suivi échéances'!A1"`.
- Pour l'aspect de bouton : couleur de fond, texte en gras et centré, bordure blanche, ligne plus haute.

**Mes fichiers :** Tableau de Bord 2026, avec la page Accueil et le bouton de retour sur chaque feuille. Le skill « accueil-boutons-excel » le fait automatiquement sur n'importe quel classeur.

---

## Thème : Mise en forme

### 5. Couleur selon un statut (mise en forme conditionnelle)

**Objectif :** colorer une cellule ou une ligne entière selon une valeur, par exemple rouge si « URGENT ».

**Étapes**
1. Sélectionner la plage, par exemple A3:W62.
2. Accueil > Mise en forme conditionnelle > Nouvelle règle > « Utiliser une formule… ».
3. Saisir la formule pour la **1re ligne** de la plage, par exemple :
   ```
   =$W3="URGENT"
   ```
4. Cliquer sur Format… et choisir la couleur.

**La règle du `$`**
- `$W3` : la colonne est bloquée et la ligne s'adapte. Toute la ligne prend la couleur selon la colonne W.
- `W3` sans `$` : chaque cellule teste sa propre colonne.

**Variante avec une date :** `=ET(ESTNUM(K3);K3-AUJOURDHUI()<=7)` colore une date qui tombe dans 7 jours ou moins.

**Mes fichiers :** Tableau de Bord 2026, avec les statuts URGENT / À venir / Planifié dans Échéances et Planning.

---

## Thème : Calcul

### 6. Compter selon plusieurs critères

**Objectif :** compter les lignes qui remplissent plusieurs conditions à la fois.

```
=NB.SI.ENS(Échéances!$W$3:$W$62;"URGENT";Échéances!$Z$3:$Z$62;"Oui")
```
EN : `=COUNTIFS(Échéances!$W$3:$W$62,"URGENT",Échéances!$Z$3:$Z$62,"Oui")`

**Usage :** les critères vont par paires (plage ; critère). On peut ajouter autant de paires que nécessaire. Pour une comparaison, mettre l'opérateur entre guillemets : `">="&AUJOURDHUI()`.

**Proches :** `SOMME.SI.ENS` (pour additionner) et `MOYENNE.SI.ENS` (pour faire la moyenne).

**Mes fichiers :** Tableau de Bord 2026, compteurs de l'Accueil et du Planning.

---

## Thème : Recherche

### 7. Rechercher une valeur dans un tableau

**Objectif :** retrouver une information d'un client, par exemple sa date de clôture, à partir de son nom.

```
=INDEX(Paramètres!$H$3:$H$62;EQUIV(B3;Paramètres!$B$3:$B$62;0))
```
EN : `=INDEX(Paramètres!$H$3:$H$62,MATCH(B3,Paramètres!$B$3:$B$62,0))`

**Comment ça marche :** `EQUIV` trouve à quelle ligne se trouve le client ; `INDEX` va chercher la valeur de cette ligne dans la colonne voulue. Le `0` impose une correspondance exacte.

**Avantages par rapport à RECHERCHEV :** la colonne cherchée peut être à gauche, et la formule ne casse pas si on insère une colonne.

**Si le client n'existe pas :** entourer la formule de `SIERREUR(…;"")`.

**Mes fichiers :** Tableau de Bord 2026, fiche client du Planning.

---

## Thème : Astuces

### 8. Raccourcis et réglages utiles

| Action | Comment |
|---|---|
| Mettre la date du jour dans une cellule | **Ctrl + ;** |
| Changer de feuille (même avec les onglets masqués) | **Ctrl + Page suiv. / Page préc.** |
| Réafficher les onglets du classeur | Fichier > Options > Options avancées > « Afficher les onglets de classeur » |
| Voir et modifier les noms (`L_…`, `E_…`) | Formules > Gestionnaire de noms (**Ctrl + F3**) |
| Afficher les formules au lieu des résultats | **Ctrl + "** (touche ² sur certains claviers) |
| Figer les titres en haut | Affichage > Figer les volets |

---

## Journal des ajouts

| Date | Ajout |
|---|---|
| 23/11/2026 | Création : fiches 1 à 8 (listes déroulantes, dates, navigation, mise en forme conditionnelle, NB.SI.ENS, INDEX/EQUIV, raccourcis) |

---

## Pistes pour un futur outil (Claude Code)

Si le référentiel est utile, il pourra devenir un outil, par exemple :
- un classeur « Boîte à outils Excel » avec une feuille par thème et des exemples qui fonctionnent ;
- ou un petit programme qui applique une fiche à un fichier, comme le skill « accueil-boutons-excel ».

Ce document servira alors de cahier des charges : chaque fiche décrit déjà l'objectif, la formule et un exemple.
