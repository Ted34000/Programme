# Appli « Ma Trésorerie »

Outil web (page Claude) de suivi et de planification de la trésorerie personnelle, utilisable sur téléphone et ordinateur.

- Lien : https://claude.ai/artifact/2rj8KXnPRwtb8SKGfkGpcY (privé, accessible avec ton compte Claude)
- `tresorerie.html` : la page publiée.
- `exemple_seed.py` : génère le jeu de données d'exemple chargé au départ (bouton « Effacer les exemples » dans l'appli).

## Onglets (critères du cahier des charges)
| Besoin | Onglet |
|---|---|
| Ajouter / enlever / modifier des prélèvements | Prélèvements (revenus, prélèvements, virements entre comptes ; mensuel, trimestriel, annuel, hebdo, ponctuel ; suspension possible) |
| Dates des prélèvements sur un planning | Planning (calendrier mensuel + liste du mois) |
| Alerte avant prélèvement | Tableau de bord → Alertes (délai réglable) + alerte de découvert ou de seuil par compte |
| Comptes et leur utilité | Comptes |
| Objectifs avec indicateurs | Objectifs (montant d'épargne, solde minimum, plafond de dépense, taux d'épargne) + rappel sur le tableau de bord |
| Simuler la trésorerie future | Prévision (courbe et tableau mensuel, de 1 à 24 mois, par compte ou au total) |
| Simuler l'impact d'un achat | Prévision → Simuler un achat (comptant ou en plusieurs fois, apport, frais mensuels) |
| Solde à une date précise et évolution | Soldes (solde à une date, courbe d'évolution, écart entre solde saisi et solde prévu) |
| Moyenne des consommations variables | Dépenses variables (saisie, moyenne sur 1 à 12 mois, budget intégré aux prévisions) |

## Où sont les données
Dans la base de l'artefact (document `budget/main`), lisible et modifiable uniquement par le propriétaire.
Les modifications faites sur un appareil apparaissent sur les autres.
