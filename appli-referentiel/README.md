# Appli « Référentiel Excel »

Application web (page Claude) qui affiche les fiches de `../boite-a-outils-excel/referentiel-excel.md` sur téléphone et ordinateur.

- Lien : https://claude.ai/artifact/Ra1jKfKjkiAUiwVeWcL5V9 (privé, accessible avec ton compte Claude)
- `referentiel-excel.html` : la page publiée.
- `sync.py` : convertit le référentiel Markdown en fiches (`fiches.json`).

## Comment les fiches se mettent à jour
Les fiches sont stockées dans la base de données de l'appli (collection `fiches`, documents `f01`, `f02`…), pas dans la page elle-même.
Une fiche ajoutée apparaît donc tout de suite sur tous les appareils, sans republier la page.

- **Par Claude** : dire « ajoute ça au référentiel Excel ». Claude met à jour le Markdown, lance `python3 sync.py`, puis écrit la fiche dans la base de l'appli.
- **Depuis le téléphone** : bouton « + Fiche » dans l'appli. Demander ensuite à Claude de recopier ces fiches (marquées `source: appli`) dans le Markdown.
