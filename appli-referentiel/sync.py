#!/usr/bin/env python3
"""Convertit le référentiel Markdown en fiches pour l'application « Référentiel Excel ».

Usage : python3 sync.py [referentiel.md] [fiches.json]

Chaque « ### N. Titre » sous un « ## Thème : X » devient une fiche :
{"num": N, "theme": X, "titre": ..., "corps": <markdown de la fiche>, "maj": "AAAA-MM-JJ"}
Le fichier JSON produit sert à remplir la base de l'application (collection « fiches »,
document « fNN »). Claude s'en charge quand on lui demande d'ajouter une fiche.
"""
import json
import re
import sys
from datetime import date
from pathlib import Path

ici = Path(__file__).resolve().parent
src = Path(sys.argv[1]) if len(sys.argv) > 1 else ici.parent / "boite-a-outils-excel" / "referentiel-excel.md"
dst = Path(sys.argv[2]) if len(sys.argv) > 2 else ici / "fiches.json"

theme, fiche, fiches = None, None, []
for ligne in src.read_text(encoding="utf-8").splitlines():
    m_theme = re.match(r"^## Thème : (.+)$", ligne)
    m_fiche = re.match(r"^### (\d+)\. (.+)$", ligne)
    if m_theme or m_fiche or ligne.startswith("## "):
        if fiche:
            fiches.append(fiche)
            fiche = None
        if m_theme:
            theme = m_theme.group(1).strip()
        elif m_fiche and theme:
            fiche = {"num": int(m_fiche.group(1)), "theme": theme, "titre": m_fiche.group(2).strip(), "lignes": []}
        elif not m_theme:
            theme = None
        continue
    if fiche is not None:
        fiche["lignes"].append(ligne)
if fiche:
    fiches.append(fiche)

aujourdhui = date.today().isoformat()
docs = {}
for f in fiches:
    lignes = f.pop("lignes")
    while lignes and lignes[-1].strip() in ("", "---"):
        lignes.pop()
    while lignes and not lignes[0].strip():
        lignes.pop(0)
    f["corps"] = "\n".join(lignes)
    f["maj"] = aujourdhui
    docs["f%02d" % f["num"]] = f

dst.write_text(json.dumps(docs, ensure_ascii=False, indent=1), encoding="utf-8")
print("%d fiches -> %s" % (len(docs), dst))
