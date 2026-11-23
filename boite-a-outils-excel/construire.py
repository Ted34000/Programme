#!/usr/bin/env python3
"""Construit le classeur « Boîte à outils Excel » à partir du référentiel Excel.

Usage : python3 construire.py [sortie.xlsx]

Une feuille par thème du référentiel, chaque fiche avec : objectif, formule FR / EN,
zone « À TOI DE JOUER » qui fonctionne (cellules jaune pâle à modifier), pièges.
Une feuille « Clients démo » réunit toutes les techniques sur un même tableau,
comme dans le Tableau de Bord 2026. Charte visuelle du cabinet.
"""
import sys
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L

SORTIE = sys.argv[1] if len(sys.argv) > 1 else "Boite_a_outils_Excel.xlsx"

# ------------------------------------------------------------------ charte cabinet
POLICE = "Aptos Narrow"
TXT, MARINE, SOUS, AIDE, TECH, PARAM = "1F2D3D", "2B4A6F", "4A5A6B", "6B7C8F", "8A94A0", "1F4E9E"
F_BANDEAU, F_ENTETE, F_LIBELLE, F_SAISIE, F_FORMULE = "BCD9EE", "CFE3F3", "EAF3FA", "FFFBE6", "F4F6F8"
FIN = Side(style="thin", color="C5D5E4")
CADRE = Border(left=FIN, right=FIN, top=FIN, bottom=FIN)
CADRE_ENTETE = Border(left=FIN, right=FIN, top=FIN, bottom=Side(style="medium", color="8FB5D6"))
DATE, MOIS, EURO = "DD/MM/YYYY", "mmmm yyyy", '#,##0.00 "€"'
NB_LIGNES_CLIENTS = 40          # lignes 3 à 42 de « Clients démo »
DERN = 2 + NB_LIGNES_CLIENTS


def fill(c):
    return PatternFill("solid", fgColor=c)


def font(sz=11, color=TXT, b=False, i=False):
    return Font(name=POLICE, size=sz, color=color, bold=b, italic=i)


def texte(cell, valeur):
    """Écrit une chaîne telle quelle, même si elle commence par « = »."""
    cell.value = valeur
    cell.data_type = "s"


# ------------------------------------------------------------------ blocs de mise en page
COL_FIN = 8   # les blocs s'étendent de B à H


def fusion(ws, r, c1=2, c2=COL_FIN):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)


def hauteur(txt, largeur=95, base=15):
    lignes = sum(max(1, -(-len(p) // largeur)) for p in txt.split("\n"))
    return max(base, 15 * lignes + 3)


def bouton(ws, ref, cible, libelle, fond=MARINE, encre="FFFFFF", taille=11):
    c = ws[ref]
    c.value = '=HYPERLINK("#\'%s\'!A1","%s")' % (cible, libelle)
    c.font = font(taille, encre, b=True)
    c.fill = fill(fond)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = Border(*(Side(style="thin", color="FFFFFF"),) * 4)


def feuille(wb, nom, titre, onglet, freeze="A2"):
    ws = wb.create_sheet(nom)
    ws.sheet_properties.tabColor = onglet
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 26
    for c in range(3, COL_FIN + 1):
        ws.column_dimensions[L(c)].width = 16
    ws.row_dimensions[1].height = 26
    ws.merge_cells("B1:C1")
    bouton(ws, "B1", "Accueil", "◄ Accueil")
    ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False
    fusion(ws, 2)
    c = ws["B2"]
    c.value = titre.upper()
    c.font = font(16, MARINE, b=True)
    c.fill = fill(F_BANDEAU)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 34
    return ws


def bloc(ws, r, txt):
    """Titre de fiche ou de bloc."""
    r += 1
    fusion(ws, r)
    c = ws.cell(r, 2, txt.upper())
    c.font = font(12, MARINE, b=True)
    c.fill = fill(F_ENTETE)
    c.alignment = Alignment(vertical="center", indent=1)
    for k in range(2, COL_FIN + 1):
        ws.cell(r, k).border = Border(bottom=Side(style="medium", color="8FB5D6"))
    ws.row_dimensions[r].height = 24
    return r + 1


def ligne(ws, r, libelle, txt, genre="texte"):
    """Libellé en B, texte fusionné C:H. genre : texte, formule, en, note."""
    b = ws.cell(r, 2, libelle)
    b.font = font(11, MARINE, b=True)
    b.fill = fill(F_LIBELLE)
    b.alignment = Alignment(vertical="top", wrap_text=True)
    b.border = CADRE
    fusion(ws, r, 3)
    c = ws.cell(r, 3)
    texte(c, txt)
    styles = {"texte": (font(), None), "formule": (font(11, TXT, b=True), F_FORMULE),
              "en": (font(10, SOUS, i=True), F_FORMULE), "note": (font(10, AIDE), None)}
    c.font, f = styles[genre]
    if f:
        c.fill = fill(f)
    c.alignment = Alignment(vertical="top", wrap_text=True)
    for k in range(3, COL_FIN + 1):
        ws.cell(r, k).border = CADRE
    ws.row_dimensions[r].height = hauteur(txt, 88 if genre in ("formule", "en") else 98)
    return r + 1


def entetes(ws, r, titres, c0=2):
    for k, t in enumerate(titres):
        c = ws.cell(r, c0 + k, t)
        c.font = font(11, MARINE, b=True)
        c.fill = fill(F_ENTETE)
        c.border = CADRE_ENTETE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 32
    return r + 1


def saisie(ws, ref, v=None, fmt=None, gauche=False):
    c = ws[ref]
    c.value = v
    c.fill = fill(F_SAISIE)
    c.font = font()
    c.border = CADRE
    c.alignment = Alignment(horizontal="left" if gauche else "center", vertical="center")
    if fmt:
        c.number_format = fmt
    return c


def calcul(ws, ref, formule, fmt=None, gras=False, gauche=False):
    c = ws[ref]
    c.value = formule
    c.font = font(11, MARINE if gras else TXT, b=gras)
    c.border = CADRE
    c.alignment = Alignment(horizontal="left" if gauche else "center", vertical="center", wrap_text=True)
    if fmt:
        c.number_format = fmt
    return c


def libelle(ws, ref, txt):
    c = ws[ref]
    c.value = txt
    c.font = font(11, TXT, b=True)
    c.fill = fill(F_LIBELLE)
    c.border = CADRE
    c.alignment = Alignment(vertical="center", wrap_text=True)
    return c


def note(ws, ref, txt):
    c = ws[ref]
    c.value = txt
    c.font = font(10, AIDE, i=True)
    c.alignment = Alignment(vertical="center")
    return c


def liste(ws, plage, nom, aide=None):
    dv = DataValidation(type="list", formula1=nom, allow_blank=True)
    if aide:
        dv.promptTitle, dv.prompt = "Liste", aide
        dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(plage)


def nom_defini(wb, nom, ref):
    wb.defined_names[nom] = DefinedName(nom, attr_text=ref)


def joker(ws, r):
    """Petit bandeau « À TOI DE JOUER »."""
    fusion(ws, r)
    c = ws.cell(r, 2, "À TOI DE JOUER  ·  modifie les cellules jaune pâle, les cellules blanches se recalculent")
    c.font = font(11, "7F6000", b=True)
    c.fill = fill("FFF2CC")
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[r].height = 22
    return r + 1


def regle_mfc(ws, plage, formule, fond, encre=None, gras=False):
    ws.conditional_formatting.add(plage, FormulaRule(
        formula=[formule], fill=fill(fond), font=Font(color=encre, bold=gras) if (encre or gras) else None))


def mfc_statuts(ws, plage, col1):
    """Couleurs de statut de la charte, sur la colonne col1 (ex. « $I3 »)."""
    for vals, fond, encre in ((("URGENT", "En retard"), "F6C7CC", "9C2B38"),
                              (("À venir", "À envoyer"), "FFE2B8", None),
                              (("En cours",), "D3E4F7", None),
                              (("Planifié", "Terminé"), "CDEBD5", None)):
        f = "OR(%s)" % ",".join('%s="%s"' % (col1, v) for v in vals) if len(vals) > 1 else '%s="%s"' % (col1, vals[0])
        regle_mfc(ws, plage, f, fond, encre, gras=bool(encre))


# ------------------------------------------------------------------ jours fériés (France métropolitaine)
def paques(an):
    a, b, c = an % 19, an // 100, an % 100
    d, e = b // 4, b % 4
    g = (8 * b + 13) // 25
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 19 * l) // 433
    mois = (h + l - 7 * m + 90) // 25
    return date(an, mois, (h + l - 7 * m + 33 * mois + 19) % 32)


def feries(an):
    p = paques(an)
    return sorted([date(an, 1, 1), p + timedelta(1), date(an, 5, 1), date(an, 5, 8), p + timedelta(39),
                   p + timedelta(50), date(an, 7, 14), date(an, 8, 15), date(an, 11, 1), date(an, 11, 11),
                   date(an, 12, 25)])


# ================================================================== construction
wb = Workbook()
accueil = wb.active
accueil.title = "Accueil"

# ------------------------------------------------------------------ Paramètres
par = feuille(wb, "Paramètres", "Paramètres", "C9CED6")
par.column_dimensions["B"].width = 28
par.column_dimensions["C"].width = 14
par.column_dimensions["D"].width = 24
par.column_dimensions["E"].width = 3
par.column_dimensions["F"].width = 3
for col in "GHIJK":
    par.column_dimensions[col].width = 15
par.column_dimensions["L"].width = 3
par.column_dimensions["M"].width = 14
par.unmerge_cells("B2:H2")
par.merge_cells("B2:D2")
r = 4
fusion(par, r, 2, 4)
c = par.cell(r, 2, "RÉGLAGES GÉNÉRAUX")
c.font, c.fill = font(11, MARINE, b=True), fill(F_BANDEAU)
libelle(par, "B5", "Date de référence")
saisie(par, "C5", None, DATE).font = font(11, PARAM, b=True)
note(par, "D5", "← vide = date du jour")
libelle(par, "B6", "Date utilisée (tech.)").font = font(9, TECH, i=True)
calcul(par, "C6", '=IF(C5="",TODAY(),C5)', DATE).font = font(9, TECH, i=True)
libelle(par, "B7", "Seuil d'urgence (jours)")
saisie(par, "C7", 7).font = font(11, PARAM, b=True)
note(par, "D7", "← utilisé par la fiche 5")
nom_defini(wb, "E_Ref", "Paramètres!$C$6")
nom_defini(wb, "E_Urgent", "Paramètres!$C$7")

r = 10
fusion(par, r, 2, 4)
c = par.cell(r, 2, "MODE D'EMPLOI")
c.font, c.fill = font(11, MARINE, b=True), fill(F_BANDEAU)
for k, t in enumerate(["Jaune pâle = à renseigner ; texte bleu = paramètre ; blanc = calculé.",
                       "Listes de référence à droite (colonnes G à K) : une valeur par cellule, sans ligne vide.",
                       "Elles alimentent les listes déroulantes via les noms L_… (fiche 1).",
                       "Jours fériés 2026-2030 en colonne M (nom E_Feries, fiche 3).",
                       "Formules > Gestionnaire de noms (Ctrl + F3) pour voir tous les noms."]):
    fusion(par, r + 1 + k, 2, 4)
    note(par, "B%d" % (r + 1 + k), t).font = font(10, AIDE)

LISTES = [("G", "Statuts", "L_Statuts", ["URGENT", "À venir", "Planifié", "En cours", "Terminé"]),
          ("H", "Oui / Non", "L_OuiNon", ["Oui", "Non"]),
          ("I", "Formes juridiques", "L_Formes", ["SARL", "SAS", "SASU", "EURL", "SCI", "EI"]),
          ("J", "Logiciels", "L_Logiciels", ["Logiciel A", "Logiciel B", "Logiciel C"]),
          ("K", "Ma liste (essai)", "L_Essai", ["Choix 1", "Choix 2", "Choix 3"])]
for col, titre, nom, vals in LISTES:
    entetes(par, 3, [titre], ord(col) - 64)
    for k in range(40):
        ref = "%s%d" % (col, 4 + k)
        saisie(par, ref, vals[k] if k < len(vals) else None)
    nom_defini(wb, nom, "OFFSET(Paramètres!$%s$4,0,0,MAX(1,COUNTA(Paramètres!$%s$4:$%s$43)),1)" % (col, col, col))
entetes(par, 3, ["Jours fériés"], 13)
fer = [d for an in range(2026, 2031) for d in feries(an)]
for k, d in enumerate(fer):
    saisie(par, "M%d" % (4 + k), d, DATE)
nom_defini(wb, "E_Feries", "Paramètres!$M$4:$M$%d" % (3 + len(fer)))
note(par, "G2", "Listes de référence (ajouter les valeurs à la suite, sans ligne vide)")
par.row_dimensions[3].height = 32
par.freeze_panes = "A4"

# ------------------------------------------------------------------ Clients démo
cli = feuille(wb, "Clients démo", "Clients démo – toutes les fiches sur un même tableau", "A8D5BA", freeze="D4")
cli.unmerge_cells("B2:H2")
cli.merge_cells("B2:L2")
largeurs = {"B": 17, "C": 22, "D": 13, "E": 17, "F": 26, "G": 13, "H": 13, "I": 13, "J": 9, "K": 13, "L": 30}
for k, v in largeurs.items():
    cli.column_dimensions[k].width = v
cli.column_dimensions["A"].width = 2
TITRES = ["Forme juridique", "Client", "Date de clôture", "Dernier mois saisi", "Période manquante\n(fiche 2)",
          "Échéance théorique", "Échéance reportée\n(fiche 3)", "Statut\n(fiches 1 et 5)", "Envoyé", "Honoraires",
          "Remarques"]
entetes(cli, 3, TITRES)
DEMO = [("SARL", "Client Alpha", date(2026, 12, 31), date(2026, 10, 1), date(2026, 11, 28), "URGENT", "Non", 1200),
        ("SAS", "Client Bêta", date(2027, 3, 31), date(2026, 9, 1), date(2026, 12, 5), "À venir", "Non", 2400),
        ("SCI", "Client Gamma", date(2026, 12, 31), date(2026, 7, 1), date(2026, 12, 15), "Planifié", "Non", 850),
        ("EI", "Client Delta", date(2026, 12, 31), date(2026, 10, 1), date(2026, 12, 25), "À venir", "Oui", 600),
        ("SASU", "Client Epsilon", date(2027, 6, 30), date(2026, 8, 1), date(2027, 1, 1), "Planifié", "Non", 3100),
        ("EURL", "Client Zêta", date(2026, 12, 31), date(2026, 10, 1), date(2026, 11, 30), "URGENT", "Oui", 1500),
        ("SARL", "Client Êta", date(2027, 9, 30), date(2026, 11, 1), date(2027, 5, 1), "En cours", "Non", 900),
        ("SAS", "Client Thêta", date(2026, 12, 31), date(2026, 10, 1), date(2027, 7, 14), "Planifié", "Non", 1750),
        ("SCI", "Client Iota", date(2026, 12, 31), date(2026, 5, 1), date(2026, 12, 20), "À venir", "Non", 450),
        ("SAS", "Client Kappa", date(2027, 3, 31), date(2026, 10, 1), date(2027, 1, 15), "Terminé", "Oui", 2000)]
P = "IF(E{r}=\"\",\"\",IF(YEAR(E{r})*12+MONTH(E{r})>=YEAR(E_Ref)*12+MONTH(E_Ref)-1,\"\"," \
    "IF(YEAR(E{r})*12+MONTH(E{r})=YEAR(E_Ref)*12+MONTH(E_Ref)-2,TEXT(EDATE(E{r},1),\"mmmm aaaa\")," \
    "TEXT(EDATE(E{r},1),\"mmmm aaaa\")&\" à \"&TEXT(EOMONTH(E_Ref,-1),\"mmmm aaaa\"))))"
for k in range(NB_LIGNES_CLIENTS):
    r = 4 + k
    d = DEMO[k] if k < len(DEMO) else (None,) * 8
    saisie(cli, "B%d" % r, d[0], gauche=True)
    c = saisie(cli, "C%d" % r, d[1], gauche=True)
    c.fill, c.font = fill(F_LIBELLE), font(11, TXT, b=True)
    saisie(cli, "D%d" % r, d[2], DATE)
    saisie(cli, "E%d" % r, d[3], MOIS)
    calcul(cli, "F%d" % r, "=" + P.format(r=r))
    saisie(cli, "G%d" % r, d[4], DATE)
    calcul(cli, "H%d" % r, '=IF(G{r}="","",WORKDAY(G{r}-1,1,E_Feries))'.format(r=r), DATE)
    saisie(cli, "I%d" % r, d[5])
    saisie(cli, "J%d" % r, d[6])
    saisie(cli, "K%d" % r, d[7], EURO)
    saisie(cli, "L%d" % r, None, gauche=True)
cli.row_dimensions[3].height = 46
# on décale le tableau d'une ligne : ligne 3 = en-têtes, données 4 à 43
DERN = 3 + NB_LIGNES_CLIENTS
CL = "'Clients démo'!$%s$4:$%s$" + str(DERN)
liste(cli, "B4:B%d" % DERN, "L_Formes")
liste(cli, "I4:I%d" % DERN, "L_Statuts", "URGENT / À venir / Planifié / En cours / Terminé")
liste(cli, "J4:J%d" % DERN, "L_OuiNon")
cli.auto_filter.ref = "B3:L%d" % DERN
ZONE = "B4:L%d" % DERN
regle_mfc(cli, ZONE, '$I4="URGENT"', "F6C7CC", "9C2B38")
mfc_statuts(cli, "I4:I%d" % DERN, "$I4")
regle_mfc(cli, "H4:H%d" % DERN, "AND(ISNUMBER(H4),H4<>G4)", "FFE2B8")
regle_mfc(cli, "F4:F%d" % DERN, 'F4<>""', "F6C7CC", "9C2B38")
cli.sheet_view.showGridLines = True
note(cli, "B%d" % (DERN + 2), "Jaune pâle = à saisir · blanc = calculé. Période manquante rouge = salaires en retard ; "
     "échéance reportée orange = la date théorique tombait un week-end ou un jour férié.")
nom_defini(wb, "L_Clients", "OFFSET('Clients démo'!$C$4,0,0,MAX(1,COUNTA('Clients démo'!$C$4:$C$%d)),1)" % DERN)

# ------------------------------------------------------------------ Fiche 1 : Saisie
ws = feuille(wb, "Saisie", "Thème : Saisie", "F4B6C2")
r = bloc(ws, 3, "Fiche 1 · Liste déroulante dynamique")
r = ligne(ws, r, "Objectif", "Proposer une liste de choix dans une colonne. La liste s'allonge toute seule quand on ajoute une valeur.")
r = ligne(ws, r, "1. La liste", "Dans la feuille Paramètres, utiliser la première colonne libre. Titre en haut, valeurs en dessous, "
          "une par cellule, sans ligne vide.")
r = ligne(ws, r, "2. Le nom", "Formules > Gestionnaire de noms > Nouveau. Nom : L_MaListe (préfixe L_, sans espace ni accent). "
          "Fait référence à :")
r = ligne(ws, r, "Formule (FR)", "=DECALER(Paramètres!$AC$3;0;0;MAX(1;NBVAL(Paramètres!$AC$3:$AC$42));1)", "formule")
r = ligne(ws, r, "Version EN", "=OFFSET(Paramètres!$AC$3,0,0,MAX(1,COUNTA(Paramètres!$AC$3:$AC$42)),1)", "en")
r = ligne(ws, r, "3. Le menu", "Sélectionner les cellules cibles > Données > Validation des données > Autoriser : Liste > "
          "Source : =L_MaListe. Facultatif : un message d'aide dans l'onglet « Message de saisie ».")
r = ligne(ws, r, "Comment ça marche", "DECALER part de la 1re valeur et prend autant de lignes qu'il y a de valeurs remplies (NBVAL), "
          "jusqu'à 40. Le MAX(1;…) évite une erreur quand la liste est vide.")
r += 1
r = joker(ws, r)
libelle(ws, "B%d" % r, "Choisis un statut")
saisie(ws, "C%d" % r)
liste(ws, "C%d" % r, "L_Statuts", "Liste L_Statuts (Paramètres, colonne G)")
note(ws, "D%d" % r, "← menu alimenté par L_Statuts (Paramètres, colonne G)")
r += 1
libelle(ws, "B%d" % r, "Choisis dans « Ma liste »")
saisie(ws, "C%d" % r)
liste(ws, "C%d" % r, "L_Essai", "Liste L_Essai (Paramètres, colonne K)")
note(ws, "D%d" % r, "← ajoute « Choix 4 » dans Paramètres, colonne K : il apparaît ici")
r += 1
libelle(ws, "B%d" % r, "Taille de « Ma liste »")
calcul(ws, "C%d" % r, '=COUNTA(Paramètres!$K$4:$K$43)&" valeur(s)"', gras=True)
note(ws, "D%d" % r, "← se met à jour quand tu ajoutes ou supprimes une valeur")
r += 1
libelle(ws, "B%d" % r, "Choisis un client")
saisie(ws, "C%d" % r, gauche=True)
ws.merge_cells("C%d:D%d" % (r, r))
liste(ws, "C%d" % r, "L_Clients", "Liste L_Clients (Clients démo, colonne C)")
note(ws, "E%d" % r, "← liste L_Clients, tirée de « Clients démo »")
r += 2
r = bloc(ws, r, "Pièges")
r = ligne(ws, r, "Cellule vide", "Une cellule vide au milieu de la liste coupe la fin de la liste (NBVAL compte les cellules "
          "remplies, DECALER prend ce nombre de lignes depuis le haut).", "note")
r = ligne(ws, r, "Plus de 40 valeurs", "Remplacer $AC$42 par une ligne plus basse.", "note")
r = ligne(ws, r, "Mes fichiers", "Tableau de Bord 2026 : L_Clients, L_Paie, L_Logiciels, L_OuiNon… (Paramètres, colonnes S à AB).")

# ------------------------------------------------------------------ Fiches 2 et 3 : Dates
ws = feuille(wb, "Dates", "Thème : Dates", "F6C98D")
r = bloc(ws, 3, "Fiche 2 · Période manquante par rapport au mois précédent")
r = ligne(ws, r, "Objectif", "À partir du dernier mois saisi, ne rien afficher si on est à jour. Sinon, afficher le ou les "
          "mois manquants jusqu'au mois précédent.")
r = ligne(ws, r, "Formule (FR)", '=SI(C3="";"";SI(ANNEE(C3)*12+MOIS(C3)>=ANNEE(AUJOURDHUI())*12+MOIS(AUJOURDHUI())-1;"";'
          'SI(ANNEE(C3)*12+MOIS(C3)=ANNEE(AUJOURDHUI())*12+MOIS(AUJOURDHUI())-2;TEXTE(MOIS.DECALER(C3;1);"mmmm aaaa");'
          'TEXTE(MOIS.DECALER(C3;1);"mmmm aaaa")&" à "&TEXTE(FIN.MOIS(AUJOURDHUI();-1);"mmmm aaaa"))))', "formule")
r = ligne(ws, r, "Version EN", '=IF(C3="","",IF(YEAR(C3)*12+MONTH(C3)>=YEAR(TODAY())*12+MONTH(TODAY())-1,"",'
          'IF(YEAR(C3)*12+MONTH(C3)=YEAR(TODAY())*12+MONTH(TODAY())-2,TEXT(EDATE(C3,1),"mmmm yyyy"),'
          'TEXT(EDATE(C3,1),"mmmm yyyy")&" à "&TEXT(EOMONTH(TODAY(),-1),"mmmm yyyy"))))', "en")
r = ligne(ws, r, "L'astuce à retenir", "ANNEE(date)*12+MOIS(date) transforme une date en « numéro de mois ». Comparer deux "
          "numéros revient à comparer des mois sans s'occuper des jours ; la soustraction donne l'écart en mois.")
r = ligne(ws, r, "Fonctions utiles", "MOIS.DECALER(date;n) : même jour, n mois plus tard · FIN.MOIS(date;-1) : dernier jour "
          "du mois précédent · TEXTE(date;\"mmmm aaaa\") : affiche « octobre 2026 ».")
r += 1
r = joker(ws, r)
libelle(ws, "B%d" % r, "Date du jour utilisée")
calcul(ws, "C%d" % r, "=E_Ref", DATE, gras=True)
note(ws, "D%d" % r, "← modifiable dans Paramètres (C5) pour simuler une autre date")
r += 1
r = entetes(ws, r, ["Dernier mois saisi", "N° de mois (tech.)", "Écart en mois", "Résultat"])
ws.merge_cells("E%d:H%d" % (r - 1, r - 1))
for d in (date(2026, 10, 1), date(2026, 9, 1), date(2026, 7, 1), date(2025, 12, 1), None):
    saisie(ws, "B%d" % r, d, MOIS)
    calcul(ws, "C%d" % r, '=IF(B{r}="","",YEAR(B{r})*12+MONTH(B{r}))'.format(r=r)).font = font(9, TECH, i=True)
    calcul(ws, "D%d" % r, '=IF(B{r}="","",YEAR(E_Ref)*12+MONTH(E_Ref)-C{r})'.format(r=r))
    ws.merge_cells("E%d:H%d" % (r, r))
    calcul(ws, "E%d" % r, "=" + P.format(r=r).replace("E%d" % r, "B%d" % r), gras=True)
    for k in range(6, 9):
        ws.cell(r, k).border = CADRE
    r += 1
note(ws, "B%d" % r, "Écart de 1 = à jour (on saisit le mois précédent). Dans « Clients démo », colonne F.")
r += 1
r = bloc(ws, r, "Fiche 3 · Reporter une date au jour ouvré suivant")
r = ligne(ws, r, "Objectif", "Une échéance qui tombe un samedi, un dimanche ou un jour férié est reportée au jour ouvré suivant.")
r = ligne(ws, r, "Formule (FR)", "=SERIE.JOUR.OUVRE(A2-1;1;Feries)", "formule")
r = ligne(ws, r, "Version EN", "=WORKDAY(A2-1,1,Feries)", "en")
r = ligne(ws, r, "Comment ça marche", "On recule d'un jour, puis on avance d'un jour ouvré. Si A2 est déjà un jour ouvré, on "
          "retombe dessus ; sinon, on obtient le jour ouvré suivant.")
r = ligne(ws, r, "Prérequis", "Une liste des jours fériés nommée (ici E_Feries, Paramètres colonne M, 2026-2030).")
r += 1
r = joker(ws, r)
r = entetes(ws, r, ["Échéance théorique", "Jour (tech.)", "Échéance reportée", "Jour", "Reportée ?"])
for d in (date(2026, 11, 28), date(2026, 12, 25), date(2026, 12, 15), date(2027, 5, 1), date(2027, 7, 14), None):
    saisie(ws, "B%d" % r, d, DATE)
    calcul(ws, "C%d" % r, '=IF(B{r}="","",TEXT(B{r},"jjjj"))'.format(r=r)).font = font(9, TECH, i=True)
    calcul(ws, "D%d" % r, '=IF(B{r}="","",WORKDAY(B{r}-1,1,E_Feries))'.format(r=r), DATE, gras=True)
    calcul(ws, "E%d" % r, '=IF(D{r}="","",TEXT(D{r},"jjjj"))'.format(r=r))
    calcul(ws, "F%d" % r, '=IF(B{r}="","",IF(D{r}<>B{r},"Oui","Non"))'.format(r=r))
    r += 1
regle_mfc(ws, "F%d:F%d" % (r - 6, r - 1), '$F%d="Oui"' % (r - 6), "FFE2B8")
note(ws, "B%d" % r, "Dans « Clients démo », colonne H.")
r += 2
r = ligne(ws, r, "Mes fichiers", "Fiche 2 : Tableau de Bord 2026, feuille Salaires, colonne D. Fiche 3 : feuille Échéances, "
          "fériés dans « Réf. échéances » (nom E_Feries).")

# ------------------------------------------------------------------ Fiche 4 : Navigation
ws = feuille(wb, "Navigation", "Thème : Navigation", "CDB8E8")
r = bloc(ws, 3, "Fiche 4 · Bouton pour aller sur une autre feuille")
r = ligne(ws, r, "Objectif", "Une cellule cliquable qui emmène sur une autre feuille, par exemple un bouton « ◄ Accueil ».")
r = ligne(ws, r, "Formule (FR)", '=LIEN_HYPERTEXTE("#\'Accueil\'!A1";"◄ Accueil")', "formule")
r = ligne(ws, r, "Version EN", '=HYPERLINK("#\'Accueil\'!A1","◄ Accueil")', "en")
r = ligne(ws, r, "Règles", "Le # indique « dans ce classeur ». Le nom de la feuille va entre apostrophes s'il contient un "
          "espace : \"#'Suivi échéances'!A1\". Aspect de bouton : couleur de fond, texte gras et centré, bordure "
          "blanche, ligne plus haute.")
r += 1
r = joker(ws, r)
ws.row_dimensions[r].height = 40
bouton(ws, "B%d" % r, "Clients démo", "Aller sur Clients démo ►", "A8D5BA", "1E4D2B", 12)
bouton(ws, "D%d" % r, "Paramètres", "Aller sur Paramètres ►", "C9CED6", MARINE, 12)
ws.merge_cells("D%d:E%d" % (r, r))
bouton(ws, "G%d" % r, "Dates", "Fiche 2 ►", "F6C98D", "843C0C", 12)
ws.merge_cells("G%d:H%d" % (r, r))
r += 1
note(ws, "B%d" % r, "Clique sur un bouton, puis reviens avec « ◄ Accueil » (en haut de chaque feuille).")
r += 2
libelle(ws, "B%d" % r, "Fabrique ton lien : feuille")
saisie(ws, "C%d" % r, "Mise en forme", gauche=True)
ws.merge_cells("C%d:D%d" % (r, r))
note(ws, "E%d" % r, "← tape le nom exact d'une feuille")
r += 1
libelle(ws, "B%d" % r, "Texte du bouton")
saisie(ws, "C%d" % r, "Voir la fiche 5", gauche=True)
ws.merge_cells("C%d:D%d" % (r, r))
r += 1
libelle(ws, "B%d" % r, "Résultat")
ws.row_dimensions[r].height = 30
c = ws["C%d" % r]
c.value = '=HYPERLINK("#\'"&C{a}&"\'!A1",C{b})'.format(a=r - 2, b=r - 1)
c.font, c.fill = font(11, "FFFFFF", b=True), fill(MARINE)
c.alignment = Alignment(horizontal="center", vertical="center")
ws.merge_cells("C%d:D%d" % (r, r))
r += 2
r = ligne(ws, r, "Mes fichiers", "Tableau de Bord 2026 : page Accueil et bouton de retour sur chaque feuille. Le skill "
          "« accueil-boutons-excel » le fait automatiquement sur n'importe quel classeur.")

# ------------------------------------------------------------------ Fiche 5 : Mise en forme
ws = feuille(wb, "Mise en forme", "Thème : Mise en forme", "A7DCE0")
r = bloc(ws, 3, "Fiche 5 · Couleur selon un statut (mise en forme conditionnelle)")
r = ligne(ws, r, "Objectif", "Colorer une cellule ou une ligne entière selon une valeur, par exemple rouge si « URGENT ».")
r = ligne(ws, r, "Étapes", "1. Sélectionner la plage (ex. A3:W62).  2. Accueil > Mise en forme conditionnelle > Nouvelle "
          "règle > « Utiliser une formule… ».  3. Saisir la formule pour la 1re ligne de la plage.  4. Format… > couleur.")
r = ligne(ws, r, "Formule (FR = EN)", '=$W3="URGENT"', "formule")
r = ligne(ws, r, "La règle du $", "$W3 : colonne bloquée, la ligne s'adapte → toute la ligne prend la couleur selon la "
          "colonne W.   W3 sans $ : chaque cellule teste sa propre colonne.")
r += 1
r = joker(ws, r)
note(ws, "B%d" % r, "Avec $ : règle =$E%d=\"URGENT\" sur B:E → toute la ligne" % (r + 2))
note(ws, "F%d" % r, "Sans $ : règle =F%d=\"URGENT\" sur F:H → la cellule seule" % (r + 2))
r += 1
entetes(ws, r, ["Client", "Dossier", "Échéance", "Statut"])
entetes(ws, r, ["Statut A", "Statut B", "Statut C"], 6)
r += 1
r0 = r
data = [("Client Alpha", "TVA", date(2026, 11, 25), "URGENT", "Planifié", "URGENT", "Terminé"),
        ("Client Bêta", "Bilan", date(2026, 11, 27), "À venir", "URGENT", "En cours", "À venir"),
        ("Client Gamma", "CFE", date(2026, 12, 15), "Planifié", "Terminé", "À venir", "URGENT"),
        ("Client Delta", "Paie", date(2026, 11, 30), "URGENT", "En cours", "Planifié", "Planifié")]
for d in data:
    saisie(ws, "B%d" % r, d[0], gauche=True)
    saisie(ws, "C%d" % r, d[1])
    saisie(ws, "D%d" % r, d[2], DATE)
    for k, v in enumerate(d[3:]):
        saisie(ws, "%s%d" % ("EFGH"[k], r), v)
    r += 1
liste(ws, "E%d:H%d" % (r0, r - 1), "L_Statuts")
regle_mfc(ws, "B%d:E%d" % (r0, r - 1), '$E%d="URGENT"' % r0, "F6C7CC", "9C2B38")
regle_mfc(ws, "F%d:H%d" % (r0, r - 1), 'F%d="URGENT"' % r0, "F6C7CC", "9C2B38")
regle_mfc(ws, "D%d:D%d" % (r0, r - 1), "AND(ISNUMBER(D{0}),D{0}-E_Ref<=E_Urgent)".format(r0), "FFE2B8", None, True)
note(ws, "B%d" % r, "Change un statut en URGENT (ou l'inverse) pour voir la différence. Colonne Échéance : variante avec une date ↓")
r += 2
r = ligne(ws, r, "Variante date (FR)", "=ET(ESTNUM(K3);K3-AUJOURDHUI()<=7)", "formule")
r = ligne(ws, r, "Version EN", "=AND(ISNUMBER(K3),K3-TODAY()<=7)", "en")
r = ligne(ws, r, "Ici", "La colonne Échéance ci-dessus est en orange si la date tombe dans les E_Urgent jours (Paramètres, C7).", "note")
r = ligne(ws, r, "Couleurs cabinet", "URGENT / En retard : rose F6C7CC + texte rouge · À venir : orange FFE2B8 · En cours : "
          "bleu D3E4F7 · Planifié / Terminé : vert CDEBD5 · Non concerné : gris E4E7EB.", "note")
r = ligne(ws, r, "Mes fichiers", "Tableau de Bord 2026 : statuts URGENT / À venir / Planifié dans Échéances et Planning. "
          "Ici : « Clients démo », colonne I.")

# ------------------------------------------------------------------ Fiche 6 : Calcul
ws = feuille(wb, "Calcul", "Thème : Calcul", "F2DC8C")
r = bloc(ws, 3, "Fiche 6 · Compter selon plusieurs critères")
r = ligne(ws, r, "Objectif", "Compter les lignes qui remplissent plusieurs conditions à la fois.")
r = ligne(ws, r, "Formule (FR)", '=NB.SI.ENS(Échéances!$W$3:$W$62;"URGENT";Échéances!$Z$3:$Z$62;"Oui")', "formule")
r = ligne(ws, r, "Version EN", '=COUNTIFS(Échéances!$W$3:$W$62,"URGENT",Échéances!$Z$3:$Z$62,"Oui")', "en")
r = ligne(ws, r, "Usage", "Les critères vont par paires (plage ; critère), autant que nécessaire. Pour une comparaison, "
          "l'opérateur va entre guillemets : \">=\"&AUJOURDHUI().  Proches : SOMME.SI.ENS (additionner), "
          "MOYENNE.SI.ENS (moyenne).")
r += 1
r = joker(ws, r)
note(ws, "B%d" % r, "Les calculs portent sur le tableau « Clients démo ».")
r += 1
libelle(ws, "B%d" % r, "Statut")
saisie(ws, "C%d" % r, "URGENT")
liste(ws, "C%d" % r, "L_Statuts")
rs = r
r += 1
libelle(ws, "B%d" % r, "Envoyé")
saisie(ws, "C%d" % r, "Non")
liste(ws, "C%d" % r, "L_OuiNon")
re_ = r
r += 2
r = entetes(ws, r, ["Calcul", "Résultat", "Formule utilisée (FR)"])
ws.merge_cells("D%d:H%d" % (r - 1, r - 1))
calc6 = [("Nombre (Statut + Envoyé)", "=COUNTIFS({I},C%d,{J},C%d)" % (rs, re_), None,
          '=NB.SI.ENS(Statut;C%d;Envoyé;C%d)' % (rs, re_)),
         ("Total honoraires", "=SUMIFS({K},{I},C%d,{J},C%d)" % (rs, re_), EURO,
          '=SOMME.SI.ENS(Honoraires;Statut;C%d;Envoyé;C%d)' % (rs, re_)),
         ("Honoraires moyens", '=IFERROR(AVERAGEIFS({K},{I},C%d,{J},C%d),"aucun")' % (rs, re_), EURO,
          '=SIERREUR(MOYENNE.SI.ENS(Honoraires;Statut;C%d;Envoyé;C%d);"aucun")' % (rs, re_)),
         ("Échéances à partir d'aujourd'hui", '=COUNTIFS({H},">="&E_Ref,{I},C%d)' % rs, None,
          '=NB.SI.ENS(Échéance reportée;">="&AUJOURDHUI();Statut;C%d)' % rs)]
cols = {k: CL % (k, k) for k in "HIJK"}
for lib, f, fmt, ffr in calc6:
    libelle(ws, "B%d" % r, lib)
    calcul(ws, "C%d" % r, f.format(**cols), fmt, gras=True)
    ws.merge_cells("D%d:H%d" % (r, r))
    c = ws["D%d" % r]
    texte(c, ffr)
    c.font, c.fill, c.alignment = font(10, SOUS), fill(F_FORMULE), Alignment(vertical="center")
    for k in range(4, 9):
        ws.cell(r, k).border = CADRE
    r += 1
note(ws, "B%d" % r, "Les noms Statut, Envoyé… dans la colonne « Formule » remplacent les plages 'Clients démo'!$I$4:$I$43, etc.")
r += 2
r = bloc(ws, r, "Compteur par statut (comme l'Accueil du Tableau de Bord)")
r = entetes(ws, r, ["Statut", "Dossiers", "Dont envoyés", "Honoraires"])
for k in range(5):
    calcul(ws, "B%d" % r, '=IF(Paramètres!G%d="","",Paramètres!G%d)' % (4 + k, 4 + k), gras=True)
    calcul(ws, "C%d" % r, '=IF(B{r}="","",COUNTIFS({I},B{r}))'.format(r=r, **cols))
    calcul(ws, "D%d" % r, '=IF(B{r}="","",COUNTIFS({I},B{r},{J},"Oui"))'.format(r=r, **cols))
    calcul(ws, "E%d" % r, '=IF(B{r}="","",SUMIFS({K},{I},B{r}))'.format(r=r, **cols), EURO)
    r += 1
mfc_statuts(ws, "B%d:B%d" % (r - 5, r - 1), "$B%d" % (r - 5))
libelle(ws, "B%d" % r, "Total")
calcul(ws, "C%d" % r, "=SUM(C%d:C%d)" % (r - 5, r - 1), gras=True)
calcul(ws, "D%d" % r, "=SUM(D%d:D%d)" % (r - 5, r - 1), gras=True)
calcul(ws, "E%d" % r, "=SUM(E%d:E%d)" % (r - 5, r - 1), EURO, gras=True)
r += 2
r = ligne(ws, r, "Mes fichiers", "Tableau de Bord 2026 : compteurs de l'Accueil et du Planning.")

# ------------------------------------------------------------------ Fiche 7 : Recherche
ws = feuille(wb, "Recherche", "Thème : Recherche", "9CC7E8")
r = bloc(ws, 3, "Fiche 7 · Rechercher une valeur dans un tableau")
r = ligne(ws, r, "Objectif", "Retrouver une information d'un client (par exemple sa date de clôture) à partir de son nom.")
r = ligne(ws, r, "Formule (FR)", "=INDEX(Paramètres!$H$3:$H$62;EQUIV(B3;Paramètres!$B$3:$B$62;0))", "formule")
r = ligne(ws, r, "Version EN", "=INDEX(Paramètres!$H$3:$H$62,MATCH(B3,Paramètres!$B$3:$B$62,0))", "en")
r = ligne(ws, r, "Comment ça marche", "EQUIV trouve à quelle ligne se trouve le client ; INDEX va chercher la valeur de cette "
          "ligne dans la colonne voulue. Le 0 impose une correspondance exacte.")
r = ligne(ws, r, "Avantages", "Par rapport à RECHERCHEV : la colonne cherchée peut être à gauche, et la formule ne casse pas "
          "si on insère une colonne. Si le client n'existe pas : entourer de SIERREUR(…;\"\").")
r += 1
r = joker(ws, r)
libelle(ws, "B%d" % r, "Client")
saisie(ws, "C%d" % r, "Client Gamma", gauche=True)
ws.merge_cells("C%d:D%d" % (r, r))
liste(ws, "C%d" % r, "L_Clients", "Choisis un client, ou tape un nom qui n'existe pas")
note(ws, "E%d" % r, "← choisis dans la liste ou tape un nom inconnu")
rc = r
r += 1
libelle(ws, "B%d" % r, "Ligne trouvée (EQUIV)")
calcul(ws, "C%d" % r, '=IFERROR(MATCH(C{0},{1},0),"introuvable")'.format(rc, CL % ("C", "C"))).font = font(9, TECH, i=True)
note(ws, "D%d" % r, "← position du client dans la colonne Client (tech.)")
r += 2
r = entetes(ws, r, ["Information", "Résultat", "Colonne lue (INDEX)"])
ws.merge_cells("D%d:E%d" % (r - 1, r - 1))
for lib, col, fmt in (("Forme juridique", "B", None), ("Date de clôture", "D", DATE), ("Échéance reportée", "H", DATE),
                      ("Statut", "I", None), ("Honoraires", "K", EURO)):
    libelle(ws, "B%d" % r, lib)
    calcul(ws, "C%d" % r, '=IFERROR(INDEX({0},MATCH($C${1},{2},0)),"")'.format(CL % (col, col), rc, CL % ("C", "C")),
           fmt, gras=True)
    ws.merge_cells("D%d:E%d" % (r, r))
    c = calcul(ws, "D%d" % r, "Clients démo, colonne %s" % col)
    c.font = font(10, SOUS)
    ws.cell(r, 5).border = CADRE
    r += 1
note(ws, "B%d" % r, "Avec SIERREUR, un client inconnu donne des cellules vides au lieu d’un message d’erreur.")
r += 2
r = ligne(ws, r, "Mes fichiers", "Tableau de Bord 2026 : fiche client du Planning.")

# ------------------------------------------------------------------ Fiche 8 : Astuces
ws = feuille(wb, "Astuces", "Thème : Astuces", "E8C547")
r = bloc(ws, 3, "Fiche 8 · Raccourcis et réglages utiles")
r = entetes(ws, r, ["Action", "Comment"])
ws.merge_cells("C%d:H%d" % (r - 1, r - 1))
ws.column_dimensions["B"].width = 40
for a, b in [("Mettre la date du jour dans une cellule", "Ctrl + ;"),
             ("Changer de feuille (même avec les onglets masqués)", "Ctrl + Page suiv. / Page préc."),
             ("Réafficher les onglets du classeur", "Fichier > Options > Options avancées > « Afficher les onglets de classeur »"),
             ("Voir et modifier les noms (L_…, E_…)", "Formules > Gestionnaire de noms (Ctrl + F3)"),
             ("Afficher les formules au lieu des résultats", "Ctrl + \" (touche ² sur certains claviers)"),
             ("Figer les titres en haut", "Affichage > Figer les volets")]:
    libelle(ws, "B%d" % r, a)
    ws.merge_cells("C%d:H%d" % (r, r))
    calcul(ws, "C%d" % r, b, gras=True, gauche=True)
    for k in range(4, 9):
        ws.cell(r, k).border = CADRE
    ws.row_dimensions[r].height = 22
    r += 1
r += 1
r = joker(ws, r)
libelle(ws, "B%d" % r, "Essaie Ctrl + ; ici")
saisie(ws, "C%d" % r, None, DATE)
r += 1
libelle(ws, "B%d" % r, "Essaie Ctrl + \" puis regarde")
calcul(ws, "C%d" % r, "=E_Ref", DATE)
note(ws, "D%d" % r, "← la formule apparaît à la place de la date ; Ctrl + \" à nouveau pour revenir")
r += 1
libelle(ws, "B%d" % r, "Ctrl + F3")
note(ws, "C%d" % r, "Ouvre la liste des noms de ce classeur : L_Statuts, L_Clients, E_Ref, E_Feries…")

# ------------------------------------------------------------------ Accueil
ws = accueil
ws.sheet_properties.tabColor = MARINE
ws.sheet_view.showGridLines = False
ws.sheet_view.showRowColHeaders = False
for col, w in zip("ABCDEF", (45, 30, 2, 30, 2, 30)):
    ws.column_dimensions[col].width = w


def bande(r, val, sz, fond, encre, h, b=True):
    ws.merge_cells("B%d:F%d" % (r, r))
    c = ws["B%d" % r]
    c.value = val
    c.font = font(sz, encre, b=b)
    if fond:
        c.fill = fill(fond)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = h


bande(1, "BOÎTE À OUTILS EXCEL", 18, F_BANDEAU, MARINE, 38)
bande(2, '="Situation au "&RIGHT("0"&DAY(E_Ref),2)&"/"&RIGHT("0"&MONTH(E_Ref),2)&"/"&YEAR(E_Ref)', 10, None, SOUS, 18, False)
bande(3, "Les 8 fiches du référentiel, chacune avec un exemple qui fonctionne.", 10, None, AIDE, 18, False)
GROUPES = [("Saisie et dates", ("BDD7EE", "DDEBF7", "1F4E79"),
            [("Saisie", "1 · Liste déroulante dynamique"), ("Dates", "2 · Période manquante\n3 · Jour ouvré suivant")]),
           ("Présentation", ("C6E0B4", "E2EFDA", "375623"),
            [("Navigation", "4 · Bouton vers une feuille"), ("Mise en forme", "5 · Couleur selon un statut")]),
           ("Calculs et recherches", ("D9D2E9", "EDE7F6", "5B3F86"),
            [("Calcul", "6 · Compter selon plusieurs critères"), ("Recherche", "7 · Rechercher une valeur"),
             ("Astuces", "8 · Raccourcis et réglages")]),
           ("Exemple complet et réglages", ("FFE699", "FFF2CC", "7F6000"),
            [("Clients démo", "Clients démo\n(toutes les fiches)"), ("Paramètres", "Paramètres\n(listes, fériés, date)")])]
r = 5
for nom, (h, b, ink), btns in GROUPES:
    bande(r, nom.upper(), 11, h, ink, 26)
    r += 1
    for k in range(0, len(btns), 3):
        for col, (sh, lab) in zip("BDF", btns[k:k + 3]):
            c = ws["%s%d" % (col, r)]
            c.value = '=HYPERLINK("#\'%s\'!A1","%s")' % (sh, lab.replace("\n", '"&CHAR(10)&"'))
            c.font = font(12, ink, b=True)
            c.fill = fill(b)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            m = Side(style="medium", color="FFFFFF")
            c.border = Border(left=m, right=m, top=m, bottom=m)
        ws.row_dimensions[r].height = 46
        ws.row_dimensions[r + 1].height = 8
        r += 2
    ws.row_dimensions[r].height = 12
    r += 1
bande(r, "MODE D'EMPLOI", 11, F_ENTETE, MARINE, 22)
r += 1
for t in ["Clique sur un bouton pour ouvrir une fiche ; « ◄ Accueil » en haut de chaque feuille pour revenir.",
          "Jaune pâle = à modifier pour essayer · blanc = calculé · bleu = titres et libellés.",
          "Chaque fiche donne la formule en français (à recopier dans Excel FR) et sa version anglaise.",
          "Onglets masqués : Ctrl + Page suiv. / Page préc. pour passer d'une feuille à l'autre."]:
    bande(r, t, 10, None, AIDE, 17, False)
    r += 1
r += 1
bande(r, "JOURNAL DES AJOUTS", 11, F_ENTETE, MARINE, 22)
r += 1
ws.merge_cells("C%d:F%d" % (r, r))
for ref, v in (("B%d" % r, "23/11/2026"), ("C%d" % r, "Création : fiches 1 à 8 et tableau Clients démo")):
    ws[ref].value = v
    ws[ref].font = font(10, TXT)
    ws[ref].alignment = Alignment(horizontal="center" if ref[0] == "B" else "left", vertical="center")
    ws[ref].border = CADRE
ws.print_area = "B1:F%d" % r
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = 9

# ------------------------------------------------------------------ finitions
from openpyxl.worksheet.properties import PageSetupProperties
for s in wb.worksheets:
    s.sheet_view.zoomScale = 100
    s.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    s.page_setup.fitToWidth, s.page_setup.fitToHeight = 1, 0
    s.page_setup.paperSize = 9
    if s.title in ("Clients démo", "Paramètres"):
        s.page_setup.orientation = "landscape"
wb.active = 0
wb.views[0].showSheetTabs = False
wb.calculation.fullCalcOnLoad = True
wb.save(SORTIE)
print("OK ->", SORTIE)
