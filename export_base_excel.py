"""
NAP SARL — Export « Base Excel » : un seul classeur avec TOUTES les données reliées
(déchargements, paiements, bons de commande, fournisseurs, journal) pour travailler hors du site.

À placer à la RACINE du projet (à côté de dechargement_data.py).
Utilisé par la page 7_Export_Excel.py
"""
from datetime import datetime
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

import dechargement_data as dd

LIGNES_RELEVE = 400          # nb de lignes du relevé fournisseur (par fournisseur)
DERNIERE_LIGNE = 5000        # plage lue dans la feuille Dechargements pour le relevé
BLEU = "1F3864"
FMT_DATE, FMT_NB, FMT_PCT = "DD/MM/YYYY", "#,##0", "0.0%"


# ---------- Outils ----------
def _v(x):
    """Valeur pandas -> valeur Excel propre."""
    if x is None:
        return None
    try:
        if pd.isna(x):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(x, pd.Timestamp):
        return x.to_pydatetime()
    if hasattr(x, "item"):
        return x.item()
    return x


def _table(ws, nom, entetes, lignes, formats=None, largeurs=None):
    """Écrit un tableau Excel nommé (en-têtes ligne 1) et le met en forme."""
    formats, largeurs = formats or {}, largeurs or {}
    ws.append(entetes)
    for l in lignes:
        ws.append(l)
    if not lignes:
        ws.append([None] * len(entetes))
    n = max(len(lignes), 1)
    t = Table(displayName=nom, ref=f"A1:{get_column_letter(len(entetes))}{n + 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    for j, h in enumerate(entetes, start=1):
        if h in formats:
            for r in range(2, n + 2):
                ws.cell(r, j).number_format = formats[h]
        ws.column_dimensions[get_column_letter(j)].width = largeurs.get(h, min(max(12, len(h) + 3), 28))
    ws.freeze_panes = "A2"


def _rassembler():
    """Charge TOUS les produits (le produit actif de la session est restauré)."""
    ancien = dd.get_produit()
    dechs, affs, bons = [], [], []
    try:
        for p in dd.PRODUITS:
            dd.set_produit(p)
            d, a = dd.charger_complet()
            d, a = d.copy(), a.copy()
            d["Produit"], a["Produit"] = p, p
            dechs.append(d)
            affs.append(a)
            bons.append(dd.charger_bons_commande().assign(Produit_Code=p))
    finally:
        dd.set_produit(ancien)

    def union(frames):
        pleins = [f for f in frames if not f.empty]
        return pd.concat(pleins, ignore_index=True) if pleins else frames[0]

    return union(dechs), union(affs), union(bons)


# ---------- Construction ----------
def construire_base(utilisateur=""):
    """Retourne les octets du classeur .xlsx."""
    dech, aff, bons = _rassembler()
    journal = dd.charger_journal()

    # Fournisseurs : fournisseurs.xlsx + ceux des déchargements + ceux des bons
    frs = dict(dd.tous_les_fournisseurs(dech))
    nom_vers_cle = {}
    for k, n in frs.items():
        nom_vers_cle.setdefault(n.casefold(), k)

    def cle_bon(code, nom):
        if code and code in frs:
            return code
        k = nom_vers_cle.get(nom.casefold())
        if k:
            return k
        if nom:
            frs[nom] = nom
            nom_vers_cle[nom.casefold()] = nom
            return nom
        return ""

    dech = dech.sort_values(["Produit", "Date_Dechargement", "Transaction"]).reset_index(drop=True)
    dech["ID_Fournisseur"] = [c or n for c, n in zip(dech["Code_Frs"], dech["Fournisseur"])]
    # (liste en Python : fonctionne quel que soit le type texte de pandas)
    dech["ID_Livraison"] = [f"{p}-{t}" for p, t in zip(dech["Produit"], dech["Transaction"])]
    bons = bons.sort_values("Reference").reset_index(drop=True)
    bons["ID_Fournisseur"] = [cle_bon(c, n) for c, n in zip(bons["Code_Fournisseur"], bons["Nom_Fournisseur"])]

    wb = Workbook()
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    ws_lisez = wb.active
    ws_lisez.title = "LISEZ_MOI"
    ws_rel = wb.create_sheet("Releve_Fournisseur")
    ws_d = wb.create_sheet("Dechargements")
    ws_p = wb.create_sheet("Paiements")
    ws_l = wb.create_sheet("Liens_Bons")
    ws_b = wb.create_sheet("Bons_Commande")
    ws_f = wb.create_sheet("Fournisseurs")
    ws_j = wb.create_sheet("Journal_Bons")

    # ---------------- Dechargements ----------------
    ent_d = ["ID_Livraison", "Produit", "Transaction", "Date_Dechargement", "Ref_BC", "ID_Fournisseur", "Fournisseur",
             "N_Camion", "Camion_BC", "Alerte_Camion", "Weighment_Ticket", "Poids_Net_KG", "Nb_Sacs", "PU_FCFA_KG",
             "Montant_Du", "Montant_Paye", "Solde", "Statut_Paiement", "Bons_Paiement", "Lieu_Stockage",
             "Statut_Livraison", "Remarque"]
    lig_d = []
    for i, r in enumerate(dech.itertuples(index=False), start=2):
        lig_d.append([
            r.ID_Livraison, r.Produit, _v(r.Transaction), _v(r.Date_Dechargement), r.Ref_BC or None, r.ID_Fournisseur,
            r.Fournisseur, r.N_Camion or None, r.Camion_BC or None, "⚠ Différent" if r.Camion_Diff else None,
            r.Weighment_Ticket or None, _v(r.Poids_Net_KG), _v(r.Nb_Sacs), _v(r.PU_FCFA_KG),
            f"=ROUND(L{i}*N{i},0)",
            f"=SUMIFS(Paiements!$G:$G,Paiements!$B:$B,A{i})",
            f"=O{i}-P{i}",
            f'=IF(ROUND(Q{i},0)=0,"PAYE",IF(ROUND(Q{i},0)<0,"SURPLUS",IF(P{i}<=0,"NON PAYE","PARTIEL")))',
            r.Bons_Paiement or None, r.Lieu_Stockage or None, r.Statut_Livraison or None, r.Remarque or None])
    _table(ws_d, "tbl_Dechargements", ent_d, lig_d,
           {"Date_Dechargement": FMT_DATE, "Poids_Net_KG": FMT_NB, "Nb_Sacs": FMT_NB, "PU_FCFA_KG": FMT_NB,
            "Montant_Du": FMT_NB, "Montant_Paye": FMT_NB, "Solde": FMT_NB},
           {"ID_Livraison": 16, "Ref_BC": 28, "Fournisseur": 32, "N_Camion": 22, "Camion_BC": 22, "Bons_Paiement": 42,
            "Remarque": 30, "Date_Dechargement": 16})

    # ---------------- Paiements ----------------
    ent_p = ["ID_Paiement", "ID_Livraison", "Produit", "Transaction", "Date_Approbation", "Bons_Paiement", "Montant",
             "Approuve_Par"]
    lig_p, lig_l = [], []
    aff = aff.sort_values(["Produit", "ID"]).reset_index(drop=True) if not aff.empty else aff
    for r in aff.itertuples(index=False):
        id_p = f"{r.Produit}-P{r.ID}"
        id_l = f"{r.Produit}-{r.Transaction}"
        lig_p.append([id_p, id_l, r.Produit, _v(r.Transaction), _v(r.Date_Approbation), r.Bons_Paiement or None,
                      _v(r.Montant), r.Approuve_Par or None])
        for bon in dd.liste_bons_texte(r.Bons_Paiement):
            if not bon.startswith("("):
                lig_l.append([id_p, id_l, bon])
    _table(ws_p, "tbl_Paiements", ent_p, lig_p, {"Date_Approbation": FMT_DATE, "Montant": FMT_NB},
           {"ID_Paiement": 14, "ID_Livraison": 16, "Bons_Paiement": 42, "Date_Approbation": 16})

    # ---------------- Liens_Bons ----------------
    ent_l = ["ID_Paiement", "ID_Livraison", "No_Bon_Paiement", "Montant_Bon_Journal"]
    lig_l = [l + [f'=IFERROR(INDEX(Journal_Bons!$E:$E,MATCH(C{i},Journal_Bons!$A:$A,0)),"")']
             for i, l in enumerate(lig_l, start=2)]
    _table(ws_l, "tbl_Liens_Bons", ent_l, lig_l, {"Montant_Bon_Journal": FMT_NB},
           {"ID_Paiement": 14, "ID_Livraison": 16, "No_Bon_Paiement": 20})

    # ---------------- Bons_Commande ----------------
    ent_b = ["Reference", "Produit", "Date_Emission", "ID_Fournisseur", "Nom_Fournisseur", "Quantite",
             "Qte_Commandee_KG", "Numero_Tracteur", "Chauffeur", "Poids_Livre_KG", "Nb_Camions", "Ecart_KG",
             "Pct_Livre", "Montant_Du", "Montant_Paye", "Solde", "Statut_Livraison"]
    lig_b = []
    for i, r in enumerate(bons.itertuples(index=False), start=2):
        d_em = pd.to_datetime(r.Date_Emission, dayfirst=True, errors="coerce")
        lig_b.append([
            r.Reference, r.Produit_Code, _v(d_em) if pd.notna(d_em) else (r.Date_Emission or None), r.ID_Fournisseur,
            r.Nom_Fournisseur, r.Quantite or None, _v(r.Qte_KG), r.Numero_Tracteur or None, r.Chauffeur or None,
            f"=SUMIFS(Dechargements!$L:$L,Dechargements!$E:$E,A{i})",
            f"=COUNTIFS(Dechargements!$E:$E,A{i})",
            f'=IF(G{i}>0,J{i}-G{i},"")',
            f'=IF(G{i}>0,J{i}/G{i},"")',
            f"=SUMIFS(Dechargements!$O:$O,Dechargements!$E:$E,A{i})",
            f"=SUMIFS(Dechargements!$P:$P,Dechargements!$E:$E,A{i})",
            f"=N{i}-O{i}",
            f'=IF(K{i}=0,"EN ATTENTE",IF(G{i}>0,IF(J{i}>={dd.SEUIL_LIVRE}*G{i},"LIVRE","EN COURS"),"DECHARGE"))'])
    _table(ws_b, "tbl_Bons_Commande", ent_b, lig_b,
           {"Date_Emission": FMT_DATE, "Qte_Commandee_KG": FMT_NB, "Poids_Livre_KG": FMT_NB, "Ecart_KG": FMT_NB,
            "Pct_Livre": FMT_PCT, "Montant_Du": FMT_NB, "Montant_Paye": FMT_NB, "Solde": FMT_NB},
           {"Reference": 30, "Nom_Fournisseur": 30, "Numero_Tracteur": 22, "Chauffeur": 26, "Date_Emission": 16})

    # ---------------- Fournisseurs ----------------
    ent_f = ["ID_Fournisseur", "Nom", "Libelle", "Camions", "Poids_KG", "Total_Du", "Total_Paye", "Reste_A_Payer"]
    lig_f = []
    tri = sorted(frs.items(), key=lambda kv: kv[1].casefold())
    for i, (k, n) in enumerate(tri, start=2):
        lig_f.append([k, n, dd.label_fournisseur(k, n),
                      f"=COUNTIFS(Dechargements!$F:$F,A{i})",
                      f"=SUMIFS(Dechargements!$L:$L,Dechargements!$F:$F,A{i})",
                      f"=SUMIFS(Dechargements!$O:$O,Dechargements!$F:$F,A{i})",
                      f"=SUMIFS(Dechargements!$P:$P,Dechargements!$F:$F,A{i})",
                      f"=F{i}-G{i}"])
    _table(ws_f, "tbl_Fournisseurs", ent_f, lig_f,
           {"Poids_KG": FMT_NB, "Total_Du": FMT_NB, "Total_Paye": FMT_NB, "Reste_A_Payer": FMT_NB},
           {"ID_Fournisseur": 16, "Nom": 34, "Libelle": 44})

    # ---------------- Journal_Bons ----------------
    ent_j = ["No_Bon_Paiement", "Date", "Produit", "Prestataire", "Montant", "Observation", "Mode", "Banque",
             "Rattache_Livraison"]
    lig_j = []
    if not journal.empty:
        jb = (journal.groupby("No_Transaction", sort=False)
              .agg(Date=("Date", "first"), Produit=("Produit", "first"), Prestataire=("Prestataire", "first"),
                   Montant=("Montant", "sum"), Observation=("Observation", "first"), Mode=("Mode", "first"),
                   Banque=("Banque", "first")).reset_index())
        for i, r in enumerate(jb.itertuples(index=False), start=2):
            lig_j.append([r.No_Transaction, _v(r.Date), r.Produit or None, r.Prestataire or None, _v(r.Montant),
                          r.Observation or None, r.Mode or None, r.Banque or None,
                          f'=IF(COUNTIF(Liens_Bons!$C:$C,A{i})>0,"OUI","NON")'])
    _table(ws_j, "tbl_Journal_Bons", ent_j, lig_j, {"Date": FMT_DATE, "Montant": FMT_NB},
           {"No_Bon_Paiement": 20, "Prestataire": 34, "Date": 14})

    # ---------------- Releve_Fournisseur (formules, sans macro) ----------------
    gras = Font(bold=True)
    entete = PatternFill("solid", fgColor=BLEU)
    ws_rel["A1"] = "RELEVÉ DE COMPTE FOURNISSEUR"
    ws_rel["A1"].font = Font(bold=True, size=16, color=BLEU)
    ws_rel["A3"], ws_rel["A4"], ws_rel["A5"] = "Choisir le fournisseur ▶", "ID_Fournisseur", "Nom"
    for c in ("A3", "A4", "A5"):
        ws_rel[c].font = gras
    nb_f = max(len(tri), 1)
    camions = dech.groupby("ID_Fournisseur").size() if not dech.empty else pd.Series(dtype=int)
    defaut = max(tri, key=lambda kv: camions.get(kv[0], 0)) if tri else None
    ws_rel["B3"] = dd.label_fournisseur(*defaut) if defaut else ""
    ws_rel["B3"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws_rel["B3"].font = Font(bold=True)
    dv = DataValidation(type="list", formula1=f"=Fournisseurs!$C$2:$C${nb_f + 1}", allow_blank=False)
    ws_rel.add_data_validation(dv)
    dv.add("B3")
    ws_rel["B4"] = '=IFERROR(INDEX(Fournisseurs!$A:$A,MATCH($B$3,Fournisseurs!$C:$C,0)),"")'
    ws_rel["B5"] = '=IFERROR(INDEX(Fournisseurs!$B:$B,MATCH($B$3,Fournisseurs!$C:$C,0)),"")'
    rng = f"Dechargements!$F$2:$F${DERNIERE_LIGNE}"
    kpis = [("Camions", '=COUNTIFS(Dechargements!$F:$F,$B$4)', FMT_NB),
            ("Poids net (kg)", '=SUMIFS(Dechargements!$L:$L,Dechargements!$F:$F,$B$4)', FMT_NB),
            ("Total dû (FCFA)", '=SUMIFS(Dechargements!$O:$O,Dechargements!$F:$F,$B$4)', FMT_NB),
            ("Total payé (FCFA)", '=SUMIFS(Dechargements!$P:$P,Dechargements!$F:$F,$B$4)', FMT_NB),
            ("Reste à payer (FCFA)", "=C8-D8", FMT_NB)]
    # KPI : libellés ligne 7, valeurs ligne 8 (colonnes A à E ; reste à payer = dû C8 - payé D8)
    for j, (lib, formule, fmt) in enumerate(kpis, start=1):
        lc, vc = ws_rel.cell(7, j), ws_rel.cell(8, j)
        lc.value, lc.font, lc.fill = lib, Font(bold=True, color="FFFFFF"), entete
        lc.alignment = Alignment(horizontal="center", wrap_text=True)
        vc.value, vc.number_format, vc.font = formule, fmt, Font(bold=True, size=13)
        vc.alignment = Alignment(horizontal="center")

    ent_r = ["Ligne source", "Date", "Transaction", "Produit", "N° Bon de commande", "Camion livré", "Poids net (kg)",
             "Dû", "Payé", "Solde ligne", "Solde cumulé", "Statut", "Bon(s) de paiement", "", "(aide)"]
    for j, h in enumerate(ent_r, start=1):
        c = ws_rel.cell(11, j)
        c.value = h or None
        if h:
            c.font, c.fill = Font(bold=True, color="FFFFFF"), entete
            c.alignment = Alignment(horizontal="center", wrap_text=True)
    cols_src = {"B": ("D", FMT_DATE, False), "C": ("C", None, False), "D": ("B", None, True), "E": ("E", None, True),
                "F": ("H", None, True), "G": ("L", FMT_NB, False), "H": ("O", FMT_NB, False),
                "I": ("P", FMT_NB, False), "J": ("Q", FMT_NB, False), "L": ("R", None, True),
                "M": ("S", None, True)}
    for k in range(LIGNES_RELEVE):
        r = 12 + k
        ws_rel[f"O{r}"] = (f'=IF($B$4="","",SUMPRODUCT(SMALL(ROW({rng})+({rng}<>$B$4)*10000000,ROWS(O$12:O{r}))))')
        ws_rel[f"A{r}"] = f'=IF(O{r}="","",IF(O{r}>=10000000,"",O{r}))'
        for col, (src, fmt, texte) in cols_src.items():
            f = f'=IF($A{r}="","",INDEX(Dechargements!${src}:${src},$A{r})' + ('&""' if texte else "") + ")"
            ws_rel[f"{col}{r}"] = f
            if fmt:
                ws_rel[f"{col}{r}"].number_format = fmt
        ws_rel[f"K{r}"] = f'=IF($A{r}="","",SUM(J$12:J{r}))'
        ws_rel[f"K{r}"].number_format = FMT_NB
        ws_rel[f"A{r}"].font = Font(color="999999")
    for col, w in {"A": 14, "B": 34, "C": 16, "D": 16, "E": 28, "F": 22, "G": 15, "H": 15, "I": 15, "J": 15,
                   "K": 16, "L": 12, "M": 46}.items():
        ws_rel.column_dimensions[col].width = w
    ws_rel.column_dimensions["O"].hidden = True
    ws_rel.freeze_panes = "A12"
    ws_rel["A9"] = (f"Le relevé affiche jusqu'à {LIGNES_RELEVE} livraisons par fournisseur "
                    f"(lignes 2 à {DERNIERE_LIGNE} de la feuille Dechargements).")
    ws_rel["A9"].font = Font(italic=True, color="666666")

    # ---------------- LISEZ_MOI ----------------
    now = datetime.now().strftime("%d/%m/%Y à %H:%M")
    lignes = [
        ("NAP SARL — BASE EXCEL RELIÉE (déchargements · paiements · bons · fournisseurs)", "titre"),
        (f"Exportée le {now} par {utilisateur or '—'}. C'est une PHOTO des données du site à cet instant.", None),
        ("⚠ Ce qui est modifié dans ce classeur ne remonte PAS dans le site. Pour mettre à jour : régénérez l'export.", "alerte"),
        ("", None),
        ("FEUILLES", "titre2"),
        ("Releve_Fournisseur : choisissez un fournisseur (cellule jaune B3) → totaux + liste de ses livraisons avec solde cumulé.", None),
        ("Dechargements : 1 ligne = 1 camion (tous produits). Montant dû, payé, solde et statut sont des FORMULES.", None),
        ("Paiements : 1 ligne = 1 paiement approuvé sur une livraison. Ajoutez une ligne ici → le payé et le statut se recalculent.", None),
        ("Liens_Bons : un paiement peut utiliser plusieurs bons de paiement (CHQ-… / CC-…) : 1 ligne par bon.", None),
        ("Bons_Commande : bons émis, avec poids livré, écart et statut de livraison (formules). EN ATTENTE = pas encore livré.", None),
        ("Fournisseurs : liste + camions, poids, dû, payé, reste à payer par fournisseur (formules).", None),
        ("Journal_Bons : bons de paiement du journal « Suivi de paiement » (une ligne par N° de bon). Rattache_Livraison = OUI/NON.", None),
        ("", None),
        ("LIAISONS (clés communes) — utilisables dans Power Query / Power Pivot (Modèle de données › Relations)", "titre2"),
        ("Dechargements[ID_Fournisseur]  →  Fournisseurs[ID_Fournisseur]", None),
        ("Dechargements[Ref_BC]  →  Bons_Commande[Reference]      (anciens numéros « BC-001 » : pas de bon correspondant)", None),
        ("Paiements[ID_Livraison]  →  Dechargements[ID_Livraison]      (ID_Livraison = PRODUIT-Transaction, ex. SHEA-1004)", None),
        ("Liens_Bons[ID_Livraison]  →  Dechargements[ID_Livraison]", None),
        ("Liens_Bons[No_Bon_Paiement]  →  Journal_Bons[No_Bon_Paiement]", None),
        ("Les tableaux portent des noms : tbl_Dechargements, tbl_Paiements, tbl_Liens_Bons, tbl_Bons_Commande, tbl_Fournisseurs, tbl_Journal_Bons.", None),
        ("", None),
        ("CONSEILS", "titre2"),
        ("Ne renommez pas les feuilles ni les en-têtes de colonnes (les formules s'y réfèrent).", None),
        ("Pour ajouter une ligne dans un tableau : copiez la ligne du dessus (les formules suivent) puis changez les valeurs.", None),
        ("Les numéros de transaction peuvent se répéter d'un produit à l'autre : utilisez toujours ID_Livraison comme clé.", None),
    ]
    for i, (txt, style) in enumerate(lignes, start=1):
        c = ws_lisez.cell(i, 1, txt)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if style == "titre":
            c.font = Font(bold=True, size=15, color=BLEU)
        elif style == "titre2":
            c.font = Font(bold=True, size=12, color="FFFFFF")
            c.fill = entete
        elif style == "alerte":
            c.font = Font(bold=True, color="C00000")
    ws_lisez.column_dimensions["A"].width = 130
    ws_lisez.sheet_view.showGridLines = False

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
