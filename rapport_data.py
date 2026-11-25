"""
NAP SARL — Rapports JOURNALIERS / HEBDOMADAIRES (PDF + Excel) et registre COURRIER & ACTIVITÉS.
À placer à la RACINE du projet (à côté de dechargement_data.py).
Utilisé par la page 8_Rapports.py

Sources lues :  donnees_bons.xlsx (bons de commande) · Suivi de paiement.xlsx (bons de paiement)
                dechargements_<PRODUIT>.xlsx (déchargements + paiements approuvés) · courrier.xlsx (page Courrier : arrivée / départ)
PDF : nécessite  pip install reportlab
"""
import numbers
import os
import re
from datetime import datetime, timedelta
from io import BytesIO
from xml.sax.saxutils import escape

import pandas as pd

import dechargement_data as dd

VERSION_RAPPORT = 5   # v5 = le courrier vient uniquement de la page Courrier (courrier.xlsx)

# ==================== CONFIGURATION ====================
FICHIER_COURRIER = "courrier.xlsx"      # registre de la page « Courrier » (Arrivée / Départ) : lu automatiquement
TYPES_FORMULAIRE = ["Autre activité"]    # les courriers se saisissent dans la page « Courrier »
GENERAL = "Général"   # courrier / activité non lié à un produit (visible seulement dans « Tous les produits »)
ENTETE_SOCIETE = "NAP SARL"
JOURS = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]
MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre",
        "novembre", "décembre"]
# Charte BLEU CIEL (voir theme_bleu_ciel.py : on y change la teinte pour tous les documents)
try:
    import theme_bleu_ciel as _th
    CIEL, CIEL_BORD, CIEL_TEXTE, CIEL_FONCE, CIEL_CLAIR = _th.CIEL, _th.CIEL_BORD, _th.CIEL_TEXTE, _th.CIEL_FONCE, _th.CIEL_CLAIR
except Exception:  # module absent : valeurs par défaut
    CIEL, CIEL_BORD, CIEL_TEXTE, CIEL_FONCE, CIEL_CLAIR = "87CEEB", "5DADE2", "2E86C1", "1B4F72", "EAF6FD"
BLEU = "#" + CIEL
# (libellé, clé, décimales)
KPIS = [
    ("Bons de commande émis", "nb_bons", 0),
    ("Quantité commandée (T)", "qte_bons_t", 1),
    ("Camions déchargés", "nb_camions", 0),
    ("Poids net déchargé (T)", "poids_t", 1),
    ("Montant des livraisons déchargées (FCFA)", "montant_decharge", 0),
    ("Bons de paiement émis", "nb_bons_paiement", 0),
    ("Montant des bons de paiement (FCFA)", "montant_bons_paiement", 0),
    ("Paiements approuvés sur livraisons", "nb_approbations", 0),
    ("Montant approuvé sur livraisons (FCFA)", "montant_approuve", 0),
    ("Courriers reçus", "courriers_recus", 0),
    ("Courriers envoyés", "courriers_envoyes", 0),
]
# =======================================================


# ==================== TRADUCTION (impression du rapport en anglais) ====================
JOURS_EN = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
MOIS_EN = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
           "November", "December"]
LANGUES = {"fr": "Français", "en": "English"}

# étiquettes (titres, colonnes, indicateurs) : français -> anglais
_EN = {
    # indicateurs
    "Bons de commande émis": "Purchase orders issued", "Quantité commandée (T)": "Quantity ordered (T)",
    "Camions déchargés": "Trucks unloaded", "Poids net déchargé (T)": "Net weight unloaded (T)",
    "Montant des livraisons déchargées (FCFA)": "Value of unloaded deliveries (FCFA)",
    "Bons de paiement émis": "Payment vouchers issued", "Montant des bons de paiement (FCFA)": "Payment vouchers amount (FCFA)",
    "Paiements approuvés sur livraisons": "Payments approved on deliveries",
    "Montant approuvé sur livraisons (FCFA)": "Amount approved on deliveries (FCFA)",
    "Courriers reçus": "Incoming mail", "Courriers envoyés": "Outgoing mail",
    # situation à la fin de la période
    "Bons de commande émis (cumul)": "Purchase orders issued (cumulative)",
    "Bons de commande pas encore livrés": "Purchase orders not yet delivered",
    "Quantité en attente de livraison (T)": "Quantity awaiting delivery (T)",
    "Camions déchargés (cumul)": "Trucks unloaded (cumulative)",
    "Poids net déchargé (cumul, T)": "Net weight unloaded (cumulative, T)",
    "Montant total dû aux fournisseurs (FCFA)": "Total amount due to suppliers (FCFA)",
    "Montant total payé (FCFA)": "Total amount paid (FCFA)",
    "Reste à payer aux fournisseurs (FCFA)": "Balance payable to suppliers (FCFA)",
    # colonnes
    "Jour": "Day", "Bons de commande": "Purchase orders", "Camions": "Trucks", "Poids net (T)": "Net weight (T)",
    "Montant livraisons (FCFA)": "Deliveries value (FCFA)", "Bons de paiement": "Payment vouchers",
    "Montant bons de paiement (FCFA)": "Payment vouchers amount (FCFA)",
    "Paiements approuvés (FCFA)": "Approved payments (FCFA)",
    "N° Bon de commande": "Purchase order no.", "Heure": "Time", "Produit": "Product", "Fournisseur": "Supplier",
    "Quantité": "Quantity", "Camion (bon)": "Truck (order)", "Chauffeur": "Driver", "Créé par": "Created by",
    "N° Bon de paiement": "Payment voucher no.", "Bénéficiaire": "Payee", "Montant (FCFA)": "Amount (FCFA)",
    "Mode": "Method", "Banque": "Bank", "Observation": "Remarks",
    "Camion livré": "Truck delivered", "Poids net (kg)": "Net weight (kg)", "Sacs": "Bags", "PU": "Unit price",
    "Montant dû (FCFA)": "Amount due (FCFA)", "Statut paiement (actuel)": "Payment status (current)", "Alerte": "Alert",
    "Bon(s) de paiement": "Payment voucher(s)", "Approuvé par": "Approved by",
    "Nature": "Category", "Priorité": "Priority",
    "N°": "No.", "Objet": "Subject", "Correspondant": "Correspondent", "Référence": "Reference", "Statut": "Status",
    "Saisi par": "Entered by",
    # titres et libellés du PDF / de l'Excel
    "Semaine précédente": "Previous week", "Jour précédent": "Previous day", "Indicateur": "Indicator",
    "Période": "Period", "Évolution": "Change", "Situation à la fin de la période": "Position at end of period",
    "Valeur": "Value", "Répartition par produit": "Breakdown by product",
    "1. Synthèse de la période": "1. Period summary",
    "2. Évolution sur la semaine": "2. Trend over the week",
    "2. Évolution sur les 7 derniers jours": "2. Trend over the last 7 days",
    "Poids net déchargé (tonnes)": "Net weight unloaded (tonnes)",
    "Montant des bons de paiement (millions FCFA)": "Payment vouchers amount (million FCFA)",
    "3. Bons de commande émis": "3. Purchase orders issued", "4. Bons de paiement émis": "4. Payment vouchers issued",
    "5. Déchargements": "5. Unloadings", "Déchargements par fournisseur": "Unloadings by supplier",
    "6. Paiements approuvés sur livraisons": "6. Payments approved on deliveries",
    "7. Courrier (arrivée et départ)": "7. Mail (incoming and outgoing)",
    "Aucune activité enregistrée sur la période.": "No activity recorded for this period.",
    "Établi par": "Prepared by", "le": "on", "Rapport": "Report", "Généré le": "Generated on", "Champ": "Field",
    # feuilles Excel
    "Infos": "Info", "Synthese": "Summary", "Situation_fin_periode": "Position_end_period",
    "Par_produit": "By_product", "Evolution_jours": "Daily_trend", "Bons_commande": "Purchase_orders",
    "Bons_paiement": "Payment_vouchers", "Dechargements": "Unloadings", "Par_fournisseur": "By_supplier",
    "Paiements_approuves": "Approved_payments", "Courrier": "Mail",
}
# valeurs (statuts, types, jours…) : français -> anglais
_VAL_EN = {
    "PAYE": "PAID", "NON PAYE": "UNPAID", "PARTIEL": "PARTIAL", "SURPLUS": "OVERPAID",
    "Camion différent du bon": "Truck differs from order",
    "Courrier reçu": "Incoming mail", "Courrier envoyé": "Outgoing mail", "Autre activité": "Other activity",
    "À traiter": "To process", "En cours": "In progress", "Traité": "Done", "Général": "General",
    "Autres dépenses (hors produits)": "Other expenses (non-product)",
    "Lettre": "Letter", "Facture": "Invoice", "BL / Connaissement": "BL / Bill of lading", "Certificat": "Certificate",
    "Déclaration": "Declaration", "Contrat": "Contract", "Chèque / Paiement": "Cheque / Payment", "Relance": "Reminder",
    "Administratif": "Administrative", "Autre": "Other", "Normale": "Normal", "Urgente": "Urgent",
    "Confidentielle": "Confidential", "Reçu": "Received", "En traitement": "Processing", "Transmis": "Forwarded",
    "Archivé": "Archived", "Répondu": "Replied", "Envoyé": "Sent",
    "Lundi": "Monday", "Mardi": "Tuesday", "Mercredi": "Wednesday", "Jeudi": "Thursday", "Vendredi": "Friday",
    "Samedi": "Saturday", "Dimanche": "Sunday", "nouveau": "new",
}


def _t(texte, langue):
    """Traduit un libellé (français -> anglais) ; en français, renvoie le texte tel quel."""
    if langue == "en":
        return _EN.get(texte, _VAL_EN.get(texte, texte))
    return texte


def date_en(d, avec_jour=True):
    s = f"{d.day} {MOIS_EN[d.month - 1]} {d.year}"
    return f"{JOURS_EN[d.weekday()]} {s}" if avec_jour else s


def periode_en(d0, d1):
    if d0 == d1:
        return date_en(d0)
    if d0.month == d1.month and d0.year == d1.year:
        return f"from {JOURS_EN[d0.weekday()]} {d0.day} to {JOURS_EN[d1.weekday()]} {d1.day} {MOIS_EN[d1.month - 1]} {d1.year}"
    return f"from {date_en(d0)} to {date_en(d1)}"


# ---------- Dates en français ----------
def date_fr(d, avec_jour=True):
    s = f"{d.day} {MOIS[d.month - 1]} {d.year}"
    return f"{JOURS[d.weekday()]} {s}" if avec_jour else s


def periode_fr(d0, d1):
    if d0 == d1:
        return date_fr(d0)
    if d0.month == d1.month and d0.year == d1.year:
        return f"du {JOURS[d0.weekday()]} {d0.day} au {JOURS[d1.weekday()]} {d1.day} {MOIS[d1.month - 1]} {d1.year}"
    return f"du {date_fr(d0)} au {date_fr(d1)}"


def _nb(v, dec=0):
    try:
        return f"{float(v):,.{dec}f}".replace(",", " ")
    except Exception:
        return str(v)


def evolution_txt(cur, prev):
    if not cur and not prev:
        return "="
    if not prev:
        return "nouveau"
    pct = round((cur - prev) / prev * 100)
    return "0 %" if pct == 0 else f"{pct:+d} %"


# ==================== COLLECTE DES DONNÉES ====================
def _union(frames):
    pleins = [f for f in frames if not f.empty]
    return pd.concat(pleins, ignore_index=True) if pleins else frames[0]


def _bons_tous():
    cols = ["Reference", "Date_Emission", "Heure_Creation", "Utilisateur", "Produit", "Nom_Fournisseur", "Quantite",
            "Numero_Tracteur", "Chauffeur"]
    if os.path.exists(dd.FICHIER_BONS):
        df = pd.read_excel(dd.FICHIER_BONS)
    else:
        df = pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols].copy()
    for c in cols:
        df[c] = df[c].map(dd._s)
    df = df[df["Reference"] != ""].reset_index(drop=True)
    df["Date"] = pd.to_datetime(df["Date_Emission"], dayfirst=True, errors="coerce")
    df["Produit_Code"] = [(re.search(r"-BC-([A-Za-z]+)", r).group(1).upper() if re.search(r"-BC-([A-Za-z]+)", r) else "")
                          for r in df["Reference"]]
    df["Qte_KG"] = df["Quantite"].map(dd.quantite_kg)
    return df


def _bons_paiement_tous():
    j = dd.charger_journal()
    cols = ["No_Bon_Paiement", "Date", "Produit", "Prestataire", "Montant", "Mode", "Banque", "Observation"]
    if j.empty:
        return pd.DataFrame(columns=cols)
    unique = lambda s: " / ".join(dict.fromkeys(x for x in s if x))
    g = (j.groupby("No_Transaction", sort=False)
         .agg(Date=("Date", "min"), Produit=("Produit", "first"), Prestataire=("Prestataire", unique),
              Montant=("Montant", "sum"), Mode=("Mode", "first"), Banque=("Banque", "first"),
              Observation=("Observation", "first")).reset_index()
         .rename(columns={"No_Transaction": "No_Bon_Paiement"}))
    return g[cols]


# Mots-clés pour rattacher un courrier à un produit (un seul produit trouvé = rattaché, sinon « Général »)
_MOTS_PRODUITS = {"SHEA": r"\b(shea|karit[eé])\b", "SOYABEAN": r"\b(soja|soya|soyabean|soybean)\b",
                  "COTTON": r"\b(coton|cotton)\b"}
_COLS_ACT_RAPPORT = ["ID_Affiche", "Date", "Type", "Produit", "Nature", "Objet", "Correspondant", "Reference",
                     "Priorite", "Statut", "Cree_Par"]


def _produit_depuis_texte(*textes):
    t = " ".join(textes)
    trouves = [code for code, motif in _MOTS_PRODUITS.items() if re.search(motif, t, re.IGNORECASE)]
    return trouves[0] if len(trouves) == 1 else ""


def charger_courrier_page():
    """Courriers enregistrés dans la page « Courrier » (courrier.xlsx), au format du rapport.
    Arrivée -> « Courrier reçu » ; Départ -> « Courrier envoyé »."""
    if not os.path.exists(FICHIER_COURRIER):
        return pd.DataFrame(columns=_COLS_ACT_RAPPORT)
    try:
        df = pd.read_excel(FICHIER_COURRIER, dtype=str)
    except Exception:
        return pd.DataFrame(columns=_COLS_ACT_RAPPORT)
    for c in ["ID", "Date", "Type", "Nature", "Objet", "Expediteur", "Destinataire", "Reference", "Priorite",
              "Statut", "Service", "Observations", "Utilisateur"]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].map(dd._s)
    lignes = []
    for r in df.itertuples(index=False):
        if not (r.Objet or r.ID):
            continue
        t = r.Type.lower()
        if t.startswith("arriv"):
            type_, correspondant = "Courrier reçu", r.Expediteur
        elif t.startswith("dép") or t.startswith("dep"):
            type_, correspondant = "Courrier envoyé", r.Destinataire
        else:
            type_, correspondant = "Autre activité", r.Expediteur or r.Destinataire
        lignes.append({"ID_Affiche": r.ID, "Date": r.Date, "Type": type_,
                       "Produit": _produit_depuis_texte(r.Objet, r.Reference, r.Observations, r.Service),
                       "Nature": r.Nature, "Objet": r.Objet, "Correspondant": correspondant,
                       "Reference": r.Reference, "Priorite": r.Priorite, "Statut": r.Statut, "Cree_Par": r.Utilisateur})
    out = pd.DataFrame(lignes, columns=_COLS_ACT_RAPPORT)
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce", dayfirst=True)
    return out


def charger_donnees():
    """Toutes les sources (tous produits). Le produit actif de la session est restauré."""
    ancien = dd.get_produit()
    dechs, affs = [], []
    try:
        for p in dd.PRODUITS:
            dd.set_produit(p)
            d, a = dd.charger_complet()
            dechs.append(d.assign(Produit=p))
            affs.append(a.assign(Produit=p))
    finally:
        dd.set_produit(ancien)
    courriers = charger_courrier_page()     # courriers saisis dans la page « Courrier »
    return {"dech": _union(dechs), "aff": _union(affs), "bons": _bons_tous(),
            "paiements": _bons_paiement_tous(), "act": courriers}


# ==================== CALCULS ====================
def _dans(serie, d0, d1):
    s = pd.to_datetime(serie, errors="coerce")
    return (s >= pd.Timestamp(d0)) & (s < pd.Timestamp(d1) + pd.Timedelta(days=1))


def _filtrer(data, d0, d1, produit):
    bons, dech, aff, pay, act = data["bons"], data["dech"], data["aff"], data["paiements"], data["act"]
    fb = bons[_dans(bons["Date"], d0, d1)]
    fd = dech[_dans(dech["Date_Dechargement"], d0, d1)]
    # Paiements approuvés réellement saisis (les lignes reprises par import n'ont pas d'heure : ignorées)
    fa = aff[(aff["Heure"] != "") & _dans(aff["Date_Approbation"], d0, d1)]
    fp = pay[_dans(pay["Date"], d0, d1)]
    fc = act[_dans(act["Date"], d0, d1)]
    if produit:
        fb = fb[fb["Produit_Code"] == produit]
        fd = fd[fd["Produit"] == produit]
        fa = fa[fa["Produit"] == produit]
        fp = fp[dd.masque(fp, [dd._s(p).upper() == produit for p in fp["Produit"]])]
        fc = fc[fc["Produit"] == produit]
    if not fa.empty:
        fa = fa.merge(dech[["Produit", "Transaction", "Ref_BC", "Fournisseur"]], on=["Produit", "Transaction"], how="left")
    else:
        fa = fa.assign(Ref_BC="", Fournisseur="")
    return {"bons": fb, "dech": fd, "appro": fa, "paiements": fp, "act": fc}


def _totaux(f):
    t = {
        "nb_bons": len(f["bons"]), "qte_bons_t": f["bons"]["Qte_KG"].sum() / 1000,
        "nb_camions": len(f["dech"]), "poids_t": f["dech"]["Poids_Net_KG"].sum() / 1000,
        "montant_decharge": f["dech"]["Montant_Du"].sum(),
        "nb_bons_paiement": len(f["paiements"]), "montant_bons_paiement": f["paiements"]["Montant"].sum(),
        "nb_approbations": len(f["appro"]), "montant_approuve": f["appro"]["Montant"].sum(),
        "courriers_recus": int((f["act"]["Type"] == "Courrier reçu").sum()),
        "courriers_envoyes": int((f["act"]["Type"] == "Courrier envoyé").sum()),
    }
    return {k: float(v) for k, v in t.items()}


def _etat_fin(data, fin, produit):
    limite = pd.Timestamp(fin) + pd.Timedelta(days=1)
    d = data["dech"][pd.to_datetime(data["dech"]["Date_Dechargement"], errors="coerce") < limite]
    a = data["aff"][pd.to_datetime(data["aff"]["Date_Approbation"], errors="coerce") < limite]
    b = data["bons"][pd.to_datetime(data["bons"]["Date"], errors="coerce") < limite]
    if produit:
        d, a, b = d[d["Produit"] == produit], a[a["Produit"] == produit], b[b["Produit_Code"] == produit]
    livres = set(d["Ref_BC"]) - {""}
    attente = b[~b["Reference"].isin(livres)]
    du, paye = float(d["Montant_Du"].sum()), float(a["Montant"].sum())
    return [
        ("Bons de commande émis (cumul)", _nb(len(b))),
        ("Bons de commande pas encore livrés", _nb(len(attente))),
        ("Quantité en attente de livraison (T)", _nb(attente["Qte_KG"].sum() / 1000, 1)),
        ("Camions déchargés (cumul)", _nb(len(d))),
        ("Poids net déchargé (cumul, T)", _nb(d["Poids_Net_KG"].sum() / 1000, 1)),
        ("Montant total dû aux fournisseurs (FCFA)", _nb(du)),
        ("Montant total payé (FCFA)", _nb(paye)),
        ("Reste à payer aux fournisseurs (FCFA)", _nb(du - paye)),
    ]


def _evolution(data, jours, produit):
    rows = []
    for j in jours:
        t = _totaux(_filtrer(data, j, j, produit))
        rows.append({"Date": pd.Timestamp(j), "Jour": JOURS[j.weekday()].capitalize(),
                     "Bons de commande": int(t["nb_bons"]), "Camions": int(t["nb_camions"]),
                     "Poids net (T)": round(t["poids_t"], 1), "Montant livraisons (FCFA)": round(t["montant_decharge"]),
                     "Bons de paiement": int(t["nb_bons_paiement"]),
                     "Montant bons de paiement (FCFA)": round(t["montant_bons_paiement"]),
                     "Paiements approuvés (FCFA)": round(t["montant_approuve"]),
                     "Courriers reçus": int(t["courriers_recus"]), "Courriers envoyés": int(t["courriers_envoyes"])})
    return pd.DataFrame(rows)


def _par_produit(data, debut, fin):
    """Tableau « tous produits » : une ligne par produit + 'Autres' (bons de paiement hors produits d'achat)."""
    rows, tot = [], _totaux(_filtrer(data, debut, fin, None))
    cumul = {"nb_bons": 0, "nb_camions": 0, "poids_t": 0, "montant_decharge": 0, "nb_bons_paiement": 0,
             "montant_bons_paiement": 0}
    for p in dd.PRODUITS:
        t = _totaux(_filtrer(data, debut, fin, p))
        for k in cumul:
            cumul[k] += t[k]
        rows.append({"Produit": dd.libelle_produit(p), "Bons de commande": int(t["nb_bons"]), "Camions": int(t["nb_camions"]),
                     "Poids net (T)": round(t["poids_t"], 1), "Montant livraisons (FCFA)": round(t["montant_decharge"]),
                     "Bons de paiement": int(t["nb_bons_paiement"]),
                     "Montant bons de paiement (FCFA)": round(t["montant_bons_paiement"])})
    if tot["nb_bons_paiement"] - cumul["nb_bons_paiement"] > 0:
        rows.append({"Produit": "Autres dépenses (hors produits)", "Bons de commande": 0, "Camions": 0,
                     "Poids net (T)": 0.0, "Montant livraisons (FCFA)": 0,
                     "Bons de paiement": int(tot["nb_bons_paiement"] - cumul["nb_bons_paiement"]),
                     "Montant bons de paiement (FCFA)": round(tot["montant_bons_paiement"] - cumul["montant_bons_paiement"])})
    return pd.DataFrame(rows)


def generer_rapport(data, type_, jour, produit=None, observations="", auteur="", langue="fr"):
    """type_ : 'Journalier' ou 'Hebdomadaire'. jour : un jour quelconque de la période (date).
    langue : 'fr' (défaut) ou 'en' (rapport imprimable en anglais)."""
    if type_ == "Hebdomadaire":
        debut = jour - timedelta(days=jour.weekday())
        fin = debut + timedelta(days=6)
        prec = (debut - timedelta(days=7), fin - timedelta(days=7))
        jours = [debut + timedelta(days=i) for i in range(7)]
        titre = f"RAPPORT HEBDOMADAIRE — {dd.libelle_produit(produit) if produit else 'TOUS PRODUITS'} — Semaine {debut.isocalendar()[1]}"
        lib_prec = "Semaine précédente"
    else:
        debut = fin = jour
        prec = (jour - timedelta(days=1), jour - timedelta(days=1))
        jours = [jour - timedelta(days=6 - i) for i in range(7)]
        titre = f"RAPPORT JOURNALIER — {dd.libelle_produit(produit) if produit else 'TOUS PRODUITS'}"
        lib_prec = "Jour précédent"

    f = _filtrer(data, debut, fin, produit)
    cur, prv = _totaux(f), _totaux(_filtrer(data, prec[0], prec[1], produit))
    kpis = [{"label": lab, "cur": cur[k], "prev": prv[k], "evol": evolution_txt(cur[k], prv[k]), "dec": dec}
            for lab, k, dec in KPIS]

    # ----- détails (colonnes lisibles) -----
    fb = f["bons"].sort_values(["Date", "Heure_Creation"])
    bons = pd.DataFrame({
        "N° Bon de commande": fb["Reference"], "Date": fb["Date"], "Heure": fb["Heure_Creation"],
        "Produit": fb["Produit_Code"], "Fournisseur": fb["Nom_Fournisseur"], "Quantité": fb["Quantite"],
        "Camion (bon)": fb["Numero_Tracteur"], "Chauffeur": fb["Chauffeur"], "Créé par": fb["Utilisateur"]})
    fp = f["paiements"].sort_values(["Date", "No_Bon_Paiement"])
    paiements = pd.DataFrame({
        "N° Bon de paiement": fp["No_Bon_Paiement"], "Date": fp["Date"], "Produit": fp["Produit"],
        "Bénéficiaire": fp["Prestataire"], "Montant (FCFA)": fp["Montant"], "Mode": fp["Mode"],
        "Banque": fp["Banque"], "Observation": fp["Observation"]})
    fd = f["dech"].sort_values(["Date_Dechargement", "Produit", "Transaction"])
    dech = pd.DataFrame({
        "Date": fd["Date_Dechargement"], "Produit": fd["Produit"], "Transaction": fd["Transaction"],
        "N° Bon de commande": fd["Ref_BC"], "Fournisseur": fd["Fournisseur"], "Camion livré": fd["N_Camion"],
        "Poids net (kg)": fd["Poids_Net_KG"], "Sacs": fd["Nb_Sacs"], "PU": fd["PU_FCFA_KG"],
        "Montant dû (FCFA)": fd["Montant_Du"], "Statut paiement (actuel)": fd["Statut_Paiement"],
        "Alerte": ["Camion différent du bon" if x else "" for x in fd["Camion_Diff"]]})
    par_frs = (f["dech"].groupby("Fournisseur").agg(Camions=("Transaction", "count"), **{
        "Poids net (kg)": ("Poids_Net_KG", "sum"), "Montant dû (FCFA)": ("Montant_Du", "sum")})
        .reset_index().sort_values("Poids net (kg)", ascending=False)) if not f["dech"].empty else pd.DataFrame(
        columns=["Fournisseur", "Camions", "Poids net (kg)", "Montant dû (FCFA)"])
    fa = f["appro"].sort_values(["Date_Approbation", "ID"]) if not f["appro"].empty else f["appro"]
    appro = pd.DataFrame({
        "Date": fa["Date_Approbation"], "Produit": fa["Produit"], "Transaction": fa["Transaction"],
        "N° Bon de commande": fa["Ref_BC"], "Fournisseur": fa["Fournisseur"], "Bon(s) de paiement": fa["Bons_Paiement"],
        "Montant (FCFA)": fa["Montant"], "Approuvé par": fa["Approuve_Par"]})
    fc = f["act"].sort_values(["Date", "ID_Affiche"])
    act = pd.DataFrame({
        "N°": fc["ID_Affiche"], "Date": fc["Date"], "Produit": fc["Produit"].replace("", GENERAL), "Type": fc["Type"],
        "Nature": fc["Nature"], "Objet": fc["Objet"], "Correspondant": fc["Correspondant"],
        "Référence": fc["Reference"], "Priorité": fc["Priorite"], "Statut": fc["Statut"], "Saisi par": fc["Cree_Par"]})

    rap = {
        "langue": "fr", "produit_code": produit,
        "titre": titre, "type": type_, "debut": debut, "fin": fin, "periode": periode_fr(debut, fin),
        "produit": dd.libelle_produit(produit) if produit else "Tous les produits", "auteur": auteur, "observations": observations.strip(),
        "genere_le": datetime.now().strftime("%d/%m/%Y à %H:%M"), "lib_prec": lib_prec,
        "periode_prec": periode_fr(*prec), "kpis": kpis, "etat": _etat_fin(data, fin, produit),
        "evolution": _evolution(data, jours, produit),
        "par_produit": _par_produit(data, debut, fin) if not produit else pd.DataFrame(),
        "bons": bons, "paiements": paiements, "dech": dech, "par_frs": par_frs, "appro": appro, "act": act,
    }
    return traduire_rapport(rap) if langue == "en" else rap


def traduire_rapport(rap):
    """Version ANGLAISE d'un rapport : titres, indicateurs, colonnes, statuts, dates.
    Les textes saisis (observations, objets de courrier, noms) restent tels quels."""
    if rap.get("langue") == "en":
        return rap
    produit, debut, fin = rap.get("produit_code"), rap["debut"], rap["fin"]
    hebdo = rap["type"] == "Hebdomadaire"
    prec = (debut - timedelta(days=7), fin - timedelta(days=7)) if hebdo else (debut - timedelta(days=1),) * 2
    nom_produit = dd.libelle_produit(produit) if produit else "ALL PRODUCTS"
    en = dict(rap)
    en["langue"] = "en"
    en["titre"] = (f"WEEKLY REPORT — {nom_produit} — Week {debut.isocalendar()[1]}" if hebdo
                   else f"DAILY REPORT — {nom_produit}")
    en["periode"], en["periode_prec"] = periode_en(debut, fin), periode_en(*prec)
    en["produit"] = dd.libelle_produit(produit) if produit else "All products"
    en["lib_prec"] = _t(rap["lib_prec"], "en")
    en["genere_le"] = rap["genere_le"].replace(" à ", " at ")
    en["kpis"] = [{**k, "label": _t(k["label"], "en"), "evol": _t(k["evol"], "en")} for k in rap["kpis"]]
    en["etat"] = [(_t(lab, "en"), val) for lab, val in rap["etat"]]

    def trad_df(df):
        df = df.copy()
        for col in ("Jour", "Statut paiement (actuel)", "Alerte", "Type", "Statut", "Produit", "Nature", "Priorité"):
            if col in df.columns:
                df[col] = df[col].map(lambda v: _VAL_EN.get(v, v) if isinstance(v, str) else v)
        return df.rename(columns={c: _t(c, "en") for c in df.columns})

    for cle in ("evolution", "par_produit", "bons", "paiements", "dech", "par_frs", "appro", "act"):
        en[cle] = trad_df(rap[cle])
    return en


# ==================== EXCEL ====================
def construire_excel(rap):
    lg = rap.get("langue", "fr")
    t = lambda x: _t(x, lg)
    infos = pd.DataFrame({t("Champ"): [t("Rapport"), t("Période"), t("Produit"), t("Établi par"), t("Généré le"),
                                       "Observations"],
                          t("Valeur"): [rap["titre"], rap["periode"], rap["produit"], rap["auteur"], rap["genere_le"],
                                        rap["observations"]]})
    synth = pd.DataFrame({t("Indicateur"): [k["label"] for k in rap["kpis"]],
                          t("Période"): [round(k["cur"], 1) for k in rap["kpis"]],
                          rap["lib_prec"]: [round(k["prev"], 1) for k in rap["kpis"]],
                          t("Évolution"): [k["evol"] for k in rap["kpis"]]})
    etat = pd.DataFrame(rap["etat"], columns=[t("Situation à la fin de la période"), t("Valeur")])
    return dd.to_excel_bytes({
        t("Infos"): infos, t("Synthese"): synth, t("Situation_fin_periode"): etat,
        t("Par_produit"): rap["par_produit"], t("Evolution_jours"): rap["evolution"],
        t("Bons_commande"): rap["bons"], t("Bons_paiement"): rap["paiements"], t("Dechargements"): rap["dech"],
        t("Par_fournisseur"): rap["par_frs"], t("Paiements_approuves"): rap["appro"],
        t("Courrier"): rap["act"]})


# ==================== PDF ====================
def construire_pdf(rap):
    """Rapport PDF A4 paysage. Nécessite reportlab (pip install reportlab)."""
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    from reportlab.graphics.shapes import Drawing, String
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    lg = rap.get("langue", "fr")
    T = lambda x: _t(x, lg)       # traduit un libellé (sans effet en français)

    bleu, gris = colors.HexColor("#" + CIEL), colors.HexColor("#" + CIEL_CLAIR)
    sur_ciel, bleu_texte, trait = colors.HexColor("#" + CIEL_FONCE), colors.HexColor("#" + CIEL_TEXTE), colors.HexColor("#9FD0EE")
    L = 273 * mm  # largeur utile (A4 paysage, marges 12 mm)
    base = getSampleStyleSheet()["Normal"]
    st_cell = ParagraphStyle("cell", parent=base, fontName="Helvetica", fontSize=7.5, leading=9)
    st_cell_r = ParagraphStyle("cellr", parent=st_cell, alignment=2)
    st_head = ParagraphStyle("head", parent=st_cell, fontName="Helvetica-Bold", textColor=sur_ciel)
    st_head_r = ParagraphStyle("headr", parent=st_head, alignment=2)
    st_titre = ParagraphStyle("titre", parent=base, fontName="Helvetica-Bold", fontSize=17, leading=20, textColor=bleu_texte)
    st_sous = ParagraphStyle("sous", parent=base, fontSize=10, leading=13, textColor=colors.HexColor("#444444"))
    st_sec = ParagraphStyle("sec", parent=base, fontName="Helvetica-Bold", fontSize=10.5, leading=13,
                            textColor=sur_ciel, backColor=bleu, borderPadding=(3, 4, 3, 4), spaceBefore=10,
                            spaceAfter=6)
    st_vide = ParagraphStyle("vide", parent=base, fontName="Helvetica-Oblique", fontSize=8.5, textColor=colors.grey)
    st_obs = ParagraphStyle("obs", parent=base, fontSize=9, leading=12)

    def txt(v, num=False):
        try:
            if pd.isna(v):
                return ""
        except (TypeError, ValueError):
            pass
        if isinstance(v, (pd.Timestamp, datetime)):
            return v.strftime("%d/%m/%Y")
        if num and isinstance(v, numbers.Number):
            return dd.fmt(v)
        return str(v)

    def tableau(df, poids, num_cols=()):
        entetes = list(df.columns)
        data = [[Paragraph(escape(h), st_head_r if h in num_cols else st_head) for h in entetes]]
        for row in df.itertuples(index=False):
            data.append([Paragraph(escape(txt(v, c in num_cols)), st_cell_r if c in num_cols else st_cell)
                         for v, c in zip(row, entetes)])
        tot = float(sum(poids))
        t = Table(data, colWidths=[L * p / tot for p in poids], repeatRows=1)
        style = [("BACKGROUND", (0, 0), (-1, 0), bleu), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                 ("GRID", (0, 0), (-1, -1), 0.25, trait),
                 ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                 ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3)]
        for i in range(2, len(data), 2):
            style.append(("BACKGROUND", (0, i), (-1, i), gris))
        t.setStyle(TableStyle(style))
        return t

    def section(titre, df, poids, num_cols=()):
        out = [Paragraph(escape(titre), st_sec)]
        out.append(tableau(df, poids, num_cols) if not df.empty
                   else Paragraph(T("Aucune activité enregistrée sur la période."), st_vide))
        return out

    def graphique(titre, cats, vals, fmt):
        d = Drawing(130 * mm, 52 * mm)
        d.add(String(0, 47 * mm, titre, fontName="Helvetica-Bold", fontSize=8.5, fillColor=bleu_texte))
        c = VerticalBarChart()
        c.x, c.y, c.width, c.height = 12 * mm, 9 * mm, 112 * mm, 33 * mm
        c.data = [list(vals)]
        c.categoryAxis.categoryNames = cats
        c.categoryAxis.labels.fontSize = 6.5
        c.valueAxis.labels.fontSize = 6.5
        c.valueAxis.valueMin = 0
        if max(list(vals) + [0]) <= 0:      # aucune donnée : axe lisible (0 à 1) au lieu de 0.002, 0.004…
            c.valueAxis.valueMax = 1
            c.valueAxis.valueStep = 0.5
        c.bars[0].fillColor = colors.HexColor("#" + CIEL_BORD)
        c.barLabelFormat = fmt
        c.barLabels.fontSize = 6.5
        c.barLabels.nudge = 5
        c.valueAxis.visibleGrid = True
        c.valueAxis.gridStrokeColor = colors.HexColor("#DDDDDD")
        d.add(c)
        return d

    # ----------------- contenu -----------------
    el = [Paragraph(f"{ENTETE_SOCIETE} — {escape(rap['titre'])}", st_titre),
          Paragraph(escape(rap["periode"][:1].upper() + rap["periode"][1:]) +
                    f" &nbsp;·&nbsp; {T('Produit')} : {escape(rap['produit'])}", st_sous),
          Paragraph(f"{T('Établi par')} {escape(rap['auteur'] or '—')} {T('le')} {rap['genere_le']}", st_sous),
          Spacer(1, 4)]

    # 1. Synthèse + situation
    el.append(Paragraph(T("1. Synthèse de la période"), st_sec))
    kp = [[Paragraph(T("Indicateur"), st_head), Paragraph(T("Période"), st_head_r),
           Paragraph(escape(rap["lib_prec"]), st_head_r), Paragraph(T("Évolution"), st_head_r)]]
    for k in rap["kpis"]:
        kp.append([Paragraph(escape(k["label"]), st_cell), Paragraph(_nb(k["cur"], k["dec"]), st_cell_r),
                   Paragraph(_nb(k["prev"], k["dec"]), st_cell_r), Paragraph(escape(k["evol"]), st_cell_r)])
    t1 = Table(kp, colWidths=[78 * mm, 28 * mm, 30 * mm, 24 * mm])
    s1 = [("BACKGROUND", (0, 0), (-1, 0), bleu), ("GRID", (0, 0), (-1, -1), 0.25, trait),
          ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]
    s1 += [("BACKGROUND", (0, i), (-1, i), gris) for i in range(2, len(kp), 2)]
    t1.setStyle(TableStyle(s1))
    et = [[Paragraph(T("Situation à la fin de la période"), st_head), Paragraph(T("Valeur"), st_head_r)]]
    for lab, val in rap["etat"]:
        et.append([Paragraph(escape(lab), st_cell), Paragraph(val, st_cell_r)])
    t2 = Table(et, colWidths=[70 * mm, 36 * mm])
    s2 = [("BACKGROUND", (0, 0), (-1, 0), bleu), ("GRID", (0, 0), (-1, -1), 0.25, trait),
          ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]
    s2 += [("BACKGROUND", (0, i), (-1, i), gris) for i in range(2, len(et), 2)]
    t2.setStyle(TableStyle(s2))
    cote = Table([[t1, t2]], colWidths=[165 * mm, 108 * mm])
    cote.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    el.append(cote)

    if rap["observations"]:
        el.append(Paragraph("Observations", st_sec))
        el.append(Paragraph(escape(rap["observations"]).replace("\n", "<br/>"), st_obs))

    if not rap["par_produit"].empty:
        el.append(Paragraph(T("Répartition par produit"), st_sec))
        el.append(tableau(rap["par_produit"], [22, 10, 9, 10, 17, 10, 19],
                          num_cols=[c for c in rap["par_produit"].columns if c != T("Produit")]))

    # 2. Évolution
    ev = rap["evolution"]
    el.append(Paragraph(T("2. Évolution sur la semaine" if rap["type"] == "Hebdomadaire"
                          else "2. Évolution sur les 7 derniers jours"), st_sec))
    cats = [f"{str(j)[:3]} {d.strftime('%d/%m')}" for j, d in zip(ev[T("Jour")], ev["Date"])]
    graph = Table([[graphique(T("Poids net déchargé (tonnes)"), cats, ev[T("Poids net (T)")], "%0.1f"),
                    graphique(T("Montant des bons de paiement (millions FCFA)"), cats,
                              ev[T("Montant bons de paiement (FCFA)")] / 1e6, "%0.1f")]], colWidths=[136 * mm, 137 * mm])
    el.append(KeepTogether([graph]))
    el.append(Spacer(1, 4))
    ev_pdf = ev.copy()
    ev_pdf["Date"] = ev_pdf["Date"].dt.strftime("%d/%m/%Y")
    el.append(tableau(ev_pdf, [11, 7, 8, 6, 7, 12, 7, 13, 12, 7, 7],
                      num_cols=[c for c in ev_pdf.columns if c not in ("Date", T("Jour"))]))

    # 3 à 8. Détails
    el += section(T("3. Bons de commande émis"), rap["bons"], [20, 8, 7, 9, 17, 8, 14, 14, 8])
    el += section(T("4. Bons de paiement émis"), rap["paiements"], [14, 8, 9, 22, 11, 7, 10, 20], [T("Montant (FCFA)")])
    el += section(T("5. Déchargements"), rap["dech"], [10, 10, 11, 19, 16, 13, 9, 6, 6, 11, 10, 12],
                  [T("Poids net (kg)"), T("Sacs"), T("PU"), T("Montant dû (FCFA)")])
    if not rap["par_frs"].empty:
        el.append(Spacer(1, 6))
        el.append(Paragraph(T("Déchargements par fournisseur"), st_cell))
        el.append(Spacer(1, 2))
        el.append(tableau(rap["par_frs"], [40, 10, 15, 18], [T("Camions"), T("Poids net (kg)"), T("Montant dû (FCFA)")]))
    el += section(T("6. Paiements approuvés sur livraisons"), rap["appro"], [8, 8, 8, 18, 18, 22, 11, 9], [T("Montant (FCFA)")])
    el += section(T("7. Courrier (arrivée et départ)"), rap["act"], [6, 8, 8, 10, 9, 26, 15, 12, 8, 8, 9])

    def pied(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.grey)
        canvas.drawString(12 * mm, 7 * mm, f"{ENTETE_SOCIETE} — {rap['titre']} — {rap['periode']}")
        canvas.drawRightString(285 * mm, 7 * mm, f"Page {doc.page}")
        canvas.restoreState()

    buf = BytesIO()
    SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm, topMargin=12 * mm,
                      bottomMargin=14 * mm, title=f"{rap['titre']} — {rap['periode']}",
                      author=rap["auteur"] or ENTETE_SOCIETE).build(el, onFirstPage=pied, onLaterPages=pied)
    return buf.getvalue()
