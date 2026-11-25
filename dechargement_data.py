"""
NAP SARL — Couche de données DÉCHARGEMENTS
Partagée par :  pages/5_Dechargements.py  et  pages/6_Releve_Fournisseur.py
À placer à la RACINE du projet (à côté de app.py).

Un fichier PAR PRODUIT, créé automatiquement : dechargements_SHEA.xlsx, dechargements_SOYABEAN.xlsx, ...
  - feuille « Dechargements »       : 1 ligne = 1 camion déchargé
  - feuille « Paiements_Affectes »  : 1 ligne = 1 paiement approuvé (bon(s) de paiement + montant)
Le montant payé, le solde et le statut sont CALCULÉS à partir des paiements approuvés.
"""
import os
import re
import glob
import unicodedata
import threading
from datetime import datetime
from io import BytesIO

import pandas as pd

VERSION_DATA = 18  # incrémentée à chaque mise à jour : les pages vérifient qu'elles ont la bonne version

# ==================== CONFIGURATION ====================
# --- Produits : un fichier de déchargements par produit ---
# >>> Pour AJOUTER un produit : 1) ajoutez son code ici et son nom lisible plus bas ;
#     2) ajoutez-le aussi dans PRODUITS_SUFFIXES (app_bons.py / 2_Bons_de_Commande.py) avec le même code
#        ex. "Cajou (Noix de cajou)": "-BC-CAJOU/NAP-2026"  ->  code "CAJOU"
PRODUITS = [
    "SHEA",
    "TOURTEAU_KARITE",
    "SOYABEAN",
    "HUILE_SOJA",
    "TOURTEAU_SOJA",
    "COTTON",
    "HUILE_COTTON",
    "TOURTEAU_COTTON",
]
PRODUITS_LIBELLES = {
    "SHEA": "Shea / Noix de Karité",
    "TOURTEAU_KARITE": "Tourteau de Karité",
    "SOYABEAN": "Graine de Soja",
    "HUILE_SOJA": "Huile de Soja",
    "TOURTEAU_SOJA": "Tourteau de Soja",
    "COTTON": "Graine de Coton",
    "HUILE_COTTON": "Huile de Coton",
    "TOURTEAU_COTTON": "Tourteau de Coton",
}
# Autres écritures possibles (journal paiements, imports) -> code produit
PRODUITS_ALIAS = {
    "KARITE": "SHEA", "KARITÉ": "SHEA", "SHEA NUT": "SHEA", "NOIX DE KARITE": "SHEA",
    "TOURTEAU DE KARITE": "TOURTEAU_KARITE", "TOURTEAU KARITE": "TOURTEAU_KARITE",
    "SHEA MEAL": "TOURTEAU_KARITE", "SHEA DOC": "TOURTEAU_KARITE",
    "SOJA": "SOYABEAN", "SOYA": "SOYABEAN", "SOYBEAN": "SOYABEAN", "GRAINE DE SOJA": "SOYABEAN",
    "HUILE DE SOJA": "HUILE_SOJA", "HUILE SOJA": "HUILE_SOJA", "SOYA OIL": "HUILE_SOJA", "SOYA CRUDE OIL": "HUILE_SOJA",
    "TOURTEAU DE SOJA": "TOURTEAU_SOJA", "TOURTEAU SOJA": "TOURTEAU_SOJA", "SOYA MEAL": "TOURTEAU_SOJA",
    "COTTON SEED": "COTTON", "COTON": "COTTON", "GRAINE DE COTON": "COTTON", "GRAINES DE COTON": "COTTON",
    "HUILE DE COTTON": "HUILE_COTTON", "HUILE DE COTON": "HUILE_COTTON", "HUILE COTTON": "HUILE_COTTON",
    "COTTON OIL": "HUILE_COTTON",
    "TOURTEAU DE COTTON": "TOURTEAU_COTTON", "TOURTEAU DE COTON": "TOURTEAU_COTTON",
    "TOURTEAU COTTON": "TOURTEAU_COTTON", "COTTON MEAL": "TOURTEAU_COTTON",
}
PRODUIT_DEFAUT = "SOYABEAN"        # produit affiché à l'ouverture (le nouveau produit)
PRODUIT_ANCIEN = "SHEA"            # produit auquel appartiennent les données déjà saisies
FICHIER_ANCIEN = "dechargements.xlsx"   # ancien fichier unique -> renommé automatiquement en dechargements_<PRODUIT_ANCIEN>.xlsx
FEUILLE_DECH = "Dechargements"
FEUILLE_AFF = "Paiements_Affectes"
FEUILLE_CORBEILLE_DECH = "Corbeille_Dechargements"
FEUILLE_CORBEILLE_AFF = "Corbeille_Paiements"
COLS_CORBEILLE_META = ["Supprime_Le", "Supprime_Par", "Motif_Suppression"]

FICHIER_BONS = "donnees_bons.xlsx"
FICHIER_FRS = "fournisseurs.xlsx"
FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"
FEUILLE_DEPENSES = "DEPENSES"

# Comptes autorisés à approuver / annuler un paiement
ROLES_APPROBATION = ["admin", "Manager"]
# Un bon est « LIVRE » quand le poids déchargé atteint ce % de la quantité commandée
SEUIL_LIVRE = 0.98

COLS_DECH = [
    "Transaction", "Type_Mouvement", "Date_Dechargement", "Ref_BC", "Code_Frs", "Fournisseur",
    "Fournisseur_BC", "Code_Frs_BC", "Affectation_BC",  # échange de bons : ticket ≠ BC
    "N_Camion", "Camion_BC", "Weighment_Ticket", "Poids_Net_KG", "Nb_Sacs", "PU_FCFA_KG",
    "Montant_Du", "Lieu_Stockage", "Statut_Livraison", "Remarque", "Cree_Par", "Heure_Creation",
]
COLS_AFF = ["ID", "Transaction", "Date_Approbation", "Bons_Paiement", "Montant", "Approuve_Par", "Heure"]
# =======================================================


def libelle_produit(code):
    """SHEA -> 'Shea (Karité)' (nom lisible pour l'affichage)."""
    return PRODUITS_LIBELLES.get(code, str(code).capitalize())


# ---------- Produit actif (propre à chaque utilisateur connecté) ----------
_ctx = threading.local()


def set_produit(code):
    """À appeler en haut de chaque page : fixe le produit pour cette session."""
    _ctx.produit = code if code in PRODUITS else PRODUIT_DEFAUT


def get_produit():
    return getattr(_ctx, "produit", PRODUIT_DEFAUT)


def fichier_actif():
    """Chemin du fichier du produit actif (migre l'ancien fichier unique la 1re fois)."""
    p = get_produit()
    chemin = f"dechargements_{p}.xlsx"
    if p == PRODUIT_ANCIEN and not os.path.exists(chemin) and os.path.exists(FICHIER_ANCIEN):
        os.replace(FICHIER_ANCIEN, chemin)
    return chemin


# ---------- Utilitaires ----------
def _s(v):
    """Valeur -> texte propre ('' si vide)."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none", "nat") else t


def masque(df, valeurs):
    """Masque booléen aligné sur df : fiable même quand df est vide (une liste vide serait lue comme des colonnes)."""
    return pd.Series(list(valeurs), index=df.index, dtype=bool)


def _n(v, defaut=0.0):
    try:
        x = float(v)
        return defaut if pd.isna(x) else x
    except Exception:
        return defaut


def fmt(n, unite=""):
    """12345678 -> '12 345 678' (+ unité)."""
    try:
        t = f"{float(n):,.0f}".replace(",", " ")
    except Exception:
        t = str(n)
    return f"{t} {unite}".strip()


def _cle_excel(t):
    t = _s(t)
    return int(t) if t.isdigit() else t


def liste_bons_texte(texte):
    """'CHQ-1 & CC-2' -> ['CHQ-1', 'CC-2']"""
    out = []
    for p in re.split(r"[&;,\n]", _s(texte)):
        p = p.strip()
        if p and p not in out:
            out.append(p)
    return out


def _tokens(nom):
    t = unicodedata.normalize("NFKD", _s(nom)).encode("ascii", "ignore").decode().upper()
    return {m for m in re.split(r"[^A-Z]+", t) if len(m) >= 3 and m != "MME"}


def meme_nom(a, b):
    """Rapprochement souple de noms (ex. MEHSAN DANIEL JOEL ~ MENSAH DANIEL JOEL)."""
    ta, tb = _tokens(a), _tokens(b)
    if not ta or not tb:
        return False
    return len(ta & tb) >= min(2, len(ta), len(tb))


def quantite_kg(texte):
    """'50 MT' -> 50000 ; '44' -> 44000 ; '30000 KG' -> 30000"""
    t = _s(texte).upper().replace(",", ".")
    m = re.search(r"(\d+(?:\.\d+)?)", t)
    if not m:
        return 0.0
    v = float(m.group(1))
    return v if "KG" in t else v * 1000


def camion_different(camion_bon, camion_livraison):
    """True si le camion de la livraison n'est pas celui prévu sur le bon (ordre tracteur/remorque ignoré)."""
    def plaques(t):
        return {x for x in re.split(r"[/&,;\s]+", _s(t).upper()) if x}
    a, b = plaques(camion_bon), plaques(camion_livraison)
    return bool(a and b and a != b)


def statut_paiement(du, paye):
    solde = round(_n(du) - _n(paye))
    if solde == 0:
        return "PAYE"
    if solde < 0:
        return "SURPLUS"
    if _n(paye) <= 0:
        return "NON PAYE"
    return "PARTIEL"


# ---------- Lecture / écriture ----------
def _lire_feuille(feuille, colonnes):
    chemin = fichier_actif()
    if not os.path.exists(chemin):
        return pd.DataFrame(columns=colonnes)
    try:
        df = pd.read_excel(chemin, sheet_name=feuille)
    except ValueError:
        return pd.DataFrame(columns=colonnes)
    for c in colonnes:
        if c not in df.columns:
            df[c] = ""
    return df[colonnes].copy()


def charger_dechargements():
    df = _lire_feuille(FEUILLE_DECH, COLS_DECH)
    for c in ["Transaction", "Type_Mouvement", "Ref_BC", "Code_Frs", "Fournisseur",
              "Fournisseur_BC", "Code_Frs_BC", "Affectation_BC",
              "N_Camion", "Camion_BC", "Weighment_Ticket",
              "Lieu_Stockage", "Statut_Livraison", "Remarque", "Cree_Par", "Heure_Creation"]:
        df[c] = df[c].map(_s)
    # Normaliser Affectation_BC
    df["Affectation_BC"] = df["Affectation_BC"].map(
        lambda x: "OUI" if _s(x).upper() in ("OUI", "YES", "1", "TRUE") else ("NON" if _s(x) else "")
    )
    for c in ["Poids_Net_KG", "Nb_Sacs", "PU_FCFA_KG", "Montant_Du"]:
        df[c] = df[c].map(_n)
    df["Date_Dechargement"] = pd.to_datetime(df["Date_Dechargement"], errors="coerce", dayfirst=True)
    # Type ACHAT / VENTE (anciennes lignes sans colonne = ACHAT)
    df["Type_Mouvement"] = df["Type_Mouvement"].map(
        lambda x: "VENTE" if _s(x).upper().startswith("V") else "ACHAT"
    )
    return df[df["Transaction"] != ""].reset_index(drop=True)


def charger_affectations():
    df = _lire_feuille(FEUILLE_AFF, COLS_AFF)
    if not isinstance(df, pd.DataFrame):
        return pd.DataFrame(columns=COLS_AFF)
    for c in COLS_AFF:
        if c not in df.columns:
            df[c] = "" if c not in ("ID", "Montant") else 0
    for c in ["Transaction", "Bons_Paiement", "Approuve_Par", "Heure"]:
        df[c] = df[c].map(_s)
    df["Montant"] = df["Montant"].map(_n)
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
    df["Date_Approbation"] = pd.to_datetime(df["Date_Approbation"], errors="coerce", dayfirst=True)
    return df[df["Transaction"] != ""].reset_index(drop=True)


def _lire_feuille_nom(chemin, feuille, colonnes):
    if not os.path.exists(chemin):
        return pd.DataFrame(columns=colonnes)
    try:
        df = pd.read_excel(chemin, sheet_name=feuille)
    except ValueError:
        return pd.DataFrame(columns=colonnes)
    for c in colonnes:
        if c not in df.columns:
            df[c] = ""
    return df


def sauver(dech, aff, corb_dech=None, corb_aff=None):
    """Écriture atomique du classeur (aucune perte si l'écriture échoue).
    Préserve les feuilles corbeille si non fournies.
    """
    cible = fichier_actif()
    tmp = cible.replace(".xlsx", ".tmp.xlsx")

    if not isinstance(dech, pd.DataFrame):
        dech = pd.DataFrame(columns=COLS_DECH)
    if not isinstance(aff, pd.DataFrame):
        aff = pd.DataFrame(columns=COLS_AFF)

    dech = dech.copy()
    aff = aff.copy()

    for c in COLS_DECH:
        if c not in dech.columns:
            dech[c] = ""
    for c in COLS_AFF:
        if c not in aff.columns:
            if c == "ID":
                aff[c] = 0
            elif c == "Montant":
                aff[c] = 0.0
            else:
                aff[c] = ""

    # Corbeille : garder l'existant si non passé
    cols_cd = COLS_DECH + COLS_CORBEILLE_META
    cols_ca = COLS_AFF + COLS_CORBEILLE_META
    if corb_dech is None:
        corb_dech = _lire_feuille_nom(cible, FEUILLE_CORBEILLE_DECH, cols_cd)
    if corb_aff is None:
        corb_aff = _lire_feuille_nom(cible, FEUILLE_CORBEILLE_AFF, cols_ca)
    if not isinstance(corb_dech, pd.DataFrame):
        corb_dech = pd.DataFrame(columns=cols_cd)
    if not isinstance(corb_aff, pd.DataFrame):
        corb_aff = pd.DataFrame(columns=cols_ca)
    for c in cols_cd:
        if c not in corb_dech.columns:
            corb_dech[c] = ""
    for c in cols_ca:
        if c not in corb_aff.columns:
            corb_aff[c] = "" if c != "Montant" else 0.0

    d = dech.loc[:, COLS_DECH].copy()
    d["Transaction"] = d["Transaction"].map(_cle_excel)

    a = aff.loc[:, COLS_AFF].copy()
    a["Transaction"] = a["Transaction"].map(_cle_excel)

    cd = corb_dech.loc[:, [c for c in cols_cd if c in corb_dech.columns]].copy()
    ca = corb_aff.loc[:, [c for c in cols_ca if c in corb_aff.columns]].copy()

    with pd.ExcelWriter(tmp, engine="openpyxl", datetime_format="DD/MM/YYYY") as w:
        d.to_excel(w, sheet_name=FEUILLE_DECH, index=False)
        a.to_excel(w, sheet_name=FEUILLE_AFF, index=False)
        cd.to_excel(w, sheet_name=FEUILLE_CORBEILLE_DECH, index=False)
        ca.to_excel(w, sheet_name=FEUILLE_CORBEILLE_AFF, index=False)
        for ws in w.book.worksheets:
            for col in ws.columns:
                largeur = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(12, largeur + 2), 45)
            ws.freeze_panes = "A2"
    os.replace(tmp, cible)


def calculer(dech, aff):
    """Ajoute Montant_Paye, Solde, Statut_Paiement, Bons_Paiement à chaque déchargement."""
    df = dech.copy()
    if aff.empty:
        df["Montant_Paye"] = 0.0
        df["Bons_Paiement"] = ""
    else:
        paye = aff.groupby("Transaction")["Montant"].sum()
        bons = aff.groupby("Transaction")["Bons_Paiement"].apply(
            lambda s: " & ".join(dict.fromkeys(b for t in s for b in liste_bons_texte(t)))
        )
        df["Montant_Paye"] = df["Transaction"].map(paye).fillna(0.0)
        df["Bons_Paiement"] = df["Transaction"].map(bons).fillna("")
    df["Solde"] = df["Montant_Du"] - df["Montant_Paye"]
    df["Camion_Diff"] = [camion_different(a, b) for a, b in zip(df["Camion_BC"], df["N_Camion"])]
    df["Statut_Paiement"] = [statut_paiement(d, p) for d, p in zip(df["Montant_Du"], df["Montant_Paye"])]
    # Affectation BC : si colonnes absentes (ancien fichier), déduire
    if "Fournisseur_BC" not in df.columns:
        df["Fournisseur_BC"] = ""
    if "Affectation_BC" not in df.columns:
        df["Affectation_BC"] = ""
    # Si Affectation vide mais Fournisseur_BC renseigné et différent → OUI
    def _aff(row):
        a = _s(row.get("Affectation_BC", "")).upper()
        if a in ("OUI", "NON"):
            return a
        fb = _s(row.get("Fournisseur_BC", ""))
        ft = _s(row.get("Fournisseur", ""))
        if fb and ft and fb.casefold() != ft.casefold():
            return "OUI"
        return "NON" if (fb or ft) else ""
    df["Affectation_BC"] = df.apply(_aff, axis=1)
    return df


def charger_complet():
    """Retourne (déchargements calculés, affectations)."""
    aff = charger_affectations()
    return calculer(charger_dechargements(), aff), aff


# ---------- Actions ----------

def _norm_type_mvt(v):
    """Normalise ACHAT / VENTE."""
    t = _s(v).upper()
    return "VENTE" if t.startswith("V") else "ACHAT"


def tickets_pesee_existants(exclure_transaction=None, type_mouvement=None):
    """Tickets non vides déjà enregistrés (filtrable par ACHAT/VENTE)."""
    dech = charger_dechargements()
    if dech.empty or "Weighment_Ticket" not in dech.columns:
        return set()
    if exclure_transaction:
        dech = dech[dech["Transaction"].map(_s) != _s(exclure_transaction)]
    if type_mouvement and "Type_Mouvement" in dech.columns:
        tm = _norm_type_mvt(type_mouvement)
        dech = dech[dech["Type_Mouvement"].map(_norm_type_mvt) == tm]
    return {t for t in dech["Weighment_Ticket"].map(_s) if t}


def ticket_pesee_deja_utilise(ticket, exclure_transaction=None, type_mouvement=None):
    """
    Doublon si même Weighment Ticket + même type (ACHAT ou VENTE).
    Un ticket ACHAT peut coexister avec le même N° en VENTE.
    Retourne (bool, message, liste_transactions).
    """
    ticket = _s(ticket)
    if not ticket:
        return False, "", []
    dech = charger_dechargements()
    if dech.empty:
        return False, "", []
    mask = dech["Weighment_Ticket"].map(lambda x: _s(x).casefold() == ticket.casefold())
    if exclure_transaction:
        mask = mask & (dech["Transaction"].map(_s) != _s(exclure_transaction))
    tm = _norm_type_mvt(type_mouvement) if type_mouvement else None
    if tm and "Type_Mouvement" in dech.columns:
        mask = mask & (dech["Type_Mouvement"].map(_norm_type_mvt) == tm)
    hits = dech.loc[mask]
    if hits.empty:
        return False, "", []
    trs = hits["Transaction"].map(_s).tolist()
    types = (
        hits["Type_Mouvement"].map(_norm_type_mvt).tolist()
        if "Type_Mouvement" in hits.columns
        else ["ACHAT"] * len(trs)
    )
    dates = []
    for _, r in hits.iterrows():
        d = r.get("Date_Dechargement")
        if pd.notna(d) and hasattr(d, "strftime"):
            dates.append(d.strftime("%d/%m/%Y"))
        else:
            dates.append(_s(d)[:10])
    frs = hits["Fournisseur"].map(_s).tolist()
    detail = ", ".join(
        f"Trans. {t} [{ty}] ({fr or '—'}, {dt or '—'})"
        for t, ty, fr, dt in zip(trs, types, frs, dates)
    )
    lib_type = tm or "ACHAT/VENTE"
    msg = (
        f"⚠️ DOUBLON TICKET DE PESÉE ({lib_type}) : le N° « {ticket} » est déjà utilisé "
        f"pour un déchargement **{lib_type}** : {detail}. "
        "Même ticket interdit en double pour le même type (ACHAT ou VENTE). "
        "Un N° en ACHAT peut coexister avec le même N° en VENTE."
    )
    return True, msg, trs


def prochain_transaction(dech):
    nums = [int(t) for t in dech["Transaction"] if str(t).isdigit()]
    return str(max(nums) + 1) if nums else "1001"


def ajouter_dechargement(d, utilisateur):
    dech, aff = charger_dechargements(), charger_affectations()
    tr = _s(d["Transaction"])
    if tr in set(dech["Transaction"]):
        return False, f"La transaction {tr} existe déjà."
    ticket = _s(d.get("Weighment_Ticket"))
    if ticket:
        type_mvt = _norm_type_mvt(d.get("Type_Mouvement"))
        doublon, msg_ticket, _trs = ticket_pesee_deja_utilise(ticket, type_mouvement=type_mvt)
        if doublon:
            return False, msg_ticket
    poids, pu = _n(d["Poids_Net_KG"]), _n(d["PU_FCFA_KG"])
    frs_ticket = _s(d.get("Fournisseur"))
    frs_bc = _s(d.get("Fournisseur_BC"))
    code_bc = _s(d.get("Code_Frs_BC"))
    aff = _s(d.get("Affectation_BC")).upper()
    if aff not in ("OUI", "NON"):
        # auto : ticket ≠ BC
        aff = "OUI" if (frs_bc and frs_ticket and frs_bc.casefold() != frs_ticket.casefold()) else "NON"

    # --- Création automatique du BC si demandé et Ref_BC vide ---
    ref_bc = _s(d.get("Ref_BC"))
    msg_bc = ""
    if not ref_bc and d.get("Creer_BC_Auto"):
        ok_bc, msg_bc, ref_bc = creer_bon_depuis_dechargement(d, utilisateur)
        if not ok_bc:
            return False, msg_bc
        # le fournisseur du BC = celui du ticket
        if not frs_bc:
            frs_bc = frs_ticket
        if not code_bc:
            code_bc = _s(d.get("Code_Frs"))
        d = dict(d)
        d["Ref_BC"] = ref_bc
        d["Camion_BC"] = _s(d.get("Camion_BC") or d.get("N_Camion"))

    ligne = {
        "Transaction": tr,
        "Type_Mouvement": ("VENTE" if _s(d.get("Type_Mouvement")).upper().startswith("V") else "ACHAT"),
        "Date_Dechargement": pd.Timestamp(d["Date_Dechargement"]),
        "Ref_BC": ref_bc,
        "Code_Frs": _s(d.get("Code_Frs")),
        "Fournisseur": frs_ticket,  # compte fournisseur = nom du ticket
        "Fournisseur_BC": frs_bc,
        "Code_Frs_BC": code_bc,
        "Affectation_BC": aff,
        "N_Camion": _s(d.get("N_Camion")),
        "Camion_BC": _s(d.get("Camion_BC")),
        "Weighment_Ticket": _s(d.get("Weighment_Ticket")),
        "Poids_Net_KG": poids,
        "Nb_Sacs": _n(d.get("Nb_Sacs")),
        "PU_FCFA_KG": pu,
        "Montant_Du": round(poids * pu),
        "Lieu_Stockage": _s(d.get("Lieu_Stockage")),
        "Statut_Livraison": "DECHARGE",
        "Remarque": _s(d.get("Remarque")),
        "Cree_Par": utilisateur,
        "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }
    for c in COLS_DECH:
        if c not in dech.columns:
            dech[c] = ""
        if c not in ligne:
            ligne[c] = "" if c not in ("Poids_Net_KG", "Nb_Sacs", "PU_FCFA_KG", "Montant_Du") else 0
    dech = pd.concat([dech, pd.DataFrame([ligne])], ignore_index=True)
    # Recharger aff proprement au cas où (ne jamais passer autre chose qu'un DataFrame)
    if not isinstance(aff, pd.DataFrame):
        aff = charger_affectations()
    sauver(dech, aff)
    msg = f"Déchargement {tr} enregistré."
    if msg_bc:
        msg = f"{msg} {msg_bc}"
    return True, msg


def charger_corbeille():
    """Retourne (déchargements_corbeille, affectations_corbeille) pour le produit actif."""
    cible = fichier_actif()
    cols_cd = COLS_DECH + COLS_CORBEILLE_META
    cols_ca = COLS_AFF + COLS_CORBEILLE_META
    cd = _lire_feuille_nom(cible, FEUILLE_CORBEILLE_DECH, cols_cd)
    ca = _lire_feuille_nom(cible, FEUILLE_CORBEILLE_AFF, cols_ca)
    if not cd.empty and "Transaction" in cd.columns:
        cd = cd[cd["Transaction"].map(_s) != ""].copy()
        if "Date_Dechargement" in cd.columns:
            cd["Date_Dechargement"] = pd.to_datetime(cd["Date_Dechargement"], errors="coerce", dayfirst=True)
        if "Type_Mouvement" in cd.columns:
            cd["Type_Mouvement"] = cd["Type_Mouvement"].map(
                lambda x: "VENTE" if _s(x).upper().startswith("V") else "ACHAT"
            )
    if not ca.empty and "Transaction" in ca.columns:
        ca = ca[ca["Transaction"].map(_s) != ""].copy()
        ca["Montant"] = ca["Montant"].map(_n)
    return cd.reset_index(drop=True), ca.reset_index(drop=True)


def supprimer_dechargement(transaction, utilisateur="", motif=""):
    """
    Supprime un déchargement en le déplaçant en corbeille (annulable).
    Les paiements affectés suivent en corbeille.
    """
    dech, aff = charger_dechargements(), charger_affectations()
    tr = _s(transaction)
    if tr not in set(dech["Transaction"].map(_s)):
        return False, "Transaction introuvable."
    ligne = dech[dech["Transaction"].map(_s) == tr].copy()
    aff_l = aff[aff["Transaction"].map(_s) == tr].copy()
    now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    user = _s(utilisateur) or "—"
    motif = _s(motif) or "Suppression manuelle"
    for c in COLS_CORBEILLE_META:
        ligne[c] = ""
        if not aff_l.empty:
            aff_l[c] = ""
    ligne["Supprime_Le"] = now
    ligne["Supprime_Par"] = user
    ligne["Motif_Suppression"] = motif
    if not aff_l.empty:
        aff_l["Supprime_Le"] = now
        aff_l["Supprime_Par"] = user
        aff_l["Motif_Suppression"] = motif

    corb_d, corb_a = charger_corbeille()
    corb_d = pd.concat([corb_d, ligne], ignore_index=True)
    if not aff_l.empty:
        corb_a = pd.concat([corb_a, aff_l], ignore_index=True)

    dech2 = dech[dech["Transaction"].map(_s) != tr]
    aff2 = aff[aff["Transaction"].map(_s) != tr]
    sauver(dech2, aff2, corb_dech=corb_d, corb_aff=corb_a)
    return True, (
        f"Déchargement {tr} déplacé en corbeille "
        f"(avec {len(aff_l)} paiement(s) affecté(s)). "
        "Vous pouvez l'annuler depuis la corbeille."
    )


def restaurer_dechargement(transaction):
    """Restaure un déchargement (et ses paiements) depuis la corbeille."""
    dech, aff = charger_dechargements(), charger_affectations()
    corb_d, corb_a = charger_corbeille()
    tr = _s(transaction)
    if corb_d.empty or tr not in set(corb_d["Transaction"].map(_s)):
        return False, "Transaction introuvable dans la corbeille."
    if tr in set(dech["Transaction"].map(_s)):
        return False, (
            f"Impossible de restaurer {tr} : une transaction active porte déjà ce N°. "
            "Renommez ou supprimez d'abord la ligne active."
        )
    # Contrôle ticket de pesée + type avant restauration
    row = corb_d[corb_d["Transaction"].map(_s) == tr].iloc[0]
    ticket = _s(row.get("Weighment_Ticket"))
    type_mvt = _norm_type_mvt(row.get("Type_Mouvement"))
    if ticket:
        doublon, msg_ticket, _ = ticket_pesee_deja_utilise(ticket, type_mouvement=type_mvt)
        if doublon:
            return False, f"Restauration refusée — {msg_ticket}"

    ligne = corb_d[corb_d["Transaction"].map(_s) == tr].copy()
    aff_l = corb_a[corb_a["Transaction"].map(_s) == tr].copy() if not corb_a.empty else pd.DataFrame(columns=COLS_AFF)
    for c in COLS_CORBEILLE_META:
        if c in ligne.columns:
            ligne = ligne.drop(columns=[c])
        if not aff_l.empty and c in aff_l.columns:
            aff_l = aff_l.drop(columns=[c])
    for c in COLS_DECH:
        if c not in ligne.columns:
            ligne[c] = ""
    for c in COLS_AFF:
        if c not in aff_l.columns:
            aff_l[c] = 0 if c in ("ID", "Montant") else ""

    dech2 = pd.concat([dech, ligne[COLS_DECH]], ignore_index=True)
    if not aff_l.empty:
        # réattribuer des ID uniques si collision
        max_id = int(aff["ID"].max()) if not aff.empty and "ID" in aff.columns else 0
        try:
            max_id = int(pd.to_numeric(aff["ID"], errors="coerce").max() or 0) if not aff.empty else 0
        except Exception:
            max_id = 0
        aff_l = aff_l.copy()
        for i in aff_l.index:
            max_id += 1
            aff_l.at[i, "ID"] = max_id
        aff2 = pd.concat([aff, aff_l[COLS_AFF]], ignore_index=True)
    else:
        aff2 = aff
    corb_d2 = corb_d[corb_d["Transaction"].map(_s) != tr]
    corb_a2 = corb_a[corb_a["Transaction"].map(_s) != tr] if not corb_a.empty else corb_a
    sauver(dech2, aff2, corb_dech=corb_d2, corb_aff=corb_a2)
    return True, f"Déchargement {tr} restauré depuis la corbeille (paiements inclus)."


def purger_corbeille(transaction=None):
    """Suppression définitive : une transaction ou toute la corbeille."""
    corb_d, corb_a = charger_corbeille()
    dech, aff = charger_dechargements(), charger_affectations()
    if transaction:
        tr = _s(transaction)
        if tr not in set(corb_d["Transaction"].map(_s)):
            return False, "Transaction introuvable dans la corbeille."
        corb_d = corb_d[corb_d["Transaction"].map(_s) != tr]
        corb_a = corb_a[corb_a["Transaction"].map(_s) != tr] if not corb_a.empty else corb_a
        sauver(dech, aff, corb_dech=corb_d, corb_aff=corb_a)
        return True, f"Transaction {tr} effacée définitivement de la corbeille."
    sauver(dech, aff, corb_dech=pd.DataFrame(columns=COLS_DECH + COLS_CORBEILLE_META),
           corb_aff=pd.DataFrame(columns=COLS_AFF + COLS_CORBEILLE_META))
    return True, "Corbeille entièrement vidée."


def modifier_dechargement(transaction, champs, utilisateur=""):
    """Met à jour un déchargement existant.
    champs = dict avec les clés à modifier (Poids_Net_KG, PU_FCFA_KG, N_Camion, etc.).
    Recalcule automatiquement Montant_Du = Poids × PU.
    """
    dech, aff = charger_dechargements(), charger_affectations()
    if transaction not in set(dech["Transaction"]):
        return False, "Transaction introuvable."
    idx = dech[dech["Transaction"] == transaction].index
    if len(idx) == 0:
        return False, "Transaction introuvable."
    i = idx[0]
    # Anti-doublon ticket de pesée + type ACHAT/VENTE
    if "Weighment_Ticket" in champs or "Type_Mouvement" in champs:
        ticket_new = _s(champs.get("Weighment_Ticket"))
        if not ticket_new:
            try:
                ticket_new = _s(dech.at[i, "Weighment_Ticket"])
            except Exception:
                ticket_new = ""
        type_mvt = champs.get("Type_Mouvement")
        if not type_mvt:
            try:
                type_mvt = dech.at[i, "Type_Mouvement"]
            except Exception:
                type_mvt = "ACHAT"
        if ticket_new:
            doublon, msg_ticket, _trs = ticket_pesee_deja_utilise(
                ticket_new, exclure_transaction=transaction, type_mouvement=type_mvt
            )
            if doublon:
                return False, msg_ticket
    # Champs texte
    for cle in ["Type_Mouvement", "Ref_BC", "Code_Frs", "Fournisseur", "Fournisseur_BC", "Code_Frs_BC",
                "Affectation_BC", "N_Camion", "Camion_BC",
                "Weighment_Ticket", "Lieu_Stockage", "Remarque", "Statut_Livraison"]:
        if cle in champs:
            dech.at[i, cle] = _s(champs[cle])
    # Date
    if "Date_Dechargement" in champs and champs["Date_Dechargement"] is not None:
        dech.at[i, "Date_Dechargement"] = pd.Timestamp(champs["Date_Dechargement"])
    # Numériques
    for cle in ["Poids_Net_KG", "Nb_Sacs", "PU_FCFA_KG"]:
        if cle in champs:
            dech.at[i, cle] = _n(champs[cle])
    # Recalcul montant dû
    poids = _n(dech.at[i, "Poids_Net_KG"])
    pu = _n(dech.at[i, "PU_FCFA_KG"])
    dech.at[i, "Montant_Du"] = round(poids * pu)
    # Recalcul affectation si Fournisseur / Fournisseur_BC ont changé
    frs_t = _s(dech.at[i, "Fournisseur"])
    frs_b = _s(dech.at[i, "Fournisseur_BC"]) if "Fournisseur_BC" in dech.columns else ""
    if frs_b and frs_t and frs_b.casefold() != frs_t.casefold():
        dech.at[i, "Affectation_BC"] = "OUI"
    elif "Affectation_BC" in dech.columns:
        # ne force NON que si les deux sont renseignés et égaux, ou si BC vide
        if not frs_b or (frs_t and frs_b.casefold() == frs_t.casefold()):
            dech.at[i, "Affectation_BC"] = "NON"
    if utilisateur:
        dech.at[i, "Cree_Par"] = _s(dech.at[i, "Cree_Par"]) or utilisateur
    sauver(dech, aff)
    return True, (f"Déchargement {transaction} modifié — "
                  f"Poids {fmt(poids, 'kg')} × PU {fmt(pu)} = {fmt(poids * pu)} FCFA.")


def supprimer_affectation(id_aff):
    """Supprime un paiement affecté (une ligne de Paiements_Affectes) par son ID."""
    dech, aff = charger_dechargements(), charger_affectations()
    id_aff = int(id_aff)
    if id_aff not in set(aff["ID"].astype(int)):
        return False, f"Paiement ID {id_aff} introuvable."
    ligne = aff[aff["ID"] == id_aff].iloc[0]
    tr = ligne["Transaction"]
    montant = ligne["Montant"]
    sauver(dech, aff[aff["ID"] != id_aff])
    return True, f"Paiement de {fmt(montant)} FCFA retiré de la livraison {tr}."


def ajouter_affectation(transaction, bons_texte, montant, utilisateur):
    """Approuve un paiement : lie bon(s) de paiement + montant à un déchargement."""
    dech, aff = charger_dechargements(), charger_affectations()
    if transaction not in set(dech["Transaction"]):
        return False, "Transaction introuvable."
    if _n(montant) <= 0:
        return False, "Le montant doit être supérieur à 0."
    nouvel_id = int(aff["ID"].max()) + 1 if not aff.empty else 1
    ligne = {
        "ID": nouvel_id, "Transaction": transaction,
        "Date_Approbation": pd.Timestamp(datetime.now().date()),
        "Bons_Paiement": " & ".join(liste_bons_texte(bons_texte)),
        "Montant": float(montant), "Approuve_Par": utilisateur,
        "Heure": datetime.now().strftime("%H:%M:%S"),
    }
    aff = pd.concat([aff, pd.DataFrame([ligne])], ignore_index=True)
    sauver(dech, aff)
    return True, "Paiement approuvé."


def annuler_paiements(transaction):
    """Remet la livraison en NON PAYÉ (supprime les paiements affectés)."""
    dech, aff = charger_dechargements(), charger_affectations()
    n = int((aff["Transaction"] == transaction).sum())
    sauver(dech, aff[aff["Transaction"] != transaction])
    return True, f"{n} paiement(s) retiré(s) : la livraison {transaction} est NON PAYÉE."


# ---------- Données liées ----------
def charger_bons_commande():
    cols = ["Reference", "Date_Emission", "Produit", "Code_Fournisseur", "Nom_Fournisseur",
            "Quantite", "Numero_Tracteur", "Chauffeur"]
    if not os.path.exists(FICHIER_BONS):
        df = pd.DataFrame(columns=cols + ["Qte_KG"])
        return df
    df = pd.read_excel(FICHIER_BONS)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols].copy()
    for c in cols:
        df[c] = df[c].map(_s)
    df = df[df["Reference"] != ""]
    # seuls les bons du produit actif (ex. 0007-BC-SOYABEAN/NAP-2026)
    motif = f"-BC-{get_produit()}"
    df = df[masque(df, [motif in r.upper() for r in df["Reference"]])]
    df["Qte_KG"] = df["Quantite"].map(quantite_kg)
    return df.reset_index(drop=True)


def prochain_numero_bc(produit=None):
    """Prochain N° séquentiel pour XXX-BC-{PRODUIT}/NAP-{année} (tous produits dans le fichier)."""
    produit = (produit or get_produit()).upper()
    annee = datetime.now().year
    max_n = 0
    if os.path.exists(FICHIER_BONS):
        try:
            df = pd.read_excel(FICHIER_BONS)
            if "Reference" in df.columns:
                for r in df["Reference"].map(_s):
                    m = re.match(r"^(\d+)-BC-", str(r).upper())
                    if m:
                        max_n = max(max_n, int(m.group(1)))
        except Exception:
            pass
    return max_n + 1


def creer_bon_depuis_dechargement(d, utilisateur=""):
    """Crée un bon de commande dans donnees_bons.xlsx à partir d'un déchargement sans BC.

    Retourne (ok: bool, message: str, reference: str).
    """
    produit = get_produit()
    annee = datetime.now().year
    n = prochain_numero_bc(produit)
    ref = f"{n:04d}-BC-{produit}/NAP-{annee}"

    frs = _s(d.get("Fournisseur") or d.get("Nom_Fournisseur"))
    code = _s(d.get("Code_Frs") or d.get("Code_Fournisseur"))
    poids = _n(d.get("Poids_Net_KG"))
    # Quantité en MT si poids >= 100, sinon afficher en KG
    if poids >= 100:
        qte_txt = f"{round(poids / 1000, 2)} MT".replace(".0 MT", " MT")
        if qte_txt.endswith(".0 MT"):
            qte_txt = qte_txt.replace(".0 MT", " MT")
    else:
        qte_txt = f"{int(poids)} KG" if poids else ""

    camion = _s(d.get("N_Camion") or d.get("Numero_Tracteur"))
    date_d = d.get("Date_Dechargement") or date.today()
    if hasattr(date_d, "strftime"):
        date_str = date_d.strftime("%d/%m/%Y")
    else:
        date_str = _s(date_d) or datetime.now().strftime("%d/%m/%Y")

    libelle = libelle_produit(produit)
    ligne = {
        "Reference": ref,
        "Date_Emission": date_str,
        "Heure_Creation": datetime.now().strftime("%H:%M:%S"),
        "Utilisateur": _s(utilisateur) or "Auto-déchargement",
        "Produit": libelle,
        "Code_Fournisseur": code,
        "Nom_Fournisseur": frs,
        "Telephone_Fournisseur": _s(d.get("Telephone_Fournisseur")),
        "Commune_Origine": _s(d.get("Commune_Origine")),
        "Quantite": qte_txt,
        "Numero_Tracteur": camion,
        "Chauffeur": _s(d.get("Chauffeur")),
        "Numero_Permis": _s(d.get("Numero_Permis")),
        "Telephone_Chauffeur": _s(d.get("Telephone_Chauffeur")),
    }

    cols_all = list(ligne.keys())
    if os.path.exists(FICHIER_BONS):
        try:
            df = pd.read_excel(FICHIER_BONS)
        except Exception:
            df = pd.DataFrame(columns=cols_all)
    else:
        df = pd.DataFrame(columns=cols_all)

    for c in cols_all:
        if c not in df.columns:
            df[c] = ""
    # Éviter doublon de référence
    if ref in set(df["Reference"].map(_s)):
        return False, f"La référence {ref} existe déjà.", ""

    df = pd.concat([df, pd.DataFrame([ligne])], ignore_index=True)
    tmp = FICHIER_BONS.replace(".xlsx", ".tmp.xlsx")
    df.to_excel(tmp, index=False)
    os.replace(tmp, FICHIER_BONS)
    return True, f"Bon de commande créé automatiquement : {ref}", ref


def charger_fournisseurs_df():
    """fournisseurs.xlsx -> colonnes Code, Nom (fournisseurs actifs)."""
    if not os.path.exists(FICHIER_FRS):
        return pd.DataFrame(columns=["Code", "Nom"])
    df = pd.read_excel(FICHIER_FRS)
    if "Actif" in df.columns:
        df = df[df["Actif"].astype(str).str.upper() == "OUI"]
    df = pd.DataFrame({
        "Code": df.get("Code", pd.Series(dtype=str)).map(_s),
        "Nom": df.get("Nom_Fournisseur", pd.Series(dtype=str)).map(_s),
    })
    return df[(df["Code"] != "") & (df["Nom"] != "")].reset_index(drop=True)


def tous_les_fournisseurs(dech):
    """dict { clé (code ou nom) : nom } = fournisseurs.xlsx + ceux déjà présents dans les déchargements."""
    out = {}
    for code, nom in charger_fournisseurs_df().itertuples(index=False):
        out[code] = nom
    for code, nom in dech[["Code_Frs", "Fournisseur"]].drop_duplicates().itertuples(index=False):
        cle = code or nom
        if cle and cle not in out:
            out[cle] = nom or cle
    return out


def fournisseurs_connus():
    """fournisseurs.xlsx + fournisseurs déjà vus dans les déchargements de TOUS les produits."""
    ancien = get_produit()
    frames = []
    try:
        for p in PRODUITS:
            set_produit(p)
            frames.append(charger_dechargements()[["Code_Frs", "Fournisseur"]])
    finally:
        set_produit(ancien)
    pleins = [f for f in frames if not f.empty]
    tous = pd.concat(pleins, ignore_index=True) if pleins else pd.DataFrame(columns=["Code_Frs", "Fournisseur"])
    return tous_les_fournisseurs(tous)


def resoudre_fournisseur_bon(frs, code_bon, nom_bon):
    """Trouve (clé, nom, connu) du fournisseur d'un bon de commande dans le dict `frs`.
    Ordre : code FRS du bon -> même nom -> nom très proche (si un seul candidat) -> sinon fournisseur « nouveau »."""
    code_bon, nom_bon = _s(code_bon), _s(nom_bon)
    if code_bon and code_bon in frs:
        return code_bon, frs[code_bon], True
    for k, n in frs.items():
        if nom_bon and n.casefold() == nom_bon.casefold():
            return k, n, True
    proches = [(k, n) for k, n in frs.items() if nom_bon and meme_nom(n, nom_bon)]
    if len(proches) == 1:
        return proches[0][0], proches[0][1], True
    if nom_bon or code_bon:
        return (code_bon or nom_bon), (nom_bon or code_bon), False
    return None, None, False


def label_fournisseur(cle, nom):
    return nom if cle == nom else f"{cle} — {nom}"


def filtrer_fournisseur(df, cle, nom):
    meme_nom_ = df["Fournisseur"].map(lambda x: _s(x).casefold() == nom.casefold())
    return df[(df["Code_Frs"] == cle) | ((df["Code_Frs"] == "") & meme_nom_)]


def charger_journal():
    """Journal des paiements (feuille DEPENSES), regroupé par N° de bon de paiement."""
    chemin = FICHIER_PAIEMENTS if os.path.exists(FICHIER_PAIEMENTS) else next(iter(glob.glob("Suivi*aiement*.xlsx")), None)
    vide = pd.DataFrame(columns=["No_Transaction", "Date", "Produit", "Prestataire", "Ref_Doc", "Montant",
                                 "Observation", "Mode", "Num_Cheque", "Banque"])
    if not chemin:
        return vide
    try:
        xl = pd.ExcelFile(chemin)
        feuille = FEUILLE_DEPENSES if FEUILLE_DEPENSES in xl.sheet_names else xl.sheet_names[0]
        df = None
        for hdr in (2, 1, 0, 3):
            tmp = pd.read_excel(chemin, sheet_name=feuille, header=hdr)
            if any("transaction" in str(c).lower() for c in tmp.columns):
                df = tmp
                break
        if df is None:
            return vide
    except Exception:
        return vide
    alias = {"date": "Date", "no de transaction": "No_Transaction", "produit": "Produit",
             "prestataires/fournisseurs": "Prestataire", "ref doc": "Ref_Doc", "montant": "Montant",
             "observation": "Observation", "mode paiement": "Mode", "numero de cheque": "Num_Cheque",
             "banque": "Banque"}
    df = df.rename(columns={c: alias[str(c).strip().lower()] for c in df.columns if str(c).strip().lower() in alias})
    for c in vide.columns:
        if c not in df.columns:
            df[c] = ""
    df = df[vide.columns].copy()
    for c in ["No_Transaction", "Produit", "Prestataire", "Ref_Doc", "Observation", "Mode", "Num_Cheque", "Banque"]:
        df[c] = df[c].map(_s)
    df["Produit"] = df["Produit"].map(lambda x: PRODUITS_ALIAS.get(x.upper(), x))
    df["Montant"] = df["Montant"].map(_n)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)
    df = df[df["No_Transaction"] != ""]
    return df.reset_index(drop=True)


def bons_paiement_achats(journal):
    """1 ligne par N° de bon de paiement, uniquement pour le produit actif."""
    j = journal[masque(journal, [_s(p).upper() == get_produit() for p in journal["Produit"]])]
    if j.empty:
        return pd.DataFrame(columns=["No_Transaction", "Date", "Prestataire", "Montant", "Observation", "Mode"])
    return (j.groupby("No_Transaction", sort=False)
              .agg(Date=("Date", "first"), Prestataire=("Prestataire", "first"),
                   Montant=("Montant", "sum"), Observation=("Observation", "first"), Mode=("Mode", "first"))
              .reset_index())


# ---------- Rapports ----------




def modifier_quantite_bon(reference, quantite_texte=None, poids_kg=None, utilisateur=""):
    """
    Modifie la quantité commandée d'un bon de commande dans donnees_bons.xlsx.
    - quantite_texte : ex. "48.5 MT" ou "48500 KG"
    - poids_kg : si fourni et quantite_texte vide, convertit en "X.XX MT"
    Retourne (ok, message).
    """
    ref = _s(reference)
    if not ref:
        return False, "Référence de bon vide."
    if not os.path.exists(FICHIER_BONS):
        return False, f"Fichier {FICHIER_BONS} introuvable."
    try:
        df = pd.read_excel(FICHIER_BONS)
    except Exception as e:
        return False, f"Lecture impossible : {e}"
    if "Reference" not in df.columns:
        return False, "Colonne Reference absente dans donnees_bons.xlsx."
    mask = df["Reference"].map(_s) == ref
    if not mask.any():
        return False, f"Bon {ref} introuvable dans donnees_bons.xlsx."

    if quantite_texte and _s(quantite_texte):
        qte = _s(quantite_texte)
    elif poids_kg is not None and _n(poids_kg) > 0:
        pk = _n(poids_kg)
        if pk >= 100:
            qte = f"{round(pk / 1000, 3):.3f} MT".rstrip("0").rstrip(".") + " MT"
            qte = qte.replace(" MT MT", " MT")
            if not qte.endswith("MT"):
                qte = f"{round(pk / 1000, 3)} MT"
        else:
            qte = f"{int(round(pk))} KG"
    else:
        return False, "Indiquez une quantité (texte) ou un poids en kg."

    idx = df.index[mask][0]
    ancienne = _s(df.at[idx, "Quantite"]) if "Quantite" in df.columns else ""
    if "Quantite" not in df.columns:
        df["Quantite"] = ""
    df.at[idx, "Quantite"] = qte
    # trace optionnelle
    if "Remarque_Qte" not in df.columns:
        df["Remarque_Qte"] = ""
    note = f"{datetime.now().strftime('%d/%m/%Y %H:%M')} {_s(utilisateur)}: {ancienne or '—'} → {qte}"
    prev = _s(df.at[idx, "Remarque_Qte"])
    df.at[idx, "Remarque_Qte"] = (prev + " | " + note) if prev else note

    tmp = FICHIER_BONS.replace(".xlsx", ".tmp.xlsx")
    try:
        df.to_excel(tmp, index=False)
        os.replace(tmp, FICHIER_BONS)
    except PermissionError:
        return False, "Fichier donnees_bons.xlsx ouvert dans Excel — fermez-le puis réessayez."
    except Exception as e:
        return False, f"Écriture impossible : {e}"

    return True, (
        f"Quantité du bon **{ref}** mise à jour : "
        f"{ancienne or '—'} → **{qte}**. "
        "Le suivi commandé ↔ livré utilisera cette nouvelle quantité."
    )


def poids_livre_sur_bc(reference, dech=None):
    """Poids net total déjà déchargé sur ce N° BC (produit actif)."""
    ref = _s(reference)
    if dech is None:
        dech = charger_dechargements()
    if dech.empty or not ref:
        return 0.0
    m = dech["Ref_BC"].map(_s) == ref
    return float(dech.loc[m, "Poids_Net_KG"].map(_n).sum())


def rechercher_dechargement(requete, type_mouvement=None):
    """
    Recherche un déchargement par N° camion ou Weighment Ticket / Ticket pont.
    Retourne DataFrame des lignes correspondantes (produit actif).
    """
    q = _s(requete).upper()
    q_compact = re.sub(r"[\s\-/._]", "", q)
    if len(q_compact) < 2:
        return pd.DataFrame()
    dech = charger_dechargements()
    if dech.empty:
        return pd.DataFrame()
    if type_mouvement:
        tm = _norm_type_mvt(type_mouvement)
        dech = dech[dech["Type_Mouvement"].map(_norm_type_mvt) == tm]
    if dech.empty:
        return pd.DataFrame()

    def _match(val):
        t = _s(val).upper()
        if not t:
            return False
        if q in t:
            return True
        t2 = re.sub(r"[\s\-/._]", "", t)
        return q_compact in t2 if q_compact else False

    mask = dech["N_Camion"].map(_match) | dech["Weighment_Ticket"].map(_match)
    # aussi transaction si purement numérique
    if q_compact.isdigit():
        mask = mask | dech["Transaction"].map(lambda x: q_compact in re.sub(r"\D", "", _s(x)))
    return dech.loc[mask].reset_index(drop=True)


def suivi_par_bc(dech, bons):
    """Un bon de commande = une ligne : commandé vs livré, dû / payé / solde, statuts."""
    d = dech[dech["Ref_BC"] != ""]
    if d.empty:
        g = pd.DataFrame(columns=["Reference", "Nb_Camions", "Poids_Livre_KG", "Nb_Sacs", "Montant_Du",
                                  "Montant_Paye", "Derniere_Date", "Code_Frs", "Fournisseur_Dech"])
    else:
        g = d.groupby("Ref_BC").agg(
            Nb_Camions=("Transaction", "count"), Poids_Livre_KG=("Poids_Net_KG", "sum"),
            Nb_Sacs=("Nb_Sacs", "sum"), Montant_Du=("Montant_Du", "sum"), Montant_Paye=("Montant_Paye", "sum"),
            Derniere_Date=("Date_Dechargement", "max"), Code_Frs=("Code_Frs", "first"),
            Fournisseur_Dech=("Fournisseur", "first"),
        ).reset_index().rename(columns={"Ref_BC": "Reference"})
    b = bons[["Reference", "Nom_Fournisseur", "Produit", "Qte_KG"]].copy() if not bons.empty else pd.DataFrame(
        columns=["Reference", "Nom_Fournisseur", "Produit", "Qte_KG"])
    m = b.merge(g, on="Reference", how="outer")
    for c in ["Nb_Camions", "Poids_Livre_KG", "Nb_Sacs", "Montant_Du", "Montant_Paye", "Qte_KG"]:
        m[c] = pd.to_numeric(m[c], errors="coerce").fillna(0.0)
    for c in ["Nom_Fournisseur", "Fournisseur_Dech", "Code_Frs", "Produit"]:
        m[c] = m[c].map(_s)
    m["Fournisseur"] = pd.Series([a or b_ for a, b_ in zip(m["Nom_Fournisseur"], m["Fournisseur_Dech"])],
                                 index=m.index, dtype="object")   # texte, même quand il n'y a aucune ligne
    m["Solde"] = m["Montant_Du"] - m["Montant_Paye"]
    m["Ecart_KG"] = [(p - q) if q > 0 else None for p, q in zip(m["Poids_Livre_KG"], m["Qte_KG"])]
    m["Pct_Livre"] = [round(p / q * 100, 1) if q > 0 else None for p, q in zip(m["Poids_Livre_KG"], m["Qte_KG"])]

    def liv(r):
        if r["Nb_Camions"] == 0:
            return "EN ATTENTE"
        if r["Qte_KG"] > 0:
            return "LIVRE" if r["Poids_Livre_KG"] >= SEUIL_LIVRE * r["Qte_KG"] else "EN COURS"
        return "DECHARGE"

    m["Statut_Livraison"] = (m.apply(liv, axis=1) if not m.empty else pd.Series([], index=m.index, dtype="object"))
    m["Statut_Paiement"] = pd.Series([statut_paiement(d_, p_) if n > 0 else "—"
                                      for d_, p_, n in zip(m["Montant_Du"], m["Montant_Paye"], m["Nb_Camions"])],
                                     index=m.index, dtype="object")
    m = m.rename(columns={"Qte_KG": "Qte_Commandee_KG"})
    cols = ["Reference", "Fournisseur", "Code_Frs", "Produit", "Qte_Commandee_KG", "Poids_Livre_KG", "Ecart_KG",
            "Pct_Livre", "Nb_Camions", "Nb_Sacs", "Montant_Du", "Montant_Paye", "Solde",
            "Statut_Livraison", "Statut_Paiement", "Derniere_Date"]
    return m[cols].sort_values("Reference", ascending=False).reset_index(drop=True)


def importer_excel_dechargement(fichier, utilisateur="Import Excel"):
    """Importe le fichier « COMMANDES DECHARGEES & PAIEMENTS ». Les transactions déjà présentes sont ignorées.
    Retourne (ajoutés, ignorés, paiements_importés)."""
    brut = pd.read_excel(fichier, sheet_name=0, header=None)
    hdr = None
    for i in range(min(15, len(brut))):
        if "transaction" in [_s(v).lower() for v in brut.iloc[i].tolist()]:
            hdr = i
            break
    if hdr is None:
        raise ValueError("Colonne « Transaction » introuvable dans les 15 premières lignes.")
    df = brut.iloc[hdr + 1:].copy()
    df.columns = [_s(v) for v in brut.iloc[hdr].tolist()]
    alias = {"datedechargement": "Date_Dechargement", "nbc": "Ref_BC", "transaction": "Transaction",
             "codefrs": "Code_Frs", "fournisseur": "Fournisseur", "ncamion": "N_Camion",
             "weighmentticket": "Weighment_Ticket", "poidsnet(kg)": "Poids_Net_KG", "nbsacs": "Nb_Sacs",
             "pu(fcfa/kg)": "PU_FCFA_KG", "montantpaye": "Paye_Import", "bondepaiement": "Bons_Import",
             "statutlivraison": "Statut_Livraison", "lieustockage": "Lieu_Stockage"}
    df = df.rename(columns={c: alias[re.sub(r"[^a-z0-9()/]", "", c.lower())]
                            for c in df.columns if re.sub(r"[^a-z0-9()/]", "", c.lower()) in alias})
    for c in ["Date_Dechargement", "Ref_BC", "Transaction", "Code_Frs", "Fournisseur", "N_Camion",
              "Weighment_Ticket", "Poids_Net_KG", "Nb_Sacs", "PU_FCFA_KG", "Paye_Import", "Bons_Import",
              "Statut_Livraison", "Lieu_Stockage"]:
        if c not in df.columns:
            df[c] = ""

    dech, aff = charger_dechargements(), charger_affectations()
    existantes = set(dech["Transaction"])
    noms = {code: nom for code, nom in charger_fournisseurs_df().itertuples(index=False)}
    nouvelles, nouv_aff, ignores = [], [], 0
    nid = int(aff["ID"].max()) + 1 if not aff.empty else 1
    for _, r in df.iterrows():
        tr = _s(r["Transaction"])
        if not tr:
            continue
        if tr in existantes:
            ignores += 1
            continue
        type_imp = _norm_type_mvt(r["Type_Mouvement"]) if "Type_Mouvement" in df.columns else "ACHAT"
        ticket_imp = _s(r["Weighment_Ticket"])
        if ticket_imp:
            # doublon ticket dans le fichier déjà chargé ou dans les nouvelles de ce lot
            tickets_mem = { _s(x) for x in dech["Weighment_Ticket"].map(_s) if _s(x) }
            tickets_mem |= { _s(n.get("Weighment_Ticket")) for n in nouvelles if _s(n.get("Weighment_Ticket")) }
            if ticket_imp.casefold() in {t.casefold() for t in tickets_mem}:
                ignores += 1
                continue
        existantes.add(tr)
        code = _s(r["Code_Frs"])
        poids, pu = _n(r["Poids_Net_KG"]), _n(r["PU_FCFA_KG"])
        date_d = pd.to_datetime(r["Date_Dechargement"], errors="coerce", dayfirst=True)
        frs_imp = _s(r["Fournisseur"]) or noms.get(code, code)
        nouvelles.append({
            "Transaction": tr, "Type_Mouvement": type_imp,
            "Date_Dechargement": date_d, "Ref_BC": _s(r["Ref_BC"]), "Code_Frs": code,
            "Fournisseur": frs_imp,
            "Fournisseur_BC": "", "Code_Frs_BC": "", "Affectation_BC": "NON",
            "N_Camion": _s(r["N_Camion"]), "Camion_BC": "",
            "Weighment_Ticket": _s(r["Weighment_Ticket"]), "Poids_Net_KG": poids, "Nb_Sacs": _n(r["Nb_Sacs"]),
            "PU_FCFA_KG": pu, "Montant_Du": round(poids * pu), "Lieu_Stockage": _s(r["Lieu_Stockage"]),
            "Statut_Livraison": _s(r["Statut_Livraison"]) or "DECHARGE", "Cree_Par": utilisateur,
            "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M:%S"), "Remarque": "",
        })
        paye = _n(r["Paye_Import"])
        if paye > 0:
            nouv_aff.append({
                "ID": nid, "Transaction": tr, "Date_Approbation": date_d,
                "Bons_Paiement": " & ".join(liste_bons_texte(r["Bons_Import"])) or "(import Excel)",
                "Montant": paye, "Approuve_Par": utilisateur, "Heure": "",
            })
            nid += 1
    if nouvelles:
        dech = pd.concat([dech, pd.DataFrame(nouvelles)], ignore_index=True)
    if nouv_aff:
        aff = pd.concat([aff, pd.DataFrame(nouv_aff)], ignore_index=True)
    sauver(dech, aff)
    return len(nouvelles), ignores, len(nouv_aff)


def to_excel_bytes(feuilles):
    """{ 'nom feuille': DataFrame } -> octets .xlsx (en-têtes en bleu ciel)"""
    from openpyxl.styles import Font, PatternFill
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="DD/MM/YYYY") as w:
        for nom, df in feuilles.items():
            df.to_excel(w, sheet_name=nom[:31], index=False)
            ws = w.sheets[nom[:31]]
            for col in ws.columns:
                largeur = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(10, largeur + 2), 45)
            ws.freeze_panes = "A2"
            for cell in ws[1]:   # en-têtes : fond bleu ciel, texte bleu foncé
                cell.fill = PatternFill("solid", fgColor="87CEEB")
                cell.font = Font(bold=True, color="1B4F72")
    return buf.getvalue()
