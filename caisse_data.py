# -*- coding: utf-8 -*-
"""
Caisse NAP SARL — mouvements de caisse (dépenses prioritaires).
Fichier : caisse_nap.xlsx
"""
from __future__ import annotations

import gc
import glob
import os
import re
import time
from datetime import datetime, date, timedelta
from io import BytesIO

import pandas as pd

FICHIER = "caisse_nap.xlsx"
FEUILLE_MOUV = "MOUVEMENTS"
FEUILLE_CAT = "CATEGORIES"
FEUILLE_TIERS = "TIERS"
FEUILLE_SOLDES = "SOLDES_JOUR"
DOSSIER = "caisse_exports"
os.makedirs(DOSSIER, exist_ok=True)

CATS_DEFAUT = [
    "Transport",
    "Achat",
    "DHL",
    "Certificat",
    "Frais bancaires",
    "Divers",
]

TYPES_TIERS = ["Client", "Fournisseur", "Interne", "Autre"]
TYPES_MOUV = ["DEPENSE", "ENTREE"]
MODES = ["Espèces", "Mobile Money", "Chèque", "Virement", "Autre"]

# Produits métier (filtre + saisie)
PRODUITS = [
    "Graines de soja (Soyabean)",
    "Graines de Cotton",
    "Noix de Karité",
    "Tourteau de Cotton",
    "Tourteau de Soja",
    "Tourteau de Karité",
]
PRODUITS_ALIAS = {
    "SOYABEAN": "Graines de soja (Soyabean)",
    "SOJA": "Graines de soja (Soyabean)",
    "GRAINE DE SOJA": "Graines de soja (Soyabean)",
    "GRAINES DE SOJA": "Graines de soja (Soyabean)",
    "COTTON": "Graines de Cotton",
    "COTON": "Graines de Cotton",
    "GRAINE DE COTON": "Graines de Cotton",
    "GRAINES DE COTON": "Graines de Cotton",
    "GRAINES DE COTTON": "Graines de Cotton",
    "SHEA": "Noix de Karité",
    "KARITE": "Noix de Karité",
    "KARITÉ": "Noix de Karité",
    "NOIX DE KARITE": "Noix de Karité",
    "NOIX DE KARITÉ": "Noix de Karité",
    "TOURTEAU DE COTTON": "Tourteau de Cotton",
    "TOURTEAU DE COTON": "Tourteau de Cotton",
    "TOURTEAU COTTON": "Tourteau de Cotton",
    "TOURTEAU DE SOJA": "Tourteau de Soja",
    "TOURTEAU SOJA": "Tourteau de Soja",
    "TOURTEAU DE KARITE": "Tourteau de Karité",
    "TOURTEAU DE KARITÉ": "Tourteau de Karité",
    "TOURTEAU KARITE": "Tourteau de Karité",
}


def normaliser_produit(texte):
    """Mappe une valeur libre / code vers le libellé PRODUITS, sinon conserve le texte."""
    t = _s(texte)
    if not t:
        return ""
    if t in PRODUITS:
        return t
    u = t.upper().replace("É", "E")
    if u in PRODUITS_ALIAS:
        return PRODUITS_ALIAS[u]
    for k, v in PRODUITS_ALIAS.items():
        if k in u or u in k:
            return v
    for p in PRODUITS:
        if t.casefold() in p.casefold() or p.casefold() in t.casefold():
            return p
    return t


COLS_MOUV = [
    "ID", "Date", "Type", "Categorie", "Libelle",
    "Tiers", "Type_Tiers", "Produit", "Montant", "Mode",
    "Num_Cheque", "Nom_Percepteur",
    "Reference", "Observation", "Utilisateur", "Heure_Creation",
]
COLS_CAT = ["Categorie", "Actif", "Notes"]
COLS_TIERS = ["Nom", "Type_Tiers", "Telephone", "Notes", "Actif"]
COLS_SOLDES = ["Date", "Solde_Depart", "Mode_Saisie", "Notes", "Utilisateur", "Heure_Creation"]


def _s(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none", "nat") else t


def _n(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(str(v).replace(" ", "").replace(",", "."))
    except Exception:
        return 0.0


def ensure_file():
    if os.path.exists(FICHIER):
        # migrer colonnes / feuilles manquantes
        try:
            xl = pd.ExcelFile(FICHIER)
            if FEUILLE_MOUV in xl.sheet_names:
                m = pd.read_excel(FICHIER, sheet_name=FEUILLE_MOUV)
                changed = False
                for col in ("Produit", "Num_Cheque", "Nom_Percepteur"):
                    if col not in m.columns:
                        m[col] = ""
                        changed = True
                if changed:
                    with pd.ExcelFile(FICHIER) as xls:
                        sheets = {s: xls.parse(s) for s in xls.sheet_names}
                    sheets[FEUILLE_MOUV] = m
                    tmp = FICHIER.replace(".xlsx", ".tmp.xlsx")
                    with pd.ExcelWriter(tmp, engine="openpyxl") as w:
                        for s, d in sheets.items():
                            d.to_excel(w, sheet_name=s, index=False)
                    try:
                        os.replace(tmp, FICHIER)
                    except PermissionError:
                        try:
                            os.remove(tmp)
                        except OSError:
                            pass
            xl = pd.ExcelFile(FICHIER)
            if FEUILLE_SOLDES not in xl.sheet_names:
                with pd.ExcelFile(FICHIER) as xls:
                    sheets = {s: xls.parse(s) for s in xls.sheet_names}
                sheets[FEUILLE_SOLDES] = pd.DataFrame(columns=COLS_SOLDES)
                tmp = FICHIER.replace(".xlsx", ".tmp.xlsx")
                with pd.ExcelWriter(tmp, engine="openpyxl") as w:
                    for s, d in sheets.items():
                        d.to_excel(w, sheet_name=s, index=False)
                os.replace(tmp, FICHIER)
        except Exception:
            pass
        return
    with pd.ExcelWriter(FICHIER, engine="openpyxl") as w:
        pd.DataFrame(columns=COLS_MOUV).to_excel(w, sheet_name=FEUILLE_MOUV, index=False)
        pd.DataFrame(
            [{"Categorie": c, "Actif": "OUI", "Notes": ""} for c in CATS_DEFAUT]
        ).to_excel(w, sheet_name=FEUILLE_CAT, index=False)
        pd.DataFrame(columns=COLS_TIERS).to_excel(w, sheet_name=FEUILLE_TIERS, index=False)
        pd.DataFrame(columns=COLS_SOLDES).to_excel(w, sheet_name=FEUILLE_SOLDES, index=False)


def _lire(nom, cols=None):
    ensure_file()
    try:
        df = pd.read_excel(FICHIER, sheet_name=nom)
    except ValueError:
        return pd.DataFrame(columns=cols or [])
    if cols:
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[[c for c in cols if c in df.columns]]
    return df


def _remplacer_avec_retry(tmp, cible, essais=8, pause=0.7):
    """OneDrive / Excel / antivirus verrouillent parfois le fichier un court instant."""
    gc.collect()
    for i in range(essais):
        try:
            os.replace(tmp, cible)
            return
        except PermissionError:
            if i == essais - 1:
                try:
                    os.remove(tmp)
                except OSError:
                    pass
                raise PermissionError(
                    f"Impossible d'enregistrer « {os.path.basename(cible)} » : fichier verrouillé. "
                    "Fermez-le dans Excel, attendez la fin de la synchro OneDrive, puis réessayez."
                )
            time.sleep(pause)


def _ecrire(nom, df):
    ensure_file()
    with pd.ExcelFile(FICHIER) as xls:
        sheets = {s: xls.parse(s) for s in xls.sheet_names}
    sheets[nom] = df
    # garantir les feuilles de base
    for feuille, cols in (
        (FEUILLE_MOUV, COLS_MOUV),
        (FEUILLE_CAT, COLS_CAT),
        (FEUILLE_TIERS, COLS_TIERS),
        (FEUILLE_SOLDES, COLS_SOLDES),
    ):
        if feuille not in sheets:
            sheets[feuille] = pd.DataFrame(columns=cols)
    tmp = FICHIER.replace(".xlsx", ".tmp.xlsx")
    try:
        with pd.ExcelWriter(tmp, engine="openpyxl") as w:
            for s, d in sheets.items():
                d.to_excel(w, sheet_name=s, index=False)
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise
    _remplacer_avec_retry(tmp, FICHIER)


# ---------- Catégories ----------
def charger_categories(actifs_only=False):
    df = _lire(FEUILLE_CAT, COLS_CAT)
    if df.empty:
        return df
    df["Categorie"] = df["Categorie"].map(_s)
    df["Actif"] = df["Actif"].map(lambda x: _s(x).upper() or "OUI")
    df = df[df["Categorie"] != ""]
    if actifs_only:
        df = df[df["Actif"] == "OUI"]
    return df.reset_index(drop=True)


def liste_categories():
    df = charger_categories(actifs_only=True)
    base = list(df["Categorie"]) if not df.empty else list(CATS_DEFAUT)
    # union ordonnée
    seen, out = set(), []
    for c in list(CATS_DEFAUT) + base:
        if c and c not in seen:
            seen.add(c)
            out.append(c)
    return out


def ajouter_categorie(nom, notes=""):
    nom = _s(nom)
    if not nom:
        return False, "Nom de catégorie obligatoire."
    df = charger_categories()
    if not df.empty and nom.casefold() in df["Categorie"].str.casefold().values:
        return False, f"La catégorie « {nom} » existe déjà."
    row = {"Categorie": nom, "Actif": "OUI", "Notes": _s(notes)}
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    _ecrire(FEUILLE_CAT, df[COLS_CAT])
    return True, f"Catégorie « {nom} » créée."


# ---------- Tiers (clients / fournisseurs) ----------
def charger_tiers(type_tiers=None, actifs_only=True):
    df = _lire(FEUILLE_TIERS, COLS_TIERS)
    if df.empty:
        return df
    for c in COLS_TIERS:
        df[c] = df[c].map(_s)
    if actifs_only:
        df = df[df["Actif"].str.upper() != "NON"]
    if type_tiers:
        df = df[df["Type_Tiers"].str.casefold() == str(type_tiers).casefold()]
    return df.sort_values("Nom").reset_index(drop=True)


def liste_tiers(type_tiers=None):
    df = charger_tiers(type_tiers=type_tiers)
    return sorted({_s(x) for x in df.get("Nom", []) if _s(x)}) if not df.empty else []


def ajouter_tiers(nom, type_tiers="Client", telephone="", notes=""):
    nom = _s(nom)
    if not nom:
        return False, "Nom obligatoire."
    type_tiers = _s(type_tiers) or "Autre"
    df = charger_tiers(actifs_only=False)
    if not df.empty:
        same = df[
            (df["Nom"].str.casefold() == nom.casefold())
            & (df["Type_Tiers"].str.casefold() == type_tiers.casefold())
        ]
        if not same.empty:
            return False, f"« {nom} » ({type_tiers}) existe déjà."
    row = {
        "Nom": nom,
        "Type_Tiers": type_tiers,
        "Telephone": _s(telephone),
        "Notes": _s(notes),
        "Actif": "OUI",
    }
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    _ecrire(FEUILLE_TIERS, df[COLS_TIERS])
    return True, f"{type_tiers} « {nom} » enregistré."


# ---------- Mouvements ----------
def charger_mouvements():
    df = _lire(FEUILLE_MOUV, COLS_MOUV)
    if df.empty:
        return df
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)
    df["Montant"] = df["Montant"].map(_n)
    for c in ["Type", "Categorie", "Libelle", "Tiers", "Type_Tiers", "Produit", "Mode",
              "Num_Cheque", "Nom_Percepteur",
              "Reference", "Observation", "Utilisateur", "Heure_Creation"]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].map(_s)
    df["Produit"] = df["Produit"].map(normaliser_produit)
    return df.sort_values(["Date", "ID"], ascending=[False, False]).reset_index(drop=True)


def prochain_id(df=None):
    if df is None:
        df = charger_mouvements()
    return 1 if df.empty else int(df["ID"].max()) + 1


def ajouter_mouvement(data, user=""):
    df = charger_mouvements()
    typ = _s(data.get("Type") or "DEPENSE").upper()
    if typ not in ("DEPENSE", "ENTREE"):
        typ = "DEPENSE"
    montant = abs(_n(data.get("Montant")))
    if montant <= 0:
        return False, "Montant obligatoire et > 0."
    cat = _s(data.get("Categorie"))
    if not cat:
        return False, "Catégorie obligatoire."
    # auto-création catégorie si absente
    cats = [c.casefold() for c in liste_categories()]
    if cat.casefold() not in cats:
        ajouter_categorie(cat)

    tiers = _s(data.get("Tiers"))
    type_tiers = _s(data.get("Type_Tiers") or "Autre")
    if tiers:
        existants = [t.casefold() for t in liste_tiers()]
        # créer le tiers s'il n'existe pas du tout
        all_tiers = charger_tiers(actifs_only=False)
        if all_tiers.empty or not (
            (all_tiers["Nom"].str.casefold() == tiers.casefold()).any()
        ):
            ajouter_tiers(tiers, type_tiers=type_tiers)

    produit = normaliser_produit(data.get("Produit"))
    mode = _s(data.get("Mode") or "Espèces")
    num_chq = _s(data.get("Num_Cheque"))
    nom_perc = _s(data.get("Nom_Percepteur"))
    if mode.casefold().startswith("chèque") or mode.casefold().startswith("cheque"):
        if not num_chq:
            return False, "N° de chèque obligatoire pour un paiement par chèque."
        if not nom_perc:
            return False, "Nom du percepteur obligatoire pour un paiement par chèque."
    row = {
        "ID": prochain_id(df),
        "Date": pd.Timestamp(data.get("Date") or date.today()),
        "Type": typ,
        "Categorie": cat,
        "Libelle": _s(data.get("Libelle")),
        "Tiers": tiers,
        "Type_Tiers": type_tiers,
        "Produit": produit,
        "Montant": montant,
        "Mode": mode,
        "Num_Cheque": num_chq,
        "Nom_Percepteur": nom_perc,
        "Reference": _s(data.get("Reference")) or (f"CHQ-{num_chq}" if num_chq else ""),
        "Observation": _s(data.get("Observation")),
        "Utilisateur": _s(user),
        "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    try:
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        _ecrire(FEUILLE_MOUV, df[COLS_MOUV])
    except PermissionError as e:
        return False, str(e)
    sens = "Dépense" if typ == "DEPENSE" else "Entrée"
    return True, f"{sens} #{row['ID']} enregistrée — {montant:,.0f} FCFA".replace(",", " ")


def supprimer_mouvement(eid):
    df = charger_mouvements()
    if df.empty or int(eid) not in df["ID"].values:
        return False, "Mouvement introuvable."
    df = df[df["ID"] != int(eid)]
    _ecrire(FEUILLE_MOUV, df[COLS_MOUV] if not df.empty else pd.DataFrame(columns=COLS_MOUV))
    return True, f"Mouvement #{eid} supprimé."


def filtrer(df, debut=None, fin=None, categorie=None, tiers=None,
            type_tiers=None, type_mouv=None, mode=None, produit=None):
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame(columns=COLS_MOUV)
    out = df.copy()
    if "Produit" not in out.columns:
        out["Produit"] = ""
    if debut is not None:
        out = out[out["Date"].notna() & (out["Date"].dt.date >= debut)]
    if fin is not None:
        out = out[out["Date"].notna() & (out["Date"].dt.date <= fin)]
    if categorie and categorie != "Toutes":
        out = out[out["Categorie"].str.casefold() == str(categorie).casefold()]
    if tiers and tiers != "Tous":
        out = out[out["Tiers"].str.casefold() == str(tiers).casefold()]
    if type_tiers and type_tiers != "Tous":
        out = out[out["Type_Tiers"].str.casefold() == str(type_tiers).casefold()]
    if type_mouv and type_mouv != "Tous":
        out = out[out["Type"].str.upper() == str(type_mouv).upper()]
    if mode and mode != "Tous":
        out = out[out["Mode"].str.casefold() == str(mode).casefold()]
    if produit and produit not in ("Tous", "Tous les produits", ""):
        cible = normaliser_produit(produit)
        out = out[out["Produit"].map(normaliser_produit).str.casefold() == cible.casefold()]
    return out.reset_index(drop=True)


def liste_produits_utilises():
    """Produits de référence + ceux déjà saisis."""
    df = charger_mouvements()
    seen, out = set(), []
    for p in PRODUITS:
        if p not in seen:
            seen.add(p)
            out.append(p)
    if not df.empty and "Produit" in df.columns:
        for p in df["Produit"].map(normaliser_produit):
            if p and p not in seen:
                seen.add(p)
                out.append(p)
    return out


def indicateurs(df):
    if df is None or df.empty:
        return {
            "nb": 0, "depenses": 0.0, "entrees": 0.0, "solde": 0.0,
            "nb_dep": 0, "nb_ent": 0,
        }
    dep = df[df["Type"].str.upper() == "DEPENSE"]["Montant"].sum()
    ent = df[df["Type"].str.upper() == "ENTREE"]["Montant"].sum()
    return {
        "nb": len(df),
        "depenses": float(dep),
        "entrees": float(ent),
        "solde": float(ent - dep),
        "nb_dep": int((df["Type"].str.upper() == "DEPENSE").sum()),
        "nb_ent": int((df["Type"].str.upper() == "ENTREE").sum()),
    }


def situation_par_tiers(df):
    """Compte de dépenses / entrées par tiers."""
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "Tiers", "Type_Tiers", "Nb", "Total_Depenses", "Total_Entrees", "Solde"
        ])
    rows = []
    for (tiers, tt), g in df.groupby(["Tiers", "Type_Tiers"], dropna=False):
        if not _s(tiers):
            continue
        dep = g[g["Type"].str.upper() == "DEPENSE"]["Montant"].sum()
        ent = g[g["Type"].str.upper() == "ENTREE"]["Montant"].sum()
        rows.append({
            "Tiers": _s(tiers),
            "Type_Tiers": _s(tt) or "—",
            "Nb": len(g),
            "Total_Depenses": float(dep),
            "Total_Entrees": float(ent),
            "Solde": float(ent - dep),
        })
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    return out.sort_values("Total_Depenses", ascending=False).reset_index(drop=True)


def situation_par_categorie(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=["Categorie", "Nb", "Total_Depenses", "Total_Entrees"])
    rows = []
    for cat, g in df.groupby("Categorie"):
        if not _s(cat):
            continue
        rows.append({
            "Categorie": _s(cat),
            "Nb": len(g),
            "Total_Depenses": float(g[g["Type"].str.upper() == "DEPENSE"]["Montant"].sum()),
            "Total_Entrees": float(g[g["Type"].str.upper() == "ENTREE"]["Montant"].sum()),
        })
    out = pd.DataFrame(rows)
    return out.sort_values("Total_Depenses", ascending=False).reset_index(drop=True) if not out.empty else out




# ---------- Soldes journaliers ----------
def charger_soldes():
    df = _lire(FEUILLE_SOLDES, COLS_SOLDES)
    if df.empty:
        return df
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)
    df["Solde_Depart"] = df["Solde_Depart"].map(_n)
    for c in ["Mode_Saisie", "Notes", "Utilisateur", "Heure_Creation"]:
        df[c] = df[c].map(_s)
    return df.sort_values("Date", ascending=False).reset_index(drop=True)


def solde_depart_du_jour(jour=None):
    """Retourne le solde de départ enregistré pour un jour, ou None."""
    jour = jour or date.today()
    if isinstance(jour, datetime):
        jour = jour.date()
    df = charger_soldes()
    if df.empty:
        return None
    m = df["Date"].dt.date == jour
    if not m.any():
        return None
    return float(df.loc[m, "Solde_Depart"].iloc[0])


def mouvements_du_jour(jour=None):
    jour = jour or date.today()
    if isinstance(jour, datetime):
        jour = jour.date()
    df = charger_mouvements()
    if df.empty:
        return df
    return df[df["Date"].notna() & (df["Date"].dt.date == jour)].copy()


def calculer_cloture(jour=None, solde_depart=None):
    """
    Solde clôture = solde départ + entrées du jour − dépenses du jour.
    """
    jour = jour or date.today()
    if isinstance(jour, datetime):
        jour = jour.date()
    if solde_depart is None:
        sd = solde_depart_du_jour(jour)
        solde_depart = float(sd) if sd is not None else 0.0
    mv = mouvements_du_jour(jour)
    kpi = indicateurs(mv)
    cloture = float(solde_depart) + kpi["entrees"] - kpi["depenses"]
    return {
        "date": jour,
        "solde_depart": float(solde_depart),
        "entrees": kpi["entrees"],
        "depenses": kpi["depenses"],
        "nb_entrees": kpi["nb_ent"],
        "nb_depenses": kpi["nb_dep"],
        "solde_cloture": cloture,
        "mouvements": mv,
    }


def cloture_hier(jour=None):
    """Solde de clôture de la veille (pour report automatique)."""
    jour = jour or date.today()
    if isinstance(jour, datetime):
        jour = jour.date()
    from datetime import timedelta
    hier = jour - timedelta(days=1)
    return calculer_cloture(hier)


def enregistrer_solde_depart(jour, montant, mode="MANUEL", notes="", user=""):
    """Enregistre ou met à jour le solde de départ du jour."""
    try:
        jour = jour or date.today()
        if isinstance(jour, datetime):
            jour = jour.date()
        montant = _n(montant)
        df = charger_soldes()
        mode = _s(mode) or "MANUEL"
        row = {
            "Date": pd.Timestamp(jour),
            "Solde_Depart": montant,
            "Mode_Saisie": mode,
            "Notes": _s(notes),
            "Utilisateur": _s(user),
            "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M"),
        }
        if not df.empty:
            m = df["Date"].dt.date == jour
            if m.any():
                for k, v in row.items():
                    df.loc[m, k] = v
                _ecrire(FEUILLE_SOLDES, df[COLS_SOLDES])
                return True, f"Solde de départ du {jour.strftime('%d/%m/%Y')} mis à jour : {montant:,.0f} FCFA".replace(",", " ")
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        _ecrire(FEUILLE_SOLDES, df[COLS_SOLDES])
        return True, f"Solde de départ du {jour.strftime('%d/%m/%Y')} enregistré : {montant:,.0f} FCFA".replace(",", " ")
    except PermissionError as e:
        return False, str(e)




# ---------- Import depuis Suivi de paiement (caisse) ----------
MODES_CAISSE = ("espèces", "especes", "cash", "mobile money", "momo", "wave", "caisse")


def _trouver_suivi_paiement():
    for n in (
        "Suivi de paiement.xlsx",
        "Suivi de paiement-5.xlsx",
        "Suivi_de_paiement.xlsx",
    ):
        if os.path.exists(n):
            return n
    for p in glob.glob("Suivi*aiement*.xlsx") + glob.glob("*paiement*.xlsx"):
        if os.path.isfile(p) and not os.path.basename(p).startswith("~$"):
            return p
    return None


def _charger_suivi_paiement(chemin=None):
    """Charge le journal DEPENSES du Suivi de paiement (colonnes normalisées)."""
    chemin = chemin or _trouver_suivi_paiement()
    if not chemin:
        return None, "Fichier Suivi de paiement.xlsx introuvable à la racine du projet."
    try:
        xl = pd.ExcelFile(chemin)
        feuille = "DEPENSES" if "DEPENSES" in xl.sheet_names else xl.sheet_names[0]
        df = None
        for hdr in (2, 1, 0, 3, 4):
            tmp = pd.read_excel(chemin, sheet_name=feuille, header=hdr)
            low = [str(c).strip().lower() for c in tmp.columns]
            if any("transaction" in c or "montant" in c for c in low):
                df = tmp
                break
        if df is None:
            df = pd.read_excel(chemin, sheet_name=feuille)
    except Exception as e:
        return None, f"Lecture Suivi de paiement : {e}"

    alias = {
        "date": "Date",
        "no de transaction": "No_Transaction",
        "n° de transaction": "No_Transaction",
        "no transaction": "No_Transaction",
        "transaction": "No_Transaction",
        "produit": "Produit",
        "prestataires/fournisseurs": "Prestataire",
        "prestataire/fournisseur": "Prestataire",
        "prestataires": "Prestataire",
        "fournisseur": "Prestataire",
        "nom du bénéficiaire": "Beneficiaire",
        "nom du beneficiaire": "Beneficiaire",
        "bénéficiaire": "Beneficiaire",
        "beneficiaire": "Beneficiaire",
        "montant": "Montant",
        "observation": "Observation",
        "mode paiement": "Mode",
        "mode de paiement": "Mode",
        "numero de cheque": "Num_Cheque",
        "numéro de chèque": "Num_Cheque",
        "banque": "Banque",
        "types opération": "Type_Operation",
        "type opération": "Type_Operation",
        "type operation": "Type_Operation",
        "descriptions/object du paiement": "Description",
        "description": "Description",
        "ref doc": "Ref_Doc",
    }
    rename = {}
    for c in df.columns:
        k = re.sub(r"\s+", " ", str(c).strip().lower())
        if k in alias:
            rename[c] = alias[k]
    df = df.rename(columns=rename)
    for c in ["Date", "No_Transaction", "Produit", "Prestataire", "Beneficiaire",
              "Montant", "Observation", "Mode", "Num_Cheque", "Banque",
              "Type_Operation", "Description", "Ref_Doc"]:
        if c not in df.columns:
            df[c] = "" if c != "Montant" else 0.0
    out = df.copy()
    out["Montant"] = out["Montant"].map(_n)
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce", dayfirst=True)
    for c in ["No_Transaction", "Produit", "Prestataire", "Beneficiaire", "Observation",
              "Mode", "Num_Cheque", "Banque", "Type_Operation", "Description", "Ref_Doc"]:
        out[c] = out[c].map(_s)
    out = out[out["Montant"] > 0].reset_index(drop=True)
    return out, None


def _est_mode_caisse(mode: str) -> bool:
    m = _s(mode).casefold()
    if not m:
        return False
    return any(k in m for k in MODES_CAISSE)


def _ref_import(row) -> str:
    """Clé anti-doublon stable pour un paiement importé."""
    no = _s(row.get("No_Transaction"))
    if no:
        return f"SP-{no}"
    d = row.get("Date")
    try:
        ds = pd.Timestamp(d).strftime("%Y%m%d") if pd.notna(d) else ""
    except Exception:
        ds = ""
    return f"SP-{ds}-{_s(row.get('Prestataire'))[:20]}-{int(_n(row.get('Montant')))}"


def apercu_import_suivi(chemin=None, seulement_caisse=True):
    """
    Retourne (df_candidats, stats, erreur).
    stats: total_suivi, candidats, deja_importes, nouveaux
    """
    journal, err = _charger_suivi_paiement(chemin)
    if err:
        return None, {}, err
    if journal is None or journal.empty:
        return journal, {"total_suivi": 0, "candidats": 0, "deja_importes": 0, "nouveaux": 0}, None

    total = len(journal)
    if seulement_caisse:
        mask = journal["Mode"].map(_est_mode_caisse)
        cand = journal.loc[mask].copy()
    else:
        cand = journal.copy()

    exist = charger_mouvements()
    refs_exist = set()
    if not exist.empty and "Reference" in exist.columns:
        refs_exist = { _s(x) for x in exist["Reference"] if _s(x).startswith("SP-") }

    cand["_Ref_Import"] = cand.apply(_ref_import, axis=1)
    cand["_Deja"] = cand["_Ref_Import"].isin(refs_exist)
    nouveaux = int((~cand["_Deja"]).sum())
    stats = {
        "total_suivi": total,
        "candidats": len(cand),
        "deja_importes": int(cand["_Deja"].sum()),
        "nouveaux": nouveaux,
        "fichier": chemin or _trouver_suivi_paiement(),
    }
    return cand, stats, None


def importer_depuis_suivi_paiement(chemin=None, seulement_caisse=True, user=""):
    """
    Importe les paiements « caisse » du Suivi de paiement en dépenses Caisse NAP.
    - Mode Espèces / Mobile Money / Caisse (si seulement_caisse=True)
    - Anti-doublon via Reference = SP-{N° transaction}
    Retourne (ok, message, nb_ajoutes).
    """
    cand, stats, err = apercu_import_suivi(chemin=chemin, seulement_caisse=seulement_caisse)
    if err:
        return False, err, 0
    if cand is None or cand.empty:
        return True, "Aucun paiement caisse à importer dans le Suivi de paiement.", 0

    a_importer = cand[~cand["_Deja"]].copy()
    if a_importer.empty:
        return True, (
            f"Rien de nouveau : {stats['deja_importes']} paiement(s) déjà en caisse "
            f"sur {stats['candidats']} candidat(s)."
        ), 0

    df = charger_mouvements()
    nid = prochain_id(df)
    rows = []
    for _, r in a_importer.iterrows():
        # Catégorie : Type_Operation sinon Produit sinon Achat
        cat = _s(r.get("Type_Operation")) or _s(r.get("Produit")) or "Achat"
        if len(cat) > 40:
            cat = cat[:40]
        # auto catégorie
        if cat.casefold() not in [c.casefold() for c in liste_categories()]:
            try:
                ajouter_categorie(cat, notes="Import Suivi paiement")
            except Exception:
                pass
        tiers = _s(r.get("Prestataire")) or _s(r.get("Beneficiaire"))
        if tiers:
            try:
                all_t = charger_tiers(actifs_only=False)
                if all_t.empty or not (all_t["Nom"].str.casefold() == tiers.casefold()).any():
                    ajouter_tiers(tiers, type_tiers="Fournisseur")
            except Exception:
                pass
        lib = _s(r.get("Description")) or _s(r.get("Observation")) or _s(r.get("Produit"))
        if not lib:
            lib = f"Paiement {_s(r.get('No_Transaction'))}"
        mode = _s(r.get("Mode")) or "Espèces"
        # normaliser mode vers liste MODES
        ml = mode.casefold()
        if "espèce" in ml or "espece" in ml or "cash" in ml:
            mode = "Espèces"
        elif "mobile" in ml or "momo" in ml or "wave" in ml:
            mode = "Mobile Money"
        elif "chèque" in ml or "cheque" in ml:
            mode = "Chèque"
        elif "virement" in ml:
            mode = "Virement"
        obs_parts = []
        if _s(r.get("Banque")):
            obs_parts.append(f"Banque: {_s(r.get('Banque'))}")
        if _s(r.get("Num_Cheque")):
            obs_parts.append(f"Chq: {_s(r.get('Num_Cheque'))}")
        if _s(r.get("Observation")) and _s(r.get("Observation")) != lib:
            obs_parts.append(_s(r.get("Observation")))
        obs = " · ".join(obs_parts)
        d = r.get("Date")
        if pd.isna(d):
            d = date.today()
        rows.append({
            "ID": nid,
            "Date": pd.Timestamp(d),
            "Type": "DEPENSE",
            "Categorie": cat,
            "Libelle": lib,
            "Tiers": tiers,
            "Type_Tiers": "Fournisseur" if tiers else "",
            "Produit": normaliser_produit(r.get("Produit")),
            "Montant": abs(_n(r.get("Montant"))),
            "Mode": mode,
            "Num_Cheque": _s(r.get("Num_Cheque")),
            "Nom_Percepteur": _s(r.get("Beneficiaire")) or tiers,
            "Reference": _s(r.get("_Ref_Import")),
            "Observation": obs,
            "Utilisateur": _s(user) or "Import SP",
            "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M"),
        })
        nid += 1

    try:
        df = pd.concat([df, pd.DataFrame(rows)], ignore_index=True)
        _ecrire(FEUILLE_MOUV, df[COLS_MOUV])
    except PermissionError as e:
        return False, str(e), 0

    msg = (
        f"{len(rows)} dépense(s) importée(s) depuis Suivi de paiement "
        f"(ignoré déjà présent : {stats['deja_importes']})."
    )
    return True, msg, len(rows)


def to_excel_bytes(df):
    buf = BytesIO()
    out = df.copy()
    if "Date" in out.columns:
        out["Date"] = pd.to_datetime(out["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="Caisse")
    buf.seek(0)
    return buf.getvalue()
