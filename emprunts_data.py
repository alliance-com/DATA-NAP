# -*- coding: utf-8 -*-
"""
NAP SARL — Gestion des emprunts / tirages bancaires & tableaux d'amortissement.
Fichier : emprunts_nap.xlsx
"""
from __future__ import annotations

import gc
import os
import time
from datetime import datetime, date
from io import BytesIO
from typing import Optional

import pandas as pd

FICHIER = "emprunts_nap.xlsx"
FEUILLE_PRETS = "PRETS"
FEUILLE_TIRAGES = "TIRAGES"
FEUILLE_ECHEANCES = "ECHEANCES"
FEUILLE_PAIEMENTS = "PAIEMENTS_EMPRUNT"

BANQUES = [
    "ECOBANK BENIN",
    "ATLANTIC BANK BENIN",
    "BIIC",
    "BGFI",
    "AUTRE",
]
TYPES_AMORT = ["ANNUITE_CONSTANTE", "AMORT_CONSTANT"]  # échéance fixe vs capital constant
PERIODICITES = ["MENSUEL", "TRIMESTRIEL", "SEMESTRIEL", "ANNUEL"]

COLS_PRETS = [
    "ID_Pret", "Banque", "Reference", "Objet", "Montant_Autorise",
    "Taux_Annuel_Pct", "Duree_Mois", "Periodicite", "Type_Amort",
    "Date_Signature", "Date_Debut", "Statut", "Notes", "Utilisateur", "Heure_Creation",
]
COLS_TIRAGES = [
    "ID_Tirage", "ID_Pret", "Date_Tirage", "Montant", "Reference_Banque",
    "Notes", "Utilisateur", "Heure_Creation",
]
COLS_ECHEANCES = [
    "ID_Echeance", "ID_Pret", "N_Echeance", "Date_Echeance",
    "Capital_Debut", "Amortissement", "Interets", "Assurance", "Echeance_Totale",
    "Capital_Fin", "Statut", "Montant_Paye", "Date_Paiement", "Notes",
]
COLS_PAIEMENTS = [
    "ID_Paiement", "ID_Pret", "ID_Echeance", "Date_Paiement", "Montant",
    "Mode", "Reference", "Notes", "Utilisateur", "Heure_Creation",
]


def _s(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none", "nat") else t


def _n(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(str(v).replace(" ", "").replace(",", ".").replace("\u00a0", ""))
    except Exception:
        return 0.0


def ensure_file():
    if os.path.exists(FICHIER):
        return
    with pd.ExcelWriter(FICHIER, engine="openpyxl") as w:
        pd.DataFrame(columns=COLS_PRETS).to_excel(w, sheet_name=FEUILLE_PRETS, index=False)
        pd.DataFrame(columns=COLS_TIRAGES).to_excel(w, sheet_name=FEUILLE_TIRAGES, index=False)
        pd.DataFrame(columns=COLS_ECHEANCES).to_excel(w, sheet_name=FEUILLE_ECHEANCES, index=False)
        pd.DataFrame(columns=COLS_PAIEMENTS).to_excel(w, sheet_name=FEUILLE_PAIEMENTS, index=False)


def _lire(nom, cols=None):
    ensure_file()
    try:
        df = pd.read_excel(FICHIER, sheet_name=nom)
    except ValueError:
        return pd.DataFrame(columns=cols or [])
    if cols:
        for c in cols:
            if c not in df.columns:
                df[c] = "" if c not in ("Montant", "Montant_Autorise", "Montant_Paye") else 0.0
        df = df[[c for c in cols if c in df.columns]]
    return df


def _remplacer(tmp, cible, essais=8, pause=0.6):
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
                    f"Fichier verrouillé : {os.path.basename(cible)}. Fermez Excel / attendez OneDrive."
                )
            time.sleep(pause)


def _ecrire(nom, df):
    ensure_file()
    with pd.ExcelFile(FICHIER) as xls:
        sheets = {s: xls.parse(s) for s in xls.sheet_names}
    sheets[nom] = df
    for f, c in (
        (FEUILLE_PRETS, COLS_PRETS),
        (FEUILLE_TIRAGES, COLS_TIRAGES),
        (FEUILLE_ECHEANCES, COLS_ECHEANCES),
        (FEUILLE_PAIEMENTS, COLS_PAIEMENTS),
    ):
        if f not in sheets:
            sheets[f] = pd.DataFrame(columns=c)
    tmp = FICHIER.replace(".xlsx", ".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as w:
        for s, d in sheets.items():
            d.to_excel(w, sheet_name=s, index=False)
    _remplacer(tmp, FICHIER)


# ---------- Prêts ----------
def charger_prets():
    df = _lire(FEUILLE_PRETS, COLS_PRETS)
    if df.empty:
        return df
    df["ID_Pret"] = pd.to_numeric(df["ID_Pret"], errors="coerce").fillna(0).astype(int)
    for c in ["Montant_Autorise", "Taux_Annuel_Pct", "Duree_Mois"]:
        df[c] = df[c].map(_n)
    df["Date_Signature"] = pd.to_datetime(df["Date_Signature"], errors="coerce", dayfirst=True)
    df["Date_Debut"] = pd.to_datetime(df["Date_Debut"], errors="coerce", dayfirst=True)
    for c in ["Banque", "Reference", "Objet", "Periodicite", "Type_Amort", "Statut", "Notes", "Utilisateur", "Heure_Creation"]:
        df[c] = df[c].map(_s)
    return df.sort_values("ID_Pret", ascending=False).reset_index(drop=True)


def prochain_id_pret(df=None):
    df = df if df is not None else charger_prets()
    return 1 if df.empty else int(df["ID_Pret"].max()) + 1


def ajouter_pret(data, user=""):
    df = charger_prets()
    montant = abs(_n(data.get("Montant_Autorise")))
    if montant <= 0:
        return False, "Montant autorisé obligatoire."
    banque = _s(data.get("Banque"))
    if not banque:
        return False, "Banque obligatoire."
    taux = _n(data.get("Taux_Annuel_Pct"))
    duree = int(_n(data.get("Duree_Mois")) or 0)
    if duree <= 0:
        return False, "Durée (mois) obligatoire."
    typ = _s(data.get("Type_Amort")) or "ANNUITE_CONSTANTE"
    if typ not in TYPES_AMORT:
        typ = "ANNUITE_CONSTANTE"
    per = _s(data.get("Periodicite")) or "MENSUEL"
    if per not in PERIODICITES:
        per = "MENSUEL"
    row = {
        "ID_Pret": prochain_id_pret(df),
        "Banque": banque,
        "Reference": _s(data.get("Reference")),
        "Objet": _s(data.get("Objet")),
        "Montant_Autorise": montant,
        "Taux_Annuel_Pct": taux,
        "Duree_Mois": duree,
        "Periodicite": per,
        "Type_Amort": typ,
        "Date_Signature": pd.Timestamp(data.get("Date_Signature") or date.today()),
        "Date_Debut": pd.Timestamp(data.get("Date_Debut") or date.today()),
        "Statut": _s(data.get("Statut") or "ACTIF"),
        "Notes": _s(data.get("Notes")),
        "Utilisateur": _s(user),
        "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    try:
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        _ecrire(FEUILLE_PRETS, df[COLS_PRETS])
    except PermissionError as e:
        return False, str(e)
    return True, f"Prêt #{row['ID_Pret']} créé — {montant:,.0f} FCFA ({banque})".replace(",", " ")


def pret_par_id(pid):
    df = charger_prets()
    if df.empty:
        return None
    m = df[df["ID_Pret"] == int(pid)]
    return None if m.empty else m.iloc[0]


# ---------- Tirages ----------
def charger_tirages(id_pret=None):
    df = _lire(FEUILLE_TIRAGES, COLS_TIRAGES)
    if df.empty:
        return df
    df["ID_Tirage"] = pd.to_numeric(df["ID_Tirage"], errors="coerce").fillna(0).astype(int)
    df["ID_Pret"] = pd.to_numeric(df["ID_Pret"], errors="coerce").fillna(0).astype(int)
    df["Montant"] = df["Montant"].map(_n)
    df["Date_Tirage"] = pd.to_datetime(df["Date_Tirage"], errors="coerce", dayfirst=True)
    for c in ["Reference_Banque", "Notes", "Utilisateur", "Heure_Creation"]:
        df[c] = df[c].map(_s)
    if id_pret is not None:
        df = df[df["ID_Pret"] == int(id_pret)]
    return df.sort_values(["Date_Tirage", "ID_Tirage"], ascending=[False, False]).reset_index(drop=True)


def total_tirages(id_pret):
    t = charger_tirages(id_pret)
    return float(t["Montant"].sum()) if not t.empty else 0.0


def ajouter_tirage(data, user=""):
    pid = int(data.get("ID_Pret") or 0)
    pret = pret_par_id(pid)
    if pret is None:
        return False, "Prêt introuvable."
    montant = abs(_n(data.get("Montant")))
    if montant <= 0:
        return False, "Montant du tirage obligatoire."
    autorise = float(pret["Montant_Autorise"])
    deja = total_tirages(pid)
    if deja + montant > autorise + 0.5:
        return False, (
            f"Tirage impossible : autorisé {autorise:,.0f}, déjà tiré {deja:,.0f}, "
            f"reste {autorise - deja:,.0f} FCFA."
        ).replace(",", " ")
    df = charger_tirages()
    nid = 1 if df.empty else int(df["ID_Tirage"].max()) + 1
    row = {
        "ID_Tirage": nid,
        "ID_Pret": pid,
        "Date_Tirage": pd.Timestamp(data.get("Date_Tirage") or date.today()),
        "Montant": montant,
        "Reference_Banque": _s(data.get("Reference_Banque")),
        "Notes": _s(data.get("Notes")),
        "Utilisateur": _s(user),
        "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    try:
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        _ecrire(FEUILLE_TIRAGES, df[COLS_TIRAGES])
    except PermissionError as e:
        return False, str(e)
    reste = autorise - deja - montant
    return True, (
        f"Tirage #{nid} enregistré — {montant:,.0f} FCFA "
        f"(reste à tirer : {reste:,.0f} FCFA)."
    ).replace(",", " ")


# ---------- Calcul amortissement ----------
def _nb_periodes(duree_mois, periodicite):
    p = str(periodicite or "MENSUEL").upper()
    if p == "MENSUEL":
        return int(duree_mois), 12
    if p == "TRIMESTRIEL":
        return max(1, int(duree_mois) // 3), 4
    if p == "SEMESTRIEL":
        return max(1, int(duree_mois) // 6), 2
    if p == "ANNUEL":
        return max(1, int(duree_mois) // 12), 1
    return int(duree_mois), 12


def _add_periodes(d0, n, periodicite):
    """Ajoute n périodes à une date."""
    p = str(periodicite or "MENSUEL").upper()
    d = pd.Timestamp(d0)
    if p == "MENSUEL":
        return d + pd.DateOffset(months=n)
    if p == "TRIMESTRIEL":
        return d + pd.DateOffset(months=3 * n)
    if p == "SEMESTRIEL":
        return d + pd.DateOffset(months=6 * n)
    if p == "ANNUEL":
        return d + pd.DateOffset(years=n)
    return d + pd.DateOffset(months=n)


def calculer_tableau(capital, taux_annuel_pct, duree_mois, periodicite="MENSUEL",
                     type_amort="ANNUITE_CONSTANTE", date_debut=None, assurance_par_echeance=0.0):
    """
    Retourne DataFrame colonnes:
    N_Echeance, Date_Echeance, Capital_Debut, Amortissement, Interets, Assurance,
    Echeance_Totale, Capital_Fin
    """
    capital = float(capital)
    if capital <= 0:
        return pd.DataFrame()
    n, freq = _nb_periodes(duree_mois, periodicite)
    i = (float(taux_annuel_pct) / 100.0) / freq  # taux période
    date_debut = pd.Timestamp(date_debut or date.today())
    ass = float(assurance_par_echeance or 0)

    rows = []
    restant = capital

    if str(type_amort).upper() == "AMORT_CONSTANT":
        amort_fixe = capital / n
        for k in range(1, n + 1):
            interet = restant * i
            amort = amort_fixe if k < n else restant  # dernière = solde
            total = amort + interet + ass
            fin = max(0.0, restant - amort)
            rows.append({
                "N_Echeance": k,
                "Date_Echeance": _add_periodes(date_debut, k, periodicite),
                "Capital_Debut": round(restant, 0),
                "Amortissement": round(amort, 0),
                "Interets": round(interet, 0),
                "Assurance": round(ass, 0),
                "Echeance_Totale": round(total, 0),
                "Capital_Fin": round(fin, 0),
            })
            restant = fin
    else:
        # Annuité constante : A = C * i / (1 - (1+i)^-n)
        if i == 0:
            annuite = capital / n
        else:
            annuite = capital * (i / (1 - (1 + i) ** (-n)))
        for k in range(1, n + 1):
            interet = restant * i
            amort = annuite - interet
            if k == n:
                amort = restant
                annuite_k = amort + interet
            else:
                annuite_k = annuite
            total = annuite_k + ass
            fin = max(0.0, restant - amort)
            rows.append({
                "N_Echeance": k,
                "Date_Echeance": _add_periodes(date_debut, k, periodicite),
                "Capital_Debut": round(restant, 0),
                "Amortissement": round(amort, 0),
                "Interets": round(interet, 0),
                "Assurance": round(ass, 0),
                "Echeance_Totale": round(total, 0),
                "Capital_Fin": round(fin, 0),
            })
            restant = fin

    return pd.DataFrame(rows)


# fix typo restant


def charger_echeances(id_pret=None):
    df = _lire(FEUILLE_ECHEANCES, COLS_ECHEANCES)
    if df.empty:
        return df
    df["ID_Echeance"] = pd.to_numeric(df["ID_Echeance"], errors="coerce").fillna(0).astype(int)
    df["ID_Pret"] = pd.to_numeric(df["ID_Pret"], errors="coerce").fillna(0).astype(int)
    df["N_Echeance"] = pd.to_numeric(df["N_Echeance"], errors="coerce").fillna(0).astype(int)
    for c in ["Capital_Debut", "Amortissement", "Interets", "Assurance", "Echeance_Totale", "Capital_Fin", "Montant_Paye"]:
        df[c] = df[c].map(_n)
    df["Date_Echeance"] = pd.to_datetime(df["Date_Echeance"], errors="coerce", dayfirst=True)
    df["Date_Paiement"] = pd.to_datetime(df["Date_Paiement"], errors="coerce", dayfirst=True)
    for c in ["Statut", "Notes"]:
        df[c] = df[c].map(_s)
    if id_pret is not None:
        df = df[df["ID_Pret"] == int(id_pret)]
    return df.sort_values(["ID_Pret", "N_Echeance"]).reset_index(drop=True)


def generer_et_sauver_tableau(id_pret, capital=None, assurance=0.0, ecraser=True, user=""):
    """Génère le tableau d'amortissement pour un prêt et l'enregistre.
    capital : si None, utilise la somme des tirages (ou montant autorisé si aucun tirage).
    """
    pret = pret_par_id(id_pret)
    if pret is None:
        return False, "Prêt introuvable.", None
    tir = total_tirages(id_pret)
    if capital is None:
        capital = tir if tir > 0 else float(pret["Montant_Autorise"])
    tab = calculer_tableau(
        capital=capital,
        taux_annuel_pct=float(pret["Taux_Annuel_Pct"]),
        duree_mois=int(pret["Duree_Mois"]),
        periodicite=_s(pret["Periodicite"]) or "MENSUEL",
        type_amort=_s(pret["Type_Amort"]) or "ANNUITE_CONSTANTE",
        date_debut=pret["Date_Debut"] if pd.notna(pret["Date_Debut"]) else date.today(),
        assurance_par_echeance=assurance,
    )
    if tab.empty:
        return False, "Impossible de calculer le tableau (capital nul ?).", None

    exist = charger_echeances()
    if ecraser and not exist.empty:
        exist = exist[exist["ID_Pret"] != int(id_pret)]
    start_id = 1 if exist.empty else int(exist["ID_Echeance"].max()) + 1
    rows = []
    for i, r in tab.iterrows():
        rows.append({
            "ID_Echeance": start_id + i,
            "ID_Pret": int(id_pret),
            "N_Echeance": int(r["N_Echeance"]),
            "Date_Echeance": r["Date_Echeance"],
            "Capital_Debut": r["Capital_Debut"],
            "Amortissement": r["Amortissement"],
            "Interets": r["Interets"],
            "Assurance": r["Assurance"],
            "Echeance_Totale": r["Echeance_Totale"],
            "Capital_Fin": r["Capital_Fin"],
            "Statut": "A_PAYER",
            "Montant_Paye": 0.0,
            "Date_Paiement": pd.NaT,
            "Notes": "",
        })
    new_df = pd.DataFrame(rows)
    out = pd.concat([exist, new_df], ignore_index=True) if not exist.empty else new_df
    try:
        _ecrire(FEUILLE_ECHEANCES, out[COLS_ECHEANCES])
    except PermissionError as e:
        return False, str(e), None
    return True, f"Tableau généré : {len(rows)} échéance(s) pour prêt #{id_pret} (capital {capital:,.0f} FCFA).".replace(",", " "), new_df


def marquer_echeance_payee(id_echeance, montant=None, date_paiement=None, notes=""):
    df = charger_echeances()
    if df.empty or int(id_echeance) not in df["ID_Echeance"].values:
        return False, "Échéance introuvable."
    m = df["ID_Echeance"] == int(id_echeance)
    total = float(df.loc[m, "Echeance_Totale"].iloc[0])
    paye = abs(_n(montant)) if montant is not None else total
    df.loc[m, "Montant_Paye"] = paye
    df.loc[m, "Date_Paiement"] = pd.Timestamp(date_paiement or date.today())
    if paye + 0.5 >= total:
        df.loc[m, "Statut"] = "PAYE"
    elif paye > 0:
        df.loc[m, "Statut"] = "PARTIEL"
    else:
        df.loc[m, "Statut"] = "A_PAYER"
    if notes:
        df.loc[m, "Notes"] = _s(notes)
    try:
        _ecrire(FEUILLE_ECHEANCES, df[COLS_ECHEANCES])
    except PermissionError as e:
        return False, str(e)
    return True, f"Échéance #{id_echeance} mise à jour ({df.loc[m, 'Statut'].iloc[0]})."


def resume_pret(id_pret):
    pret = pret_par_id(id_pret)
    if pret is None:
        return {}
    tir = total_tirages(id_pret)
    ech = charger_echeances(id_pret)
    paye = float(ech["Montant_Paye"].sum()) if not ech.empty else 0.0
    du = float(ech["Echeance_Totale"].sum()) if not ech.empty else 0.0
    reste_ech = max(0.0, du - paye)
    capital_restant = float(ech.iloc[-1]["Capital_Fin"]) if not ech.empty else (tir or float(pret["Montant_Autorise"]))
    # si des échéances payées, capital restant = capital fin de la dernière non totalement soldée
    if not ech.empty:
        non_payees = ech[ech["Statut"] != "PAYE"]
        if not non_payees.empty:
            capital_restant = float(non_payees.iloc[0]["Capital_Debut"])
        else:
            capital_restant = 0.0
    return {
        "pret": pret,
        "autorise": float(pret["Montant_Autorise"]),
        "tire": tir,
        "reste_tirer": float(pret["Montant_Autorise"]) - tir,
        "total_echeances": du,
        "total_paye": paye,
        "reste_a_payer": reste_ech,
        "capital_restant": capital_restant,
        "nb_echeances": len(ech),
        "nb_payees": int((ech["Statut"] == "PAYE").sum()) if not ech.empty else 0,
    }


def to_excel_bytes(df):
    buf = BytesIO()
    out = df.copy()
    for c in out.columns:
        if "Date" in c:
            out[c] = pd.to_datetime(out[c], errors="coerce").dt.strftime("%d/%m/%Y")
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="Amortissement")
    buf.seek(0)
    return buf.getvalue()
