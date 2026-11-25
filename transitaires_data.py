# -*- coding: utf-8 -*-
"""
NAP SARL — Couche de données TRANSITAIRES
Fichier Excel local : transitaires.xlsx
À placer à la racine du projet (à côté de app.py).
"""
import os
import re
from datetime import datetime

import pandas as pd

VERSION_TRS = 1

FICHIER_TRS = "transitaires.xlsx"
COLS_TRS = ["Code", "Nom", "Telephone", "Email", "Adresse", "RCCM", "IFU", "Actif", "Cree_Par", "Date_Creation"]


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


def chemin_fichier_trs():
    """Cherche transitaires.xlsx dans le dossier courant puis à côté de ce module."""
    if os.path.exists(FICHIER_TRS):
        return FICHIER_TRS
    alt = os.path.join(os.path.dirname(os.path.abspath(__file__)), FICHIER_TRS)
    if os.path.exists(alt):
        return alt
    return FICHIER_TRS  # sera créé ici


def charger_transitaires(actifs_seulement=False):
    chemin = chemin_fichier_trs()
    if not os.path.exists(chemin):
        df = pd.DataFrame(columns=COLS_TRS)
        df.to_excel(chemin, index=False)
        return df
    df = pd.read_excel(chemin)
    for c in COLS_TRS:
        if c not in df.columns:
            df[c] = ""
    df = df[COLS_TRS].copy()
    for c in COLS_TRS:
        df[c] = df[c].map(_s)
    if actifs_seulement:
        # Accepte OUI / Oui / oui / 1 / True
        actif = df["Actif"].astype(str).str.strip().str.upper()
        df = df[actif.isin(["OUI", "1", "TRUE", "YES", "Y", "O"])]
    return df.reset_index(drop=True)


def prochain_code(df=None):
    if df is None:
        df = charger_transitaires()
    nums = []
    for c in df["Code"].dropna():
        m = re.search(r"(\d+)", str(c))
        if m:
            nums.append(int(m.group(1)))
    n = max(nums) + 1 if nums else 1
    return f"TRS-{n:03d}"


def ajouter_transitaire(nom, telephone="", email="", adresse="", rccm="", ifu="", utilisateur=""):
    df = charger_transitaires()
    nom = _s(nom)
    if not nom:
        return False, "Le nom est obligatoire.", None
    # doublon exact
    if any(df["Nom"].str.casefold() == nom.casefold()):
        exist = df[df["Nom"].str.casefold() == nom.casefold()].iloc[0]
        return False, f"Ce transitaire existe déjà ({exist['Code']}).", exist["Code"]
    code = prochain_code(df)
    ligne = {
        "Code": code,
        "Nom": nom,
        "Telephone": _s(telephone),
        "Email": _s(email),
        "Adresse": _s(adresse),
        "RCCM": _s(rccm),
        "IFU": _s(ifu),
        "Actif": "OUI",
        "Cree_Par": _s(utilisateur),
        "Date_Creation": datetime.now().strftime("%d/%m/%Y"),
    }
    df = pd.concat([df, pd.DataFrame([ligne])], ignore_index=True)
    df.to_excel(chemin_fichier_trs(), index=False)
    return True, f"Transitaire {code} — {nom} créé.", code


def modifier_transitaire(code, champs):
    df = charger_transitaires()
    if code not in set(df["Code"]):
        return False, "Code introuvable."
    idx = df[df["Code"] == code].index[0]
    for k in ["Nom", "Telephone", "Email", "Adresse", "RCCM", "IFU", "Actif"]:
        if k in champs:
            df.at[idx, k] = _s(champs[k])
    df.to_excel(chemin_fichier_trs(), index=False)
    return True, f"Transitaire {code} modifié."


def options_selection(actifs_seulement=True):
    """{label: (code, nom)} pour selectbox."""
    df = charger_transitaires(actifs_seulement=actifs_seulement)
    if df.empty and actifs_seulement:
        # repli : tous les transitaires si aucun n'est marqué actif
        df = charger_transitaires(actifs_seulement=False)
    out = {}
    for _, r in df.iterrows():
        if not _s(r["Nom"]):
            continue
        label = f"{r['Code']} — {r['Nom']}"
        out[label] = (r["Code"], r["Nom"])
    return out


def label_trs(code, nom):
    return nom if code == nom else f"{code} — {nom}"
