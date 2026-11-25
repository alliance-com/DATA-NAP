# -*- coding: utf-8 -*-
"""
NAP SARL — Clients (Local / International)
Fichier : clients.xlsx (racine projet)
"""
from __future__ import annotations

import os
from datetime import datetime
from io import BytesIO

import pandas as pd

FICHIER = "clients.xlsx"
FEUILLE = "Clients"
VERSION_CLIENTS = 1

TYPES = ["LOCAL", "INTERNATIONAL"]

COLS = [
    "Code", "Nom", "Type", "Pays", "Ville", "Adresse",
    "Contact", "Telephone", "Email", "NIF", "Notes",
    "Actif", "Cree_Par", "Date_Creation",
]


def _s(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none", "nat") else t


def ensure_file():
    if os.path.exists(FICHIER):
        return
    with pd.ExcelWriter(FICHIER, engine="openpyxl") as w:
        pd.DataFrame(columns=COLS).to_excel(w, sheet_name=FEUILLE, index=False)


def charger(actifs_only=False):
    ensure_file()
    try:
        df = pd.read_excel(FICHIER, sheet_name=FEUILLE)
    except Exception:
        return pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in df.columns:
            df[c] = ""
    for c in COLS:
        df[c] = df[c].map(_s)
    df["Type"] = df["Type"].map(lambda x: "INTERNATIONAL" if _s(x).upper().startswith("I") else "LOCAL")
    df["Actif"] = df["Actif"].map(lambda x: "OUI" if _s(x).upper() in ("OUI", "YES", "1", "TRUE", "") else "NON")
    if actifs_only:
        df = df[df["Actif"] == "OUI"]
    return df[df["Nom"] != ""].sort_values(["Type", "Nom"]).reset_index(drop=True)


def sauver(df):
    ensure_file()
    out = df.copy()
    for c in COLS:
        if c not in out.columns:
            out[c] = ""
    out = out[COLS]
    tmp = FICHIER.replace(".xlsx", ".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as w:
        out.to_excel(w, sheet_name=FEUILLE, index=False)
    os.replace(tmp, FICHIER)


def prochain_code(type_cli="LOCAL"):
    df = charger()
    prefix = "CLI-L" if type_cli == "LOCAL" else "CLI-I"
    nums = []
    for c in df.get("Code", pd.Series(dtype=str)).astype(str):
        if c.upper().startswith(prefix):
            tail = c[len(prefix):].lstrip("-")
            if tail.isdigit():
                nums.append(int(tail))
    n = max(nums) + 1 if nums else 1
    return f"{prefix}-{n:03d}"


def ajouter(data, user=""):
    df = charger()
    nom = _s(data.get("Nom"))
    if not nom:
        return False, "Nom client obligatoire."
    if not df.empty and nom.casefold() in set(df["Nom"].str.casefold()):
        return False, f"Le client « {nom} » existe déjà."
    typ = "INTERNATIONAL" if _s(data.get("Type")).upper().startswith("I") else "LOCAL"
    code = _s(data.get("Code")) or prochain_code(typ)
    if not df.empty and code in df["Code"].values:
        return False, f"Code {code} déjà utilisé."
    row = {
        "Code": code,
        "Nom": nom,
        "Type": typ,
        "Pays": _s(data.get("Pays")),
        "Ville": _s(data.get("Ville")),
        "Adresse": _s(data.get("Adresse")),
        "Contact": _s(data.get("Contact")),
        "Telephone": _s(data.get("Telephone")),
        "Email": _s(data.get("Email")),
        "NIF": _s(data.get("NIF")),
        "Notes": _s(data.get("Notes")),
        "Actif": "OUI",
        "Cree_Par": _s(user),
        "Date_Creation": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    sauver(df)
    return True, f"Client {code} — {nom} ({typ}) ajouté."


def maj(code, data):
    df = charger()
    m = df["Code"] == _s(code)
    if not m.any():
        return False, "Client introuvable."
    for cle in ["Nom", "Type", "Pays", "Ville", "Adresse", "Contact", "Telephone", "Email", "NIF", "Notes", "Actif"]:
        if cle in data:
            val = _s(data[cle])
            if cle == "Type":
                val = "INTERNATIONAL" if val.upper().startswith("I") else "LOCAL"
            if cle == "Actif":
                val = "OUI" if val.upper() in ("OUI", "YES", "1", "TRUE") else "NON"
            df.loc[m, cle] = val
    sauver(df)
    return True, f"Client {code} mis à jour."


def options_selection(type_filtre=None):
    """dict label -> (code, nom)"""
    df = charger(actifs_only=True)
    if type_filtre in TYPES:
        df = df[df["Type"] == type_filtre]
    out = {}
    for _, r in df.iterrows():
        lab = f"{r['Code']} — {r['Nom']}"
        if r["Type"] == "INTERNATIONAL" and r["Pays"]:
            lab += f" ({r['Pays']})"
        out[lab] = (r["Code"], r["Nom"])
    return out


def liste_noms(type_filtre=None):
    df = charger(actifs_only=True)
    if type_filtre in TYPES:
        df = df[df["Type"] == type_filtre]
    return sorted({_s(x) for x in df["Nom"] if _s(x)})


def to_excel_bytes(df=None):
    if df is None:
        df = charger()
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Clients")
    buf.seek(0)
    return buf.getvalue()
