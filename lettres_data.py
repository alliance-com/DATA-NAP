# -*- coding: utf-8 -*-
"""Registre des lettres sortantes NAP SARL."""
from __future__ import annotations

import os
from datetime import datetime, date
from io import BytesIO

import pandas as pd

FICHIER = "lettres_bancaires.xlsx"
FEUILLE = "LETTRES"

TYPES = [
    "TIRAGE_CFA",
    "PAIEMENT_FOURNISSEUR",
    "EXTENSION_LC",
    "ACCEPTATION_ECART",
    "REPONSE_JUSTIFICATIFS",
    "LETTRE_GENERALE",
]

BANQUES = [
    "ECOBANK BENIN",
    "Banque Internationale pour l'Industrie et le Commerce (BIIC)",
    "UBA BENIN",
    "BGFIBank BENIN",
    "Orabank BENIN",
    "Autre",
]

COMPTES = [
    "112241067001",
    "20720970166",
]

DEVISES = ["XOF", "USD", "EUR"]

COLS = [
    "ID",
    "Reference",
    "Date_Lettre",
    "Type",
    "Banque",
    "Objet",
    "Montant",
    "Devise",
    "Compte_Credit",
    "Beneficiaire",
    "Ref_LC",
    "Ref_Facture",
    "Ref_Transaction",
    "Motif",
    "Corps_Libre",
    "Signataire",
    "Statut",
    "Fichier_Word",
    "Cree_Par",
    "Heure_Creation",
]


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


def charger():
    if not os.path.exists(FICHIER):
        return pd.DataFrame(columns=COLS)
    try:
        df = pd.read_excel(FICHIER, sheet_name=FEUILLE)
    except Exception:
        try:
            df = pd.read_excel(FICHIER)
        except Exception:
            return pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in df.columns:
            df[c] = 0 if c in ("ID", "Montant") else ""
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
    df["Montant"] = df["Montant"].map(_n)
    df["Date_Lettre"] = pd.to_datetime(df["Date_Lettre"], errors="coerce", dayfirst=True)
    for c in COLS:
        if c not in ("ID", "Montant", "Date_Lettre"):
            df[c] = df[c].map(_s)
    return df.sort_values(["Date_Lettre", "ID"], ascending=[False, False]).reset_index(drop=True)


def sauver(df):
    out = df.copy()
    for c in COLS:
        if c not in out.columns:
            out[c] = ""
    out = out[COLS]
    tmp = FICHIER.replace(".xlsx", ".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as w:
        out.to_excel(w, sheet_name=FEUILLE, index=False)
    os.replace(tmp, FICHIER)


def prochain_id(df=None):
    if df is None:
        df = charger()
    if df.empty:
        return 1
    return int(df["ID"].max()) + 1


def prochaine_ref(annee=None):
    df = charger()
    annee = annee or date.today().year
    prefix = f"REF-{annee}-"
    nums = []
    if not df.empty:
        for r in df["Reference"].astype(str):
            if r.startswith("REF-") or r.upper().startswith("REF"):
                # formats REF-036-2026 or REF -036-2026
                parts = r.replace(" ", "").replace("REF", "").strip("-").split("-")
                for p in parts:
                    if p.isdigit() and len(p) <= 4:
                        nums.append(int(p))
                        break
    n = max(nums) + 1 if nums else 1
    return f"REF-{n:03d}-{annee}"


def ajouter(data, user=""):
    df = charger()
    eid = prochain_id(df)
    ref = _s(data.get("Reference")) or prochaine_ref()
    if not df.empty and ref in df["Reference"].astype(str).values:
        return False, f"La référence {ref} existe déjà."
    row = {
        "ID": eid,
        "Reference": ref,
        "Date_Lettre": pd.Timestamp(data.get("Date_Lettre") or date.today()),
        "Type": _s(data.get("Type")),
        "Banque": _s(data.get("Banque")),
        "Objet": _s(data.get("Objet")),
        "Montant": _n(data.get("Montant")),
        "Devise": _s(data.get("Devise") or "XOF"),
        "Compte_Credit": _s(data.get("Compte_Credit")),
        "Beneficiaire": _s(data.get("Beneficiaire")),
        "Ref_LC": _s(data.get("Ref_LC")),
        "Ref_Facture": _s(data.get("Ref_Facture")),
        "Ref_Transaction": _s(data.get("Ref_Transaction")),
        "Motif": _s(data.get("Motif")),
        "Corps_Libre": _s(data.get("Corps_Libre")),
        "Signataire": _s(data.get("Signataire") or "AUTHORIZED SIGNATORY"),
        "Statut": _s(data.get("Statut") or "BROUILLON"),
        "Fichier_Word": "",
        "Cree_Par": _s(user),
        "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    sauver(df)
    return True, f"Lettre {ref} enregistrée (#{eid}).", eid, ref


def maj_fichier(eid, chemin):
    df = charger()
    m = df["ID"] == int(eid)
    if not m.any():
        return
    df.loc[m, "Fichier_Word"] = chemin
    df.loc[m, "Statut"] = "GENERE"
    sauver(df)


def to_excel_bytes(df):
    buf = BytesIO()
    out = df.copy()
    if "Date_Lettre" in out.columns:
        out["Date_Lettre"] = pd.to_datetime(out["Date_Lettre"], errors="coerce").dt.strftime("%d/%m/%Y")
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="LETTRES")
    buf.seek(0)
    return buf.getvalue()
