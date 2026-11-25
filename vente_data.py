# -*- coding: utf-8 -*-
"""
NAP SARL — Ventes (export)
Fichier : ventes_export.xlsx (racine projet)
Lié aux contrats / BL / Commercial Invoices.
"""
from __future__ import annotations

import os
from datetime import datetime, date
from io import BytesIO

import pandas as pd

_BASE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_BASE)
FICHIER = os.path.join(_ROOT, "ventes_export.xlsx")
if not os.path.isdir(_ROOT):
    FICHIER = "ventes_export.xlsx"

FEUILLE = "VENTES"
VERSION = 1

# Produits vendus (export)
PRODUITS_VENTE = [
    "Tourteau de Karité",
    "Tourteau de Soja",
    "Tourteau de Cotton",
    "Graine de Soja",
    "Huile de Soja",
    "Huile de Cotton",
    "Beurre Karité",
]

# Mapping éventuel vers codes logistiques internes
PRODUIT_CODE = {
    "Tourteau de Karité": "SHEA_MEAL",
    "Tourteau de Soja": "SOYA_MEAL",
    "Tourteau de Cotton": "COTTON_MEAL",
    "Graine de Soja": "SOYABEAN",
    "Huile de Soja": "SOYA_OIL",
    "Huile de Cotton": "COTTON_OIL",
    "Beurre Karité": "SHEA_BUTTER",
}

DEVISES = ["EUR", "USD", "FCFA"]
INCOTERMS = ["FOB COTONOU PORT", "CFR", "CIF", "EXW"]
STATUTS = ["BROUILLON", "CONFIRMEE", "EN_COURS", "EXPEDIEE", "FACTUREE", "PAYEE", "ANNULEE"]

COLS = [
    "ID", "Ref_Vente", "Date_Vente",
    "Client", "Client_Adresse", "Pays_Client",
    "Produit", "Code_Produit",
    "Quantite_MT", "PU", "Devise", "Montant_Brut",
    "Avance", "Montant_Net",
    "Incoterm", "Contrat_No", "BL_No", "Invoice_No",
    "Destination", "Shipping_Line", "Nb_Containers",
    "Statut", "Notes",
    "Cree_Par", "Heure_Creation",
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


def _n(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(str(v).replace(" ", "").replace(",", ""))
    except Exception:
        return 0.0


def ensure_file():
    if os.path.exists(FICHIER):
        return
    with pd.ExcelWriter(FICHIER, engine="openpyxl") as w:
        pd.DataFrame(columns=COLS).to_excel(w, sheet_name=FEUILLE, index=False)


def charger():
    ensure_file()
    try:
        df = pd.read_excel(FICHIER, sheet_name=FEUILLE)
    except Exception:
        return pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in df.columns:
            df[c] = 0 if c in (
                "ID", "Quantite_MT", "PU", "Montant_Brut", "Avance", "Montant_Net", "Nb_Containers"
            ) else ""
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
    df["Date_Vente"] = pd.to_datetime(df["Date_Vente"], errors="coerce", dayfirst=True)
    for c in ["Quantite_MT", "PU", "Montant_Brut", "Avance", "Montant_Net"]:
        df[c] = df[c].map(_n)
    df["Nb_Containers"] = pd.to_numeric(df["Nb_Containers"], errors="coerce").fillna(0).astype(int)
    for c in COLS:
        if c not in (
            "ID", "Date_Vente", "Quantite_MT", "PU", "Montant_Brut", "Avance", "Montant_Net", "Nb_Containers"
        ):
            df[c] = df[c].map(_s)
    return df.sort_values(["Date_Vente", "ID"], ascending=[False, False]).reset_index(drop=True)


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
    import time, gc
    gc.collect()
    last = None
    for i in range(6):
        try:
            os.replace(tmp, FICHIER)
            return
        except PermissionError as e:
            last = e
            time.sleep(0.5)
    try:
        os.remove(tmp)
    except OSError:
        pass
    raise PermissionError(
        f"Impossible d'enregistrer {os.path.basename(FICHIER)} (fichier ouvert ?). Fermez Excel / OneDrive."
    ) from last


def prochain_id(df=None):
    if df is None:
        df = charger()
    return 1 if df.empty else int(df["ID"].max()) + 1


def prochain_ref(annee=None):
    df = charger()
    annee = annee or date.today().year
    nums = []
    for r in df.get("Ref_Vente", pd.Series(dtype=str)).astype(str):
        # VE-2026-0001
        parts = r.replace(" ", "").split("-")
        if len(parts) >= 3 and parts[0].upper() == "VE" and parts[1].isdigit() and int(parts[1]) == int(annee):
            try:
                nums.append(int(parts[2]))
            except Exception:
                pass
    n = max(nums) + 1 if nums else 1
    return f"VE-{annee}-{n:04d}"


def calculer(qte, pu, avance=0.0):
    qte, pu, avance = _n(qte), _n(pu), _n(avance)
    brut = round(qte * pu, 2)
    net = round(brut - avance, 2)
    return brut, avance, net


def ajouter(data, user=""):
    df = charger()
    ref = _s(data.get("Ref_Vente")) or prochain_ref()
    if not df.empty and ref in df["Ref_Vente"].astype(str).values:
        return False, f"Référence {ref} déjà utilisée.", None
    produit = _s(data.get("Produit"))
    if produit not in PRODUITS_VENTE:
        return False, "Produit non reconnu.", None
    qte = _n(data.get("Quantite_MT"))
    pu = _n(data.get("PU"))
    if qte <= 0 or pu <= 0:
        return False, "Quantité et PU doivent être > 0.", None
    brut, av, net = calculer(qte, pu, data.get("Avance"))
    row = {
        "ID": prochain_id(df),
        "Ref_Vente": ref,
        "Date_Vente": pd.Timestamp(data.get("Date_Vente") or date.today()),
        "Client": _s(data.get("Client")),
        "Client_Adresse": _s(data.get("Client_Adresse")),
        "Pays_Client": _s(data.get("Pays_Client")),
        "Produit": produit,
        "Code_Produit": PRODUIT_CODE.get(produit, ""),
        "Quantite_MT": qte,
        "PU": pu,
        "Devise": _s(data.get("Devise") or "EUR"),
        "Montant_Brut": brut,
        "Avance": av,
        "Montant_Net": net,
        "Incoterm": _s(data.get("Incoterm") or "FOB COTONOU PORT"),
        "Contrat_No": _s(data.get("Contrat_No")),
        "BL_No": _s(data.get("BL_No")),
        "Invoice_No": _s(data.get("Invoice_No")),
        "Destination": _s(data.get("Destination")),
        "Shipping_Line": _s(data.get("Shipping_Line")),
        "Nb_Containers": int(_n(data.get("Nb_Containers"))),
        "Statut": _s(data.get("Statut") or "CONFIRMEE"),
        "Notes": _s(data.get("Notes")),
        "Cree_Par": _s(user),
        "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    try:
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        sauver(df)
    except PermissionError as e:
        return False, str(e), None
    return True, f"Vente {ref} enregistrée.", row["ID"]


def maj(eid, data):
    df = charger()
    m = df["ID"] == int(eid)
    if not m.any():
        return False, "Vente introuvable."
    produit = _s(data.get("Produit"))
    qte = _n(data.get("Quantite_MT"))
    pu = _n(data.get("PU"))
    brut, av, net = calculer(qte, pu, data.get("Avance"))
    updates = {
        "Date_Vente": pd.Timestamp(data.get("Date_Vente") or date.today()),
        "Client": _s(data.get("Client")),
        "Client_Adresse": _s(data.get("Client_Adresse")),
        "Pays_Client": _s(data.get("Pays_Client")),
        "Produit": produit,
        "Code_Produit": PRODUIT_CODE.get(produit, ""),
        "Quantite_MT": qte,
        "PU": pu,
        "Devise": _s(data.get("Devise")),
        "Montant_Brut": brut,
        "Avance": av,
        "Montant_Net": net,
        "Incoterm": _s(data.get("Incoterm")),
        "Contrat_No": _s(data.get("Contrat_No")),
        "BL_No": _s(data.get("BL_No")),
        "Invoice_No": _s(data.get("Invoice_No")),
        "Destination": _s(data.get("Destination")),
        "Shipping_Line": _s(data.get("Shipping_Line")),
        "Nb_Containers": int(_n(data.get("Nb_Containers"))),
        "Statut": _s(data.get("Statut")),
        "Notes": _s(data.get("Notes")),
    }
    for k, v in updates.items():
        df.loc[m, k] = v
    try:
        sauver(df)
    except PermissionError as e:
        return False, str(e)
    return True, f"Vente #{eid} mise à jour."


def liste_clients():
    df = charger()
    if df.empty:
        return []
    return sorted({_s(x) for x in df["Client"] if _s(x)})


def to_excel_bytes(df):
    buf = BytesIO()
    out = df.copy()
    if "Date_Vente" in out.columns:
        out["Date_Vente"] = pd.to_datetime(out["Date_Vente"], errors="coerce").dt.strftime("%d/%m/%Y")
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="VENTES")
    buf.seek(0)
    return buf.getvalue()
