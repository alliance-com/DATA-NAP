# -*- coding: utf-8 -*-
"""
Commercial Invoices NAP SARL — base Excel
Fichier : commercial_invoices.xlsx (racine projet ou dossier export)
"""
from __future__ import annotations

import os
from datetime import datetime, date
from io import BytesIO

import pandas as pd

_BASE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_BASE)
FICHIER = os.path.join(_ROOT, "commercial_invoices.xlsx")
if not os.path.isdir(_ROOT):
    FICHIER = "commercial_invoices.xlsx"

FEUILLE = "INVOICES"
VERSION = 1

COLS = [
    "ID", "Invoice_No", "Date_Invoice",
    "Buyer", "Buyer_Address",
    "Description", "Product", "Nb_Containers", "Container_Type",
    "Contract_No", "Incoterm", "BL_No", "Shipping_Line",
    "Origin", "Final_Destination",
    "Qty_MT", "Price_Per_MT", "Currency",
    "Gross_Amount", "Advance", "Less_PI_Pct", "Less_PI_Amount",
    "Net_Amount", "Amount_Words",
    "Bank_Details", "Notes",
    "Statut", "Cree_Par", "Heure_Creation",
]

DEVISES = ["EUR", "USD"]
STATUTS = ["BROUILLON", "EMISE", "PAYEE", "ANNULEE"]

BANK_DEFAULT = (
    "STE NAP SARL, A/C NO: 112241067001, "
    "IBAN NO USD BJ66BJ0620100111224106700155, "
    "SWIFT CODE: ECOCBJBJ, ECOBANK BENIN"
)

PRODUCT_TEMPLATES = {
    "CRUDE SHEA BUTTER": (
        "{nb} x 20 CONTAINER EACH CONTAINER WITH HEATING PAD "
        "RAW CRUDE SHEA BUTTER IN FLEXI TANK OF REP OF BENIN ORIGIN WEST AFRICA "
        "MATERIAL DESCRIPTION: FA010C CRUDE SHB\n"
        "CONTRACT NO:{contract}\n"
        "Prices are {incoterm}\n"
        "Bill of Lading:  {bl}\n"
        "Shipping Line: {liner}\n"
        "ORIGIN OF GOODS –{origin}\n"
        "FINAL DESTINATION – {dest}"
    ),
    "SHEA NUT MEAL": (
        "{nb} x 20 CONTAINERS\n"
        "SHEA NUT MEAL BENIN ORIGIN\n"
        "Contract No ;{contract}\n"
        "Prices are {incoterm}\n"
        "Bill of Lading: {bl}\n"
        "Shipping Line: {liner}\n"
        "ORIGIN OF GOODS –{origin}\n"
        "FINAL DESTINATION – {dest}"
    ),
}


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
                "ID", "Nb_Containers", "Qty_MT", "Price_Per_MT",
                "Gross_Amount", "Advance", "Less_PI_Pct", "Less_PI_Amount", "Net_Amount",
            ) else ""
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
    df["Date_Invoice"] = pd.to_datetime(df["Date_Invoice"], errors="coerce", dayfirst=True)
    for c in ["Qty_MT", "Price_Per_MT", "Gross_Amount", "Advance", "Less_PI_Pct", "Less_PI_Amount", "Net_Amount"]:
        df[c] = df[c].map(_n)
    df["Nb_Containers"] = pd.to_numeric(df["Nb_Containers"], errors="coerce").fillna(0).astype(int)
    for c in COLS:
        if c not in ("ID", "Date_Invoice", "Nb_Containers", "Qty_MT", "Price_Per_MT",
                     "Gross_Amount", "Advance", "Less_PI_Pct", "Less_PI_Amount", "Net_Amount"):
            df[c] = df[c].map(_s)
    return df.sort_values(["Date_Invoice", "ID"], ascending=[False, False]).reset_index(drop=True)


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
    # retry replace
    import time, gc
    gc.collect()
    last = None
    for i in range(6):
        try:
            os.replace(tmp, FICHIER)
            return
        except PermissionError as e:
            last = e
            time.sleep(0.6)
    try:
        os.remove(tmp)
    except OSError:
        pass
    raise PermissionError(
        f"Impossible d'enregistrer {os.path.basename(FICHIER)} (fichier ouvert ou OneDrive). "
        "Fermez Excel puis réessayez."
    ) from last


def prochain_id(df=None):
    if df is None:
        df = charger()
    return 1 if df.empty else int(df["ID"].max()) + 1


def prochain_numero(annee=None):
    df = charger()
    annee = annee or date.today().year
    nums = []
    if not df.empty:
        for r in df["Invoice_No"].astype(str):
            # NAP/1001/2026
            parts = r.replace(" ", "").split("/")
            if len(parts) >= 3 and parts[-1].isdigit() and int(parts[-1]) == int(annee):
                try:
                    nums.append(int(parts[1]))
                except Exception:
                    pass
            elif len(parts) >= 2 and parts[1].isdigit():
                try:
                    nums.append(int(parts[1]))
                except Exception:
                    pass
    n = max(nums) + 1 if nums else 1001
    return f"NAP/{n}/{annee}"


def calculer_montants(qty, price, advance=0.0, less_pi_pct=0.0):
    qty, price = _n(qty), _n(price)
    advance, less_pi_pct = _n(advance), _n(less_pi_pct)
    gross = round(qty * price, 2)
    base_after_adv = gross - advance
    pi_amt = round(base_after_adv * (less_pi_pct / 100.0), 2) if less_pi_pct else 0.0
    net = round(gross - advance - pi_amt, 2)
    return gross, advance, pi_amt, net


def construire_description(data):
    produit = _s(data.get("Product")) or "CRUDE SHEA BUTTER"
    tpl = PRODUCT_TEMPLATES.get(produit)
    if not tpl:
        # générique
        tpl = (
            "{nb} x 20 CONTAINERS\n{product}\n"
            "CONTRACT NO:{contract}\nPrices are {incoterm}\n"
            "Bill of Lading:  {bl}\nShipping Line: {liner}\n"
            "ORIGIN OF GOODS –{origin}\nFINAL DESTINATION – {dest}"
        )
    return tpl.format(
        nb=int(_n(data.get("Nb_Containers")) or 0),
        product=produit,
        contract=_s(data.get("Contract_No")),
        incoterm=_s(data.get("Incoterm") or "FOB COTONOU PORT"),
        bl=_s(data.get("BL_No")),
        liner=_s(data.get("Shipping_Line")),
        origin=_s(data.get("Origin") or "COTONOU BENIN"),
        dest=_s(data.get("Final_Destination")),
    )


def ajouter(data, user=""):
    df = charger()
    num = _s(data.get("Invoice_No")) or prochain_numero()
    if not df.empty and num in df["Invoice_No"].astype(str).values:
        return False, f"Invoice {num} existe déjà.", None

    qty = _n(data.get("Qty_MT"))
    price = _n(data.get("Price_Per_MT"))
    advance = _n(data.get("Advance"))
    pi_pct = _n(data.get("Less_PI_Pct"))
    gross, adv, pi_amt, net = calculer_montants(qty, price, advance, pi_pct)

    desc = _s(data.get("Description"))
    if not desc:
        desc = construire_description(data)

    words = _s(data.get("Amount_Words"))
    if not words:
        try:
            from export.ci_amount_words import amount_in_words
            words = amount_in_words(net, _s(data.get("Currency") or "EUR"))
        except Exception:
            try:
                from ci_amount_words import amount_in_words
                words = amount_in_words(net, _s(data.get("Currency") or "EUR"))
            except Exception:
                words = ""

    row = {
        "ID": prochain_id(df),
        "Invoice_No": num,
        "Date_Invoice": pd.Timestamp(data.get("Date_Invoice") or date.today()),
        "Buyer": _s(data.get("Buyer")),
        "Buyer_Address": _s(data.get("Buyer_Address")),
        "Description": desc,
        "Product": _s(data.get("Product")),
        "Nb_Containers": int(_n(data.get("Nb_Containers"))),
        "Container_Type": _s(data.get("Container_Type") or "20'"),
        "Contract_No": _s(data.get("Contract_No")),
        "Incoterm": _s(data.get("Incoterm") or "FOB COTONOU PORT"),
        "BL_No": _s(data.get("BL_No")),
        "Shipping_Line": _s(data.get("Shipping_Line")),
        "Origin": _s(data.get("Origin") or "COTONOU BENIN"),
        "Final_Destination": _s(data.get("Final_Destination")),
        "Qty_MT": qty,
        "Price_Per_MT": price,
        "Currency": _s(data.get("Currency") or "EUR"),
        "Gross_Amount": gross,
        "Advance": adv,
        "Less_PI_Pct": pi_pct,
        "Less_PI_Amount": pi_amt,
        "Net_Amount": net,
        "Amount_Words": words,
        "Bank_Details": _s(data.get("Bank_Details") or BANK_DEFAULT),
        "Notes": _s(data.get("Notes")),
        "Statut": _s(data.get("Statut") or "EMISE"),
        "Cree_Par": _s(user),
        "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    try:
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        sauver(df)
    except PermissionError as e:
        return False, str(e), None
    return True, f"Invoice {num} enregistrée.", row["ID"]


def maj(eid, data):
    df = charger()
    m = df["ID"] == int(eid)
    if not m.any():
        return False, "Invoice introuvable."
    qty = _n(data.get("Qty_MT"))
    price = _n(data.get("Price_Per_MT"))
    advance = _n(data.get("Advance"))
    pi_pct = _n(data.get("Less_PI_Pct"))
    gross, adv, pi_amt, net = calculer_montants(qty, price, advance, pi_pct)
    desc = _s(data.get("Description")) or construire_description(data)
    words = _s(data.get("Amount_Words"))
    if not words:
        try:
            from export.ci_amount_words import amount_in_words
            words = amount_in_words(net, _s(data.get("Currency") or "EUR"))
        except Exception:
            words = ""
    updates = {
        "Date_Invoice": pd.Timestamp(data.get("Date_Invoice") or date.today()),
        "Buyer": _s(data.get("Buyer")),
        "Buyer_Address": _s(data.get("Buyer_Address")),
        "Description": desc,
        "Product": _s(data.get("Product")),
        "Nb_Containers": int(_n(data.get("Nb_Containers"))),
        "Container_Type": _s(data.get("Container_Type")),
        "Contract_No": _s(data.get("Contract_No")),
        "Incoterm": _s(data.get("Incoterm")),
        "BL_No": _s(data.get("BL_No")),
        "Shipping_Line": _s(data.get("Shipping_Line")),
        "Origin": _s(data.get("Origin")),
        "Final_Destination": _s(data.get("Final_Destination")),
        "Qty_MT": qty,
        "Price_Per_MT": price,
        "Currency": _s(data.get("Currency")),
        "Gross_Amount": gross,
        "Advance": adv,
        "Less_PI_Pct": pi_pct,
        "Less_PI_Amount": pi_amt,
        "Net_Amount": net,
        "Amount_Words": words,
        "Bank_Details": _s(data.get("Bank_Details")),
        "Notes": _s(data.get("Notes")),
        "Statut": _s(data.get("Statut")),
    }
    for k, v in updates.items():
        df.loc[m, k] = v
    try:
        sauver(df)
    except PermissionError as e:
        return False, str(e)
    return True, f"Invoice #{eid} mise à jour."


def liste_buyers():
    df = charger()
    if df.empty:
        return []
    return sorted({_s(x) for x in df["Buyer"] if _s(x)})


def to_excel_bytes(df):
    buf = BytesIO()
    out = df.copy()
    if "Date_Invoice" in out.columns:
        out["Date_Invoice"] = pd.to_datetime(out["Date_Invoice"], errors="coerce").dt.strftime("%d/%m/%Y")
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="INVOICES")
    buf.seek(0)
    return buf.getvalue()
