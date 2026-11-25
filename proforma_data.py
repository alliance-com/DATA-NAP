# -*- coding: utf-8 -*-
"""Factures proforma NAP SARL."""
from __future__ import annotations

import os
from datetime import datetime, date
from io import BytesIO

import pandas as pd

FICHIER = "proformas.xlsx"
FEUILLE = "PROFORMAS"
DOSSIER = "proformas_generes"
os.makedirs(DOSSIER, exist_ok=True)

COLS = [
    "ID", "Numero", "Date_Facture", "Client", "Adresse", "IFU_Client", "Pays",
    "Produit", "Description", "Incoterm", "Conditionnement",
    "Quantite_T", "Prix_Unitaire", "Montant_HT",
    "Appliquer_AIB", "Taux_AIB", "Montant_AIB",
    "Appliquer_TVA", "Taux_TVA", "Montant_TVA",
    "Montant_TTC", "Montant_Lettres",
    "Condition_Paiement", "Validite", "Livraison", "Mode_Paiement",
    "Compte_Bancaire", "Autres_Conditions",
    "Statut", "Cree_Par", "Heure_Creation",
]

DEVISES_LABEL = "FCFA"

CONDITIONS_PAIEMENT = [
    "Paiement de 100 % d'avance en 7 jours.",
    "Paiement de 50 % d'avance, solde avant chargement.",
    "Paiement à 30 jours fin de mois.",
    "Paiement par lettre de crédit.",
]

MODES_PAIEMENT = [
    "Paiement par chèque banque sur l'un des comptes de NAP SARL",
    "Virement bancaire",
    "Mobile Money",
    "Espèces",
]

COMPTE_DEFAUT = (
    "NAP COTON ECOBANK, CODE BANQUE : BJ66BJ062, CODE GUICHET : 01001, "
    "No COMPTE : 112241067003, RIB : 49, IBAN NO : BJ66BJ0620100111224106700349"
)

AUTRES_DEFAUT = (
    "- L'exécution de cette opération est destinée à l'exportation.\n"
    "- La pratique du prix hors taxe pour ces chargements est conditionnée par le retour "
    "des déclarations d'exportations à la direction de NAP SARL.\n"
    "- Si le client fait une exportation maritime ou terrestre, elle sera au nom de NAP SARL. "
    "Après l'exportation l'acheteur fournira le connaissement et la déclaration. "
    "Les frais de l'exportation seront aux comptes de l'acheteur."
)


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


def _b(v):
    s = _s(v).upper()
    if s in ("1", "OUI", "YES", "TRUE", "VRAI"):
        return True
    if isinstance(v, (int, float)) and v:
        return True
    return False




# ---------- Montant en lettres (FR / EN) ----------
_UNITS_FR = [
    "", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf",
    "dix", "onze", "douze", "treize", "quatorze", "quinze", "seize",
    "dix-sept", "dix-huit", "dix-neuf",
]
_TENS_FR = [
    "", "", "vingt", "trente", "quarante", "cinquante",
    "soixante", "soixante", "quatre-vingt", "quatre-vingt",
]
_UNITS_EN = [
    "", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine",
    "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen",
    "seventeen", "eighteen", "nineteen",
]
_TENS_EN = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]


def _sous_1000_fr(n):
    n = int(n)
    if n == 0:
        return ""
    if n < 20:
        return _UNITS_FR[n]
    if n < 100:
        d, u = divmod(n, 10)
        if d == 7:  # soixante-dix
            return "soixante" + ("-" + _UNITS_FR[10 + u] if u else "-dix")
        if d == 9:  # quatre-vingt-dix
            return "quatre-vingt" + ("-" + _UNITS_FR[10 + u] if u else "-dix")
        if d == 8:
            return "quatre-vingt" + (("s" if u == 0 else "-" + _UNITS_FR[u]))
        base = _TENS_FR[d]
        if u == 0:
            return base
        if u == 1 and d in (2, 3, 4, 5, 6):
            return base + " et un"
        return base + "-" + _UNITS_FR[u]
    c, r = divmod(n, 100)
    cent = "cent" if c == 1 else _UNITS_FR[c] + " cent"
    if r == 0:
        return cent + ("s" if c > 1 else "")
    return cent + " " + _sous_1000_fr(r)


def _sous_1000_en(n):
    n = int(n)
    if n == 0:
        return ""
    if n < 20:
        return _UNITS_EN[n]
    if n < 100:
        d, u = divmod(n, 10)
        return _TENS_EN[d] + ("-" + _UNITS_EN[u] if u else "")
    c, r = divmod(n, 100)
    return _UNITS_EN[c] + " hundred" + ((" " + _sous_1000_en(r)) if r else "")


def montant_en_lettres(n, lang="FR", devise="FCFA"):
    """Convertit un montant entier en lettres (FR ou EN)."""
    try:
        n = int(round(float(n)))
    except Exception:
        return ""
    lang = "EN" if str(lang).upper().startswith("EN") else "FR"
    if n < 0:
        n = abs(n)
    if n == 0:
        return "ZÉRO FRANC CFA" if lang == "FR" else "ZERO CFA FRANCS"

    if lang == "FR":
        parts = []
        milliards, n = divmod(n, 1_000_000_000)
        millions, n = divmod(n, 1_000_000)
        milliers, n = divmod(n, 1_000)
        if milliards:
            if milliards == 1:
                parts.append("un milliard")
            else:
                parts.append(_sous_1000_fr(milliards) + " milliards")
        if millions:
            if millions == 1:
                parts.append("un million")
            else:
                parts.append(_sous_1000_fr(millions) + " millions")
        if milliers:
            if milliers == 1:
                parts.append("mille")
            else:
                parts.append(_sous_1000_fr(milliers) + " mille")
        if n:
            parts.append(_sous_1000_fr(n))
        texte = " ".join(parts).replace("  ", " ").strip()
        # "X millions / milliards de francs" uniquement si le reste est vide
        reste = bool(milliers or n)
        if (milliards or millions) and not reste:
            texte = texte + " de francs CFA"
        else:
            texte = texte + " francs CFA"
        return texte.upper()
    else:
        parts = []
        billions, n = divmod(n, 1_000_000_000)
        millions, n = divmod(n, 1_000_000)
        thousands, n = divmod(n, 1_000)
        if billions:
            parts.append(_sous_1000_en(billions) + " billion")
        if millions:
            parts.append(_sous_1000_en(millions) + (" million" if millions == 1 else " million"))
        if thousands:
            parts.append(_sous_1000_en(thousands) + " thousand")
        if n:
            parts.append(_sous_1000_en(n))
        texte = " ".join(parts).upper()
        return f"{texte} CFA FRANCS"



def calculer(qte, pu, appliquer_aib, taux_aib, appliquer_tva, taux_tva):
    ht = round(_n(qte) * _n(pu), 2)
    aib = round(ht * (_n(taux_aib) / 100.0), 2) if appliquer_aib else 0.0
    base_tva = ht + aib
    tva = round(base_tva * (_n(taux_tva) / 100.0), 2) if appliquer_tva else 0.0
    ttc = round(ht + aib + tva, 2)
    return ht, aib, tva, ttc


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
            df[c] = 0 if c in ("ID", "Quantite_T", "Prix_Unitaire", "Montant_HT",
                               "Taux_AIB", "Montant_AIB", "Taux_TVA", "Montant_TVA", "Montant_TTC") else ""
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
    for c in ["Quantite_T", "Prix_Unitaire", "Montant_HT", "Taux_AIB", "Montant_AIB",
              "Taux_TVA", "Montant_TVA", "Montant_TTC"]:
        df[c] = df[c].map(_n)
    df["Date_Facture"] = pd.to_datetime(df["Date_Facture"], errors="coerce", dayfirst=True)
    return df.sort_values(["Date_Facture", "ID"], ascending=[False, False]).reset_index(drop=True)


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
    return 1 if df.empty else int(df["ID"].max()) + 1


def prochain_numero(annee=None):
    df = charger()
    annee = annee or date.today().year
    nums = []
    if not df.empty:
        for r in df["Numero"].astype(str):
            # NAP/22/2026
            parts = r.replace(" ", "").split("/")
            if len(parts) >= 2 and parts[-1].isdigit() and int(parts[-1]) == annee:
                try:
                    nums.append(int(parts[1]))
                except Exception:
                    pass
            elif len(parts) >= 2 and parts[1].isdigit():
                try:
                    nums.append(int(parts[1]))
                except Exception:
                    pass
    n = max(nums) + 1 if nums else 1
    return f"NAP/{n}/{annee}"


def ajouter(data, user=""):
    df = charger()
    eid = prochain_id(df)
    num = _s(data.get("Numero")) or prochain_numero()
    if not df.empty and num in df["Numero"].astype(str).values:
        return False, f"Le numéro {num} existe déjà.", None

    qte = _n(data.get("Quantite_T"))
    pu = _n(data.get("Prix_Unitaire"))
    app_aib = bool(data.get("Appliquer_AIB"))
    app_tva = bool(data.get("Appliquer_TVA"))
    taux_aib = _n(data.get("Taux_AIB") if data.get("Taux_AIB") is not None else 1)
    taux_tva = _n(data.get("Taux_TVA") if data.get("Taux_TVA") is not None else 18)
    ht, aib, tva, ttc = calculer(qte, pu, app_aib, taux_aib, app_tva, taux_tva)

    row = {
        "ID": eid,
        "Numero": num,
        "Date_Facture": pd.Timestamp(data.get("Date_Facture") or date.today()),
        "Client": _s(data.get("Client")),
        "Adresse": _s(data.get("Adresse")),
        "IFU_Client": _s(data.get("IFU_Client")),
        "Pays": _s(data.get("Pays")),
        "Produit": _s(data.get("Produit")),
        "Description": _s(data.get("Description")),
        "Incoterm": _s(data.get("Incoterm") or "EX-USINE"),
        "Conditionnement": _s(data.get("Conditionnement") or "SACS PP DE 50KG"),
        "Quantite_T": qte,
        "Prix_Unitaire": pu,
        "Montant_HT": ht,
        "Appliquer_AIB": "OUI" if app_aib else "NON",
        "Taux_AIB": taux_aib,
        "Montant_AIB": aib,
        "Appliquer_TVA": "OUI" if app_tva else "NON",
        "Taux_TVA": taux_tva,
        "Montant_TVA": tva,
        "Montant_TTC": ttc,
        "Montant_Lettres": _s(data.get("Montant_Lettres")) or montant_en_lettres(ttc, lang="FR"),
        "Condition_Paiement": _s(data.get("Condition_Paiement")),
        "Validite": _s(data.get("Validite")),
        "Livraison": _s(data.get("Livraison")),
        "Mode_Paiement": _s(data.get("Mode_Paiement")),
        "Compte_Bancaire": _s(data.get("Compte_Bancaire") or COMPTE_DEFAUT),
        "Autres_Conditions": _s(data.get("Autres_Conditions") or AUTRES_DEFAUT),
        "Statut": _s(data.get("Statut") or "BROUILLON"),
        "Cree_Par": _s(user),
        "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    sauver(df)
    return True, f"Proforma {num} enregistrée.", eid


def maj(eid, data):
    df = charger()
    m = df["ID"] == int(eid)
    if not m.any():
        return False, "Proforma introuvable."
    qte = _n(data.get("Quantite_T"))
    pu = _n(data.get("Prix_Unitaire"))
    app_aib = bool(data.get("Appliquer_AIB"))
    app_tva = bool(data.get("Appliquer_TVA"))
    taux_aib = _n(data.get("Taux_AIB") if data.get("Taux_AIB") is not None else 1)
    taux_tva = _n(data.get("Taux_TVA") if data.get("Taux_TVA") is not None else 18)
    ht, aib, tva, ttc = calculer(qte, pu, app_aib, taux_aib, app_tva, taux_tva)
    updates = {
        "Date_Facture": pd.Timestamp(data.get("Date_Facture") or date.today()),
        "Client": _s(data.get("Client")),
        "Adresse": _s(data.get("Adresse")),
        "IFU_Client": _s(data.get("IFU_Client")),
        "Pays": _s(data.get("Pays")),
        "Produit": _s(data.get("Produit")),
        "Description": _s(data.get("Description")),
        "Incoterm": _s(data.get("Incoterm")),
        "Conditionnement": _s(data.get("Conditionnement")),
        "Quantite_T": qte,
        "Prix_Unitaire": pu,
        "Montant_HT": ht,
        "Appliquer_AIB": "OUI" if app_aib else "NON",
        "Taux_AIB": taux_aib,
        "Montant_AIB": aib,
        "Appliquer_TVA": "OUI" if app_tva else "NON",
        "Taux_TVA": taux_tva,
        "Montant_TVA": tva,
        "Montant_TTC": ttc,
        "Montant_Lettres": _s(data.get("Montant_Lettres")) or montant_en_lettres(ttc, lang="FR"),
        "Condition_Paiement": _s(data.get("Condition_Paiement")),
        "Validite": _s(data.get("Validite")),
        "Livraison": _s(data.get("Livraison")),
        "Mode_Paiement": _s(data.get("Mode_Paiement")),
        "Compte_Bancaire": _s(data.get("Compte_Bancaire")),
        "Autres_Conditions": _s(data.get("Autres_Conditions")),
    }
    for k, v in updates.items():
        df.loc[m, k] = v
    sauver(df)
    return True, f"Proforma #{eid} mise à jour."


def to_excel_bytes(df):
    buf = BytesIO()
    out = df.copy()
    if "Date_Facture" in out.columns:
        out["Date_Facture"] = pd.to_datetime(out["Date_Facture"], errors="coerce").dt.strftime("%d/%m/%Y")
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="PROFORMAS")
    buf.seek(0)
    return buf.getvalue()
