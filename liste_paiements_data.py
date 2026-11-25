# -*- coding: utf-8 -*-
"""
Liste paiements par produit / période — source : Suivi de paiement.xlsx
"""
from __future__ import annotations

import glob
import os
import re
from datetime import datetime, date
from io import BytesIO

import pandas as pd

FICHIER_CANDIDATS = [
    "Suivi de paiement.xlsx",
    "Suivi de paiement-5.xlsx",
    "Suivi_de_paiement.xlsx",
]
FEUILLE_PREF = "DEPENSES"

COLS_STD = [
    "Date", "No_Transaction", "Produit", "Prestataire", "Beneficiaire",
    "Ref_Doc", "Montant", "Observation", "Mode", "Num_Cheque", "Banque",
    "Type_Operation", "Description",
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
        return float(str(v).replace(" ", "").replace(",", "").replace("\u00a0", ""))
    except Exception:
        return 0.0


def trouver_fichier():
    for n in FICHIER_CANDIDATS:
        if os.path.exists(n):
            return n
    for p in glob.glob("Suivi*aiement*.xlsx") + glob.glob("*paiement*.xlsx"):
        if os.path.isfile(p) and not p.startswith("~$"):
            return p
    return None


def charger_paiements(chemin=None):
    chemin = chemin or trouver_fichier()
    vide = pd.DataFrame(columns=COLS_STD)
    if not chemin or not os.path.exists(chemin):
        return vide, None, "Fichier Suivi de paiement introuvable à la racine du projet."

    try:
        xl = pd.ExcelFile(chemin)
        feuille = FEUILLE_PREF if FEUILLE_PREF in xl.sheet_names else xl.sheet_names[0]
        df = None
        for hdr in (2, 1, 0, 3, 4):
            tmp = pd.read_excel(chemin, sheet_name=feuille, header=hdr)
            cols_l = [str(c).strip().lower() for c in tmp.columns]
            if any("transaction" in c or "montant" in c or "date" in c for c in cols_l):
                df = tmp
                break
        if df is None:
            df = pd.read_excel(chemin, sheet_name=feuille)
    except Exception as e:
        return vide, chemin, f"Lecture impossible : {e}"

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
        "ref doc": "Ref_Doc",
        "réf doc": "Ref_Doc",
        "ref. doc": "Ref_Doc",
        "montant": "Montant",
        "observation": "Observation",
        "mode paiement": "Mode",
        "mode de paiement": "Mode",
        "numero de cheque": "Num_Cheque",
        "numéro de chèque": "Num_Cheque",
        "n° chèque": "Num_Cheque",
        "banque": "Banque",
        "types opération": "Type_Operation",
        "type opération": "Type_Operation",
        "type operation": "Type_Operation",
        "descriptions/object du paiement": "Description",
        "description": "Description",
        "objet": "Description",
    }
    rename = {}
    for c in df.columns:
        k = re.sub(r"\s+", " ", str(c).strip().lower())
        if k in alias:
            rename[c] = alias[k]
    df = df.rename(columns=rename)
    for c in COLS_STD:
        if c not in df.columns:
            df[c] = "" if c != "Montant" else 0.0

    out = df[COLS_STD].copy()
    for c in ["No_Transaction", "Produit", "Prestataire", "Beneficiaire", "Ref_Doc",
              "Observation", "Mode", "Num_Cheque", "Banque", "Type_Operation", "Description"]:
        out[c] = out[c].map(_s)
    out["Montant"] = out["Montant"].map(_n)
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce", dayfirst=True)
    # Écarter dates absurdes (erreurs Excel / parsing)
    try:
        y = out["Date"].dt.year
        out.loc[(y < 2000) | (y > 2100), "Date"] = pd.NaT
    except Exception:
        pass
    # ignorer lignes totalement vides
    out = out[~(
        (out["No_Transaction"] == "") & (out["Montant"] == 0) & (out["Prestataire"] == "")
    )].reset_index(drop=True)
    return out, chemin, None


def liste_produits(df):
    if df is None or df.empty:
        return []
    return sorted({_s(p) for p in df["Produit"] if _s(p)}, key=str.upper)


def filtrer(df, produit=None, date_debut=None, date_fin=None):
    if df is None or df.empty:
        return df
    out = df.copy()
    if produit and str(produit).strip() and str(produit) != "Tous":
        p = str(produit).strip().casefold()
        out = out[out["Produit"].map(lambda x: _s(x).casefold() == p)]
    if date_debut is not None:
        out = out[out["Date"].notna() & (out["Date"] >= pd.Timestamp(date_debut))]
    if date_fin is not None:
        out = out[out["Date"].notna() & (out["Date"] <= pd.Timestamp(date_fin) + pd.Timedelta(days=0))]
        # include full end day
        out = out[out["Date"] <= pd.Timestamp(date_fin) + pd.Timedelta(hours=23, minutes=59)]
    return out.reset_index(drop=True)


def to_excel_bytes(df, produit="Tous", periode=""):
    """Excel formaté : titre, en-têtes bleu ciel, montants, total, largeurs."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from datetime import datetime as _dt

    buf = BytesIO()
    out = df.copy() if df is not None else pd.DataFrame(columns=COLS_STD)
    if "Date" in out.columns:
        out["Date"] = pd.to_datetime(out["Date"], errors="coerce").dt.strftime("%d/%m/%Y")

    # Colonnes affichage ordonnées + libellés
    mapping = [
        ("Date", "Date"),
        ("No_Transaction", "N° Transaction"),
        ("Produit", "Produit"),
        ("Prestataire", "Prestataire / Fournisseur"),
        ("Beneficiaire", "Bénéficiaire"),
        ("Montant", "Montant (FCFA)"),
        ("Mode", "Mode de paiement"),
        ("Banque", "Banque"),
        ("Num_Cheque", "N° Chèque"),
        ("Ref_Doc", "Réf. Doc"),
        ("Observation", "Observation"),
        ("Description", "Description / Objet"),
        ("Type_Operation", "Type opération"),
    ]
    cols_ok = [(c, lab) for c, lab in mapping if c in out.columns]
    data = out[[c for c, _ in cols_ok]].copy()
    data.columns = [lab for _, lab in cols_ok]

    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        # Ligne titre / méta puis tableau à partir de la ligne 5
        data.to_excel(w, index=False, sheet_name="Paiements", startrow=4)
        ws = w.sheets["Paiements"]

        # Styles
        fill_header = PatternFill("solid", fgColor="D6EAF8")
        fill_title = PatternFill("solid", fgColor="1A5276")
        fill_total = PatternFill("solid", fgColor="EAF2F8")
        font_title = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
        font_sub = Font(name="Calibri", size=10, color="1A5276")
        font_header = Font(name="Calibri", bold=True, size=10, color="1A5276")
        font_cell = Font(name="Calibri", size=10)
        font_total = Font(name="Calibri", bold=True, size=11, color="1A5276")
        thin = Border(
            left=Side(style="thin", color="B0B0B0"),
            right=Side(style="thin", color="B0B0B0"),
            top=Side(style="thin", color="B0B0B0"),
            bottom=Side(style="thin", color="B0B0B0"),
        )
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        right = Alignment(horizontal="right", vertical="center")

        n_cols = max(len(cols_ok), 1)
        # Titre fusionné
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
        c1 = ws.cell(1, 1, "STE NAP SARL — LISTE DES PAIEMENTS")
        c1.font = font_title
        c1.fill = fill_title
        c1.alignment = center
        for col in range(1, n_cols + 1):
            ws.cell(1, col).fill = fill_title

        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
        filt = f"Produit : {produit or 'Tous'}"
        if periode:
            filt += f"    |    Période : {periode}"
        c2 = ws.cell(2, 1, filt)
        c2.font = font_sub
        c2.alignment = center

        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=n_cols)
        c3 = ws.cell(3, 1, f"Généré le {_dt.now().strftime('%d/%m/%Y à %H:%M')}  ·  {len(data)} ligne(s)")
        c3.font = Font(name="Calibri", size=9, italic=True, color="5D6D7E")
        c3.alignment = center

        # En-têtes (ligne 5)
        header_row = 5
        for col in range(1, n_cols + 1):
            cell = ws.cell(header_row, col)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = center
            cell.border = thin

        # Corps
        montant_col = None
        for i, (_, lab) in enumerate(cols_ok, start=1):
            if lab == "Montant (FCFA)":
                montant_col = i
        for r in range(header_row + 1, header_row + 1 + len(data)):
            for col in range(1, n_cols + 1):
                cell = ws.cell(r, col)
                cell.font = font_cell
                cell.border = thin
                if montant_col and col == montant_col:
                    cell.alignment = right
                    cell.number_format = '#,##0'
                else:
                    cell.alignment = left
            # zébrage
            if (r - header_row) % 2 == 0:
                for col in range(1, n_cols + 1):
                    if ws.cell(r, col).fill.fgColor is None or ws.cell(r, col).fill.fgColor.rgb == "00000000":
                        ws.cell(r, col).fill = PatternFill("solid", fgColor="F8FBFD")

        # Ligne total
        total_row = header_row + 1 + len(data)
        total_mt = float(out["Montant"].sum()) if "Montant" in out.columns and len(out) else 0.0
        ws.cell(total_row, 1, "TOTAL").font = font_total
        for col in range(1, n_cols + 1):
            cell = ws.cell(total_row, col)
            cell.fill = fill_total
            cell.border = thin
            cell.font = font_total
        if montant_col:
            cell_m = ws.cell(total_row, montant_col, total_mt)
            cell_m.number_format = '#,##0'
            cell_m.font = font_total
            cell_m.alignment = right
            cell_m.fill = fill_total
            cell_m.border = thin
        # Libellé total sur 1re colonne non montant
        ws.cell(total_row, 1, f"TOTAL ({len(data)} paiements)")

        # Largeurs colonnes
        widths = {
            "Date": 12,
            "N° Transaction": 16,
            "Produit": 14,
            "Prestataire / Fournisseur": 28,
            "Bénéficiaire": 22,
            "Montant (FCFA)": 16,
            "Mode de paiement": 14,
            "Banque": 14,
            "N° Chèque": 14,
            "Réf. Doc": 14,
            "Observation": 28,
            "Description / Objet": 28,
            "Type opération": 14,
        }
        for i, (_, lab) in enumerate(cols_ok, start=1):
            ws.column_dimensions[get_column_letter(i)].width = widths.get(lab, 14)

        ws.row_dimensions[1].height = 22
        ws.row_dimensions[header_row].height = 20
        ws.freeze_panes = "A6"
        ws.print_title_rows = "5:5"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

    buf.seek(0)
    return buf.getvalue()
