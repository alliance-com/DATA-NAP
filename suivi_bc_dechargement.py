# -*- coding: utf-8 -*-
"""
NAP SARL — Suivi Bons de Commande ↔ Déchargements
Excel conditionnel + PDF A3 + rattachement BC
À placer à la racine du projet.
"""
from __future__ import annotations

import os
import re
from datetime import datetime
from io import BytesIO

import pandas as pd

import dechargement_data as dd

VERSION_SUIVI_BC = 1

FICHIER_BONS = getattr(dd, "FICHIER_BONS", "donnees_bons.xlsx")
SEUIL_LIVRE = getattr(dd, "SEUIL_LIVRE", 0.98)


def _s(v):
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


def _n(v, defaut=0.0):
    try:
        x = float(v)
        return defaut if pd.isna(x) else x
    except Exception:
        return defaut


def _code_produit_from_ref(ref, produit_txt=""):
    ref_u = _s(ref).upper().replace("\\", "/")
    for code in dd.PRODUITS:
        if f"-BC-{code}" in ref_u:
            return code
    # alias texte produit
    p = _s(produit_txt).upper()
    for alias, code in dd.PRODUITS_ALIAS.items():
        if alias in p:
            return code
    for code, lib in dd.PRODUITS_LIBELLES.items():
        if code in p or _s(lib).upper() in p:
            return code
    return ""


def charger_bons_bruts():
    if not os.path.exists(FICHIER_BONS):
        return pd.DataFrame()
    try:
        df = pd.read_excel(FICHIER_BONS)
    except Exception:
        return pd.DataFrame()
    if df.empty:
        return df
    for c in [
        "Reference", "Date_Emission", "Produit", "Code_Fournisseur", "Nom_Fournisseur",
        "Telephone_Fournisseur", "Commune_Origine", "Quantite", "Numero_Tracteur",
        "Chauffeur", "Numero_Permis", "Telephone_Chauffeur", "Utilisateur",
    ]:
        if c not in df.columns:
            df[c] = ""
    df["Reference"] = df["Reference"].map(_s)
    df = df[df["Reference"] != ""].copy()
    df["Date_Emission"] = pd.to_datetime(df["Date_Emission"], errors="coerce", dayfirst=True)
    df["Qte_KG"] = df["Quantite"].map(dd.quantite_kg)
    df["Code_Produit"] = [
        _code_produit_from_ref(r, p) for r, p in zip(df["Reference"], df["Produit"])
    ]
    return df.reset_index(drop=True)


def _dechargements_tous_produits(produit_filtre=None):
    """Charge et concat extène les déchargements (un ou tous produits)."""
    ancien = dd.get_produit()
    frames = []
    try:
        codes = [produit_filtre] if produit_filtre and produit_filtre in dd.PRODUITS else list(dd.PRODUITS)
        for p in codes:
            dd.set_produit(p)
            d = dd.charger_dechargements()
            if d is None or d.empty:
                continue
            d = d.copy()
            d["Produit_Fichier"] = p
            frames.append(d)
    finally:
        dd.set_produit(ancien)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def construire_suivi(produit=None, statut=None, date_debut=None, date_fin=None):
    """
    Une ligne par bon de commande + situation de déchargement.
    statut: None | 'DECHARGE' | 'NON_DECHARGE' | 'PARTIEL' | 'CAMION_DIFF'
    """
    bons = charger_bons_bruts()
    if bons.empty:
        return pd.DataFrame()

    if produit and produit in dd.PRODUITS:
        bons = bons[bons["Code_Produit"] == produit]

    if date_debut is not None:
        bons = bons[bons["Date_Emission"].notna() & (bons["Date_Emission"].dt.date >= date_debut)]
    if date_fin is not None:
        bons = bons[bons["Date_Emission"].notna() & (bons["Date_Emission"].dt.date <= date_fin)]

    dech = _dechargements_tous_produits(produit if produit in dd.PRODUITS else None)

    lignes = []
    for _, b in bons.iterrows():
        ref = _s(b["Reference"])
        camion_bc = _s(b.get("Numero_Tracteur"))
        qte = _n(b.get("Qte_KG"))

        sub = pd.DataFrame()
        if not dech.empty and "Ref_BC" in dech.columns:
            sub = dech[dech["Ref_BC"].map(_s) == ref].copy()

        nb = len(sub)
        poids = float(sub["Poids_Net_KG"].map(_n).sum()) if nb else 0.0
        tickets = " | ".join(dict.fromkeys(_s(x) for x in sub["Weighment_Ticket"] if _s(x))) if nb else ""
        transactions = " | ".join(dict.fromkeys(_s(x) for x in sub["Transaction"] if _s(x))) if nb else ""
        camions_livres = " | ".join(dict.fromkeys(_s(x) for x in sub["N_Camion"] if _s(x))) if nb else ""
        dates_d = ""
        if nb:
            ds = []
            for d in sub["Date_Dechargement"]:
                if pd.notna(d) and hasattr(d, "strftime"):
                    ds.append(d.strftime("%d/%m/%Y"))
            dates_d = " | ".join(dict.fromkeys(ds))

        camion_diff = False
        if nb and camion_bc:
            for c_liv in sub["N_Camion"].map(_s):
                if c_liv and dd.camion_different(camion_bc, c_liv):
                    camion_diff = True
                    break

        if nb == 0:
            st_liv = "NON DECHARGE"
        elif qte > 0 and poids >= SEUIL_LIVRE * qte:
            st_liv = "DECHARGE"
        elif nb > 0 and qte > 0 and poids < SEUIL_LIVRE * qte:
            st_liv = "PARTIEL"
        else:
            st_liv = "DECHARGE"  # déchargé sans qté commandée connue

        if camion_diff and st_liv in ("DECHARGE", "PARTIEL"):
            st_aff = "CAMION DIFFERENT"
        else:
            st_aff = st_liv

        lignes.append({
            "Reference": ref,
            "Date_Emission": b.get("Date_Emission"),
            "Produit": _s(b.get("Produit")) or dd.libelle_produit(_s(b.get("Code_Produit"))),
            "Code_Produit": _s(b.get("Code_Produit")),
            "Code_Fournisseur": _s(b.get("Code_Fournisseur")),
            "Fournisseur": _s(b.get("Nom_Fournisseur")),
            "Quantite": _s(b.get("Quantite")),
            "Qte_KG": qte,
            "Camion_BC": camion_bc,
            "Chauffeur": _s(b.get("Chauffeur")),
            "Nb_Camions_Decharges": nb,
            "Poids_Livre_KG": poids,
            "Ecart_KG": (poids - qte) if qte > 0 else None,
            "Pct_Livre": round(poids / qte * 100, 1) if qte > 0 else None,
            "Camions_Livres": camions_livres,
            "Transactions": transactions,
            "Tickets_Pesee": tickets,
            "Dates_Dechargement": dates_d,
            "Statut": st_aff,
            "Statut_Livraison": st_liv,
            "Camion_Diff": camion_diff,
        })

    out = pd.DataFrame(lignes)
    if out.empty:
        return out

    # Filtre statut
    if statut and statut != "TOUS":
        if statut == "DECHARGE":
            out = out[out["Statut_Livraison"] == "DECHARGE"]
        elif statut == "NON_DECHARGE":
            out = out[out["Statut_Livraison"] == "NON DECHARGE"]
        elif statut == "PARTIEL":
            out = out[out["Statut_Livraison"] == "PARTIEL"]
        elif statut == "CAMION_DIFF":
            out = out[out["Camion_Diff"] == True]  # noqa: E712

    return out.sort_values("Reference", ascending=False).reset_index(drop=True)


def excel_conditionnel(df: pd.DataFrame) -> bytes:
    """Excel avec mise en forme conditionnelle par statut."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.dataframe import dataframe_to_rows

    cols = [
        "Reference", "Date_Emission", "Produit", "Fournisseur", "Quantite", "Qte_KG",
        "Camion_BC", "Camions_Livres", "Nb_Camions_Decharges", "Poids_Livre_KG",
        "Ecart_KG", "Pct_Livre", "Transactions", "Tickets_Pesee", "Dates_Dechargement",
        "Statut", "Chauffeur",
    ]
    cols = [c for c in cols if c in df.columns]
    data = df[cols].copy() if not df.empty else pd.DataFrame(columns=cols)
    if "Date_Emission" in data.columns:
        data["Date_Emission"] = pd.to_datetime(data["Date_Emission"], errors="coerce").dt.strftime("%d/%m/%Y")

    wb = Workbook()
    ws = wb.active
    ws.title = "Suivi BC"

    header_fill = PatternFill("solid", fgColor="1A5276")
    header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    thin = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    green = PatternFill("solid", fgColor="C6EFCE")
    yellow = PatternFill("solid", fgColor="FFEB9C")
    orange = PatternFill("solid", fgColor="FCE4D6")
    white_font = Font(name="Calibri", size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Titre
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(cols), 1))
    ws["A1"] = f"NAP SARL — Suivi Bons de Commande / Déchargements — {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A1"].font = Font(bold=True, size=14, color="1A5276", name="Calibri")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22

    # Légende
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(len(cols), 1))
    ws["A2"] = (
        "VERT = Déchargé  |  JAUNE = Déchargé avec camion différent du bon  |  "
        "ORANGE = Partiel  |  Blanc = Non déchargé"
    )
    ws["A2"].font = Font(size=9, italic=True, color="555555")

    # En-têtes ligne 4
    headers = {
        "Reference": "N° BC",
        "Date_Emission": "Date émission",
        "Produit": "Produit",
        "Fournisseur": "Fournisseur",
        "Quantite": "Qté commandée",
        "Qte_KG": "Qté (kg)",
        "Camion_BC": "Camion sur BC",
        "Camions_Livres": "Camion(s) livré(s)",
        "Nb_Camions_Decharges": "Nb camions",
        "Poids_Livre_KG": "Poids livré (kg)",
        "Ecart_KG": "Écart (kg)",
        "Pct_Livre": "% livré",
        "Transactions": "N° Transaction",
        "Tickets_Pesee": "Ticket(s) pesée",
        "Dates_Dechargement": "Date(s) déch.",
        "Statut": "Statut",
        "Chauffeur": "Chauffeur BC",
    }
    for col_idx, c in enumerate(cols, 1):
        cell = ws.cell(row=4, column=col_idx, value=headers.get(c, c))
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin
    ws.row_dimensions[4].height = 28
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(cols))}4"

    # Données
    for r_idx, row in enumerate(data.itertuples(index=False), 5):
        statut = ""
        for c_idx, (col_name, val) in enumerate(zip(cols, row), 1):
            if col_name == "Statut":
                statut = _s(val)
            v = val
            if col_name in ("Qte_KG", "Poids_Livre_KG", "Ecart_KG", "Pct_Livre", "Nb_Camions_Decharges"):
                try:
                    v = float(val) if val is not None and not (isinstance(val, float) and pd.isna(val)) else None
                except Exception:
                    v = val
            cell = ws.cell(row=r_idx, column=c_idx, value=v if v is not None and not (isinstance(v, float) and pd.isna(v)) else "")
            cell.font = white_font
            cell.border = thin
            cell.alignment = center if c_idx > 4 else Alignment(vertical="center", wrap_text=True)

        # Couleur ligne selon statut
        fill = None
        if statut in ("DECHARGE",):
            fill = green
        elif statut in ("CAMION DIFFERENT",):
            fill = yellow
        elif statut in ("PARTIEL",):
            fill = orange
        if fill:
            for c_idx in range(1, len(cols) + 1):
                ws.cell(row=r_idx, column=c_idx).fill = fill

    # Largeurs
    widths = {
        "Reference": 28, "Date_Emission": 12, "Produit": 18, "Fournisseur": 22,
        "Quantite": 12, "Qte_KG": 10, "Camion_BC": 14, "Camions_Livres": 16,
        "Nb_Camions_Decharges": 10, "Poids_Livre_KG": 12, "Ecart_KG": 10, "Pct_Livre": 9,
        "Transactions": 16, "Tickets_Pesee": 14, "Dates_Dechargement": 14, "Statut": 16, "Chauffeur": 14,
    }
    for i, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 12)

    # Feuille légende / filtres appliqués
    ws2 = wb.create_sheet("Legende")
    ws2["A1"] = "Légende des couleurs"
    ws2["A1"].font = Font(bold=True, size=12)
    ws2["A3"] = "Déchargé (complet)"
    ws2["A3"].fill = green
    ws2["A4"] = "Déchargé — camion différent du bon"
    ws2["A4"].fill = yellow
    ws2["A5"] = "Partiel (livraison en cours)"
    ws2["A5"].fill = orange
    ws2["A6"] = "Non déchargé"
    ws2["B3"] = "Toute la ligne en vert"
    ws2["B4"] = "Toute la ligne en jaune — vérifier / rattacher le BC si besoin"
    ws2["B5"] = "Ligne orange"
    ws2["B6"] = "Sans couleur"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def pdf_a3(df: pd.DataFrame, titre_filtre: str = "") -> bytes:
    """Rapport PDF paysage A3."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, HRFlowable,
    )
    from xml.sax.saxutils import escape

    buf = BytesIO()
    page = landscape(A3)
    doc = SimpleDocTemplate(
        buf, pagesize=page,
        leftMargin=10 * mm, rightMargin=10 * mm,
        topMargin=10 * mm, bottomMargin=12 * mm,
        title="Suivi BC Déchargements NAP",
        author="NAP SARL",
    )
    W = page[0] - 20 * mm
    navy = colors.HexColor("#1A5276")
    green = colors.HexColor("#C6EFCE")
    yellow = colors.HexColor("#FFEB9C")
    orange = colors.HexColor("#FCE4D6")

    st_t = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=14, alignment=TA_CENTER, textColor=navy)
    st_s = ParagraphStyle("s", fontName="Helvetica", fontSize=9, alignment=TA_CENTER, textColor=colors.HexColor("#333"))
    st_c = ParagraphStyle("c", fontName="Helvetica", fontSize=7, leading=9, alignment=TA_CENTER)
    st_l = ParagraphStyle("l", fontName="Helvetica", fontSize=7, leading=9)
    st_h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=7, textColor=colors.white, alignment=TA_CENTER)

    el = [
        Paragraph("NAP SARL — SUIVI BONS DE COMMANDE / DÉCHARGEMENTS", st_t),
        Paragraph(
            f"{escape(titre_filtre)} · Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
            st_s,
        ),
        Spacer(1, 2 * mm),
        HRFlowable(width="100%", thickness=1.2, color=colors.HexColor("#5DADE2"), spaceAfter=3),
        Paragraph(
            "Légende : <font color='#1E8449'><b>VERT</b></font> = Déchargé &nbsp; "
            "<font color='#B7950B'><b>JAUNE</b></font> = Camion différent &nbsp; "
            "<font color='#D35400'><b>ORANGE</b></font> = Partiel &nbsp; "
            "Blanc = Non déchargé",
            st_s,
        ),
        Spacer(1, 2 * mm),
    ]

    headers = [
        "N° BC", "Date", "Produit", "Fournisseur", "Qté", "Camion BC",
        "Camion livré", "Nb", "Poids kg", "%", "Transaction", "Ticket", "Statut",
    ]
    data = [[Paragraph(h, st_h) for h in headers]]

    def _pct(v):
        try:
            return f"{float(v):.0f}%" if v is not None and not pd.isna(v) else ""
        except Exception:
            return ""

    rows_meta = []  # statut pour couleur
    if df is not None and not df.empty:
        for _, r in df.iterrows():
            date_txt = ""
            d = r.get("Date_Emission")
            if pd.notna(d) and hasattr(d, "strftime"):
                date_txt = d.strftime("%d/%m/%Y")
            row = [
                Paragraph(escape(_s(r.get("Reference"))), st_c),
                Paragraph(escape(date_txt), st_c),
                Paragraph(escape(_s(r.get("Produit"))[:18]), st_l),
                Paragraph(escape(_s(r.get("Fournisseur"))[:20]), st_l),
                Paragraph(escape(_s(r.get("Quantite"))), st_c),
                Paragraph(escape(_s(r.get("Camion_BC"))), st_c),
                Paragraph(escape(_s(r.get("Camions_Livres"))[:22]), st_c),
                Paragraph(str(int(_n(r.get("Nb_Camions_Decharges")))), st_c),
                Paragraph(f"{_n(r.get('Poids_Livre_KG')):,.0f}".replace(",", " "), st_c),
                Paragraph(_pct(r.get("Pct_Livre")), st_c),
                Paragraph(escape(_s(r.get("Transactions"))[:18]), st_c),
                Paragraph(escape(_s(r.get("Tickets_Pesee"))[:14]), st_c),
                Paragraph(escape(_s(r.get("Statut"))), st_c),
            ]
            data.append(row)
            rows_meta.append(_s(r.get("Statut")))

    col_w = [
        0.12 * W, 0.06 * W, 0.08 * W, 0.11 * W, 0.06 * W, 0.07 * W,
        0.08 * W, 0.04 * W, 0.07 * W, 0.04 * W, 0.09 * W, 0.08 * W, 0.10 * W,
    ]
    t = Table(data, colWidths=col_w, repeatRows=1)
    sty = [
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONT", (0, 1), (-1, -1), "Helvetica", 7),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#AAB7B8")),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for i, st_row in enumerate(rows_meta, 1):
        if st_row == "DECHARGE":
            sty.append(("BACKGROUND", (0, i), (-1, i), green))
        elif st_row == "CAMION DIFFERENT":
            sty.append(("BACKGROUND", (0, i), (-1, i), yellow))
        elif st_row == "PARTIEL":
            sty.append(("BACKGROUND", (0, i), (-1, i), orange))
    t.setStyle(TableStyle(sty))
    el.append(t)

    def pied(canvas, d_):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#666"))
        canvas.drawString(10 * mm, 6 * mm, "NAP SARL — Suivi BC / Déchargements — Document interne")
        canvas.drawRightString(page[0] - 10 * mm, 6 * mm, f"Page {d_.page}")
        canvas.restoreState()

    doc.build(el, onFirstPage=pied, onLaterPages=pied)
    buf.seek(0)
    return buf.getvalue()


def rattacher_bc_a_transaction(transaction, nouvelle_ref_bc, produit_code=None, utilisateur=""):
    """
    Met à jour Ref_BC (et si possible Camion_BC) sur un déchargement existant.
    produit_code : fichier produit où chercher la transaction.
    """
    tr = _s(transaction)
    ref = _s(nouvelle_ref_bc)
    if not tr or not ref:
        return False, "Transaction et N° BC obligatoires."

    ancien = dd.get_produit()
    try:
        codes = [produit_code] if produit_code in dd.PRODUITS else list(dd.PRODUITS)
        trouve = False
        for p in codes:
            dd.set_produit(p)
            dech = dd.charger_dechargements()
            aff = dd.charger_affectations()
            if dech.empty:
                continue
            m = dech["Transaction"].map(_s) == tr
            if not m.any():
                continue
            # Camion prévu sur le bon
            camion_bc = ""
            bons = charger_bons_bruts()
            if not bons.empty:
                br = bons[bons["Reference"].map(_s) == ref]
                if not br.empty:
                    camion_bc = _s(br.iloc[0].get("Numero_Tracteur"))
            dech.loc[m, "Ref_BC"] = ref
            if camion_bc and "Camion_BC" in dech.columns:
                dech.loc[m, "Camion_BC"] = camion_bc
            # note
            if "Remarque" in dech.columns:
                prev = _s(dech.loc[m, "Remarque"].iloc[0])
                note = f"{datetime.now().strftime('%d/%m/%Y %H:%M')} {_s(utilisateur)}: BC rattaché → {ref}"
                dech.loc[m, "Remarque"] = (prev + " | " + note) if prev else note
            dd.sauver(dech, aff)
            trouve = True
            break
        if not trouve:
            return False, f"Transaction « {tr} » introuvable dans les déchargements."
        return True, f"N° BC **{ref}** rattaché à la transaction **{tr}**."
    finally:
        dd.set_produit(ancien)


def dechargements_sans_bc(produit=None):
    """Déchargements ACHAT sans Ref_BC (candidats au rattachement)."""
    dech = _dechargements_tous_produits(produit if produit in dd.PRODUITS else None)
    if dech.empty:
        return dech
    if "Type_Mouvement" in dech.columns:
        dech = dech[dech["Type_Mouvement"].map(lambda x: not _s(x).upper().startswith("V"))]
    return dech[dech["Ref_BC"].map(_s) == ""].reset_index(drop=True)
