# -*- coding: utf-8 -*-
"""NAP SARL — Rapport des paiements par fournisseur / bénéficiaire (Excel + PDF A3)."""
from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime
from io import BytesIO
from xml.sax.saxutils import escape

import pandas as pd

MOTS_BENEF = ["prestataire", "fournisseur", "beneficiaire", "benef"]
MOTS_ENTETE = MOTS_BENEF + ["date", "montant", "moyen", "compte", "objet", "designation", "motif"]


# ───────────────────────── Utilitaires ─────────────────────────
def norm(s) -> str:
    t = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return t.lower().strip()


def trouver(cols, mots, exclure=()):
    for c in cols:
        n = norm(c)
        if any(m in n for m in mots) and not any(e in n for e in exclure):
            return c
    return None


def to_num(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    t = (
        s.astype(str)
        .str.replace(r"[^\d,.\-]", "", regex=True)
        .str.replace(",", ".", regex=False)
        .str.replace(r"\.(?=.*\.)", "", regex=True)
    )
    return pd.to_numeric(t, errors="coerce")


def fmt_montant(x) -> str:
    if x is None or pd.isna(x):
        return ""
    x = float(x)
    if abs(x - round(x)) < 0.005:
        return f"{x:,.0f}".replace(",", " ")
    return f"{x:,.2f}".replace(",", " ")


def texte_cellule(v, est_montant=False) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if est_montant:
        return fmt_montant(v)
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def table_texte(df: pd.DataFrame, cols, col_montant=None) -> pd.DataFrame:
    out = pd.DataFrame(index=df.index)
    for c in cols:
        out[c] = [texte_cellule(v, c == col_montant) for v in df[c]]
    return out.reset_index(drop=True)


def synthese(df, col_benef, col_montant) -> pd.DataFrame:
    g = df.groupby(col_benef, dropna=False)
    s = pd.DataFrame({"Bénéficiaire": g.size().index, "Nb paiements": g.size().values})
    if col_montant:
        s["Total"] = g[col_montant].sum().values
        s = s.sort_values("Total", ascending=False)
    return s.reset_index(drop=True)


# ───────────────────────── Lecture du classeur ─────────────────────────
def lister_feuilles(data: bytes):
    with pd.ExcelFile(BytesIO(data)) as xls:
        return list(xls.sheet_names)


def detecter_entete(brut: pd.DataFrame) -> int:
    meilleur, score_max = 0, -1
    for i in range(min(15, len(brut))):
        cells = [norm(c) for c in brut.iloc[i].tolist() if pd.notna(c)]
        score = sum(any(m in c for m in MOTS_ENTETE) for c in cells)
        if score > score_max:
            meilleur, score_max = i, score
    return meilleur


def lire_feuille(data: bytes, feuille: str) -> pd.DataFrame:
    brut = pd.read_excel(BytesIO(data), sheet_name=feuille, header=None)
    brut = brut.dropna(how="all").reset_index(drop=True)
    if brut.empty:
        return pd.DataFrame()
    h = detecter_entete(brut)
    noms, vus = [], {}
    for j, v in enumerate(brut.iloc[h].tolist()):
        n = str(v).strip() if pd.notna(v) else ""
        if not n:
            n = f"Col{j + 1}"
        if n in vus:
            vus[n] += 1
            n = f"{n}_{vus[n]}"
        else:
            vus[n] = 1
        noms.append(n)
    df = brut.iloc[h + 1:].copy()
    df.columns = noms
    return df.dropna(axis=1, how="all").dropna(how="all").reset_index(drop=True)


# ───────────────────────── Export Excel ─────────────────────────
def _py(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, "item"):
        return v.item()
    return v


def excel_bytes(df, cols, col_benef, col_montant, titre_benef, filtres, synth=None) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.properties import PageSetupProperties

    wb = Workbook()
    ws = wb.active
    ws.title = "Rapport"
    n = len(cols)
    navy = "1A5276"
    fin = Side(style="thin", color="AAB7B8")
    bord = Border(left=fin, right=fin, top=fin, bottom=fin)

    def ligne_titre(r, texte, taille, gras=True, couleur="000000", h=22):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n)
        c = ws.cell(row=r, column=1, value=texte)
        c.font = Font(name="Calibri", size=taille, bold=gras, color=couleur)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = h

    ligne_titre(1, "NAP SARL", 20, True, navy, 30)
    ligne_titre(2, "RAPPORT DES PAIEMENTS", 14, True, "2E86C1", 24)
    ligne_titre(3, titre_benef, 13, True, "000000", 24)
    ligne_titre(4, f"{filtres}  —  Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
                10, False, "555555", 20)

    H = 6  # ligne d'en-tête du tableau
    for j, c in enumerate(cols, 1):
        cell = ws.cell(row=H, column=j, value=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = bord
    ws.row_dimensions[H].height = 26

    for i, (_, row) in enumerate(df[cols].iterrows(), 1):
        r = H + i
        for j, c in enumerate(cols, 1):
            v = _py(row[c])
            cell = ws.cell(row=r, column=j, value=v)
            cell.border = bord
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if c == col_montant:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif isinstance(v, datetime):
                cell.number_format = "DD/MM/YYYY"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if i % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F4F9FC")

    fin_data = H + len(df)
    if col_montant in cols:
        r = fin_data + 1
        jm = cols.index(col_montant) + 1
        for j in range(1, n + 1):
            c = ws.cell(row=r, column=j)
            c.fill = PatternFill("solid", fgColor="D6EAF8")
            c.font = Font(bold=True)
            c.border = bord
        ws.cell(row=r, column=1, value="TOTAL")
        cm = ws.cell(row=r, column=jm,
                     value=f"=SUM({get_column_letter(jm)}{H + 1}:{get_column_letter(jm)}{fin_data})")
        cm.number_format = "#,##0"
        cm.alignment = Alignment(horizontal="right")

    # largeurs
    for j, c in enumerate(cols, 1):
        longueur = max([len(str(c))] + [len(texte_cellule(v, c == col_montant)) for v in df[c].head(300)])
        ws.column_dimensions[get_column_letter(j)].width = min(max(longueur + 3, 10), 45)

    ws.freeze_panes = ws.cell(row=H + 1, column=1)
    ws.print_title_rows = f"{H}:{H}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.oddFooter.center.text = "Page &P / &N"

    if synth is not None and len(synth) > 1:
        w2 = wb.create_sheet("Synthèse")
        w2.append(list(synth.columns))
        for _, r in synth.iterrows():
            w2.append([_py(v) for v in r.tolist()])
        for c in w2[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=navy)
        w2.column_dimensions["A"].width = 40
        w2.column_dimensions["B"].width = 16
        w2.column_dimensions["C"].width = 20
        for row in w2.iter_rows(min_row=2, min_col=3, max_col=3):
            for c in row:
                c.number_format = "#,##0"

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ───────────────────────── Export PDF A3 ─────────────────────────
def pdf_a3(df, cols, col_benef, col_montant, titre_benef, filtres, synth=None) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        HRFlowable, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
    )

    buf = BytesIO()
    page = landscape(A3)
    doc = SimpleDocTemplate(
        buf, pagesize=page,
        leftMargin=12 * mm, rightMargin=12 * mm, topMargin=14 * mm, bottomMargin=14 * mm,
        title="Rapport des paiements — NAP SARL", author="NAP SARL",
    )
    W = page[0] - 24 * mm
    navy = colors.HexColor("#1A5276")
    fs = 10 if len(cols) <= 8 else 9 if len(cols) <= 12 else 8

    st_t = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=28, leading=36,
                          alignment=TA_CENTER, textColor=navy, spaceAfter=4)
    st_p = ParagraphStyle("p", fontName="Helvetica-Bold", fontSize=18, leading=26,
                          alignment=TA_CENTER, textColor=colors.HexColor("#2E86C1"), spaceAfter=8)
    st_b = ParagraphStyle("b", fontName="Helvetica-Bold", fontSize=15, leading=22,
                          alignment=TA_CENTER, textColor=colors.HexColor("#222222"), spaceAfter=6)
    st_i = ParagraphStyle("i", fontName="Helvetica", fontSize=10, leading=15,
                          alignment=TA_CENTER, textColor=colors.HexColor("#555555"))
    st_h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=fs, leading=fs + 3,
                          textColor=colors.white, alignment=TA_CENTER)
    st_l = ParagraphStyle("l", fontName="Helvetica", fontSize=fs, leading=fs + 3)
    st_r = ParagraphStyle("r", parent=st_l, alignment=TA_RIGHT)
    st_tot = ParagraphStyle("tot", parent=st_r, fontName="Helvetica-Bold")
    st_totl = ParagraphStyle("totl", parent=st_l, fontName="Helvetica-Bold")

    total = float(df[col_montant].sum()) if col_montant else None

    el = [
        Paragraph("NAP SARL", st_t),
        Paragraph("RAPPORT DES PAIEMENTS", st_p),
        Spacer(1, 2 * mm),
        Paragraph(escape(titre_benef), st_b),
        Paragraph(escape(filtres) + f" &nbsp;·&nbsp; Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}", st_i),
        Spacer(1, 4 * mm),
        HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#5DADE2"),
                   spaceBefore=2, spaceAfter=8),
    ]

    # Données
    data = [[Paragraph("N°", st_h)] + [Paragraph(escape(str(c)), st_h) for c in cols]]
    textes = []
    for idx, (_, row) in enumerate(df[cols].iterrows(), 1):
        ligne = [texte_cellule(row[c], c == col_montant) for c in cols]
        textes.append(ligne)
        cells = [Paragraph(str(idx), ParagraphStyle("n", parent=st_l, alignment=TA_CENTER))]
        for c, t in zip(cols, ligne):
            cells.append(Paragraph(escape(t), st_r if c == col_montant else st_l))
        data.append(cells)

    if col_montant in cols:
        tot = [Paragraph("", st_l)] * (len(cols) + 1)
        tot = list(tot)
        tot[1] = Paragraph("TOTAL", st_totl)
        tot[cols.index(col_montant) + 1] = Paragraph(fmt_montant(total), st_tot)
        data.append(tot)

    # Largeurs proportionnelles au contenu
    w_num = 12 * mm
    poids = []
    for j, c in enumerate(cols):
        longueur = max([len(str(c))] + [len(t[j]) for t in textes[:300]] + [4])
        poids.append(min(max(longueur, 6), 38))
    reste = W - w_num
    col_w = [w_num] + [reste * p / sum(poids) for p in poids]

    t = Table(data, colWidths=col_w, repeatRows=1)
    sty = [
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#AAB7B8")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    n_lignes = len(df)
    for i in range(1, n_lignes + 1):
        if i % 2 == 0:
            sty.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F4F9FC")))
    if col_montant in cols:
        sty += [
            ("BACKGROUND", (0, n_lignes + 1), (-1, n_lignes + 1), colors.HexColor("#D6EAF8")),
            ("LINEABOVE", (0, n_lignes + 1), (-1, n_lignes + 1), 1.2, navy),
        ]
    t.setStyle(TableStyle(sty))
    el.append(t)

    el.append(Spacer(1, 6 * mm))
    resume = f"<b>{n_lignes}</b> paiement(s)"
    if total is not None:
        resume += f" &nbsp;·&nbsp; Total : <b>{fmt_montant(total)} FCFA</b>"
    el.append(Paragraph(resume, ParagraphStyle("res", parent=st_i, fontSize=12, leading=18,
                                               textColor=colors.HexColor("#222222"))))

    # Synthèse par bénéficiaire
    if synth is not None and len(synth) > 1:
        rows = [[Paragraph("Bénéficiaire", st_h), Paragraph("Nb paiements", st_h)]]
        if "Total" in synth.columns:
            rows[0].append(Paragraph("Total", st_h))
        for _, r in synth.iterrows():
            ligne = [Paragraph(escape(str(r["Bénéficiaire"])), st_l),
                     Paragraph(str(int(r["Nb paiements"])), st_r)]
            if "Total" in synth.columns:
                ligne.append(Paragraph(fmt_montant(r["Total"]), st_r))
            rows.append(ligne)
        ws_ = [110 * mm, 35 * mm] + ([50 * mm] if "Total" in synth.columns else [])
        ts = Table(rows, colWidths=ws_, repeatRows=1, hAlign="LEFT")
        ts.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), navy),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#AAB7B8")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        el.append(Spacer(1, 8 * mm))
        el.append(KeepTogether([
            Paragraph("<b>Synthèse par fournisseur / bénéficiaire</b>",
                      ParagraphStyle("sy", parent=st_l, fontSize=12, leading=18, spaceAfter=4)),
            ts,
        ]))

    def pied(canvas, d_):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#666666"))
        canvas.drawString(12 * mm, 7 * mm, "NAP SARL — Rapport des paiements — Document interne")
        canvas.drawRightString(page[0] - 12 * mm, 7 * mm, f"Page {d_.page}")
        canvas.restoreState()

    doc.build(el, onFirstPage=pied, onLaterPages=pied)
    return buf.getvalue()
