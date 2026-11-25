# -*- coding: utf-8 -*-
"""Impression A3 — Liste des paiements par produit / période."""
from __future__ import annotations

from datetime import datetime
from io import BytesIO

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable,
)


def _s(v):
    if v is None:
        return ""
    return str(v).strip()


def _fmt_m(n):
    try:
        return f"{float(n):,.0f}".replace(",", " ")
    except Exception:
        return _s(n)


def _fmt_d(d):
    try:
        import pandas as pd
        ts = pd.to_datetime(d, errors="coerce", dayfirst=True)
        if pd.isna(ts):
            return _s(d)
        return ts.strftime("%d/%m/%Y")
    except Exception:
        return _s(d)


def generer_pdf_a3(df, titre_filtre="", periode="", produit=""):
    buf = BytesIO()
    page = landscape(A3)
    doc = SimpleDocTemplate(
        buf, pagesize=page,
        leftMargin=12 * mm, rightMargin=12 * mm,
        topMargin=12 * mm, bottomMargin=14 * mm,
    )
    styles = getSampleStyleSheet()
    s_title = ParagraphStyle(
        "t", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=16, textColor=colors.HexColor("#1A5276"), alignment=TA_CENTER, spaceAfter=2,
    )
    s_sub = ParagraphStyle(
        "s", parent=styles["Normal"], fontName="Helvetica",
        fontSize=9, textColor=colors.HexColor("#2C3E50"), alignment=TA_CENTER, spaceAfter=2,
    )
    s_meta = ParagraphStyle(
        "m", parent=styles["Normal"], fontName="Helvetica",
        fontSize=8, textColor=colors.HexColor("#5D6D7E"), alignment=TA_CENTER, spaceAfter=6,
    )
    s_cell = ParagraphStyle(
        "c", parent=styles["Normal"], fontName="Helvetica", fontSize=7, leading=9,
    )
    s_cell_b = ParagraphStyle(
        "cb", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7, leading=9,
    )
    s_foot = ParagraphStyle(
        "f", parent=styles["Normal"], fontName="Helvetica", fontSize=7,
        textColor=colors.HexColor("#7F8C8D"), alignment=TA_CENTER,
    )

    story = []
    story.append(Paragraph("STE NAP SARL", s_title))
    story.append(Paragraph(
        "PLOT NO 14 GLODJIGBE ECONOMIC ZONE · RCCM RB/COT/21 B 30710 · IFU 3202113334014 · TEL +229 94 94 04 14",
        s_sub,
    ))
    story.append(Spacer(1, 2 * mm))
    story.append(HRFlowable(width="100%", thickness=1.2, color=colors.HexColor("#5DADE2"), spaceAfter=3))
    story.append(Paragraph("LISTE DES PAIEMENTS", s_title))
    filtres = []
    if produit and produit != "Tous":
        filtres.append(f"Produit : <b>{_s(produit)}</b>")
    if periode:
        filtres.append(f"Période : <b>{_s(periode)}</b>")
    if titre_filtre:
        filtres.append(_s(titre_filtre))
    story.append(Paragraph(" · ".join(filtres) if filtres else "Tous produits · Toutes périodes", s_meta))
    story.append(Paragraph(
        f"Document généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
        s_meta,
    ))

    headers = [
        "Date", "N° Transaction", "Produit", "Prestataire / Fournisseur",
        "Bénéficiaire", "Montant", "Mode", "Banque", "N° Chèque", "Observation",
    ]
    data = [[Paragraph(h, s_cell_b) for h in headers]]
    total = 0.0
    if df is not None and not df.empty:
        for _, r in df.iterrows():
            mt = float(r.get("Montant") or 0)
            total += mt
            data.append([
                Paragraph(_fmt_d(r.get("Date")), s_cell),
                Paragraph(_s(r.get("No_Transaction")), s_cell),
                Paragraph(_s(r.get("Produit")), s_cell),
                Paragraph(_s(r.get("Prestataire")), s_cell),
                Paragraph(_s(r.get("Beneficiaire")), s_cell),
                Paragraph(_fmt_m(mt), s_cell_b),
                Paragraph(_s(r.get("Mode")), s_cell),
                Paragraph(_s(r.get("Banque")), s_cell),
                Paragraph(_s(r.get("Num_Cheque")), s_cell),
                Paragraph(_s(r.get("Observation") or r.get("Description")), s_cell),
            ])

    # largeurs A3 paysage ~ 400 mm utiles
    w = [22*mm, 28*mm, 22*mm, 45*mm, 35*mm, 28*mm, 22*mm, 25*mm, 22*mm, 50*mm]
    t = Table(data, colWidths=w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D6EAF8")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1A5276")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ALIGN", (5, 1), (5, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#B0B0B0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FBFD")]),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(t)
    story.append(Spacer(1, 4 * mm))

    # Totaux
    nb = 0 if df is None or df.empty else len(df)
    tot_tbl = Table([
        [
            Paragraph(f"<b>Nombre de paiements : {nb}</b>", s_cell_b),
            Paragraph(f"<b>MONTANT TOTAL : {_fmt_m(total)} FCFA</b>", ParagraphStyle(
                "tot", parent=s_cell_b, fontSize=10, textColor=colors.HexColor("#1A5276"), alignment=TA_RIGHT,
            )),
        ]
    ], colWidths=[120*mm, 180*mm])
    tot_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EAF2F8")),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#5DADE2")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(tot_tbl)
    story.append(Spacer(1, 8 * mm))
    story.append(HRFlowable(width="100%", thickness=0.6, color=colors.HexColor("#CCCCCC"), spaceAfter=3))
    story.append(Paragraph(
        "STE NAP SARL — Liste des paiements — Document interne — Impression A3",
        s_foot,
    ))

    def _footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#888888"))
        canvas.drawCentredString(page[0] / 2, 8 * mm, f"Page {doc_.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    buf.seek(0)
    return buf.getvalue()
